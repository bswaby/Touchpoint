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

"""
Mission Dashboard 4.1.0
======================================================
Purpose: Comprehensive mission trip management dashboard with sidebar navigation
Author: Ben Swaby
Email: bswaby@fbchtn.org

Features:
- Sidebar navigation with expandable trip sections
- Role-based views (Admins see all trips, Leaders see only their trips)
- Trip-specific sections (Overview, Team, Meetings, Budget, Documents, Messages, Tasks)
- Mobile-responsive collapsible sidebar
- Prominent upcoming deadlines at top
- Performance optimized queries (~1 second)
- AJAX loading for detailed data

Role-Based Access:
- Users with "Edit" role: Full admin access to all trips, can VIEW all financial data
- Trip Leaders: Access to their assigned trips only (via LeaderId or MemberTypeId 140/310/320)
- Leaders see payment totals but NOT individual giver information (privacy)

Financial Permissions:
- Fee adjustments require one of: Finance, FinanceAdmin, or ManageTransactions role
- Users with "Edit" role can VIEW financial data but cannot MODIFY without finance role

Change Log:
v4.1.0 - April 2026
  - Added: Persistent configuration stored in Special Content (survives code updates)
  - Added: Settings UI with tabbed layout (Approval Workflow, Quick Email Templates, Dashboard Configuration)
  - Added: All Config class values editable from Dashboard Configuration tab
  - Added: Dynamic church URL via model.CmsHost (removed hardcoded domain references)
  - Added: Configurable "Share My Missions" link in settings
  - Fixed: let/var conflict with queryStartTime declaration
  - Changed: Config class now serves as defaults; saved settings in Special Content take priority
  - Note: Existing code customizations to Config class continue to work until settings are saved
    from the UI. After that, the UI-saved values take priority.

v4.0.0 - March 2026
  - Sidebar navigation redesign
  - Role-based views
  - Trip-specific sections
  - Mobile-responsive collapsible sidebar
  - Performance optimized queries

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name: Mission_Dashboard
4. Paste all this code
5. Test and optionally add to menu
--Upload Instructions End--
"""


# Debug statements removed - they break AJAX JSON responses
# To debug, uncomment and add conditional: if not is_ajax_request():
# print("<!-- Debug: Script file is executing -->")

# SCRIPT-LEVEL DEBUG REMOVED - model.Data not available at script level

#####################################################################
# CONFIGURATION SECTION
#####################################################################
# These values serve as DEFAULTS. Once you save settings from the
# Dashboard Configuration tab (Settings > Dashboard Configuration),
# your saved values are stored in Special Content as JSON and take
# priority over these defaults on every script load.
#
# This means:
#   1. You can safely update this script without losing your settings
#   2. Settings saved from the UI survive code updates
#   3. Any value NOT saved from the UI falls back to these defaults
#   4. To reset a setting to default, delete the Special Content
#      file named "TPxi_MissionsDashboard_Config"
#####################################################################

# ::START:: Configuration
class Config:
    """Default configuration values. Overridden by saved settings in Special Content."""
    
    # Organization Settings
    ACTIVE_ORG_STATUS_ID = 30  # Organization status ID for active missions
    MISSION_TRIP_FLAG = 1      # IsMissionTrip flag value
    
    # Member Type IDs (adjust based on your lookup.MemberType table)
    MEMBER_TYPE_LEADER = 230   # Leader member type to exclude from counts
    ATTENDANCE_TYPE_LEADER = 10  # Leader attendance type
    
    # Date Formats
    DATE_FORMAT_SHORT = 'short'  # For format_date function
    DATE_FORMAT_DISPLAY = 'M/d/yyyy'
    
    # UI Settings
    ITEMS_PER_PAGE = 20  # For pagination
    SHOW_CLOSED_BY_DEFAULT = False
    ENABLE_DEVELOPER_VISUALIZATION = False
    ENABLE_SQL_DEBUG = False  # Set to False to disable SQL debug output by default
    
    # Email Settings
    SEND_PAYMENT_REMINDERS = False
    PAYMENT_REMINDER_TEMPLATE = "MissionPaymentReminder" #future feature
    
    # Church Sharing Link (shown to admins/leaders to share with team members)
    # Change this to your church's public-facing URL for the missions page
    MY_MISSIONS_LINK = ""  # e.g., "https://yourchurch.com/MyMissions" - leave empty to auto-generate from CmsHost

    # Currency Settings
    CURRENCY_SYMBOL = "$"
    USE_THOUSANDS_SEPARATOR = True
    
    # Security Roles - any of these grants full admin access to all trips
    ADMIN_ROLES = ["Edit", "Admin", "Finance", "MissionsDirector"]
    LEADER_MEMBER_TYPES = [140, 310, 320]  # MemberTypeIds that indicate leadership

    # Finance Roles - required for making financial adjustments (fees, transactions)
    # Users with admin roles can VIEW financial data but need one of these to MODIFY
    FINANCE_ROLES = ["Finance", "FinanceAdmin", "ManageTransactions"]

    # Menu Access Roles - users with these roles can see this script in the TouchPoint menu
    # Leaders without these roles need a direct link to access the dashboard
    MENU_ACCESS_ROLES = ["Edit", "Admin", "Finance", "MissionsDirector", "Access"]

    # Sidebar Settings
    SIDEBAR_WIDTH = 280  # Sidebar width in pixels
    SIDEBAR_COLLAPSED_WIDTH = 60  # Collapsed sidebar width
    SIDEBAR_BG_COLOR = "#1a1a2e"  # Dark sidebar background
    SIDEBAR_ACTIVE_COLOR = "#0f3460"  # Active item background

    # Trip Section Configuration
    TRIP_SECTIONS = [
        {'id': 'overview', 'label': 'Overview', 'icon': '&#128202;'},  # chart icon
        {'id': 'team', 'label': 'Team Members', 'icon': '&#128101;'},  # people icon
        {'id': 'meetings', 'label': 'Meetings', 'icon': '&#128197;'},  # calendar icon
        {'id': 'budget', 'label': 'Budget & Fundraising', 'icon': '&#128176;'},  # money icon
        # {'id': 'documents', 'label': 'Documents', 'icon': '&#128196;'},  # document icon - coming later
        {'id': 'messages', 'label': 'Messages', 'icon': '&#9993;'},  # envelope icon
        # {'id': 'tasks', 'label': 'Tasks & Goals', 'icon': '&#9989;'}  # checkmark icon - coming later
    ]
    
    # Cache Settings (in seconds)
    CACHE_DURATION = 300  # 5 minutes
    USE_CACHING = True  # Enable caching for performance
    
    # Application-specific organization IDs to exclude (adjust for your setup)
    APPLICATION_ORG_IDS = [2736, 2737, 2738, 3032, 3117, 3304, 3361]  # Organizations used for applications
    
    # Dashboard Sections
    ENABLE_STATS_TAB = True
    ENABLE_FINANCE_TAB = True
    ENABLE_MESSAGES_TAB = True
    DEFAULT_TAB = "dashboard"
    
    # Performance Settings
    USE_SIMPLE_QUERIES = True  # Use optimized queries for better performance
    MAX_QUERY_TIMEOUT = 30  # Maximum seconds for queries
    BATCH_SIZE = 100  # Batch size for queries
    
    # Performance Tuning Notes:
    # - Original query using custom.MissionTripTotals_Optimized view was taking 47 seconds
    # - Intermediate optimization (filtering closed trips): ~29 seconds (39% improvement)
    # - Current optimized version using CTE instead of view: ~1 second (98% improvement!)
    # - This version shows only open trips in all totals by default
    # - The CTE approach replaces the inefficient scalar function in the view
    # - All queries now use the optimized CTE pattern for consistent performance

    # Email Template Settings
    # Templates are stored in TouchPoint Special Content with prefix "MissionsDashboard_Template_"
    # Placeholders: {{PersonName}}, {{TripName}}, {{MyGivingLink}}, {{SupportLink}}, {{TripCost}}, {{Outstanding}}
    TEMPLATE_CONTENT_PREFIX = "MissionsDashboard_Template_"

    # Default Email Templates (used if no custom template is saved)
    DEFAULT_TEMPLATES = {
        'goal_reminder_team': {
            'name': 'Goal Reminder (Team)',
            'description': 'Sent to all team members as a fundraising reminder',
            'subject': '{{TripName}} - Fundraising Goal Reminder',
            'body': '''Hi {{PersonName}},

This is a friendly reminder about your fundraising goals for {{TripName}}.

Please check your current fundraising status and reach out if you need any help meeting your goal.

View your payment status here:
{{MyGivingLink}}

Share this link with friends and family who want to support your trip:
{{SupportLink}}

Thank you for your commitment to this mission trip!

Blessings'''
        },
        'goal_reminder_individual': {
            'name': 'Goal Reminder (Individual)',
            'description': 'Personalized reminder with specific cost details',
            'subject': '{{TripName}} - Fundraising Goal Reminder',
            'body': '''Hi {{PersonName}},

This is a friendly reminder about your fundraising goal for {{TripName}}.

Your trip cost: {{TripCost}}
Amount still needed: {{Outstanding}}

View your payment status here:
{{MyGivingLink}}

Share this link with friends and family who want to support your trip:
{{SupportLink}}

Please reach out if you need any help meeting your goal.

Thank you for your commitment to this mission trip!

Blessings'''
        },
        'passport_request': {
            'name': 'Passport Information Request',
            'description': 'Request passport details from team member',
            'subject': '{{TripName}} - Passport Information Needed',
            'body': '''Hi {{PersonName}},

We need your passport information for the upcoming {{TripName}} mission trip.

Please complete the passport form at the link below:

{{ChurchUrl}}/OnlineReg/3421

This form will collect your passport number, expiration date, and other travel details we need for trip planning.

Please complete this as soon as possible so we can finalize travel arrangements.

Thank you!

Blessings'''
        },
        'meeting_reminder': {
            'name': 'Meeting Reminder',
            'description': 'Reminder about upcoming team meeting',
            'subject': '{{TripName}} - Meeting Reminder: {{MeetingDescription}}',
            'body': '''Hi Team,

This is a reminder about our upcoming meeting:

{{MeetingDescription}}
{{MeetingDate}} at {{MeetingTime}}
{{MeetingLocation}}

Please make sure to attend. If you cannot make it, please let us know as soon as possible.

See you there!

Blessings'''
        }
    }

# ::END:: Configuration

#####################################################################
# LIBRARY IMPORTS AND INITIALIZATION
#####################################################################

# ::START:: Initialization
import datetime
import re

# Church base URL - dynamically set from TouchPoint
CHURCH_URL = str(model.CmsHost) if hasattr(model, 'CmsHost') and model.CmsHost else 'https://myfbch.com'
# Ensure it has https://
if not CHURCH_URL.startswith('http'):
    CHURCH_URL = 'https://' + CHURCH_URL
# Remove trailing slash
CHURCH_URL = CHURCH_URL.rstrip('/')

# Resolve MyMissions sharing link
MY_MISSIONS_LINK = Config.MY_MISSIONS_LINK or (CHURCH_URL + '/PyScriptForm/Mission_Dashboard')
# Display-friendly version (without https://)
MY_MISSIONS_DISPLAY = MY_MISSIONS_LINK.replace('https://', '').replace('http://', '')

# Set page header (no debug output - breaks AJAX JSON responses)
try:
    model.Header = "Missions Dashboard"
except Exception as e:
    pass  # Silently ignore if not in TouchPoint context

# Initialize configuration with persistent settings from Special Content
CONFIG_CONTENT_NAME = 'TPxi_MissionsDashboard_Config'

def load_config():
    """Load config from Special Content, merge with defaults"""
    config = Config()
    try:
        stored = model.TextContent(CONFIG_CONTENT_NAME)
        if stored and stored.strip():
            import json as _json
            saved = _json.loads(stored)
            if isinstance(saved, dict):
                for key, val in saved.items():
                    if hasattr(config, key) and val is not None:
                        default_val = getattr(config, key)
                        # Ensure type consistency with defaults
                        if isinstance(default_val, list) and isinstance(val, str):
                            # Convert comma-separated string back to list
                            if default_val and isinstance(default_val[0], int):
                                val = [int(x.strip()) for x in val.split(',') if x.strip()]
                            else:
                                val = [x.strip() for x in val.split(',') if x.strip()]
                        elif isinstance(default_val, int) and not isinstance(val, int):
                            try:
                                val = int(val)
                            except:
                                continue
                        elif isinstance(default_val, bool) and not isinstance(val, bool):
                            val = str(val).lower() in ('true', '1', 'yes')
                        setattr(config, key, val)
    except Exception as e:
        pass
    return config

def save_config(config):
    """Save current config to Special Content"""
    import json as _json
    # Build dict of all configurable settings
    data = {
        'ACTIVE_ORG_STATUS_ID': config.ACTIVE_ORG_STATUS_ID,
        'MISSION_TRIP_FLAG': config.MISSION_TRIP_FLAG,
        'MEMBER_TYPE_LEADER': config.MEMBER_TYPE_LEADER,
        'ATTENDANCE_TYPE_LEADER': config.ATTENDANCE_TYPE_LEADER,
        'ITEMS_PER_PAGE': config.ITEMS_PER_PAGE,
        'SHOW_CLOSED_BY_DEFAULT': config.SHOW_CLOSED_BY_DEFAULT,
        'ENABLE_SQL_DEBUG': config.ENABLE_SQL_DEBUG,
        'SEND_PAYMENT_REMINDERS': config.SEND_PAYMENT_REMINDERS,
        'MY_MISSIONS_LINK': config.MY_MISSIONS_LINK,
        'CURRENCY_SYMBOL': config.CURRENCY_SYMBOL,
        'ADMIN_ROLES': config.ADMIN_ROLES,
        'FINANCE_ROLES': config.FINANCE_ROLES,
        'LEADER_MEMBER_TYPES': config.LEADER_MEMBER_TYPES,
        'SIDEBAR_BG_COLOR': config.SIDEBAR_BG_COLOR,
        'SIDEBAR_ACTIVE_COLOR': config.SIDEBAR_ACTIVE_COLOR,
        'APPLICATION_ORG_IDS': config.APPLICATION_ORG_IDS,
        'ENABLE_STATS_TAB': config.ENABLE_STATS_TAB,
        'ENABLE_FINANCE_TAB': config.ENABLE_FINANCE_TAB,
        'ENABLE_MESSAGES_TAB': config.ENABLE_MESSAGES_TAB,
    }
    model.WriteContentText(CONFIG_CONTENT_NAME, _json.dumps(data, indent=2), "")

config = load_config()

# Check if function library exists and load it
try:
    function_library = model.TextContent('_FunctionLibrary')
    function_library = function_library.replace('\r', '')  # Normalize line endings
    exec(function_library)
except:
    # Define essential functions if library not found
    def format_currency(value, use_symbol=True, use_separator=True):
        """Format currency with proper symbols and separators"""
        if value is None:
            return "$0.00"
        try:
            formatted = "{:,.2f}".format(float(value)) if use_separator else "{:.2f}".format(float(value))
            return config.CURRENCY_SYMBOL + formatted if use_symbol else formatted
        except:
            return "$0.00"

def format_date(date_str, format_type='short'):
    """Format date strings consistently - handle TouchPoint SQL dates properly"""
    if not date_str:
        return ""
    
    try:
        date_str = str(date_str)
        
        # Handle TouchPoint datetime objects directly
        if hasattr(date_str, 'strftime'):
            if format_type == 'short':
                return date_str.strftime('%-m/%-d')
            else:
                return date_str.strftime('%-m/%-d/%Y')
        
        # Remove time portion if present (handle SQL datetime format)
        if ' ' in date_str:
            date_str = date_str.split(' ')[0]
        
        # Parse different date formats
        parsed_date = None
        
        # Try TouchPoint's ParseDate first
        if hasattr(model, 'ParseDate'):
            try:
                parsed_date = model.ParseDate(date_str)
                if parsed_date:
                    if hasattr(model, 'FmtDate'):
                        if format_type == 'short':
                            return model.FmtDate(parsed_date, 'M/d')
                        else:
                            return model.FmtDate(parsed_date, 'M/d/yyyy')
            except:
                pass
        
        # Manual parsing for common formats
        if not parsed_date:
            # Handle YYYY-MM-DD format
            if '-' in date_str and len(date_str.split('-')) == 3:
                parts = date_str.split('-')
                if len(parts[0]) == 4:  # YYYY-MM-DD
                    month = int(parts[1])
                    day = int(parts[2])
                    year = int(parts[0])
                    if format_type == 'short':
                        return "{0}/{1}".format(month, day)
                    else:
                        return "{0}/{1}/{2}".format(month, day, year)
            
            # Handle MM/DD/YYYY or M/D/YYYY format
            elif '/' in date_str:
                parts = date_str.split('/')
                if len(parts) >= 2:
                    if format_type == 'short':
                        return "{0}/{1}".format(int(parts[0]), int(parts[1]))
                    else:
                        return date_str
        
        # If all else fails, return the date portion
        return date_str
        
    except Exception as e:
        # In case of any error, return the original string
        return str(date_str) if date_str else ""

# SQL Debug mode - check for developer role and debug parameter
SQL_DEBUG = config.ENABLE_SQL_DEBUG  # Use config default
if config.ENABLE_DEVELOPER_VISUALIZATION and model.UserIsInRole("SuperAdmin"):
    # Allow SuperAdmin to override via URL parameter
    if hasattr(model.Data, 'debug'):
        SQL_DEBUG = model.Data.debug == '1'

def debug_sql(query, label="Query"):
    """Output SQL query in HTML comments for debugging"""
    if SQL_DEBUG or (hasattr(model.Data, 'debug') and model.Data.debug == '1'):
        # Clean up the query for safe HTML comment output
        clean_query = query.replace('-->', '-- >')
        print '''
<!-- SQL DEBUG: {0}
========================================
{1}
========================================
End SQL DEBUG: {0} -->
        '''.format(label, clean_query)

def execute_query_with_debug(query, label="Query", query_type="list"):
    """Execute query with optional debug output"""
    debug_sql(query, label)
    
    try:
        if query_type == "top1":
            return q.QuerySqlTop1(query)
        elif query_type == "sql":
            return q.QuerySql(query)
        elif query_type == "count":
            return q.QuerySqlInt(query)
        else:
            return q.QuerySql(query)
    except Exception as e:
        if SQL_DEBUG:
            print "<!-- SQL ERROR in {0}: {1} -->".format(label, str(e))
        raise

def get_mission_trip_totals_cte(include_closed=False):
    """Generate the optimized CTE for mission trip totals that replaces custom.MissionTripTotals_Optimized
    This runs in ~1 second vs 47+ seconds for the original view
    
    Args:
        include_closed: If False (default), only includes open trips (Close date > today)
    """
    
    # Add closed trip filter if needed
    # Exclude trips that have a 'Close' field with a date in the past
    # Trips without a 'Close' field are included (they're not explicitly closed)
    closed_filter = ''
    if not include_closed:
        closed_filter = '''
            AND NOT EXISTS (
                SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
                WHERE oe.OrganizationId = o.OrganizationId
                  AND oe.Field = 'Close'
                  AND oe.DateValue IS NOT NULL
                  AND oe.DateValue <= GETDATE()
            )'''
    
    # Build the complete CTE
    cte = '''
    WITH ActiveTrips AS (
        SELECT OrganizationId, OrganizationName AS Trip
        FROM dbo.Organizations o
        WHERE IsMissionTrip = 1 AND OrganizationStatusId = 30''' + closed_filter + '''
    ),
    TripGoers AS (
        SELECT
            at.OrganizationId,
            at.Trip,
            om.PeopleId,
            p.Name,
            p.Name2 AS SortOrder,
            ts.IndAmt AS TripCost,
            ts.IndDue AS IndDue,  -- Individual due (includes fee adjustments)
            om.TranId
        FROM ActiveTrips at
        INNER JOIN dbo.OrganizationMembers om ON om.OrganizationId = at.OrganizationId
        INNER JOIN dbo.OrgMemMemTags omm ON omm.OrgId = om.OrganizationId AND omm.PeopleId = om.PeopleId
        INNER JOIN dbo.MemberTags mt ON mt.Id = omm.MemberTagId AND mt.Name = 'Goer'
        INNER JOIN dbo.People p ON p.PeopleId = om.PeopleId
        LEFT JOIN dbo.TransactionSummary ts ON ts.OrganizationId = om.OrganizationId
            AND ts.PeopleId = om.PeopleId AND ts.RegId = om.TranId
    ),
    PaymentCalculations AS (
        SELECT
            tg.OrganizationId,
            tg.Trip,
            tg.PeopleId,
            tg.Name,
            tg.SortOrder,
            tg.TripCost,
            tg.IndDue,  -- Pass through IndDue (includes fee adjustments)
            -- Individual payments (replaces first part of TotalPaid function)
            ISNULL((SELECT TOP(1) ts.IndPaid
                    FROM dbo.TransactionSummary ts
                    WHERE ts.RegId = tg.TranId AND ts.PeopleId = tg.PeopleId
                    ORDER BY ts.TranDate DESC), 0) AS IndPaid,
            -- Supporter payments (replaces second part of TotalPaid function)
            ISNULL((SELECT SUM(gsa.Amount)
                    FROM dbo.GoerSenderAmounts gsa
                    INNER JOIN dbo.OrganizationMembers om ON om.OrganizationId = tg.OrganizationId AND om.PeopleId = tg.PeopleId
                    WHERE gsa.GoerId = tg.PeopleId
                    AND gsa.OrgId = tg.OrganizationId
                    AND gsa.Created > om.EnrollmentDate
                    AND gsa.SupporterId <> tg.PeopleId), 0) AS SupporterPaid
        FROM TripGoers tg
    ),
    PaymentTotals AS (
        SELECT
            OrganizationId,
            Trip,
            PeopleId,
            Name,
            SortOrder,
            TripCost,
            IndDue,  -- Individual due (includes fee adjustments)
            IndPaid + SupporterPaid AS TotalPaid,
            SupporterPaid
        FROM PaymentCalculations
    ),
    UndesignatedFunds AS (
        SELECT
            at.OrganizationId,
            at.Trip,
            NULL AS PeopleId,
            'Undesignated' AS Name,
            'YZZZZ' AS SortOrder,
            NULL AS TripCost,
            NULL AS IndDue,  -- Undesignated funds don't have individual due
            ISNULL(SUM(gsa.Amount), 0) AS TotalPaid,
            0 AS SupporterPaid  -- Not applicable for undesignated
        FROM ActiveTrips at
        LEFT JOIN dbo.GoerSenderAmounts gsa ON gsa.OrgId = at.OrganizationId
            AND gsa.GoerId IS NULL AND ISNULL(gsa.InActive, 0) = 0
        GROUP BY at.OrganizationId, at.Trip
    ),
    MissionTripTotals AS (
        -- Final result - all people (including those who have paid in full)
        -- Due calculation uses IndDue (which includes fee adjustments) minus supporter payments
        SELECT
            OrganizationId AS InvolvementId,
            Trip,
            PeopleId,
            Name,
            SortOrder,
            TripCost,
            TotalPaid AS Raised,
            -- Use IndDue (includes fee adjustments) minus SupporterPaid for accurate outstanding
            CASE
                WHEN IndDue IS NOT NULL THEN ISNULL(IndDue, 0) - ISNULL(SupporterPaid, 0)
                ELSE 0  -- For undesignated funds
            END AS Due
        FROM (
            SELECT * FROM PaymentTotals
            UNION ALL
            SELECT * FROM UndesignatedFunds
        ) combined
    )
    '''
    
    return cte

# ::END:: Initialization

#####################################################################
# ROLE & ACCESS CONTROL - Sidebar Navigation Support
#####################################################################

# ::START:: Role Detection
def get_user_role_and_trips():
    """
    Determine user's role and accessible trips for sidebar navigation.

    Returns:
        dict: {
            'is_admin': bool,           # Has Edit role (full access to all trips)
            'can_manage_finance': bool, # Has Finance/FinanceAdmin/ManageTransactions role
            'is_trip_leader': bool,     # Leader of at least one trip
            'is_trip_member': bool,     # Member of at least one trip (not leader)
            'accessible_trips': list,   # List of trips user can access (for leaders)
            'member_trips': list,       # List of trips where user is a member (not leader)
            'user_id': int              # Current user's PeopleId
        }
    """
    user_id = model.UserPeopleId

    # Check admin roles - any of these grants full access to all trips
    is_admin = False
    for role in config.ADMIN_ROLES:
        if model.UserIsInRole(role):
            is_admin = True
            break

    # Check finance roles - required for making financial adjustments
    can_manage_finance = False
    for role in config.FINANCE_ROLES:
        if model.UserIsInRole(role):
            can_manage_finance = True
            break

    # Get trips where user is a leader (org leader or leader member type)
    leader_member_types = ','.join(str(x) for x in config.LEADER_MEMBER_TYPES)

    leader_trips_sql = '''
    SELECT DISTINCT
        o.OrganizationId,
        o.OrganizationName,
        o.MemberCount,
        oe_start.DateValue AS StartDate,
        oe_end.DateValue AS EndDate,
        CASE
            WHEN oe_start.DateValue > GETDATE() THEN 'upcoming'
            WHEN oe_end.DateValue < GETDATE() THEN 'completed'
            ELSE 'active'
        END AS TripStatus
    FROM Organizations o WITH (NOLOCK)
    LEFT JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
        AND om.PeopleId = {0}
        AND om.MemberTypeId IN ({1})
        AND om.InactiveDate IS NULL
    LEFT JOIN OrganizationExtra oe_start ON o.OrganizationId = oe_start.OrganizationId
        AND oe_start.Field = 'Main Event Start'
    LEFT JOIN OrganizationExtra oe_end ON o.OrganizationId = oe_end.OrganizationId
        AND oe_end.Field = 'Main Event End'
    WHERE o.IsMissionTrip = {2}
      AND o.OrganizationStatusId = {3}
      AND (o.LeaderId = {0} OR om.PeopleId IS NOT NULL)
    ORDER BY oe_start.DateValue, o.OrganizationName
    '''.format(user_id, leader_member_types, config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)

    leader_trips = list(q.QuerySql(leader_trips_sql))
    leader_trip_ids = set(t.OrganizationId for t in leader_trips)

    # Get trips where user is a MEMBER (not a leader) - for My Missions portal
    member_trips_sql = '''
    SELECT DISTINCT
        o.OrganizationId,
        o.OrganizationName,
        o.MemberCount,
        oe_start.DateValue AS StartDate,
        oe_end.DateValue AS EndDate,
        CASE
            WHEN oe_start.DateValue > GETDATE() THEN 'upcoming'
            WHEN oe_end.DateValue < GETDATE() THEN 'completed'
            ELSE 'active'
        END AS TripStatus,
        om.MemberTypeId
    FROM Organizations o WITH (NOLOCK)
    INNER JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
        AND om.PeopleId = {0}
        AND om.InactiveDate IS NULL
    LEFT JOIN OrganizationExtra oe_start ON o.OrganizationId = oe_start.OrganizationId
        AND oe_start.Field = 'Main Event Start'
    LEFT JOIN OrganizationExtra oe_end ON o.OrganizationId = oe_end.OrganizationId
        AND oe_end.Field = 'Main Event End'
    WHERE o.IsMissionTrip = {1}
      AND o.OrganizationStatusId = {2}
      AND om.MemberTypeId NOT IN ({3})  -- Exclude leader types
      AND o.LeaderId != {0}  -- Not the org leader
    ORDER BY oe_start.DateValue, o.OrganizationName
    '''.format(user_id, config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, leader_member_types)

    member_trips = list(q.QuerySql(member_trips_sql))

    return {
        'is_admin': is_admin,
        'can_manage_finance': can_manage_finance,
        'is_trip_leader': len(leader_trips) > 0,
        'is_trip_member': len(member_trips) > 0,
        'accessible_trips': leader_trips,
        'member_trips': member_trips,
        'user_id': user_id
    }


def has_trip_access(user_role, org_id):
    """
    Check if user can access a specific trip.

    Args:
        user_role: dict from get_user_role_and_trips()
        org_id: Organization ID to check

    Returns:
        bool: True if user has access to this trip
    """
    if user_role['is_admin']:
        return True

    # Check if org_id is in user's accessible trips
    for trip in user_role['accessible_trips']:
        if str(trip.OrganizationId) == str(org_id):
            return True

    return False


def get_all_trips_for_sidebar(include_closed=False):
    """
    Get all trips for admin sidebar display.

    Args:
        include_closed: Whether to include closed trips

    Returns:
        list: All mission trips with basic info for sidebar
    """
    closed_filter = ""
    if not include_closed:
        # Exclude trips that have a 'Close' field with a date in the past
        # Trips without a 'Close' field are included (they're not explicitly closed)
        closed_filter = '''
            AND NOT EXISTS (
                SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
                WHERE oe.OrganizationId = o.OrganizationId
                  AND oe.Field = 'Close'
                  AND oe.DateValue IS NOT NULL
                  AND oe.DateValue <= GETDATE()
            )'''

    trips_sql = '''
    SELECT
        o.OrganizationId,
        o.OrganizationName,
        o.MemberCount,
        oe_start.DateValue AS StartDate,
        oe_end.DateValue AS EndDate,
        CASE
            WHEN oe_start.DateValue > GETDATE() THEN 'upcoming'
            WHEN oe_end.DateValue < GETDATE() THEN 'completed'
            ELSE 'active'
        END AS TripStatus,
        ISNULL((
            SELECT SUM(mtt.Due)
            FROM (
                SELECT
                    om2.OrganizationId,
                    om2.PeopleId,
                    ISNULL(ts.IndAmt, 0) - ISNULL(ts.IndPaid, 0) - ISNULL((
                        SELECT SUM(gsa.Amount)
                        FROM GoerSenderAmounts gsa
                        WHERE gsa.GoerId = om2.PeopleId
                          AND gsa.OrgId = om2.OrganizationId
                          AND gsa.Created > om2.EnrollmentDate
                          AND gsa.SupporterId <> om2.PeopleId
                    ), 0) AS Due
                FROM OrganizationMembers om2 WITH (NOLOCK)
                INNER JOIN OrgMemMemTags omm ON omm.OrgId = om2.OrganizationId AND omm.PeopleId = om2.PeopleId
                INNER JOIN MemberTags mt ON mt.Id = omm.MemberTagId AND mt.Name = 'Goer'
                LEFT JOIN TransactionSummary ts ON ts.OrganizationId = om2.OrganizationId
                    AND ts.PeopleId = om2.PeopleId AND ts.RegId = om2.TranId
                WHERE om2.OrganizationId = o.OrganizationId
            ) mtt
            WHERE mtt.Due > 0
        ), 0) AS Outstanding
    FROM Organizations o WITH (NOLOCK)
    LEFT JOIN OrganizationExtra oe_start ON o.OrganizationId = oe_start.OrganizationId
        AND oe_start.Field = 'Main Event Start'
    LEFT JOIN OrganizationExtra oe_end ON o.OrganizationId = oe_end.OrganizationId
        AND oe_end.Field = 'Main Event End'
    WHERE o.IsMissionTrip = {0}
      AND o.OrganizationStatusId = {1}
      {2}
    ORDER BY
        CASE
            WHEN oe_start.DateValue > GETDATE() THEN 2  -- upcoming
            WHEN oe_end.DateValue < GETDATE() THEN 3   -- completed
            ELSE 1  -- active
        END,
        oe_start.DateValue,
        o.OrganizationName
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, closed_filter)

    return list(q.QuerySql(trips_sql))


def print_access_denied():
    """Display access denied message"""
    print '''
    <div style="padding: 40px; text-align: center;">
        <h2 style="color: #c62828;">Access Denied</h2>
        <p>You do not have permission to view this trip.</p>
        <p>If you believe this is an error, please contact your administrator.</p>
        <a href="?" class="btn btn-primary">Return to Dashboard</a>
    </div>
    '''

# ::END:: Role Detection

#####################################################################
# SIDEBAR NAVIGATION
#####################################################################

# ::START:: Sidebar Navigation
def render_sidebar(user_role, current_trip=None, current_section='overview', current_view=None):
    """
    Render the navigation sidebar with trip list and sections.

    Args:
        user_role: Dict with 'is_admin', 'accessible_trips' from get_user_role_and_trips()
        current_trip: Currently selected trip org_id (string or int)
        current_section: Currently active section within trip (default 'overview')
        current_view: Currently active admin view ('home', 'finance', 'stats')

    Returns:
        HTML string for the complete sidebar
    """
    current_trip = str(current_trip) if current_trip else None
    html = []

    # Mobile toggle button (outside sidebar)
    html.append('<button class="mobile-toggle" onclick="MissionsDashboard.toggleMobile()">&#9776;<span class="mobile-toggle-label">Menu</span></button>')

    # Start sidebar
    html.append('<nav class="nav-sidebar" id="navSidebar">')

    # Sidebar Header
    html.append('''
        <div class="sidebar-header">
            <h3>Missions</h3>
            <button class="sidebar-toggle" onclick="MissionsDashboard.toggleSidebar()" title="Toggle Sidebar">
                &#9664;
            </button>
        </div>
    ''')

    # Admin Navigation (only for admins)
    if user_role.get('is_admin', False):
        html.append('<div class="admin-nav">')

        # Dashboard Home
        home_active = 'active' if current_view == 'home' and not current_trip else ''
        html.append('''
            <a href="?" class="nav-link {0}">
                <span class="icon">&#127968;</span>
                <span class="sidebar-text">Dashboard Home</span>
            </a>
        '''.format(home_active))

        # Review (Pending Approvals)
        review_active = 'active' if current_view == 'review' and not current_trip else ''
        html.append('''
            <a href="?view=review" class="nav-link {0}">
                <span class="icon">&#9998;</span>
                <span class="sidebar-text">Review</span>
            </a>
        '''.format(review_active))

        # All Finances
        finance_active = 'active' if current_view == 'finance' and not current_trip else ''
        html.append('''
            <a href="?view=finance" class="nav-link {0}">
                <span class="icon">&#128176;</span>
                <span class="sidebar-text">All Finances</span>
            </a>
        '''.format(finance_active))

        # Statistics
        stats_active = 'active' if current_view == 'stats' and not current_trip else ''
        html.append('''
            <a href="?view=stats" class="nav-link {0}">
                <span class="icon">&#128200;</span>
                <span class="sidebar-text">Statistics</span>
            </a>
        '''.format(stats_active))

        # Calendar
        calendar_active = 'active' if current_view == 'calendar' and not current_trip else ''
        html.append('''
            <a href="?view=calendar" class="nav-link {0}">
                <span class="icon">&#128197;</span>
                <span class="sidebar-text">Calendar</span>
            </a>
        '''.format(calendar_active))

        # Settings (admin only)
        settings_active = 'active' if current_view == 'settings' and not current_trip else ''
        html.append('''
            <a href="?view=settings" class="nav-link {0}">
                <span class="icon">&#9881;</span>
                <span class="sidebar-text">Settings</span>
            </a>
        '''.format(settings_active))

        html.append('</div>')  # End admin-nav

    # Trips Header
    html.append('''
        <div class="trips-header">
            <h4>Mission Trips</h4>
        </div>
    ''')

    # Get trips to display based on role
    if user_role.get('is_admin', False):
        all_trips = get_all_trips_for_sidebar()
    else:
        all_trips = user_role.get('accessible_trips', [])

    # Separate regular trips from application orgs
    trips = [t for t in all_trips if t.OrganizationId not in config.APPLICATION_ORG_IDS]
    app_orgs = [t for t in all_trips if t.OrganizationId in config.APPLICATION_ORG_IDS]

    # Check if a trip is currently selected
    trips_section_expanded = current_trip in [str(t.OrganizationId) for t in trips] or current_trip is None or current_trip == ''

    # Mission Trips Section (collapsible)
    html.append('''
        <div class="trips-section">
            <div class="trips-header" onclick="MissionsDashboard.toggleTripsSection()">
                <span class="expand-icon {0}">&#9654;</span>
                <span>Trips</span>
                <span class="trips-count">{1}</span>
            </div>
            <ul class="trip-list" style="display: {2};">
    '''.format('expanded' if trips_section_expanded else '', len(trips), 'block' if trips_section_expanded else 'none'))

    for trip in trips:
        org_id = str(trip.OrganizationId)
        org_name = trip.OrganizationName

        # Determine if this trip is expanded (currently selected)
        is_expanded = org_id == current_trip
        expanded_class = 'expanded' if is_expanded else ''
        active_class = 'active' if is_expanded else ''

        # Get team member count for badge
        member_count = getattr(trip, 'MemberCount', 0) or 0

        html.append('<li class="trip-item {0}" data-trip-id="{1}">'.format(expanded_class, org_id))

        # Get trip dates for display
        trip_start = getattr(trip, 'StartDate', None)
        trip_end = getattr(trip, 'EndDate', None)
        date_display = _format_trip_date_range(trip_start, trip_end)

        # Trip Header (clickable to expand/navigate)
        html.append('''
            <div class="trip-header {0}" onclick="MissionsDashboard.toggleTrip('{1}')" title="{2}">
                <span class="expand-icon">&#9654;</span>
                <div class="trip-info">
                    <a href="?trip={1}" class="trip-link" onclick="event.stopPropagation();">{2}</a>
                    <span class="trip-dates">{4}</span>
                </div>
                <span class="trip-member-badge" title="{3} team members">&#128101; {3}</span>
            </div>
        '''.format(active_class, org_id, _escape_html(org_name), member_count, date_display))

        # Trip Sections (expandable)
        html.append('<div class="trip-sections">')

        # Check if user is admin for section visibility
        is_admin = user_role.get('is_admin', False)

        for section in config.TRIP_SECTIONS:
            section_id = section['id']
            section_label = section['label']
            section_icon = section['icon']

            # Hide Messages section for non-admins
            if section_id == 'messages' and not is_admin:
                continue

            # Check if this section is active
            section_active = 'active' if is_expanded and current_section == section_id else ''

            html.append('''
                <a href="?trip={0}&section={1}" class="section-link {2}"
                   onclick="MissionsDashboard.loadSection('{0}', '{1}', event)">
                    <span class="section-icon">{3}</span>
                    <span class="sidebar-text">{4}</span>
                </a>
            '''.format(org_id, section_id, section_active, section_icon, section_label))

        html.append('</div>')  # End trip-sections
        html.append('</li>')  # End trip-item

    html.append('</ul>')  # End trip-list
    html.append('</div>')  # End trips-section

    # Application Orgs Section (collapsible, minimized by default)
    if app_orgs and user_role.get('is_admin', False):
        # Check if any app org is currently selected
        app_org_section_expanded = current_trip in [str(t.OrganizationId) for t in app_orgs]

        html.append('''
            <div class="app-orgs-section">
                <div class="app-orgs-header" onclick="MissionsDashboard.toggleAppOrgs()">
                    <span class="expand-icon {0}">&#9654;</span>
                    <span>Applications</span>
                    <span class="app-org-count">{1}</span>
                </div>
                <ul class="app-orgs-list" style="display: {2};">
        '''.format('expanded' if app_org_section_expanded else '', len(app_orgs), 'block' if app_org_section_expanded else 'none'))

        for app_org in app_orgs:
            org_id = str(app_org.OrganizationId)
            org_name = app_org.OrganizationName

            # Determine if this app org is expanded (currently selected)
            is_expanded = org_id == current_trip
            expanded_class = 'expanded' if is_expanded else ''
            active_class = 'active' if is_expanded else ''

            # Get member count for badge
            member_count = getattr(app_org, 'MemberCount', 0) or 0

            html.append('<li class="trip-item app-org-item {0}" data-trip-id="{1}">'.format(expanded_class, org_id))

            # App Org Header (clickable to expand/navigate)
            html.append('''
                <div class="trip-header {0}" onclick="MissionsDashboard.toggleTrip('{1}')" title="{2}">
                    <span class="expand-icon">&#9654;</span>
                    <div class="trip-info">
                        <a href="?trip={1}" class="trip-link" onclick="event.stopPropagation();">{2}</a>
                    </div>
                    <span class="trip-member-badge" title="{3} members">&#128101; {3}</span>
                </div>
            '''.format(active_class, org_id, _escape_html(org_name), member_count))

            # App Org Sections (expandable) - same sections as trips
            html.append('<div class="trip-sections">')

            for section in config.TRIP_SECTIONS:
                section_id = section['id']
                section_label = section['label']
                section_icon = section['icon']

                # Hide Messages section for non-admins (app orgs are admin-only but keep for consistency)
                if section_id == 'messages' and not user_role.get('is_admin', False):
                    continue

                # Check if this section is active
                section_active = 'active' if is_expanded and current_section == section_id else ''

                html.append('''
                    <a href="?trip={0}&section={1}" class="section-link {2}"
                       onclick="MissionsDashboard.loadSection('{0}', '{1}', event)">
                        <span class="section-icon">{3}</span>
                        <span class="sidebar-text">{4}</span>
                    </a>
                '''.format(org_id, section_id, section_active, section_icon, section_label))

            html.append('</div>')  # End trip-sections
            html.append('</li>')  # End trip-item

        html.append('</ul></div>')

    # Mobile close button at bottom of sidebar
    html.append('<div class="mobile-close-btn"><button onclick="MissionsDashboard.closeMobile()">&#10005; Close Menu</button></div>')

    # Close sidebar
    html.append('</nav>')

    # Mobile overlay (must come after nav for CSS sibling selector to work)
    html.append('<div class="mobile-overlay" onclick="MissionsDashboard.closeMobile()"></div>')

    return ''.join(html)


def _convert_to_python_date(date_val):
    """
    Convert a .NET DateTime or other date format to Python datetime.

    Args:
        date_val: Date value (may be .NET DateTime, Python datetime, or string)

    Returns:
        datetime.datetime or None
    """
    if date_val is None:
        return None

    # Already a Python datetime with strftime
    if hasattr(date_val, 'strftime'):
        return date_val

    # .NET DateTime object has Year, Month, Day properties
    if hasattr(date_val, 'Year') and hasattr(date_val, 'Month') and hasattr(date_val, 'Day'):
        try:
            hour = date_val.Hour if hasattr(date_val, 'Hour') else 0
            minute = date_val.Minute if hasattr(date_val, 'Minute') else 0
            second = date_val.Second if hasattr(date_val, 'Second') else 0
            return datetime.datetime(date_val.Year, date_val.Month, date_val.Day, hour, minute, second)
        except:
            pass

    # Try parsing from string
    date_str = str(date_val).strip()
    if not date_str or date_str.lower() in ['none', 'null', '']:
        return None

    try:
        time_hour = 0
        time_min = 0
        time_sec = 0
        date_part = date_str

        if ' ' in date_str:
            parts_split = date_str.split(' ')
            date_part = parts_split[0]
            # Parse time portion if present (e.g. "3/18/2026 2:30:00 PM" or "2026-03-18 14:30:00")
            if len(parts_split) >= 2:
                time_str = parts_split[1]
                am_pm = parts_split[2].upper() if len(parts_split) >= 3 else ''
                time_parts = time_str.split(':')
                if len(time_parts) >= 2:
                    time_hour = int(time_parts[0])
                    time_min = int(time_parts[1])
                    if len(time_parts) >= 3:
                        time_sec = int(float(time_parts[2]))
                    # Handle AM/PM
                    if am_pm == 'PM' and time_hour < 12:
                        time_hour += 12
                    elif am_pm == 'AM' and time_hour == 12:
                        time_hour = 0

        # Handle M/D/YYYY format
        if '/' in date_part:
            parts = date_part.split('/')
            if len(parts) == 3:
                return datetime.datetime(int(parts[2]), int(parts[0]), int(parts[1]), time_hour, time_min, time_sec)

        # Handle YYYY-MM-DD format
        if '-' in date_part:
            parts = date_part.split('-')
            if len(parts) == 3:
                return datetime.datetime(int(parts[0]), int(parts[1]), int(parts[2]), time_hour, time_min, time_sec)
    except:
        pass

    return None


def _format_trip_date_range(start_date, end_date):
    """
    Format a date range for display in the sidebar.

    Args:
        start_date: Start date (datetime, .NET DateTime, or None)
        end_date: End date (datetime, .NET DateTime, or None)

    Returns:
        str: Formatted date range like "Oct 10 - 15, 2025" or "Dates TBD"
    """
    try:
        # Convert to Python datetime if needed
        start_date = _convert_to_python_date(start_date)
        end_date = _convert_to_python_date(end_date)

        if not start_date and not end_date:
            return "Dates TBD"

        # Format start date
        if start_date:
            start_str = start_date.strftime('%b %d')
            start_year = start_date.year
        else:
            return "Dates TBD"

        # Format end date
        if end_date:
            end_year = end_date.year
            # If same month and year, shorten the format
            if start_date.month == end_date.month and start_year == end_year:
                end_str = str(end_date.day)
            elif start_year == end_year:
                end_str = end_date.strftime('%b %d')
            else:
                end_str = end_date.strftime('%b %d, %Y')

            # Add year to end if not already there
            if start_year == end_year and '%Y' not in end_str:
                return "{0} - {1}, {2}".format(start_str, end_str, end_year)
            else:
                return "{0} - {1}".format(start_str, end_str)
        else:
            return "{0}, {1}".format(start_str, start_year)
    except:
        return "Dates TBD"


def _get_trip_status(trip):
    """
    Determine the status of a trip based on dates.
    Returns: 'Active', 'Upcoming', or 'Completed'
    """
    try:
        today = datetime.date.today()

        # Check if trip has TripBegin and TripEnd attributes
        trip_start = getattr(trip, 'TripBegin', None) or getattr(trip, 'TripStart', None)
        trip_end = getattr(trip, 'TripEnd', None)

        if trip_start and trip_end:
            # Convert to date if datetime
            if hasattr(trip_start, 'date'):
                trip_start = trip_start.date()
            if hasattr(trip_end, 'date'):
                trip_end = trip_end.date()

            if today < trip_start:
                return 'Upcoming'
            elif today > trip_end:
                return 'Completed'
            else:
                return 'Active'

        # Default to active if no dates
        return 'Active'
    except:
        return 'Active'


def _escape_html(text):
    """Escape HTML special characters in text."""
    if not text:
        return ''
    return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&#39;')


def _escape_js_string(text):
    """
    Escape text for use in JavaScript string literals within HTML onclick attributes.
    This properly handles characters that would break JavaScript parsing.

    The order matters:
    1. Escape backslashes first (\ -> \\)
    2. Escape single quotes (' -> \')
    3. Escape newlines and other control characters
    """
    if not text:
        return ''
    result = str(text)
    result = result.replace('\\', '\\\\')  # Escape backslashes first
    result = result.replace("'", "\\'")     # Escape single quotes
    result = result.replace('\n', '\\n')    # Escape newlines
    result = result.replace('\r', '\\r')    # Escape carriage returns
    result = result.replace('\t', '\\t')    # Escape tabs
    return result


# ::END:: Sidebar Navigation

#####################################################################
# TRIP SECTION VIEWS
#####################################################################

# ::START:: Trip Section Views
def render_trip_section(org_id, section, user_role):
    """
    Dispatcher function to render the appropriate trip section.

    Args:
        org_id: Organization/Trip ID
        section: Section name (overview, team, meetings, budget, documents, messages, tasks)
        user_role: User role info from get_user_role_and_trips()
    """
    # Block non-admins from accessing the messages section
    if section == 'messages' and not user_role.get('is_admin', False):
        section = 'overview'  # Redirect to overview

    # Check if non-admin user is approved for this trip
    # Sections that require approval: budget, documents, messages
    restricted_sections = ['budget', 'documents', 'messages']
    is_admin = user_role.get('is_admin', False)

    if section in restricted_sections and not is_admin:
        # Check user's approval status
        people_id = user_role.get('people_id')
        if people_id:
            approval_status = _get_approval_status(org_id, people_id)
            if approval_status != 'approved':
                # User is not approved - show pending notice
                return _render_pending_access_notice(org_id, section, approval_status)

    # Map sections to render functions
    section_map = {
        'overview': render_trip_overview,
        'team': render_trip_team,
        'meetings': render_trip_meetings,
        'budget': render_trip_budget,
        'documents': render_trip_documents,
        'messages': render_trip_messages,
        'tasks': render_trip_tasks
    }

    render_func = section_map.get(section, render_trip_overview)
    return render_func(org_id, user_role)


def _render_pending_access_notice(org_id, section, status):
    """Render a notice for pending/denied members trying to access restricted sections."""
    section_names = {
        'budget': 'Budget & Fundraising',
        'documents': 'Documents',
        'messages': 'Messages'
    }
    section_name = section_names.get(section, section.title())

    if status == 'denied':
        status_text = "Your registration has been denied"
        status_class = "denied"
        icon = "&#10007;"
        message = "Please contact the trip leader for more information about your registration status."
    else:
        status_text = "Your registration is pending review"
        status_class = "pending"
        icon = "&#8987;"
        message = "Once your registration is approved by a trip administrator, you will have access to this section."

    return '''
    <div class="pending-access-notice">
        <style>
            .pending-access-notice {
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                min-height: 300px;
                text-align: center;
                padding: 40px;
            }
            .pending-access-icon {
                font-size: 64px;
                margin-bottom: 20px;
            }
            .pending-access-icon.pending {
                color: #fd7e14;
            }
            .pending-access-icon.denied {
                color: #dc3545;
            }
            .pending-access-title {
                font-size: 1.5rem;
                font-weight: 600;
                color: #495057;
                margin-bottom: 10px;
            }
            .pending-access-subtitle {
                font-size: 1rem;
                color: #6c757d;
                margin-bottom: 20px;
            }
            .pending-access-badge {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                padding: 8px 16px;
                border-radius: 20px;
                font-size: 0.9rem;
                font-weight: 500;
                margin-bottom: 20px;
            }
            .pending-access-badge.pending {
                background: #fff3cd;
                color: #856404;
            }
            .pending-access-badge.denied {
                background: #f8d7da;
                color: #721c24;
            }
            .pending-access-message {
                max-width: 400px;
                color: #6c757d;
                font-size: 0.95rem;
                line-height: 1.5;
            }
        </style>
        <div class="pending-access-icon {1}">{2}</div>
        <div class="pending-access-title">Access to {0} is Restricted</div>
        <div class="pending-access-badge {1}">
            {2} {3}
        </div>
        <div class="pending-access-message">
            {4}
        </div>
    </div>
    '''.format(section_name, status_class, icon, status_text, message)


def _get_trip_info(org_id):
    """Get basic trip/organization info for headers."""
    # Get basic org info from SQL
    sql = '''
    SELECT
        o.OrganizationId,
        o.OrganizationName,
        o.LeaderId,
        leader.Name2 as LeaderName,
        o.Location,
        o.PendingLoc,
        o.Description,
        o.RegSettingXml,
        (SELECT COUNT(*) FROM OrganizationMembers om WITH (NOLOCK)
         WHERE om.OrganizationId = o.OrganizationId
           AND om.MemberTypeId NOT IN (230, 311)
           AND om.InactiveDate IS NULL) as MemberCount
    FROM Organizations o WITH (NOLOCK)
    LEFT JOIN People leader WITH (NOLOCK) ON o.LeaderId = leader.PeopleId
    WHERE o.OrganizationId = {0}
    '''.format(org_id)

    results = list(q.QuerySql(sql))
    if not results:
        return None

    row = results[0]

    # Build a simple object we can add attributes to
    class TripInfo:
        pass

    trip = TripInfo()
    trip.OrganizationId = row.OrganizationId
    trip.OrganizationName = row.OrganizationName
    trip.LeaderId = row.LeaderId
    trip.LeaderName = row.LeaderName
    trip.Location = row.Location
    trip.PendingLoc = row.PendingLoc
    trip.Description = row.Description
    trip.RegSettingXml = row.RegSettingXml
    trip.MemberCount = row.MemberCount

    # Use model.ExtraValueDateOrg for dates
    # The method may return a .NET DateTime, Python datetime, string, or None
    def parse_extra_value_date(date_val):
        """Parse date from ExtraValueDateOrg which may be datetime or string"""
        if date_val is None:
            return None

        # Convert to string first to handle .NET types
        date_str = str(date_val).strip()

        # Check for empty or null-like values
        if not date_str or date_str.lower() in ['none', 'null', '']:
            return None

        # If it already has strftime, it's a Python datetime object
        if hasattr(date_val, 'strftime'):
            return date_val

        # If it has Year/Month/Day properties, it's likely a .NET DateTime
        if hasattr(date_val, 'Year') and hasattr(date_val, 'Month') and hasattr(date_val, 'Day'):
            try:
                return datetime.datetime(date_val.Year, date_val.Month, date_val.Day)
            except:
                pass

        # Parse from string - handle "M/D/YYYY h:mm:ss AM/PM" format
        try:
            if ' ' in date_str:
                date_part = date_str.split(' ')[0]
            else:
                date_part = date_str
            parts = date_part.split('/')
            if len(parts) == 3:
                month = int(parts[0])
                day = int(parts[1])
                year = int(parts[2])
                if year > 1900 and month >= 1 and month <= 12 and day >= 1 and day <= 31:
                    return datetime.datetime(year, month, day)
        except:
            pass

        # Try using model.ParseDate as fallback
        try:
            parsed = model.ParseDate(date_str)
            if parsed:
                return parsed
        except:
            pass

        return None

    try:
        raw_begin = model.ExtraValueDateOrg(int(org_id), "Main Event Start")
        trip.TripBegin = parse_extra_value_date(raw_begin)
    except Exception as e:
        trip.TripBegin = None

    try:
        raw_end = model.ExtraValueDateOrg(int(org_id), "Main Event End")
        trip.TripEnd = parse_extra_value_date(raw_end)
    except Exception as e:
        trip.TripEnd = None

    try:
        raw_close = model.ExtraValueDateOrg(int(org_id), "Close")
        trip.TripClose = parse_extra_value_date(raw_close)
    except Exception as e:
        trip.TripClose = None

    return trip


def render_trip_overview(org_id, user_role):
    """Render the trip overview section with KPIs and quick stats."""
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    html = []

    # Get team emails for email popup
    team_emails_sql = '''
    SELECT p.EmailAddress
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
      AND p.EmailAddress IS NOT NULL
      AND p.EmailAddress != ''
    '''.format(org_id)
    team_emails_result = list(q.QuerySql(team_emails_sql))
    team_emails = [r.EmailAddress for r in team_emails_result]
    team_emails_js = ','.join(team_emails) if team_emails else ''

    # Prepare date strings for Edit Dates modal (ISO format for date inputs)
    start_date_iso = ''
    end_date_iso = ''
    close_date_iso = ''
    if trip.TripBegin and hasattr(trip.TripBegin, 'strftime'):
        start_date_iso = trip.TripBegin.strftime('%Y-%m-%d')
    if trip.TripEnd and hasattr(trip.TripEnd, 'strftime'):
        end_date_iso = trip.TripEnd.strftime('%Y-%m-%d')
    if trip.TripClose and hasattr(trip.TripClose, 'strftime'):
        close_date_iso = trip.TripClose.strftime('%Y-%m-%d')

    # Calculate countdown to trip
    days_until = None
    trip_status = None
    if trip.TripBegin and hasattr(trip.TripBegin, 'date'):
        today = datetime.date.today()
        trip_date = trip.TripBegin.date() if hasattr(trip.TripBegin, 'date') else trip.TripBegin
        days_until = (trip_date - today).days
        if days_until > 0:
            trip_status = 'upcoming'
        elif days_until == 0:
            trip_status = 'today'
        elif days_until < 0 and trip.TripEnd:
            trip_end_date = trip.TripEnd.date() if hasattr(trip.TripEnd, 'date') else trip.TripEnd
            if today <= trip_end_date:
                trip_status = 'in_progress'
            else:
                trip_status = 'completed'

    # Get team demographics early for leader info
    demographics = _get_team_demographics(org_id)
    leaders = demographics.get('leaders', [])

    # Get stats
    stats = _get_trip_stats(org_id)

    is_admin = user_role.get('is_admin', False)

    # Trip dates - formatted
    trip_begin_short = ''
    trip_end_short = ''
    if trip.TripBegin and hasattr(trip.TripBegin, 'strftime'):
        trip_begin_short = trip.TripBegin.strftime('%b %d')
    if trip.TripEnd and hasattr(trip.TripEnd, 'strftime'):
        trip_end_short = trip.TripEnd.strftime('%b %d, %Y')

    # === HERO HEADER SECTION ===
    html.append('''
    <style>
        .trip-hero {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            border-radius: 16px;
            padding: 28px 32px;
            margin-bottom: 24px;
            color: white;
        }
        .trip-hero-top {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            flex-wrap: wrap;
            gap: 16px;
            margin-bottom: 20px;
        }
        .trip-hero-title h1 {
            font-size: 2.1rem;
            font-weight: 700;
            margin: 0 0 6px 0;
            color: white;
            line-height: 1.2;
        }
        .trip-hero-location {
            font-size: 1.2rem;
            color: rgba(255,255,255,0.75);
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .trip-hero-countdown {
            text-align: right;
            flex-shrink: 0;
        }
        .countdown-number {
            font-size: 3.6rem;
            font-weight: 800;
            line-height: 1;
            color: #a5b4fc;
        }
        .countdown-label {
            font-size: 0.96rem;
            color: rgba(255,255,255,0.6);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 4px;
        }
        .trip-status-badge {
            display: inline-block;
            padding: 8px 18px;
            border-radius: 20px;
            font-size: 0.96rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .status-today { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
        .status-in-progress { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
        .status-completed { background: rgba(255,255,255,0.2); }
        .trip-hero-details {
            display: flex;
            flex-wrap: wrap;
            gap: 32px;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.1);
        }
        .trip-detail-item {
            min-width: 120px;
        }
        .trip-detail-item label {
            display: block;
            font-size: 0.84rem;
            color: rgba(255,255,255,0.5);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 6px;
        }
        .trip-detail-item .value {
            font-size: 1.2rem;
            font-weight: 500;
            color: white;
        }
        .trip-leader-avatars {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
        }
        .leader-item {
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
        }
        .leader-avatar {
            width: 56px;
            height: 56px;
            border-radius: 50%;
            border: 2px solid rgba(255,255,255,0.3);
            background: #667eea;
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 600;
            font-size: 1.1rem;
            flex-shrink: 0;
        }
        .leader-avatar img {
            width: 100%;
            height: 100%;
            border-radius: 50%;
            object-fit: cover;
        }
        .leader-name {
            font-size: 0.95rem;
            font-weight: 500;
            color: white;
            max-width: 120px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .hero-actions {
            display: flex;
            gap: 10px;
            margin-top: 20px;
            flex-wrap: wrap;
        }
        .hero-btn {
            padding: 9px 16px;
            border-radius: 8px;
            font-size: 1rem;
            font-weight: 500;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            border: none;
            transition: all 0.2s;
        }
        .hero-btn-primary {
            background: #667eea;
            color: white;
        }
        .hero-btn-primary:hover { background: #5a6fd6; color: white; }
        .hero-btn-secondary {
            background: rgba(255,255,255,0.1);
            color: white;
            border: 1px solid rgba(255,255,255,0.2);
        }
        .hero-btn-secondary:hover { background: rgba(255,255,255,0.2); color: white; }
        @media (max-width: 768px) {
            .trip-hero { padding: 16px 20px; }
            .trip-hero-title h1 { font-size: 1.62rem; }
            .trip-hero-location { font-size: 1rem; }
            .trip-hero-top { gap: 12px; margin-bottom: 16px; }
            .trip-hero-countdown { text-align: left; }
            .countdown-number { font-size: 2.4rem; }
            .countdown-label { font-size: 0.8rem; }
            .trip-hero-details { gap: 16px; padding-top: 16px; }
            .trip-detail-item { min-width: 100px; }
            .trip-detail-item label { font-size: 0.7rem; }
            .trip-detail-item .value { font-size: 1rem; }
            .leader-item { flex-direction: column; text-align: center; gap: 4px; }
            .leader-avatar { width: 48px; height: 48px; font-size: 1rem; }
            .leader-name { font-size: 0.85rem; max-width: 80px; }
            .hero-actions { gap: 8px; margin-top: 16px; }
            .hero-btn { padding: 8px 12px; font-size: 0.85rem; }
        }
    </style>
    ''')

    # Build the hero section
    html.append('<div class="trip-hero">')

    # Top section with title and countdown
    html.append('<div class="trip-hero-top">')

    # Title and location
    html.append('<div class="trip-hero-title">')
    html.append('<h1>{0}</h1>'.format(_escape_html(trip.OrganizationName)))
    location_text = trip.Location or trip.PendingLoc or 'Destination TBD'
    pending_note = ' (pending)' if not trip.Location and trip.PendingLoc else ''
    html.append('<div class="trip-hero-location">&#127757; {0}{1}</div>'.format(_escape_html(location_text), pending_note))
    html.append('</div>')

    # Countdown or status
    html.append('<div class="trip-hero-countdown">')
    if trip_status == 'upcoming' and days_until is not None:
        html.append('<div class="countdown-number">{0}</div>'.format(days_until))
        html.append('<div class="countdown-label">days to go</div>')
    elif trip_status == 'today':
        html.append('<span class="trip-status-badge status-today">Departing Today!</span>')
    elif trip_status == 'in_progress':
        html.append('<span class="trip-status-badge status-in-progress">Trip In Progress</span>')
    elif trip_status == 'completed':
        html.append('<span class="trip-status-badge status-completed">Completed</span>')
    html.append('</div>')

    html.append('</div>')  # trip-hero-top

    # Details row
    html.append('<div class="trip-hero-details">')

    # Dates
    html.append('<div class="trip-detail-item">')
    html.append('<label>Trip Dates</label>')
    if trip_begin_short and trip_end_short:
        html.append('<div class="value">{0} - {1}</div>'.format(trip_begin_short, trip_end_short))
    else:
        html.append('<div class="value" style="color: rgba(255,255,255,0.5);">Dates TBD</div>')
    html.append('</div>')

    # Team Size
    html.append('<div class="trip-detail-item">')
    html.append('<label>Team Size</label>')
    html.append('<div class="value">{0} members</div>'.format(stats.get('member_count', 0)))
    html.append('</div>')

    # Trip Cost (if set)
    standard_fee = stats.get('standard_fee', 0)
    if standard_fee > 0:
        html.append('<div class="trip-detail-item">')
        html.append('<label>Trip Cost</label>')
        html.append('<div class="value">{0}</div>'.format(format_currency(standard_fee)))
        html.append('</div>')

    # Trip Close Date (when trip goes inactive)
    html.append('<div class="trip-detail-item">')
    html.append('<label>Close Date</label>')
    if trip.TripClose and hasattr(trip.TripClose, 'strftime'):
        close_date_display = trip.TripClose.strftime('%b %d, %Y')
        # Check if trip is closed/inactive
        today = datetime.date.today()
        close_date = trip.TripClose.date() if hasattr(trip.TripClose, 'date') else trip.TripClose
        if close_date < today:
            html.append('<div class="value" style="color: #ff6b6b;">{0} (Inactive)</div>'.format(close_date_display))
        else:
            days_to_close = (close_date - today).days
            if days_to_close <= 7:
                html.append('<div class="value" style="color: #ffd93d;">{0} ({1} days)</div>'.format(close_date_display, days_to_close))
            else:
                html.append('<div class="value">{0}</div>'.format(close_date_display))
    else:
        html.append('<div class="value" style="color: rgba(255,255,255,0.5);">Not set</div>')
    html.append('</div>')

    # Leaders
    html.append('<div class="trip-detail-item">')
    html.append('<label>Trip Leaders</label>')
    if leaders:
        html.append('<div class="trip-leader-avatars">')
        for leader in leaders[:3]:  # Show max 3 avatars
            picture_url = leader.get('picture', '')
            leader_name = leader.get('name', '')
            leader_initial = leader_name[0].upper() if leader_name else '?'
            # Get first name only for display
            first_name = leader_name.split()[0] if leader_name else ''

            html.append('<div class="leader-item" onclick="PersonDetails.open({0}, {1})" title="{2}">'.format(
                leader['id'], org_id, _escape_html(leader_name)))
            if picture_url:
                html.append('<div class="leader-avatar"><img src="{0}" alt="{1}" onerror="this.parentElement.innerHTML=\'{2}\'"></div>'.format(
                    picture_url, _escape_html(leader_name), leader_initial))
            else:
                html.append('<div class="leader-avatar">{0}</div>'.format(leader_initial))
            html.append('<span class="leader-name">{0}</span>'.format(_escape_html(first_name)))
            html.append('</div>')
        if len(leaders) > 3:
            html.append('<div class="leader-item"><div class="leader-avatar" style="background: rgba(255,255,255,0.2);">+{0}</div></div>'.format(len(leaders) - 3))
        html.append('</div>')
    else:
        html.append('<div class="value" style="color: rgba(255,255,255,0.5);">None assigned</div>')
    html.append('</div>')

    html.append('</div>')  # trip-hero-details

    # Action buttons
    html.append('<div class="hero-actions">')
    html.append('<button onclick="MissionsEmail.openTeam(\'{0}\', \'{1}\')" class="hero-btn hero-btn-primary">&#9993; Email Team ({2})</button>'.format(
        org_id, team_emails_js.replace("'", "\\'"), len(team_emails)))
    html.append('<a href="?trip={0}&section=team" class="hero-btn hero-btn-secondary">&#128101; View Team</a>'.format(org_id))
    # Admin-only buttons: Create Meeting and Edit Dates
    if is_admin:
        html.append('<button onclick="MissionsMeeting.openCreate(\'{0}\', \'{1}\')" class="hero-btn hero-btn-secondary">&#128197; Create Meeting</button>'.format(
            org_id, _escape_html(trip.OrganizationName).replace("'", "\\'")))
    html.append('<a href="/Org/{0}" target="_blank" class="hero-btn hero-btn-secondary">&#128279; TouchPoint</a>'.format(org_id))
    if is_admin:
        html.append('<button onclick="MissionsDates.openEdit(\'{0}\', \'{1}\', \'{2}\', \'{3}\', \'{4}\')" class="hero-btn hero-btn-secondary">&#128197; Edit Dates</button>'.format(
            org_id, _escape_html(trip.OrganizationName).replace("'", "\\'"), start_date_iso, end_date_iso, close_date_iso))
    html.append('</div>')  # hero-actions

    # === MY MISSIONS SHARE LINK (admin only) ===
    if is_admin and leaders:
        html.append('''
        <div class="leader-access-notice" style="margin-top: 15px; padding: 12px 16px; background: rgba(255,255,255,0.1); border-radius: 8px; border-left: 3px solid #ffc107;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div>
                    <strong style="color: #ffc107;">&#127758; Share My Missions Link</strong>
                    <p style="margin: 4px 0 0 0; font-size: 13px; color: rgba(255,255,255,0.7);">
                        Share <strong>{my_missions_display}</strong> with team members to view their trip details, payments, and meetings.
                    </p>
                </div>
                <button onclick="copyMyMissionsLink()" class="hero-btn" style="background: rgba(255,255,255,0.2); border: 1px solid rgba(255,255,255,0.3); padding: 6px 12px; font-size: 13px;">
                    &#128203; Copy Link
                </button>
            </div>
        </div>
        <script>
        function copyMyMissionsLink() {{
            var link = '{my_missions_link}';
            if (navigator.clipboard) {{
                navigator.clipboard.writeText(link).then(function() {{
                    alert('Link copied!\\n\\n' + link + '\\n\\nTeam members can use this link to view their trip details, payment status, and upcoming meetings.');
                }});
            }} else {{
                prompt('Copy this link to share with team members:', link);
            }}
        }}
        </script>
        '''.format(org_id, my_missions_link=MY_MISSIONS_LINK, my_missions_display=MY_MISSIONS_DISPLAY))

    html.append('</div>')  # trip-hero

    # === FINANCIAL KPIs - SINGLE CLEAN ROW ===
    html.append('''
    <style>
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            border: 1px solid #e9ecef;
        }
        .stat-card.highlight {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
        }
        .stat-card.highlight .stat-value { color: white; }
        .stat-card.highlight .stat-label { color: rgba(255,255,255,0.8); }
        .stat-card.warning { border-left: 4px solid #fd7e14; }
        .stat-card.success { border-left: 4px solid #28a745; }
        .stat-value {
            font-size: 1.75rem;
            font-weight: 700;
            color: #333;
            line-height: 1.2;
        }
        .stat-label {
            font-size: 0.8rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-top: 6px;
        }
    </style>
    ''')

    deposit = stats.get('deposit', 0)
    total_raised = stats.get('total_raised', 0)
    total_outstanding = stats.get('total_outstanding', 0)
    upcoming_meetings = stats.get('upcoming_meetings', 0)

    html.append('<div class="stats-grid">')

    html.append('''
        <div class="stat-card success">
            <div class="stat-value">{0}</div>
            <div class="stat-label">Raised</div>
        </div>
    '''.format(format_currency(total_raised)))

    if total_outstanding > 0:
        html.append('''
            <div class="stat-card warning">
                <div class="stat-value">{0}</div>
                <div class="stat-label">Outstanding</div>
            </div>
        '''.format(format_currency(total_outstanding)))

    if deposit > 0:
        html.append('''
            <div class="stat-card">
                <div class="stat-value">{0}</div>
                <div class="stat-label">Deposit Required</div>
            </div>
        '''.format(format_currency(deposit)))

    html.append('''
        <div class="stat-card">
            <div class="stat-value">{0}</div>
            <div class="stat-label">Upcoming Meetings</div>
        </div>
    '''.format(upcoming_meetings))

    # No-charge members if any
    no_charge_count = stats.get('no_charge_count', 0)
    if no_charge_count > 0:
        html.append('''
            <div class="stat-card">
                <div class="stat-value">{0}</div>
                <div class="stat-label">No-Charge Members</div>
            </div>
        '''.format(no_charge_count))

    html.append('</div>')  # stats-grid

    # === APPROVAL SETTINGS CARD (Admin Only) ===
    if is_admin:
        trip_approval_setting = get_trip_approval_setting(org_id)
        approvals_enabled = are_approvals_enabled_for_trip(org_id)
        global_settings = get_global_settings()
        global_enabled = global_settings.get('approvals_enabled', True)

        # Determine current setting label for display
        if trip_approval_setting == 'enabled':
            setting_label = 'Enabled (Override)'
            setting_color = '#4caf50'
            setting_icon = '&#10003;'
        elif trip_approval_setting == 'disabled':
            setting_label = 'Disabled (Override)'
            setting_color = '#f44336'
            setting_icon = '&#10007;'
        else:
            if global_enabled:
                setting_label = 'Enabled (Global)'
                setting_color = '#4caf50'
                setting_icon = '&#10003;'
            else:
                setting_label = 'Disabled (Global)'
                setting_color = '#f44336'
                setting_icon = '&#10007;'

        html.append('''
        <div class="card" style="margin-bottom: 24px;">
            <div class="card-body" style="padding: 16px 20px;">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
                    <div>
                        <strong style="font-size: 14px; color: #333;">&#10003; Registration Approval Workflow</strong>
                        <p style="margin: 4px 0 0 0; font-size: 13px; color: #666;">
                            Current status: <span style="color: {0}; font-weight: 600;">{1} {2}</span>
                        </p>
                    </div>
                    <button onclick="TripApprovalSettings.showModal({3}, '{4}')"
                            style="padding: 8px 16px; border: 1px solid #ddd; border-radius: 6px; background: #f8f9fa; cursor: pointer; font-size: 13px;">
                        &#9881; Configure
                    </button>
                </div>
            </div>
        </div>
        '''.format(setting_color, setting_icon, setting_label, org_id, trip_approval_setting))

        # === TRIP FEE SETTINGS CARD (Admin Only, shown when approvals enabled) ===
        if approvals_enabled:
            fee_settings = get_trip_fee_settings(org_id)
            trip_cost_override = fee_settings.get('trip_cost', 0)
            deposit_amount = fee_settings.get('deposit_amount', 0)

            # Get TouchPoint's built-in trip cost for comparison
            builtin_trip_cost = 0
            try:
                cost_sql = '''
                SELECT TOP 1 ts.IndAmt
                FROM TransactionSummary ts WITH (NOLOCK)
                JOIN OrganizationMembers om WITH (NOLOCK) ON ts.OrganizationId = om.OrganizationId
                    AND ts.PeopleId = om.PeopleId AND ts.RegId = om.TranId
                WHERE ts.OrganizationId = {0} AND ts.IndAmt > 0
                '''.format(org_id)
                cost_result = list(q.QuerySql(cost_sql))
                if cost_result and cost_result[0].IndAmt:
                    builtin_trip_cost = int(cost_result[0].IndAmt)
            except:
                pass

            # Display effective trip cost (override or builtin)
            effective_cost = trip_cost_override if trip_cost_override > 0 else builtin_trip_cost
            cost_source = 'Override' if trip_cost_override > 0 else ('TouchPoint' if builtin_trip_cost > 0 else 'Not Set')

            html.append('''
        <div class="card" style="margin-bottom: 24px;">
            <div class="card-body" style="padding: 16px 20px;">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
                    <div>
                        <strong style="font-size: 14px; color: #333;">&#128176; Trip Fee Settings</strong>
                        <p style="margin: 4px 0 0 0; font-size: 13px; color: #666;">
                            Trip Cost: <span style="font-weight: 600;">{0}</span> <span style="color: #888;">({1})</span>
                            &nbsp;&bull;&nbsp;
                            Deposit: <span style="font-weight: 600;">{2}</span>
                        </p>
                    </div>
                    <button onclick="TripFeeSettings.showModal({3}, {4}, {5})"
                            style="padding: 8px 16px; border: 1px solid #ddd; border-radius: 6px; background: #f8f9fa; cursor: pointer; font-size: 13px;">
                        &#9881; Configure
                    </button>
                </div>
            </div>
        </div>
            '''.format(
                format_currency(effective_cost) if effective_cost > 0 else 'Not Set',
                cost_source,
                format_currency(deposit_amount) if deposit_amount > 0 else 'Not Set',
                org_id,
                trip_cost_override,
                deposit_amount
            ))

    # === ITEMS NEEDING ATTENTION ===
    missing_passport = stats.get('missing_passport', 0)
    missing_bgcheck = stats.get('missing_bgcheck', 0)
    missing_photo = stats.get('missing_photo', 0)
    total_missing = missing_passport + missing_bgcheck + missing_photo

    if total_missing > 0:
        html.append('''
        <style>
            .attention-bar {
                display: flex;
                flex-wrap: wrap;
                gap: 16px;
                background: #fff5f5;
                border: 1px solid #fed7d7;
                border-radius: 12px;
                padding: 16px 20px;
                margin-bottom: 24px;
            }
            .attention-item {
                display: flex;
                align-items: center;
                gap: 10px;
                color: #c53030;
            }
            .attention-item.warning { color: #c05621; }
            .attention-count {
                font-size: 1.25rem;
                font-weight: 700;
            }
            .attention-label a {
                color: inherit;
                text-decoration: underline;
            }
        </style>
        <div class="attention-bar">
        ''')

        if missing_passport > 0:
            html.append('''
                <div class="attention-item">
                    <span class="attention-count">{0}</span>
                    <span class="attention-label"><a href="?trip={1}&section=documents">Missing Passport</a></span>
                </div>
            '''.format(missing_passport, org_id))

        if missing_bgcheck > 0:
            html.append('''
                <div class="attention-item">
                    <span class="attention-count">{0}</span>
                    <span class="attention-label"><a href="?trip={1}&section=documents">Incomplete Background Check</a></span>
                </div>
            '''.format(missing_bgcheck, org_id))

        if missing_photo > 0:
            html.append('''
                <div class="attention-item warning">
                    <span class="attention-count">{0}</span>
                    <span class="attention-label"><a href="?trip={1}&section=team">Missing Photo</a></span>
                </div>
            '''.format(missing_photo, org_id))

        html.append('</div>')

    # === TWO COLUMN SECTION: TEAM & MEETINGS ===
    html.append('''
    <style>
        .overview-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 24px;
            margin-bottom: 24px;
        }
        @media (max-width: 768px) {
            .overview-grid { grid-template-columns: 1fr; }
        }
        .overview-card {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            border: 1px solid #e9ecef;
            overflow: hidden;
        }
        .overview-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 16px 20px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }
        .overview-card-header h3 {
            margin: 0;
            font-size: 1rem;
            font-weight: 600;
            color: #333;
        }
        .overview-card-body { padding: 20px; }
        .view-all-link {
            font-size: 0.85rem;
            color: #667eea;
            text-decoration: none;
        }
        .view-all-link:hover { text-decoration: underline; }
        .demo-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 0;
            border-bottom: 1px solid #f0f0f0;
        }
        .demo-row:last-child { border-bottom: none; }
        .demo-label { color: #666; font-size: 0.95rem; }
        .demo-value { font-weight: 600; color: #333; }
        .demo-bar {
            display: flex;
            height: 8px;
            background: #e9ecef;
            border-radius: 4px;
            overflow: hidden;
            margin-top: 8px;
        }
        .demo-bar-segment { height: 100%; }
        .meeting-item {
            padding: 12px 0;
            border-bottom: 1px solid #f0f0f0;
        }
        .meeting-item:last-child { border-bottom: none; }
        .meeting-upcoming { border-left: 3px solid #28a745; padding-left: 12px; margin-left: -20px; }
        .meeting-past { opacity: 0.7; }
        .meeting-title { font-weight: 500; color: #333; }
        .meeting-meta { font-size: 0.85rem; color: #666; margin-top: 4px; }
    </style>
    ''')

    html.append('<div class="overview-grid">')

    # Team Demographics Card
    html.append('<div class="overview-card">')
    html.append('<div class="overview-card-header">')
    html.append('<h3>Team Demographics</h3>')
    html.append('<a href="?trip={0}&section=team" class="view-all-link">View Team &rarr;</a>'.format(org_id))
    html.append('</div>')
    html.append('<div class="overview-card-body">')

    total = demographics.get('total', 0)
    if total > 0:
        # Gender row with bar
        gender = demographics.get('gender', {})
        male_count = gender.get('Male', 0)
        female_count = gender.get('Female', 0)
        male_pct = int(male_count * 100 / total) if total > 0 else 0
        female_pct = int(female_count * 100 / total) if total > 0 else 0

        html.append('<div class="demo-row">')
        html.append('<span class="demo-label">Gender</span>')
        html.append('<span class="demo-value">{0}M / {1}F</span>'.format(male_count, female_count))
        html.append('</div>')
        html.append('<div class="demo-bar">')
        html.append('<div class="demo-bar-segment" style="width: {0}%; background: #4e73df;"></div>'.format(male_pct))
        html.append('<div class="demo-bar-segment" style="width: {0}%; background: #e74a9e;"></div>'.format(female_pct))
        html.append('</div>')

        # Age distribution - simple text
        age_bins = demographics.get('age_bins', {})
        age_summary = []
        for age_group in ['0-12', '13-17', '18-25', '26-40', '41-60', '60+']:
            count = age_bins.get(age_group, 0)
            if count > 0:
                age_summary.append('{0}: {1}'.format(age_group, count))

        if age_summary:
            html.append('<div class="demo-row" style="margin-top: 12px;">')
            html.append('<span class="demo-label">Ages</span>')
            html.append('<span class="demo-value" style="font-size: 0.9rem;">{0}</span>'.format(' | '.join(age_summary)))
            html.append('</div>')

        # Experience
        repeat = demographics.get('repeat_participants', 0)
        first_timers = demographics.get('first_timers', 0)
        html.append('<div class="demo-row">')
        html.append('<span class="demo-label">Experience</span>')
        html.append('<span class="demo-value">{0} returning, {1} first-time</span>'.format(repeat, first_timers))
        html.append('</div>')

        # Church membership
        membership = demographics.get('membership', {})
        member_count = membership.get('Member', 0)
        html.append('<div class="demo-row">')
        html.append('<span class="demo-label">Church Members</span>')
        html.append('<span class="demo-value">{0} of {1}</span>'.format(member_count, total))
        html.append('</div>')

    else:
        html.append('<p style="color: #999; text-align: center; padding: 20px 0;">No team members yet</p>')

    html.append('</div>')  # overview-card-body
    html.append('</div>')  # overview-card (demographics)

    # Meetings Card - Show both past and upcoming (typically only 3-4 meetings per trip)
    html.append('<div class="overview-card">')
    html.append('<div class="overview-card-header">')
    html.append('<h3>Meetings</h3>')
    html.append('<a href="?trip={0}&section=meetings" class="view-all-link">View All &rarr;</a>'.format(org_id))
    html.append('</div>')
    html.append('<div class="overview-card-body">')

    # Add CSS for expandable attendance
    html.append('''
    <style>
        .meeting-attendance-toggle {
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 0.9rem;
            color: #667eea;
            margin-left: 8px;
            padding: 4px 8px;
            border-radius: 4px;
            background: rgba(102, 126, 234, 0.1);
        }
        .meeting-attendance-toggle:hover {
            background: rgba(102, 126, 234, 0.2);
        }
        .meeting-attendance-list {
            display: none;
            margin-top: 8px;
            padding: 8px 12px;
            background: #f8f9fa;
            border-radius: 6px;
            font-size: 0.95rem;
        }
        .meeting-attendance-list.expanded {
            display: block;
        }
        .attendance-present {
            color: #28a745;
            margin-right: 8px;
            display: inline-block;
        }
        .attendance-absent {
            color: #dc3545;
            margin-right: 8px;
            display: inline-block;
        }
        .attendance-section {
            margin-bottom: 6px;
        }
        .attendance-section-label {
            font-weight: 600;
            font-size: 0.85rem;
            text-transform: uppercase;
            color: #666;
            margin-bottom: 4px;
        }
        .headcount-note {
            color: #6c757d;
            font-style: italic;
            font-size: 0.95rem;
        }
        .meeting-email-btn {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 0.85rem;
            color: white;
            background: #667eea;
            padding: 4px 10px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            margin-left: 8px;
            transition: background 0.2s;
        }
        .meeting-email-btn:hover {
            background: #5a67d8;
        }
    </style>
    ''')

    try:
        # Get all meetings for this trip with attendance counts
        # Use COALESCE to prefer extra values (set via API) over column values
        all_meetings_sql = '''
        SELECT m.MeetingId, m.MeetingDate,
               COALESCE(me_desc.Data, m.Description) as Description,
               COALESCE(me_loc.Data, m.Location) as Location,
               m.NumPresent, m.HeadCount,
               CASE WHEN m.MeetingDate >= GETDATE() THEN 1 ELSE 0 END as IsUpcoming
        FROM Meetings m WITH (NOLOCK)
        LEFT JOIN MeetingExtra me_desc WITH (NOLOCK) ON m.MeetingId = me_desc.MeetingId AND me_desc.Field = 'Description'
        LEFT JOIN MeetingExtra me_loc WITH (NOLOCK) ON m.MeetingId = me_loc.MeetingId AND me_loc.Field = 'Location'
        WHERE m.OrganizationId = {0}
        ORDER BY m.MeetingDate DESC
        '''.format(org_id)
        all_meetings_list = list(q.QuerySql(all_meetings_sql))

        # Get all org members for attendance comparison
        org_members_sql = '''
        SELECT om.PeopleId, p.Name2 as Name
        FROM OrganizationMembers om WITH (NOLOCK)
        JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
        WHERE om.OrganizationId = {0}
        ORDER BY p.Name2
        '''.format(org_id)
        org_members = list(q.QuerySql(org_members_sql))
        org_member_ids = set([m.PeopleId for m in org_members])
        org_member_names = dict([(m.PeopleId, m.Name) for m in org_members])

        if all_meetings_list:
            # Separate upcoming and past meetings
            upcoming_meetings = [m for m in all_meetings_list if m.IsUpcoming == 1]
            past_meetings = [m for m in all_meetings_list if m.IsUpcoming == 0]

            # Show upcoming meetings first (in chronological order)
            if upcoming_meetings:
                html.append('<div style="margin-bottom: 12px;">')
                html.append('<div style="font-size: 0.75rem; color: #28a745; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; font-weight: 600;">&#9650; Upcoming</div>')
                for meeting in reversed(upcoming_meetings):  # chronological order
                    meeting_date_str = ''
                    meeting_date_display = ''
                    meeting_time_display = ''
                    if meeting.MeetingDate:
                        mtg_date = _convert_to_python_date(meeting.MeetingDate)
                        if mtg_date and hasattr(mtg_date, 'strftime'):
                            meeting_date_str = mtg_date.strftime('%a, %b %d @ %I:%M %p')
                            meeting_date_display = mtg_date.strftime('%A, %B %d, %Y')
                            meeting_time_display = mtg_date.strftime('%I:%M %p')

                    description = meeting.Description if meeting.Description else 'Team Meeting'
                    location = meeting.Location if meeting.Location else ''

                    # Escape values for JavaScript
                    trip_name_escaped = _escape_html(trip.OrganizationName).replace("'", "\\'").replace('"', '\\"')
                    desc_escaped = _escape_html(description).replace("'", "\\'").replace('"', '\\"')
                    loc_escaped = _escape_html(location).replace("'", "\\'").replace('"', '\\"') if location else ''

                    html.append('<div class="meeting-item meeting-upcoming">')
                    html.append('<div class="meeting-title">{0}'.format(_escape_html(description)))
                    # Add email button for upcoming meetings
                    html.append('<button class="meeting-email-btn" onclick="MissionsEmail.openMeetingReminder(\'{0}\', \'{1}\', \'{2}\', \'{3}\', \'{4}\', \'{5}\', \'{6}\')" title="Send meeting reminder email">&#9993; Email</button>'.format(
                        org_id,
                        team_emails_js.replace("'", "\\'"),
                        trip_name_escaped,
                        meeting_date_display,
                        meeting_time_display,
                        desc_escaped,
                        loc_escaped
                    ))
                    html.append('</div>')
                    html.append('<div class="meeting-meta">{0}{1}</div>'.format(
                        meeting_date_str,
                        ' &bull; {0}'.format(_escape_html(location)) if location else ''
                    ))
                    html.append('</div>')
                html.append('</div>')

            # Show past meetings (most recent first) with expandable attendance
            if past_meetings:
                html.append('<div style="margin-top: 12px;">')
                html.append('<div style="font-size: 0.75rem; color: #6c757d; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; font-weight: 600;">&#9660; Previous</div>')

                for idx, meeting in enumerate(past_meetings[:4]):  # Show up to 4 past meetings
                    meeting_id = meeting.MeetingId
                    meeting_date_str = ''
                    if meeting.MeetingDate:
                        mtg_date = _convert_to_python_date(meeting.MeetingDate)
                        if mtg_date and hasattr(mtg_date, 'strftime'):
                            meeting_date_str = mtg_date.strftime('%a, %b %d @ %I:%M %p')

                    description = meeting.Description if meeting.Description else 'Team Meeting'
                    location = meeting.Location if meeting.Location else ''
                    num_present = meeting.NumPresent or 0
                    head_count = meeting.HeadCount or 0
                    is_headcount_only = head_count > 0 and num_present == 0

                    html.append('<div class="meeting-item meeting-past" style="opacity: 0.85;">')
                    html.append('<div class="meeting-title" style="color: #6c757d;">')
                    html.append('{0}'.format(_escape_html(description)))

                    # Attendance toggle button
                    if is_headcount_only:
                        html.append('<span class="meeting-attendance-toggle" style="background: rgba(108, 117, 125, 0.1); color: #6c757d; cursor: default;">')
                        html.append('&#128101; {0} (headcount)'.format(head_count))
                        html.append('</span>')
                    else:
                        attendance_count = num_present if num_present > 0 else 0
                        html.append('<span class="meeting-attendance-toggle" onclick="toggleAttendance({0})"'.format(meeting_id))
                        html.append(' data-meeting="{0}">'.format(meeting_id))
                        html.append('&#128101; {0}/{1}'.format(attendance_count, len(org_members)))
                        html.append(' <span class="toggle-arrow" id="arrow-{0}">&#9660;</span>'.format(meeting_id))
                        html.append('</span>')

                    html.append('</div>')
                    html.append('<div class="meeting-meta">{0}{1}</div>'.format(
                        meeting_date_str,
                        ' &bull; {0}'.format(_escape_html(location)) if location else ''
                    ))

                    # Expandable attendance list (for non-headcount meetings)
                    if not is_headcount_only:
                        # Get attendance for this meeting
                        attendance_sql = '''
                        SELECT a.PeopleId, a.AttendanceFlag, p.Name2 as Name
                        FROM Attend a WITH (NOLOCK)
                        JOIN People p WITH (NOLOCK) ON a.PeopleId = p.PeopleId
                        WHERE a.MeetingId = {0}
                        '''.format(meeting_id)
                        attendance_records = list(q.QuerySql(attendance_sql))

                        present_ids = set([a.PeopleId for a in attendance_records if a.AttendanceFlag == 1 or a.AttendanceFlag == True])
                        present_names = [org_member_names.get(pid, 'Unknown') for pid in present_ids if pid in org_member_ids]
                        absent_names = [org_member_names.get(pid, 'Unknown') for pid in org_member_ids if pid not in present_ids]

                        present_names.sort()
                        absent_names.sort()

                        html.append('<div class="meeting-attendance-list" id="attendance-{0}">'.format(meeting_id))

                        if present_names:
                            html.append('<div class="attendance-section">')
                            html.append('<div class="attendance-section-label">&#10004; Present ({0})</div>'.format(len(present_names)))
                            html.append('<div>')
                            for name in present_names:
                                html.append('<span class="attendance-present">{0}</span>'.format(_escape_html(name)))
                            html.append('</div>')
                            html.append('</div>')

                        if absent_names:
                            html.append('<div class="attendance-section">')
                            html.append('<div class="attendance-section-label">&#10008; Absent ({0})</div>'.format(len(absent_names)))
                            html.append('<div>')
                            for name in absent_names:
                                html.append('<span class="attendance-absent">{0}</span>'.format(_escape_html(name)))
                            html.append('</div>')
                            html.append('</div>')

                        if not present_names and not absent_names:
                            html.append('<div class="headcount-note">No individual attendance recorded</div>')

                        html.append('</div>')

                    html.append('</div>')

                if len(past_meetings) > 4:
                    html.append('<p style="color: #666; font-size: 0.85rem; margin-top: 8px; margin-bottom: 0;">+ {0} more past meetings</p>'.format(len(past_meetings) - 4))
                html.append('</div>')

            # If no upcoming but has past meetings, show a note
            if not upcoming_meetings and past_meetings:
                html.append('<p style="color: #999; font-size: 0.85rem; font-style: italic; margin-top: 8px; margin-bottom: 0;">No upcoming meetings scheduled</p>')
        else:
            html.append('<p style="color: #999; text-align: center; padding: 20px 0;">No meetings yet</p>')

    except Exception as e:
        html.append('<p style="color: #dc3545;">Error loading meetings: {0}</p>'.format(str(e)))

    html.append('</div>')  # overview-card-body
    html.append('</div>')  # overview-card (meetings)

    html.append('</div>')  # overview-grid

    # === RECENT SIGN-UPS SECTION ===
    html.append('''
    <style>
        .signups-table {
            width: 100%;
            border-collapse: collapse;
        }
        .signups-table th {
            text-align: left;
            padding: 12px 16px;
            background: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            font-weight: 600;
            font-size: 0.85rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .signups-table td {
            padding: 12px 16px;
            border-bottom: 1px solid #f0f0f0;
        }
        .signups-table tr:last-child td { border-bottom: none; }
        .signups-table tr:hover { background: #f8f9fa; }
        .signup-name { color: #667eea; text-decoration: none; font-weight: 500; }
        .signup-name:hover { text-decoration: underline; }
        .signup-status-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 3px 8px;
            border-radius: 10px;
            font-size: 0.75rem;
            font-weight: 500;
        }
        .signup-status-pending { background: #fff3cd; color: #856404; }
        .signup-status-approved { background: #d4edda; color: #155724; }
        .signup-status-denied { background: #f8d7da; color: #721c24; }
        .signup-review-btn {
            padding: 4px 10px;
            font-size: 0.75rem;
            background: #667eea;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .signup-review-btn:hover { background: #5a6fd6; }
    </style>
    ''')

    html.append('<div class="overview-card">')
    html.append('<div class="overview-card-header">')
    html.append('<h3>Recent Sign-ups</h3>')
    html.append('</div>')
    html.append('<div class="overview-card-body" style="padding: 0;">')

    signups_sql = '''
    WITH TripSubgroups AS (
        SELECT
            omt.OrgId,
            omt.PeopleId,
            MAX(CASE WHEN sg.Name = 'trip-approved' THEN 1 ELSE 0 END) as IsApproved,
            MAX(CASE WHEN sg.Name = 'trip-denied' THEN 1 ELSE 0 END) as IsDenied
        FROM OrgMemMemTags omt WITH (NOLOCK)
        JOIN MemberTags sg WITH (NOLOCK) ON omt.MemberTagId = sg.Id
        WHERE sg.Name IN ('trip-approved', 'trip-denied')
          AND omt.OrgId = {0}
        GROUP BY omt.OrgId, omt.PeopleId
    )
    SELECT TOP 8
        p.PeopleId,
        p.Name,
        om.EnrollmentDate,
        mt.Description as MemberType,
        COALESCE(ts.IsApproved, 0) as IsApproved,
        COALESCE(ts.IsDenied, 0) as IsDenied
    FROM OrganizationMembers om WITH (NOLOCK)
    INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
    LEFT JOIN TripSubgroups ts ON om.OrganizationId = ts.OrgId AND om.PeopleId = ts.PeopleId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
    ORDER BY om.EnrollmentDate DESC
    '''.format(org_id)

    try:
        signups = list(q.QuerySql(signups_sql))
        # Check if approvals are enabled for this trip
        show_approval_status = are_approvals_enabled_for_trip(org_id)
        if signups:
            html.append('<table class="signups-table">')
            html.append('''
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Joined</th>
                        {0}
                        <th style="text-align: center;">Action</th>
                    </tr>
                </thead>
                <tbody>
            '''.format('<th>Status</th>' if show_approval_status else ''))
            for signup in signups:
                enrollment_date = ''
                if signup.EnrollmentDate:
                    try:
                        enrollment_date = signup.EnrollmentDate.strftime('%b %d, %Y')
                    except:
                        enrollment_date = str(signup.EnrollmentDate)[:10]

                # Determine status (only shown if approvals enabled)
                status_html = ''
                if show_approval_status:
                    if signup.IsDenied == 1:
                        status_html = '<td><span class="signup-status-badge signup-status-denied">&#10006; Denied</span></td>'
                    elif signup.IsApproved == 1:
                        status_html = '<td><span class="signup-status-badge signup-status-approved">&#10004; Approved</span></td>'
                    else:
                        status_html = '<td><span class="signup-status-badge signup-status-pending">&#9203; Pending</span></td>'

                # Action button - Review is always available, but approval actions depend on settings
                if show_approval_status and signup.IsDenied == 1:
                    action_html = '<button class="signup-review-btn" onclick="ApprovalWorkflow.revokeStatus({0}, {1})">Revoke</button>'.format(signup.PeopleId, org_id)
                elif show_approval_status and signup.IsApproved == 1:
                    action_html = '<button class="signup-review-btn" onclick="ApprovalWorkflow.revokeStatus({0}, {1})">Revoke</button>'.format(signup.PeopleId, org_id)
                else:
                    action_html = '<button class="signup-review-btn" onclick="ApprovalWorkflow.showModal({0}, \'{1}\', {2}, {3})">Review</button>'.format(
                        signup.PeopleId,
                        _escape_html(signup.Name or '').replace("'", "\\'"),
                        org_id,
                        'true' if show_approval_status else 'false'
                    )

                html.append('''
                    <tr data-people-id="{0}">
                        <td><a href="javascript:void(0)" onclick="PersonDetails.open({0}, {5})" class="signup-name">{1}</a></td>
                        <td>{2}</td>
                        {3}
                        <td style="text-align: center;">{4}</td>
                    </tr>
                '''.format(
                    signup.PeopleId,
                    _escape_html(signup.Name or ''),
                    enrollment_date,
                    status_html,
                    action_html,
                    org_id
                ))
            html.append('</tbody></table>')
        else:
            html.append('<p style="color: #999; text-align: center; padding: 20px;">No sign-ups yet for this trip.</p>')
    except Exception as e:
        html.append('<p style="color: #dc3545; padding: 16px;">Error loading sign-ups: {0}</p>'.format(str(e)))

    html.append('</div>')  # overview-card-body
    html.append('</div>')  # overview-card

    return ''.join(html)


def _get_trip_stats(org_id):
    """Get quick statistics for a trip including document/photo status."""
    stats = {
        'member_count': 0,
        'upcoming_meetings': 0,
        'total_raised': 0,
        'total_outstanding': 0,
        'total_cost': 0,
        'missing_passport': 0,
        'missing_bgcheck': 0,
        'missing_photo': 0,
        'standard_fee': 0,
        'deposit': 0,
        'no_charge_count': 0,
        'no_charge_value': 0
    }

    # Member count
    member_sql = '''
    SELECT COUNT(*) as cnt
    FROM OrganizationMembers
    WHERE OrganizationId = {0}
      AND MemberTypeId NOT IN (230, 311)
      AND InactiveDate IS NULL
    '''.format(org_id)
    result = list(q.QuerySql(member_sql))
    if result:
        stats['member_count'] = result[0].cnt

    # Upcoming meetings
    meeting_sql = '''
    SELECT COUNT(*) as cnt
    FROM Meetings
    WHERE OrganizationId = {0}
      AND MeetingDate >= GETDATE()
    '''.format(org_id)
    result = list(q.QuerySql(meeting_sql))
    if result:
        stats['upcoming_meetings'] = result[0].cnt

    # Financial stats using MissionTripTotals CTE
    try:
        mission_trip_cte = get_mission_trip_totals_cte(include_closed=True)
        finance_sql = mission_trip_cte + '''
        SELECT
            ISNULL(SUM(mtt.Raised), 0) as TotalRaised,
            ISNULL(SUM(mtt.Due), 0) as TotalOutstanding,
            ISNULL(SUM(mtt.TripCost), 0) as TotalCost
        FROM MissionTripTotals mtt
        WHERE mtt.InvolvementId = {0}
          AND mtt.PeopleId IS NOT NULL
          AND mtt.SortOrder <> 'ZZZZZ'
          AND mtt.Name <> 'total'
        '''.format(org_id)
        result = list(q.QuerySql(finance_sql))
        if result:
            stats['total_raised'] = float(result[0].TotalRaised or 0)
            stats['total_outstanding'] = float(result[0].TotalOutstanding or 0)
            stats['total_cost'] = float(result[0].TotalCost or 0)
    except:
        pass

    # Document status - missing passports, background checks, and photos
    # Use subqueries to avoid duplicate counts from multiple volunteer/recrg records
    try:
        doc_status_sql = '''
        SELECT
            (SELECT COUNT(DISTINCT p2.PeopleId)
             FROM OrganizationMembers om2 WITH (NOLOCK)
             JOIN People p2 WITH (NOLOCK) ON om2.PeopleId = p2.PeopleId
             LEFT JOIN RecReg rr WITH (NOLOCK) ON rr.PeopleId = p2.PeopleId
             WHERE om2.OrganizationId = {0}
               AND om2.MemberTypeId NOT IN (230, 311)
               AND om2.InactiveDate IS NULL
               AND (rr.passportnumber IS NULL OR rr.passportnumber = '')) as MissingPassport,
            (SELECT COUNT(DISTINCT p3.PeopleId)
             FROM OrganizationMembers om3 WITH (NOLOCK)
             JOIN People p3 WITH (NOLOCK) ON om3.PeopleId = p3.PeopleId
             LEFT JOIN Volunteer v WITH (NOLOCK) ON v.PeopleId = p3.PeopleId
             LEFT JOIN lookup.VolApplicationStatus vs WITH (NOLOCK) ON vs.Id = v.StatusId
             WHERE om3.OrganizationId = {0}
               AND om3.MemberTypeId NOT IN (230, 311)
               AND om3.InactiveDate IS NULL
               AND (vs.Description IS NULL OR vs.Description <> 'Complete')) as MissingBGCheck,
            (SELECT COUNT(DISTINCT p4.PeopleId)
             FROM OrganizationMembers om4 WITH (NOLOCK)
             JOIN People p4 WITH (NOLOCK) ON om4.PeopleId = p4.PeopleId
             WHERE om4.OrganizationId = {0}
               AND om4.MemberTypeId NOT IN (230, 311)
               AND om4.InactiveDate IS NULL
               AND p4.PictureId IS NULL) as MissingPhoto
        '''.format(org_id)
        result = list(q.QuerySql(doc_status_sql))
        if result:
            stats['missing_passport'] = int(result[0].MissingPassport or 0)
            stats['missing_bgcheck'] = int(result[0].MissingBGCheck or 0)
            stats['missing_photo'] = int(result[0].MissingPhoto or 0)
    except:
        pass

    # Get standard fee, deposit, and no-charge member info
    try:
        fee_sql = '''
        WITH OrgSettings AS (
            SELECT
                OrganizationId,
                RegSettingXml.value('(/Settings/Fees/Fee)[1]', 'decimal(18,2)') as StandardFee,
                RegSettingXml.value('(/Settings/Fees/Deposit)[1]', 'decimal(18,2)') as Deposit
            FROM Organizations WITH (NOLOCK)
            WHERE OrganizationId = {0}
        )
        SELECT
            os.StandardFee,
            os.Deposit,
            COUNT(CASE WHEN ISNULL(ts.IndAmt, 0) = 0 THEN 1 END) as NoChargeCount
        FROM OrgSettings os
        LEFT JOIN OrganizationMembers om WITH (NOLOCK) ON os.OrganizationId = om.OrganizationId
            AND om.MemberTypeId NOT IN (230, 311)
            AND om.InactiveDate IS NULL
        LEFT JOIN TransactionSummary ts WITH (NOLOCK) ON ts.OrganizationId = om.OrganizationId
            AND ts.PeopleId = om.PeopleId
            AND ts.RegId = om.TranId
        GROUP BY os.StandardFee, os.Deposit
        '''.format(org_id)
        result = list(q.QuerySql(fee_sql))
        if result:
            stats['standard_fee'] = float(result[0].StandardFee or 0)
            stats['deposit'] = float(result[0].Deposit or 0)
            stats['no_charge_count'] = int(result[0].NoChargeCount or 0)
            # Calculate no-charge value (what those members would have cost at standard fee)
            stats['no_charge_value'] = stats['no_charge_count'] * stats['standard_fee']
    except:
        pass

    return stats


def _get_team_demographics(org_id):
    """Get demographic breakdown of team members."""
    demographics = {
        'gender': {'Male': 0, 'Female': 0, 'Unknown': 0},
        'age_bins': {'0-12': 0, '13-17': 0, '18-25': 0, '26-40': 0, '41-60': 0, '60+': 0},
        'membership': {'Member': 0, 'Non-Member': 0},
        'baptized': {'Yes': 0, 'No/Unknown': 0},
        'repeat_participants': 0,
        'first_timers': 0,
        'total': 0,
        'leaders': []
    }

    try:
        # Get comprehensive member data with picture from Picture table
        demo_sql = '''
        SELECT
            p.PeopleId,
            p.Name2,
            p.GenderId,
            p.Age,
            p.MemberStatusId,
            p.BaptismDate,
            om.MemberTypeId,
            mt.Description as MemberType,
            pic.MediumUrl as PictureUrl,
            (SELECT COUNT(*)
             FROM OrganizationMembers om2 WITH (NOLOCK)
             JOIN Organizations o2 WITH (NOLOCK) ON om2.OrganizationId = o2.OrganizationId
             WHERE om2.PeopleId = p.PeopleId
               AND o2.IsMissionTrip = 1
               AND o2.OrganizationId <> {0}) as PriorTripCount
        FROM OrganizationMembers om WITH (NOLOCK)
        JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
        LEFT JOIN lookup.MemberType mt WITH (NOLOCK) ON om.MemberTypeId = mt.Id
        LEFT JOIN Picture pic WITH (NOLOCK) ON p.PictureId = pic.PictureId
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId NOT IN (230, 311)
          AND om.InactiveDate IS NULL
        ORDER BY
            CASE WHEN om.MemberTypeId IN (140, 310, 320) THEN 0 ELSE 1 END,
            p.Name2
        '''.format(org_id)

        members = list(q.QuerySql(demo_sql))
        demographics['total'] = len(members)

        for member in members:
            # Gender
            if member.GenderId == 1:
                demographics['gender']['Male'] += 1
            elif member.GenderId == 2:
                demographics['gender']['Female'] += 1
            else:
                demographics['gender']['Unknown'] += 1

            # Age bins
            age = member.Age if member.Age else 0
            if age <= 12:
                demographics['age_bins']['0-12'] += 1
            elif age <= 17:
                demographics['age_bins']['13-17'] += 1
            elif age <= 25:
                demographics['age_bins']['18-25'] += 1
            elif age <= 40:
                demographics['age_bins']['26-40'] += 1
            elif age <= 60:
                demographics['age_bins']['41-60'] += 1
            else:
                demographics['age_bins']['60+'] += 1

            # Church membership (MemberStatusId 10 = Member)
            if member.MemberStatusId == 10:
                demographics['membership']['Member'] += 1
            else:
                demographics['membership']['Non-Member'] += 1

            # Baptism status
            if member.BaptismDate:
                demographics['baptized']['Yes'] += 1
            else:
                demographics['baptized']['No/Unknown'] += 1

            # Repeat participation
            prior_trips = member.PriorTripCount if member.PriorTripCount else 0
            if prior_trips > 0:
                demographics['repeat_participants'] += 1
            else:
                demographics['first_timers'] += 1

            # Track leaders (MemberTypeId 140=Leader, 310=OrgLeader, 320=AssistantLeader)
            if member.MemberTypeId in [140, 310, 320]:
                # Extract first name for display
                if ',' in member.Name2:
                    name_parts = member.Name2.split(',')
                    display_name = '{0} {1}'.format(name_parts[1].strip(), name_parts[0].strip())
                else:
                    display_name = member.Name2
                demographics['leaders'].append({
                    'id': member.PeopleId,
                    'name': display_name,
                    'role': member.MemberType or 'Leader',
                    'picture': member.PictureUrl or ''
                })

        # Also check for org's LeaderId (may not have leader member type)
        # Get picture from Picture table via People.PictureId
        org_leader_sql = '''
        SELECT p.PeopleId, p.Name2, pic.MediumUrl as PictureUrl
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN People p WITH (NOLOCK) ON o.LeaderId = p.PeopleId
        LEFT JOIN Picture pic WITH (NOLOCK) ON p.PictureId = pic.PictureId
        WHERE o.OrganizationId = {0}
          AND o.LeaderId IS NOT NULL
        '''.format(org_id)
        org_leader_result = list(q.QuerySql(org_leader_sql))

        if org_leader_result and org_leader_result[0].PeopleId:
            leader_id = org_leader_result[0].PeopleId
            leader_name = org_leader_result[0].Name2 or 'Unknown'
            leader_picture = org_leader_result[0].PictureUrl or ''

            # Check if already in leaders list
            existing_ids = [l['id'] for l in demographics['leaders']]
            if leader_id not in existing_ids:
                # Format name
                if ',' in leader_name:
                    name_parts = leader_name.split(',')
                    display_name = '{0} {1}'.format(name_parts[1].strip(), name_parts[0].strip())
                else:
                    display_name = leader_name
                # Add org leader at beginning of list
                demographics['leaders'].insert(0, {
                    'id': leader_id,
                    'name': display_name,
                    'role': 'Trip Leader',
                    'picture': leader_picture
                })

    except Exception as e:
        demographics['error'] = str(e)

    return demographics


# =====================================================
# TRIP REGISTRATION APPROVAL WORKFLOW HELPER FUNCTIONS
# =====================================================

def _get_approval_status(org_id, people_id):
    """Get approval status for a specific person in a trip.

    Returns: 'approved', 'denied', or 'pending'
    """
    try:
        org_id = int(org_id)
        people_id = int(people_id)

        # Check subgroup membership
        if model.InSubGroup(people_id, org_id, 'trip-approved'):
            return 'approved'
        elif model.InSubGroup(people_id, org_id, 'trip-denied'):
            return 'denied'
        else:
            return 'pending'
    except:
        return 'pending'


def _get_all_approval_statuses(org_id):
    """Get approval statuses for all members in a trip.

    Returns: dict mapping people_id -> {status, denial_reason (if denied)}
    """
    statuses = {}

    try:
        org_id = int(org_id)

        # Get all org members
        members_sql = '''
        SELECT om.PeopleId
        FROM OrganizationMembers om WITH (NOLOCK)
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId NOT IN (230, 311)
          AND om.InactiveDate IS NULL
        '''.format(org_id)

        members = list(q.QuerySql(members_sql))

        # Get denial reasons
        denial_reasons = _get_denial_reasons(org_id)

        for member in members:
            pid = member.PeopleId
            status = _get_approval_status(org_id, pid)

            statuses[pid] = {
                'status': status,
                'denial_reason': None
            }

            # Add denial reason if denied
            if status == 'denied':
                for denial in denial_reasons.get('denials', []):
                    if denial.get('people_id') == pid:
                        statuses[pid]['denial_reason'] = denial.get('reason', '')
                        statuses[pid]['denied_by_name'] = denial.get('denied_by_name', '')
                        statuses[pid]['denied_date'] = denial.get('denied_date', '')
                        break

    except Exception as e:
        pass

    return statuses


def _get_denial_reasons(org_id):
    """Get denial reasons JSON from org extra value.

    Returns: dict with 'denials' list
    """
    try:
        org_id = int(org_id)
        json_str = model.ExtraValueOrg(org_id, 'TripApprovalReasons')

        if json_str:
            return json.loads(json_str)
        else:
            return {'denials': []}
    except:
        return {'denials': []}


def _add_denial_reason(org_id, people_id, name, reason, denied_by, denied_by_name):
    """Add or update a denial reason for a person.

    Args:
        org_id: Organization ID
        people_id: Person ID being denied
        name: Person's name
        reason: Denial reason text
        denied_by: PeopleId of admin who denied
        denied_by_name: Name of admin who denied
    """
    try:
        org_id = int(org_id)
        people_id = int(people_id)

        # Get current denials
        data = _get_denial_reasons(org_id)

        # Remove any existing entry for this person
        data['denials'] = [d for d in data.get('denials', []) if d.get('people_id') != people_id]

        # Add new denial
        import datetime
        data['denials'].append({
            'people_id': people_id,
            'name': name,
            'reason': reason,
            'denied_by': denied_by,
            'denied_by_name': denied_by_name,
            'denied_date': datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
        })

        # Save back to org extra value
        model.AddExtraValueTextOrg(org_id, 'TripApprovalReasons', json.dumps(data))
        return True
    except Exception as e:
        return False


def _remove_denial_reason(org_id, people_id):
    """Remove a denial reason for a person.

    Args:
        org_id: Organization ID
        people_id: Person ID to remove denial for
    """
    try:
        org_id = int(org_id)
        people_id = int(people_id)

        # Get current denials
        data = _get_denial_reasons(org_id)

        # Remove entry for this person
        data['denials'] = [d for d in data.get('denials', []) if d.get('people_id') != people_id]

        # Save back to org extra value
        model.AddExtraValueTextOrg(org_id, 'TripApprovalReasons', json.dumps(data))
        return True
    except:
        return False


def _get_registration_answers(org_id, people_id):
    """Get registration form questions and answers for a person.

    Returns: list of {question, answer} dicts

    Tries multiple sources:
    1. RegQuestion/RegAnswer tables (newer registrations)
    2. RegistrationData table (older registrations - XML parsed in Python)
    """
    import re
    answers = []

    try:
        org_id = int(org_id)
        people_id = int(people_id)

        # Try RegQuestion/RegAnswer tables first (newer registrations)
        # Note: RegQuestion uses Label column for question text, AnswerValue for the answer
        sql = '''
        SELECT
            rq.Label as Question,
            rq.[Order] as QuestionOrder,
            ra.AnswerValue as Answer
        FROM Registration r WITH (NOLOCK)
        JOIN RegPeople rp WITH (NOLOCK) ON r.RegistrationId = rp.RegistrationId
        LEFT JOIN RegAnswer ra WITH (NOLOCK) ON rp.RegPeopleId = ra.RegPeopleId
        LEFT JOIN RegQuestion rq WITH (NOLOCK) ON ra.RegQuestionId = rq.RegQuestionId
        WHERE r.OrganizationId = {0}
          AND rp.PeopleId = {1}
          AND rq.Label IS NOT NULL
        ORDER BY rq.[Order]
        '''.format(org_id, people_id)

        results = list(q.QuerySql(sql))

        for row in results:
            if row.Question:
                answers.append({
                    'question': row.Question,
                    'answer': row.Answer or ''
                })

        # If no results from RegAnswer, try RegistrationData table (older registrations)
        # Query the raw XML and parse it in Python for better compatibility
        if not answers:
            rd_sql = '''
            SELECT CAST(rd.Data AS NVARCHAR(MAX)) as XmlData
            FROM RegistrationData rd WITH (NOLOCK)
            WHERE rd.OrganizationId = {0}
              AND rd.completed = 1
            ORDER BY rd.Stamp DESC
            '''.format(org_id)

            rd_results = list(q.QuerySql(rd_sql))

            # Parse each XML record looking for the person
            for row in rd_results:
                if not row.XmlData:
                    continue

                xml_data = row.XmlData

                # Check if this XML contains the person we're looking for
                # Look for <PeopleId>123</PeopleId> pattern
                people_id_pattern = '<PeopleId>{0}</PeopleId>'.format(people_id)
                if people_id_pattern not in xml_data:
                    continue

                # Found the person - extract their OnlineRegPersonModel section
                # Find the section containing this PeopleId
                person_pattern = r'<OnlineRegPersonModel[^>]*>.*?<PeopleId>{0}</PeopleId>.*?</OnlineRegPersonModel>'.format(people_id)
                person_match = re.search(person_pattern, xml_data, re.DOTALL)

                if not person_match:
                    continue

                person_xml = person_match.group(0)

                # Extract ExtraQuestion elements: <ExtraQuestion set="0" question="Question text">Answer</ExtraQuestion>
                # Note: attributes may appear in any order, so we look for question="..." anywhere
                extra_pattern = r'<ExtraQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</ExtraQuestion>'
                for match in re.finditer(extra_pattern, person_xml):
                    question = match.group(1)
                    answer = match.group(2)
                    if question:
                        answers.append({
                            'question': question,
                            'answer': answer.strip() if answer else ''
                        })

                # Extract Text elements: <Text set="0" question="Question text">Answer</Text>
                text_pattern = r'<Text[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</Text>'
                for match in re.finditer(text_pattern, person_xml):
                    question = match.group(1)
                    answer = match.group(2)
                    if question:
                        answers.append({
                            'question': question,
                            'answer': answer.strip() if answer else ''
                        })

                # Extract YesNoQuestion elements: <YesNoQuestion question="Question text">True/False</YesNoQuestion>
                yesno_pattern = r'<YesNoQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</YesNoQuestion>'
                for match in re.finditer(yesno_pattern, person_xml):
                    question = match.group(1)
                    answer_val = match.group(2)
                    if question:
                        answers.append({
                            'question': question,
                            'answer': 'Yes' if answer_val.strip() == 'True' else 'No'
                        })

                # Found the person's data, no need to check more records
                if answers:
                    break

    except Exception as e:
        pass

    return answers


def render_trip_team(org_id, user_role):
    """Render team members for a trip."""
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    # Check if user is admin for role-based access
    is_admin = user_role.get('is_admin', False)

    html = []

    # Section header
    html.append('''
        <div class="section-header">
            <div>
                <div class="breadcrumb">
                    <a href="?">Dashboard</a> &rsaquo;
                    <a href="?trip={0}">{1}</a> &rsaquo; Team
                </div>
                <h2>Team Members</h2>
            </div>
        </div>
    '''.format(org_id, _escape_html(trip.OrganizationName)))

    # Get team members - passport info is in RecReg table, background check in Volunteer table
    # Picture URL from Picture table via People.PictureId
    members_sql = '''
    SELECT
        p.PeopleId,
        p.Name2 as Name,
        p.EmailAddress,
        p.CellPhone,
        mt.Description as MemberType,
        om.MemberTypeId,
        om.EnrollmentDate,
        CASE WHEN rr.passportnumber IS NOT NULL AND rr.passportexpires IS NOT NULL THEN 1 ELSE 0 END as HasPassportInfo,
        vs.Description as BGCheckStatus,
        p.PictureId,
        pic.MediumUrl as PictureUrl,
        p.Age,
        p.FamilyId
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
    LEFT JOIN RecReg rr WITH (NOLOCK) ON rr.PeopleId = p.PeopleId
    LEFT JOIN Volunteer v WITH (NOLOCK) ON v.PeopleId = p.PeopleId
    LEFT JOIN lookup.VolApplicationStatus vs WITH (NOLOCK) ON vs.Id = v.StatusId
    LEFT JOIN Picture pic WITH (NOLOCK) ON p.PictureId = pic.PictureId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
    ORDER BY
        CASE WHEN om.MemberTypeId IN (140, 310, 320) THEN 0 ELSE 1 END,
        p.FamilyId,
        p.PositionInFamilyId,
        p.Name2
    '''.format(org_id)

    members = list(q.QuerySql(members_sql))

    # Collect team emails for the email popup
    team_emails = [m.EmailAddress for m in members if m.EmailAddress]
    team_emails_js = ','.join(team_emails) if team_emails else ''

    # Section action bar
    html.append('''
        <div class="section-actions">
            <button onclick="MissionsEmail.openTeam('{0}', '{1}')" class="section-action-btn section-action-btn-primary">
                &#9993; Email All Team Members ({2})
            </button>
            <a href="/Org/{0}#tab-registrations" target="_blank" class="section-action-btn section-action-btn-secondary">
                &#128203; View Registrations
            </a>
        </div>
    '''.format(org_id, team_emails_js.replace("'", "\\'"), len(team_emails)))

    # Check if approvals are enabled for this trip
    approvals_enabled = are_approvals_enabled_for_trip(org_id)

    # Get approval statuses for all members (admin only, and only if approvals enabled)
    approval_statuses = {}
    status_counts = {'all': len(members), 'pending': 0, 'approved': 0, 'denied': 0}
    if is_admin and approvals_enabled:
        approval_statuses = _get_all_approval_statuses(org_id)
        for pid, status_info in approval_statuses.items():
            status = status_info.get('status', 'pending')
            if status in status_counts:
                status_counts[status] += 1

    # Add approval filter tabs (admin only, only if approvals enabled)
    if is_admin and approvals_enabled:
        html.append('''
            <div class="approval-filter-tabs" data-org-id="{0}" style="display: flex; gap: 5px; flex-wrap: wrap; margin-bottom: 10px;">
                <button class="approval-filter-tab active" data-filter="all" onclick="ApprovalWorkflow.filterTable('all')">
                    All <span class="filter-count">({1})</span>
                </button>
                <button class="approval-filter-tab" data-filter="pending" onclick="ApprovalWorkflow.filterTable('pending')">
                    Pending <span class="filter-count pending-count">({2})</span>
                </button>
                <button class="approval-filter-tab" data-filter="approved" onclick="ApprovalWorkflow.filterTable('approved')">
                    Approved <span class="filter-count approved-count">({3})</span>
                </button>
                <button class="approval-filter-tab" data-filter="denied" onclick="ApprovalWorkflow.filterTable('denied')">
                    Denied <span class="filter-count denied-count">({4})</span>
                </button>
            </div>
        '''.format(org_id, status_counts['all'], status_counts['pending'], status_counts['approved'], status_counts['denied']))

    html.append('<div class="card">')
    html.append('<div class="card-body">')
    html.append('<table class="mission-table" id="team-members-table">')
    # Build table header - include Status column for admins only if approvals enabled
    if is_admin and approvals_enabled:
        html.append('''
            <thead>
                <tr>
                    <th style="width: 200px;">Name</th>
                    <th class="text-center">Status</th>
                    <th>Role</th>
                    <th class="text-center">Age</th>
                    <th>Contact</th>
                    <th class="text-center">Photo</th>
                    <th class="text-center">Passport</th>
                    <th class="text-center">Background</th>
                    <th class="text-center">Actions</th>
                </tr>
            </thead>
            <tbody>
        ''')
    else:
        # Non-admins or admins with approvals disabled - no Status column
        html.append('''
            <thead>
                <tr>
                    <th style="width: 220px;">Name</th>
                    <th>Role</th>
                    <th class="text-center">Age</th>
                    <th>Contact</th>
                    <th class="text-center">Photo</th>
                    <th class="text-center">Passport</th>
                    <th class="text-center">Background</th>
                    <th class="text-center">Actions</th>
                </tr>
            </thead>
            <tbody>
        ''')

    # Track family grouping for visual bands
    current_family_id = None
    family_band_index = 0  # Alternates 0/1 for different background colors

    for member in members:
        # Track family grouping - toggle band color when family changes
        if member.FamilyId != current_family_id:
            current_family_id = member.FamilyId
            family_band_index = (family_band_index + 1) % 2
        family_band_class = 'family-band-{0}'.format(family_band_index)

        # Determine role badge
        role_class = ''
        if member.MemberTypeId in [140, 310]:
            role_class = 'status-active'
        elif member.MemberTypeId == 320:
            role_class = 'status-medium'

        # Photo status
        has_photo = member.PictureId is not None
        photo_icon = '&#10003;' if has_photo else '&#10007;'
        photo_class = 'text-success' if has_photo else 'text-warning'

        # Passport status - make clickable if missing to send passport form request
        if member.HasPassportInfo:
            passport_icon = '&#10003;'
            passport_class = 'text-success'
            passport_html = '<span class="{0}">{1}</span>'.format(passport_class, passport_icon)
        else:
            # Make the X icon clickable to send passport form email
            passport_icon = '&#10007;'  # Keep for backward compatibility
            passport_class = 'text-danger'
            # Use _escape_js_string for JavaScript onclick handlers
            escaped_member_name_js = _escape_js_string(member.Name or '')
            escaped_trip_name_js = _escape_js_string(trip.OrganizationName or '')
            escaped_email_passport_js = _escape_js_string(member.EmailAddress or '')
            passport_html = '''<a href="#" onclick="event.stopPropagation(); MissionsEmail.openPassportRequest({0}, '{1}', '{2}', '{3}', {4}); return false;"
                              class="text-danger" style="cursor: pointer; text-decoration: none;"
                              title="Click to send passport form request">&#10007;</a>'''.format(
                member.PeopleId, escaped_email_passport_js, escaped_member_name_js, escaped_trip_name_js, org_id)

        # Background check status (from Volunteer table via VolApplicationStatus lookup)
        bg_status = member.BGCheckStatus
        if bg_status and bg_status.lower() == 'complete':
            bg_icon = '&#10003;'
            bg_class = 'text-success'
            bg_html = '<span class="{0}">{1}</span>'.format(bg_class, bg_icon)
        elif bg_status:
            bg_icon = '&#10007;'
            bg_class = 'text-warning'
            bg_html = '''<a href="/Person2/{0}#tab-volunteer" target="_blank"
                          onclick="event.stopPropagation();"
                          class="{1}" style="text-decoration: none;"
                          title="Click to open volunteer tab and run background check">{2}</a>'''.format(
                member.PeopleId, bg_class, bg_icon)
        else:
            bg_icon = '&#8211;'
            bg_class = 'text-muted'
            bg_html = '''<a href="/Person2/{0}#tab-volunteer" target="_blank"
                          onclick="event.stopPropagation();"
                          class="{1}" style="text-decoration: none;"
                          title="Click to open volunteer tab and run background check">{2}</a>'''.format(
                member.PeopleId, bg_class, bg_icon)

        # Prepare data for email popup
        email_data = member.EmailAddress or ''
        name_parts = member.Name.split(',') if member.Name else ['', '']
        first_name = name_parts[1].strip().split(' ')[0] if len(name_parts) > 1 else name_parts[0]

        # Build dropdown menu items for Actions column
        # Use _escape_js_string for JavaScript onclick handlers
        escaped_member_name_action = _escape_js_string(member.Name or '')
        escaped_trip_name_action = _escape_js_string(trip.OrganizationName or '')
        escaped_email_action = _escape_js_string(member.EmailAddress or '')

        # Passport menu item - admin only
        if is_admin:
            if member.HasPassportInfo:
                passport_menu_item = '''<button class="actions-dropdown-item" onclick="event.stopPropagation(); window.open('/Person2/{0}#tab-personal', '_blank'); ActionsDropdown.closeAll();">
                    <span class="icon">&#128274;</span>View Passport<span class="status-dot complete"></span>
                </button>'''.format(member.PeopleId)
            else:
                passport_menu_item = '''<button class="actions-dropdown-item" onclick="event.stopPropagation(); MissionsEmail.openPassportRequest({0}, '{1}', '{2}', '{3}', {4}); ActionsDropdown.closeAll();">
                    <span class="icon">&#128274;</span>Request Passport<span class="status-dot incomplete"></span>
                </button>'''.format(member.PeopleId, escaped_email_action, escaped_member_name_action, escaped_trip_name_action, org_id)
        else:
            passport_menu_item = ''

        # Background check menu item - admin only
        bg_status = member.BGCheckStatus
        if is_admin:
            if bg_status and bg_status.lower() == 'complete':
                bgcheck_menu_item = '''<button class="actions-dropdown-item" onclick="event.stopPropagation(); window.open('/Person2/{0}#tab-volunteer', '_blank'); ActionsDropdown.closeAll();">
                    <span class="icon">&#128100;</span>Background Check<span class="status-dot complete"></span>
                </button>'''.format(member.PeopleId)
            else:
                bgcheck_menu_item = '''<button class="actions-dropdown-item" onclick="event.stopPropagation(); window.open('/Person2/{0}#tab-volunteer', '_blank'); ActionsDropdown.closeAll();">
                    <span class="icon">&#128100;</span>Run Background Check<span class="status-dot warning"></span>
                </button>'''.format(member.PeopleId)
        else:
            bgcheck_menu_item = ''

        # Member type menu item - admin only
        if is_admin:
            if member.MemberTypeId in [140, 310]:
                # Currently a leader - offer to set as Member
                member_type_menu_item = '''<button class="actions-dropdown-item" onclick="event.stopPropagation(); MemberRole.set({0}, {1}, 'Member'); ActionsDropdown.closeAll();">
                    <span class="icon">&#128101;</span>Set as Member
                </button>'''.format(member.PeopleId, org_id)
            else:
                # Currently a member/other - offer to set as Leader
                member_type_menu_item = '''<button class="actions-dropdown-item" onclick="event.stopPropagation(); MemberRole.set({0}, {1}, 'Leader'); ActionsDropdown.closeAll();">
                    <span class="icon">&#11088;</span>Set as Leader
                </button>'''.format(member.PeopleId, org_id)
        else:
            member_type_menu_item = ''

        # Build clickable email HTML
        if email_data:
            # Use _escape_js_string for JavaScript onclick handlers
            escaped_email_js = _escape_js_string(email_data)
            escaped_name_js = _escape_js_string(first_name)
            email_html = '<a href="#" onclick="event.stopPropagation(); MissionsEmail.openIndividual({0}, \'{1}\', \'{2}\', {4}); return false;" style="color: #007bff;">{3}</a>'.format(
                member.PeopleId, escaped_email_js, escaped_name_js, _escape_html(email_data), org_id)
        else:
            email_html = ''

        # Build clickable phone HTML with formatting
        cell_phone = member.CellPhone or ''
        if cell_phone:
            # Format phone for display
            formatted_phone = model.FmtPhone(cell_phone, "") if cell_phone else ''
            # Clean digits for tel: link
            phone_digits = ''.join(c for c in cell_phone if c.isdigit())
            phone_html = '<a href="tel:{0}" onclick="event.stopPropagation();" style="color: #007bff;">{1}</a>'.format(phone_digits, _escape_html(formatted_phone))
        else:
            phone_html = ''

        # Use picture URL from Picture table if available
        picture_url = member.PictureUrl or ''
        if picture_url:
            img_html = '''<img src="{0}"
                             style="width: 40px; height: 40px; border-radius: 50%; object-fit: cover; border: 2px solid {1};"
                             onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';">
                         <div style="display: none; width: 40px; height: 40px; border-radius: 50%; background: #6c757d; color: white; align-items: center; justify-content: center; font-weight: bold; font-size: 14px;">{2}</div>'''.format(picture_url, '#28a745' if has_photo else '#fd7e14', member.Name[0].upper() if member.Name else '?')
        else:
            # No picture - show initials placeholder
            img_html = '''<div style="width: 40px; height: 40px; border-radius: 50%; background: #6c757d; color: white; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 14px; border: 2px solid #fd7e14;">{0}</div>'''.format(member.Name[0].upper() if member.Name else '?')

        # Build approval status cell for admins (only if approvals enabled)
        member_approval_status = approval_statuses.get(member.PeopleId, {}).get('status', 'pending') if (is_admin and approvals_enabled) else ''
        status_cell_html = ''
        if is_admin and approvals_enabled:
            escaped_name_for_js = _escape_js_string(member.Name or '')
            if member_approval_status == 'approved':
                status_cell_html = '''<td data-label="Status" class="text-center">
                    <span class="approval-status-badge approved" title="Approved">&#10003; Approved</span>
                    <button class="btn-revoke-small" onclick="event.stopPropagation(); ApprovalWorkflow.revoke({0}, {1})" title="Revoke approval">&#8635;</button>
                </td>'''.format(member.PeopleId, org_id)
            elif member_approval_status == 'denied':
                denial_info = approval_statuses.get(member.PeopleId, {})
                denial_reason = _escape_html(denial_info.get('denial_reason', '') or '')
                status_cell_html = '''<td data-label="Status" class="text-center">
                    <span class="approval-status-badge denied" title="Denied: {2}">&#10007; Denied</span>
                    <button class="btn-revoke-small" onclick="event.stopPropagation(); ApprovalWorkflow.revoke({0}, {1})" title="Revoke denial">&#8635;</button>
                </td>'''.format(member.PeopleId, org_id, denial_reason)
            else:  # pending
                status_cell_html = '''<td data-label="Status" class="text-center">
                    <span class="approval-status-badge pending">&#8987; Pending</span>
                    <button class="btn-approve-small" onclick="event.stopPropagation(); ApprovalWorkflow.showModal({0}, '{1}', {2}, true)" title="Review &amp; Approve/Deny">Review</button>
                </td>'''.format(member.PeopleId, escaped_name_for_js, org_id)

        # Build table row - different structure for admin with approvals vs others
        if is_admin and approvals_enabled:
            html.append('''
                <tr class="{21}" data-people-id="{0}" data-status="{25}">
                    <td data-label="Name">
                        <div style="display: flex; align-items: center; gap: 10px; cursor: pointer;" onclick="PersonDetails.open({0}, {16})">
                            {15}
                            <span style="color: #007bff; font-weight: 500;">{1}</span>
                        </div>
                    </td>
                    {26}
                    <td data-label="Role"><span class="status-badge {2}">{3}</span></td>
                    <td data-label="Age" class="text-center">{20}</td>
                    <td data-label="Contact">
                        <span style="font-size: 10px;">{17}</span><br>
                        <span style="font-size: 10px;">{18}</span>
                    </td>
                    <td data-label="Photo" class="text-center {6}">{7}</td>
                    <td data-label="Passport" class="text-center">{19}</td>
                    <td data-label="Background" class="text-center">{22}</td>
                    <td data-label="Actions" class="text-center">
                        <div class="actions-dropdown">
                            <button class="actions-dropdown-btn" onclick="event.stopPropagation(); ActionsDropdown.toggle(this);">
                                Actions &#9662;
                            </button>
                            <div class="actions-dropdown-content">
                                <button class="actions-dropdown-item" onclick="event.stopPropagation(); ApprovalWorkflow.showModal({0}, '{14}', {16}, true); ActionsDropdown.closeAll();">
                                    <span class="icon">&#128203;</span>Review Registration
                                </button>
                                <button class="actions-dropdown-item" onclick="event.stopPropagation(); MissionsEmail.openIndividual({0}, '{13}', '{14}', {16}); ActionsDropdown.closeAll();">
                                    <span class="icon">&#9993;</span>Send Email
                                </button>
                                {23}
                                {24}
                                {27}
                                <button class="actions-dropdown-item" onclick="event.stopPropagation(); window.open('/Person2/{0}', '_blank'); ActionsDropdown.closeAll();">
                                    <span class="icon">&#128100;</span>View Profile
                                </button>
                            </div>
                        </div>
                    </td>
                </tr>
            '''.format(
                member.PeopleId,
                _escape_html(member.Name),
                role_class,
                _escape_html(member.MemberType or 'Member'),
                _escape_html(member.EmailAddress or ''),  # 4
                _escape_html(member.CellPhone or ''),  # 5
                photo_class,
                photo_icon,
                passport_class,  # 8
                passport_icon,   # 9
                '#28a745' if has_photo else '#fd7e14',  # 10
                bg_class,  # 11
                bg_icon,   # 12
                _escape_html(email_data).replace("'", "\\'"),  # 13
                _escape_html(first_name).replace("'", "\\'"),  # 14
                img_html,  # 15
                org_id,  # 16
                email_html,  # 17
                phone_html,  # 18
                passport_html,  # 19
                member.Age if member.Age is not None else '-',  # 20
                family_band_class,  # 21
                bg_html,  # 22
                passport_menu_item,  # 23
                bgcheck_menu_item,  # 24
                member_approval_status,  # 25: data-status value
                status_cell_html,  # 26: status cell HTML
                member_type_menu_item  # 27: set leader/member menu item
            ))
        else:
            html.append('''
                <tr class="{21}">
                    <td data-label="Name">
                        <div style="display: flex; align-items: center; gap: 10px; cursor: pointer;" onclick="PersonDetails.open({0}, {16})">
                            {15}
                            <span style="color: #007bff; font-weight: 500;">{1}</span>
                        </div>
                    </td>
                    <td data-label="Role"><span class="status-badge {2}">{3}</span></td>
                    <td data-label="Age" class="text-center">{20}</td>
                    <td data-label="Contact">
                        <span style="font-size: 10px;">{17}</span><br>
                        <span style="font-size: 10px;">{18}</span>
                    </td>
                    <td data-label="Photo" class="text-center {6}">{7}</td>
                    <td data-label="Passport" class="text-center">{19}</td>
                    <td data-label="Background" class="text-center">{22}</td>
                    <td data-label="Actions" class="text-center">
                        <div class="actions-dropdown">
                            <button class="actions-dropdown-btn" onclick="event.stopPropagation(); ActionsDropdown.toggle(this);">
                                Actions &#9662;
                            </button>
                            <div class="actions-dropdown-content">
                                <button class="actions-dropdown-item" onclick="event.stopPropagation(); ApprovalWorkflow.showModal({0}, '{14}', {16}, false); ActionsDropdown.closeAll();">
                                    <span class="icon">&#128203;</span>Review Registration
                                </button>
                                <button class="actions-dropdown-item" onclick="event.stopPropagation(); MissionsEmail.openIndividual({0}, '{13}', '{14}', {16}); ActionsDropdown.closeAll();">
                                    <span class="icon">&#9993;</span>Send Email
                                </button>
                                {23}
                                {24}
                                {25}
                                <button class="actions-dropdown-item" onclick="event.stopPropagation(); window.open('/Person2/{0}', '_blank'); ActionsDropdown.closeAll();">
                                    <span class="icon">&#128100;</span>View Profile
                                </button>
                            </div>
                        </div>
                    </td>
                </tr>
            '''.format(
            member.PeopleId,
            _escape_html(member.Name),
            role_class,
            _escape_html(member.MemberType or 'Member'),
            _escape_html(member.EmailAddress or ''),  # 4: plain email (unused now)
            _escape_html(member.CellPhone or ''),  # 5: plain phone (unused now)
            photo_class,
            photo_icon,
            passport_class,
            passport_icon,
            '#28a745' if has_photo else '#fd7e14',  # 10: border color (not used now but kept for compatibility)
            bg_class,  # 11 (unused, kept for compatibility)
            bg_icon,   # 12 (unused, kept for compatibility)
            _escape_html(email_data).replace("'", "\\'"),  # 13: email for JS
            _escape_html(first_name).replace("'", "\\'"),  # 14: first name for JS
            img_html,  # 15: image HTML
            org_id,  # 16: org_id for popup
            email_html,  # 17: clickable email HTML
            phone_html,  # 18: clickable phone HTML with formatting
            passport_html,  # 19: passport status HTML (clickable if missing)
            member.Age if member.Age is not None else '-',  # 20: age
            family_band_class,  # 21: family grouping CSS class
            bg_html,  # 22: background check HTML (clickable if not complete)
            passport_menu_item,  # 23: passport dropdown menu item
            bgcheck_menu_item,  # 24: background check dropdown menu item
            member_type_menu_item  # 25: set leader/member menu item
        ))

    html.append('</tbody></table>')
    html.append('</div>')  # card-body
    html.append('</div>')  # card

    return ''.join(html)


def render_trip_meetings(org_id, user_role):
    """Render meetings for a trip with attendance details and email functionality."""
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    # Check if user is admin for role-based access
    is_admin = user_role.get('is_admin', False)

    html = []

    # Add CSS styles for attendance and email functionality
    html.append('''
    <style>
        .meeting-card {
            padding: 12px 16px;
            background: white;
            border-radius: 8px;
            margin-bottom: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .meeting-attendance-toggle {
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 0.9rem;
            color: #667eea;
            margin-left: 8px;
            padding: 4px 8px;
            border-radius: 4px;
            background: rgba(102, 126, 234, 0.1);
        }
        .meeting-attendance-toggle:hover {
            background: rgba(102, 126, 234, 0.2);
        }
        .meeting-attendance-list {
            display: none;
            margin-top: 8px;
            padding: 8px 12px;
            background: #f8f9fa;
            border-radius: 6px;
            font-size: 0.95rem;
        }
        .meeting-attendance-list.expanded {
            display: block;
        }
        .attendance-present {
            color: #28a745;
            margin-right: 8px;
            display: inline-block;
        }
        .attendance-absent {
            color: #dc3545;
            margin-right: 8px;
            display: inline-block;
        }
        .attendance-section {
            margin-bottom: 6px;
        }
        .attendance-section-label {
            font-weight: 600;
            font-size: 0.85rem;
            text-transform: uppercase;
            color: #666;
            margin-bottom: 4px;
        }
        .headcount-note {
            color: #6c757d;
            font-style: italic;
            font-size: 0.95rem;
        }
        .meeting-email-btn {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 0.85rem;
            color: white;
            background: #667eea;
            padding: 4px 10px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            margin-left: 8px;
            transition: background 0.2s;
        }
        .meeting-email-btn:hover {
            background: #5a67d8;
        }
    </style>
    ''')

    # Section header
    html.append('''
        <div class="section-header">
            <div>
                <div class="breadcrumb">
                    <a href="?">Dashboard</a> &rsaquo;
                    <a href="?trip={0}">{1}</a> &rsaquo; Meetings
                </div>
                <h2>Meetings</h2>
            </div>
        </div>
    '''.format(org_id, _escape_html(trip.OrganizationName)))

    # Get team emails for meeting reminders
    team_emails_sql = '''
    SELECT p.EmailAddress
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
      AND p.EmailAddress IS NOT NULL
      AND p.EmailAddress != ''
    '''.format(org_id)
    team_emails_result = list(q.QuerySql(team_emails_sql))
    team_emails = [r.EmailAddress for r in team_emails_result]
    team_emails_js = ','.join(team_emails) if team_emails else ''

    # Get all org members for attendance comparison
    org_members_sql = '''
    SELECT om.PeopleId, p.Name2 as Name
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
    ORDER BY p.Name2
    '''.format(org_id)
    org_members = list(q.QuerySql(org_members_sql))
    org_member_ids = set([m.PeopleId for m in org_members])
    org_member_names = dict([(m.PeopleId, m.Name) for m in org_members])

    # Action bar with Create New Meeting button - admin only
    if is_admin:
        escaped_trip_name = _escape_html(trip.OrganizationName or '').replace("'", "\\'")
        html.append('''
            <div class="section-actions">
                <button onclick="MissionsMeeting.openCreate('{0}', '{1}')" class="section-action-btn section-action-btn-primary">
                    &#128197; Create New Meeting
                </button>
                <a href="/Org/{0}#tab-Meetings-tab" target="_blank" class="section-action-btn section-action-btn-secondary">
                    &#9881; Manage All Meetings
                </a>
            </div>
        '''.format(org_id, escaped_trip_name))

    # Get meetings - use COALESCE to prefer extra values over column values
    meetings_sql = '''
    SELECT
        m.MeetingId,
        m.MeetingDate,
        COALESCE(me_desc.Data, m.Description) as Description,
        COALESCE(me_loc.Data, m.Location) as Location,
        m.NumPresent,
        m.HeadCount,
        CASE WHEN m.MeetingDate >= CAST(GETDATE() AS DATE) THEN 'upcoming' ELSE 'past' END as Status
    FROM Meetings m WITH (NOLOCK)
    LEFT JOIN MeetingExtra me_desc WITH (NOLOCK) ON m.MeetingId = me_desc.MeetingId AND me_desc.Field = 'Description'
    LEFT JOIN MeetingExtra me_loc WITH (NOLOCK) ON m.MeetingId = me_loc.MeetingId AND me_loc.Field = 'Location'
    WHERE m.OrganizationId = {0}
      AND m.MeetingDate IS NOT NULL
    ORDER BY m.MeetingDate DESC
    '''.format(org_id)

    meetings = list(q.QuerySql(meetings_sql))

    # Separate upcoming and past meetings
    upcoming = [m for m in meetings if m.Status == 'upcoming']
    past = [m for m in meetings if m.Status == 'past']

    # Upcoming Meetings
    html.append('<div class="card mb-4">')
    html.append('<div class="card-header"><h4>&#9650; Upcoming Meetings</h4></div>')
    html.append('<div class="card-body">')

    if upcoming:
        for meeting in sorted(upcoming, key=lambda x: x.MeetingDate):
            meeting_id = meeting.MeetingId
            mtg_date = _convert_to_python_date(meeting.MeetingDate)
            meeting_day = ''
            meeting_date_str = ''
            meeting_time_str = ''
            meeting_date_display = ''
            meeting_time_display = ''
            meeting_date_input = ''  # YYYY-MM-DD format for input field
            meeting_time_input = ''  # HH:MM format for input field
            if mtg_date and hasattr(mtg_date, 'strftime'):
                meeting_day = mtg_date.strftime('%A')  # Full day name (Monday, Tuesday, etc.)
                meeting_date_str = mtg_date.strftime('%B %d, %Y')  # January 14, 2026
                meeting_time_str = mtg_date.strftime('%I:%M %p')  # 06:00 PM
                meeting_date_display = mtg_date.strftime('%A, %B %d, %Y')
                meeting_time_display = mtg_date.strftime('%I:%M %p')
                meeting_date_input = mtg_date.strftime('%Y-%m-%d')  # 2026-01-14
                meeting_time_input = mtg_date.strftime('%H:%M')  # 18:00

            description = meeting.Description if meeting.Description else 'Team Meeting'
            location = meeting.Location if meeting.Location else ''

            # Escape values for JavaScript
            trip_name_escaped = _escape_html(trip.OrganizationName).replace("'", "\\'").replace('"', '\\"')
            desc_escaped = _escape_html(description).replace("'", "\\'").replace('"', '\\"')
            loc_escaped = _escape_html(location).replace("'", "\\'").replace('"', '\\"') if location else ''

            html.append('<div class="meeting-card" style="border-left: 4px solid #28a745;">')
            html.append('<div class="d-flex justify-content-between align-items-start">')
            html.append('<div>')
            html.append('<strong>{0}</strong> &bull; {1} @ {2}'.format(meeting_day, meeting_date_str, meeting_time_str))
            html.append('<p class="mb-1"><small>{0}</small></p>'.format(_escape_html(description)))
            if location:
                html.append('<small style="color: #666;">&#128205; {0}</small>'.format(_escape_html(location)))
            html.append('</div>')
            html.append('<div style="display: flex; align-items: center; gap: 8px;">')
            html.append('<span class="status-badge status-medium">Upcoming</span>')
            # Edit button
            html.append('<button class="meeting-email-btn" onclick="MissionsMeeting.openEdit({0}, {1}, \'{2}\', \'{3}\', \'{4}\', \'{5}\', \'{6}\')" title="Edit meeting" style="background: #ff9800;">&#9998; Edit</button>'.format(
                meeting_id,
                org_id,
                trip_name_escaped,
                meeting_date_input,
                meeting_time_input,
                desc_escaped,
                loc_escaped
            ))
            # Email button
            html.append('<button class="meeting-email-btn" onclick="MissionsEmail.openMeetingReminder(\'{0}\', \'{1}\', \'{2}\', \'{3}\', \'{4}\', \'{5}\', \'{6}\')" title="Send meeting reminder email">&#9993; Email Team</button>'.format(
                org_id,
                team_emails_js.replace("'", "\\'"),
                trip_name_escaped,
                meeting_date_display,
                meeting_time_display,
                desc_escaped,
                loc_escaped
            ))
            html.append('</div>')
            html.append('</div>')
            html.append('</div>')
    else:
        html.append('<p class="text-muted">No upcoming meetings scheduled.</p>')

    html.append('</div>')  # card-body
    html.append('</div>')  # card

    # Past Meetings
    html.append('<div class="card">')
    html.append('<div class="card-header"><h4>&#9660; Past Meetings</h4></div>')
    html.append('<div class="card-body">')

    if past:
        for meeting in past:
            meeting_id = meeting.MeetingId
            mtg_date = _convert_to_python_date(meeting.MeetingDate)
            meeting_day = ''
            meeting_date_str = ''
            meeting_time_str = ''
            meeting_date_display = ''
            meeting_time_display = ''
            meeting_date_input = ''  # YYYY-MM-DD format for input field
            meeting_time_input = ''  # HH:MM format for input field
            if mtg_date and hasattr(mtg_date, 'strftime'):
                meeting_day = mtg_date.strftime('%A')  # Full day name (Monday, Tuesday, etc.)
                meeting_date_str = mtg_date.strftime('%B %d, %Y')  # January 14, 2026
                meeting_time_str = mtg_date.strftime('%I:%M %p')  # 06:00 PM
                meeting_date_display = mtg_date.strftime('%A, %B %d, %Y')
                meeting_time_display = mtg_date.strftime('%I:%M %p')
                meeting_date_input = mtg_date.strftime('%Y-%m-%d')  # 2026-01-14
                meeting_time_input = mtg_date.strftime('%H:%M')  # 18:00

            description = meeting.Description if meeting.Description else 'Team Meeting'
            location = meeting.Location if meeting.Location else ''
            num_present = meeting.NumPresent or 0
            head_count = meeting.HeadCount or 0
            is_headcount_only = head_count > 0 and num_present == 0

            # Escape values for JavaScript
            trip_name_escaped = _escape_html(trip.OrganizationName).replace("'", "\\'").replace('"', '\\"')
            desc_escaped = _escape_html(description).replace("'", "\\'").replace('"', '\\"')
            loc_escaped = _escape_html(location).replace("'", "\\'").replace('"', '\\"') if location else ''

            html.append('<div class="meeting-card" style="border-left: 4px solid #6c757d; opacity: 0.9;">')
            html.append('<div class="d-flex justify-content-between align-items-start">')
            html.append('<div style="flex: 1;">')
            html.append('<strong style="color: #6c757d;">{0}</strong> &bull; {1} @ {2}'.format(meeting_day, meeting_date_str, meeting_time_str))
            html.append('<p class="mb-1"><small>{0}</small></p>'.format(_escape_html(description)))
            if location:
                html.append('<small style="color: #999;">&#128205; {0}</small>'.format(_escape_html(location)))
            html.append('</div>')

            # Edit, Email button and Attendance toggle
            html.append('<div style="display: flex; align-items: center; gap: 8px;">')
            # Edit button
            html.append('<button class="meeting-email-btn" onclick="MissionsMeeting.openEdit({0}, {1}, \'{2}\', \'{3}\', \'{4}\', \'{5}\', \'{6}\')" title="Edit meeting" style="background: #ff9800;">&#9998; Edit</button>'.format(
                meeting_id,
                org_id,
                trip_name_escaped,
                meeting_date_input,
                meeting_time_input,
                desc_escaped,
                loc_escaped
            ))
            # Email button for follow-up
            html.append('<button class="meeting-email-btn" onclick="MissionsEmail.openMeetingReminder(\'{0}\', \'{1}\', \'{2}\', \'{3}\', \'{4}\', \'{5}\', \'{6}\')" title="Email team about this meeting" style="background: #6c757d;">&#9993; Email</button>'.format(
                org_id,
                team_emails_js.replace("'", "\\'"),
                trip_name_escaped,
                meeting_date_display,
                meeting_time_display,
                desc_escaped,
                loc_escaped
            ))
            # Attendance toggle button
            if is_headcount_only:
                html.append('<span class="meeting-attendance-toggle" style="background: rgba(108, 117, 125, 0.1); color: #6c757d; cursor: default;">')
                html.append('&#128101; {0} (headcount)'.format(head_count))
                html.append('</span>')
            else:
                attendance_count = num_present if num_present > 0 else 0
                html.append('<span class="meeting-attendance-toggle" onclick="toggleAttendance({0})"'.format(meeting_id))
                html.append(' data-meeting="{0}" style="cursor: pointer;">'.format(meeting_id))
                html.append('&#128101; {0}/{1}'.format(attendance_count, len(org_members)))
                html.append(' <span class="toggle-arrow" id="arrow-{0}">&#9660;</span>'.format(meeting_id))
                html.append('</span>')
            html.append('</div>')

            html.append('</div>')

            # Expandable attendance list (for non-headcount meetings)
            if not is_headcount_only:
                # Get attendance for this meeting
                attendance_sql = '''
                SELECT a.PeopleId, a.AttendanceFlag, p.Name2 as Name
                FROM Attend a WITH (NOLOCK)
                JOIN People p WITH (NOLOCK) ON a.PeopleId = p.PeopleId
                WHERE a.MeetingId = {0}
                '''.format(meeting_id)
                attendance_records = list(q.QuerySql(attendance_sql))

                present_ids = set([a.PeopleId for a in attendance_records if a.AttendanceFlag == 1 or a.AttendanceFlag == True])
                present_names = [org_member_names.get(pid, 'Unknown') for pid in present_ids if pid in org_member_ids]
                absent_names = [org_member_names.get(pid, 'Unknown') for pid in org_member_ids if pid not in present_ids]

                present_names.sort()
                absent_names.sort()

                html.append('<div class="meeting-attendance-list" id="attendance-{0}">'.format(meeting_id))

                if present_names:
                    html.append('<div class="attendance-section">')
                    html.append('<div class="attendance-section-label">&#10004; Present ({0})</div>'.format(len(present_names)))
                    html.append('<div>')
                    for name in present_names:
                        html.append('<span class="attendance-present">{0}</span>'.format(_escape_html(name)))
                    html.append('</div>')
                    html.append('</div>')

                if absent_names:
                    html.append('<div class="attendance-section">')
                    html.append('<div class="attendance-section-label">&#10008; Absent ({0})</div>'.format(len(absent_names)))
                    html.append('<div>')
                    for name in absent_names:
                        html.append('<span class="attendance-absent">{0}</span>'.format(_escape_html(name)))
                    html.append('</div>')
                    html.append('</div>')

                if not present_names and not absent_names:
                    html.append('<div class="headcount-note">No individual attendance recorded</div>')

                html.append('</div>')

            html.append('</div>')  # meeting-card
    else:
        html.append('<p class="text-muted">No past meetings recorded.</p>')

    html.append('</div>')  # card-body
    html.append('</div>')  # card

    return ''.join(html)


def render_trip_budget(org_id, user_role):
    """
    Render budget and fundraising for a trip.
    Leaders see per-member totals only (no giver names).
    Admins see full financial details.
    Financial adjustments require Finance/FinanceAdmin/ManageTransactions role.
    """
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    is_admin = user_role.get('is_admin', False)
    can_manage_finance = user_role.get('can_manage_finance', False)

    html = []

    # Section header
    html.append('''
        <div class="section-header">
            <div>
                <div class="breadcrumb">
                    <a href="?">Dashboard</a> &rsaquo;
                    <a href="?trip={0}">{1}</a> &rsaquo; Budget
                </div>
                <h2>Budget & Fundraising</h2>
            </div>
        </div>
    '''.format(org_id, _escape_html(trip.OrganizationName)))

    # Get team emails for reminder
    team_emails_sql = '''
    SELECT p.EmailAddress
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
      AND p.EmailAddress IS NOT NULL
      AND p.EmailAddress != ''
    '''.format(org_id)
    team_emails_result = list(q.QuerySql(team_emails_sql))
    team_emails = [r.EmailAddress for r in team_emails_result]
    team_emails_js = ','.join(team_emails) if team_emails else ''

    # Quick action bar for budget
    # Note: Goal amounts are per-person (set during registration), so editing requires
    # per-member transaction adjustments via TouchPoint's org member transactions page

    html.append('''
        <div class="section-actions">
            <button onclick="MissionsEmail.openGoalReminder('{0}', '{1}', '{2}')" class="section-action-btn section-action-btn-primary">
                &#128276; Send Goal Reminder to All
            </button>
        </div>
    '''.format(org_id, team_emails_js.replace("'", "\\'"), _escape_html(trip.OrganizationName).replace("'", "\\'")))

    # Get member payment data - include ALL organization members (even those not charged yet)
    # Uses MissionTripTotals CTE for payment data, but LEFT JOINs to OrganizationMembers
    # to ensure we show everyone on the team
    mission_trip_cte = get_mission_trip_totals_cte(include_closed=True)

    budget_sql = mission_trip_cte + '''
    SELECT
        om.PeopleId,
        p.Name2 as Name,
        p.EmailAddress,
        ISNULL(mtt.TripCost, 0) as TripCost,
        ISNULL(mtt.Raised, 0) as TotalPaid,
        ISNULL(mtt.Due, 0) as Outstanding,
        CASE WHEN mtt.PeopleId IS NULL THEN 1 ELSE 0 END as NotCharged
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    LEFT JOIN MissionTripTotals mtt ON mtt.InvolvementId = om.OrganizationId
        AND mtt.PeopleId = om.PeopleId
        AND mtt.SortOrder <> 'ZZZZZ'
        AND mtt.Name <> 'total'
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
    ORDER BY p.Name2
    '''.format(org_id)

    members = list(q.QuerySql(budget_sql))

    # Calculate totals
    total_cost = sum(float(m.TripCost or 0) for m in members)
    total_paid = sum(float(m.TotalPaid or 0) for m in members)
    total_outstanding = sum(float(m.Outstanding or 0) for m in members)

    # Summary KPIs
    html.append('<div class="kpi-container mb-4">')
    html.append('''
        <div class="kpi-card">
            <div class="kpi-value">{0}</div>
            <div class="kpi-label">Total Trip Cost</div>
        </div>
    '''.format(format_currency(total_cost)))

    html.append('''
        <div class="kpi-card">
            <div class="kpi-value">{0}</div>
            <div class="kpi-label">Total Paid</div>
        </div>
    '''.format(format_currency(total_paid)))

    html.append('''
        <div class="kpi-card">
            <div class="kpi-value">{0}</div>
            <div class="kpi-label">Total Outstanding</div>
        </div>
    '''.format(format_currency(total_outstanding)))

    progress_pct = (float(total_paid) / float(total_cost) * 100) if total_cost > 0 else 0.0
    html.append('''
        <div class="kpi-card">
            <div class="kpi-value">{0:.1f}%</div>
            <div class="kpi-label">Funded</div>
        </div>
    '''.format(progress_pct))

    html.append('</div>')  # kpi-container

    # Calculate no-charge members (those with $0 trip cost)
    # Get the standard fee from org settings to calculate value
    stats = _get_trip_stats(org_id)
    standard_fee = stats.get('standard_fee', 0)

    no_charge_members = [m for m in members if float(m.TripCost or 0) == 0]
    no_charge_count = len(no_charge_members)
    no_charge_value = no_charge_count * standard_fee if standard_fee > 0 else 0

    if no_charge_count > 0:
        html.append('<div style="margin-top: 16px; margin-bottom: 8px; padding-bottom: 8px; border-bottom: 1px solid #e9ecef;">')
        html.append('<h6 style="margin: 0; color: #666; font-weight: 500; font-size: 0.9rem;">No-Charge Members</h6>')
        html.append('</div>')
        html.append('<div class="kpi-container" style="margin-top: 8px;">')

        html.append('''
            <div class="kpi-card" style="background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); color: white;">
                <div class="kpi-value" style="color: white;">{0}</div>
                <div class="kpi-label" style="color: rgba(255,255,255,0.9);">No-Charge Members</div>
            </div>
        '''.format(no_charge_count))

        html.append('''
            <div class="kpi-card" style="background: linear-gradient(135deg, #fc4a1a 0%, #f7b733 100%); color: white;">
                <div class="kpi-value" style="color: white;">{0}</div>
                <div class="kpi-label" style="color: rgba(255,255,255,0.9);">No-Charge Value</div>
            </div>
        '''.format(format_currency(no_charge_value)))

        # List no-charge member names
        no_charge_names = ', '.join([m.Name for m in no_charge_members[:5]])
        if no_charge_count > 5:
            no_charge_names += ' + {0} more'.format(no_charge_count - 5)
        html.append('''
            <div class="kpi-card" style="background: #f8f9fa; flex: 2;">
                <div class="kpi-value" style="font-size: 0.95rem; color: #333;">{0}</div>
                <div class="kpi-label">Members at No Charge</div>
            </div>
        '''.format(_escape_html(no_charge_names)))

        html.append('</div>')  # kpi-container for no-charge

    # Member payment breakdown
    html.append('<div class="card">')
    html.append('<div class="card-header"><h4>Member Payment Status</h4></div>')
    html.append('<div class="card-body">')

    html.append('<table class="mission-table">')

    # Add actions column header for users with finance permissions
    actions_header = '<th class="text-center">Actions</th>' if can_manage_finance else ''
    html.append('''
        <thead>
            <tr>
                <th>Member</th>
                <th class="text-right">Trip Cost</th>
                <th class="text-right">Paid</th>
                <th class="text-right">Outstanding</th>
                <th>Progress</th>
                <th class="text-center">Remind</th>
                {0}
            </tr>
        </thead>
        <tbody>
    '''.format(actions_header))

    # For JavaScript onclick handlers, use _escape_js_string (not _escape_html)
    trip_name_js = _escape_js_string(trip.OrganizationName or '')

    for member in members:
        cost = float(member.TripCost or 0)
        paid = float(member.TotalPaid or 0)
        outstanding = float(member.Outstanding or 0)
        not_charged = getattr(member, 'NotCharged', 0) == 1
        pct = (paid / cost * 100.0) if cost > 0 else 0.0

        # Progress bar color and status display
        if not_charged:
            bar_class = 'bg-secondary'
            status_text = 'Not Charged'
        elif pct >= 100:
            bar_class = 'bg-success'
            status_text = '{0:.0f}%'.format(pct)
        elif pct >= 50:
            bar_class = 'bg-warning'
            status_text = '{0:.0f}%'.format(pct)
        else:
            bar_class = 'bg-danger'
            status_text = '{0:.0f}%'.format(pct)

        # Actions column for users with finance permissions
        actions_cell = ''
        if can_manage_finance:
            # Use _escape_js_string for JavaScript onclick handlers
            member_name_js_adj = _escape_js_string(member.Name or '')
            actions_cell = '''
                <td class="text-center">
                    <button onclick="MissionsFee.openAdjust({6}, '{7}', '{0}', {1})"
                            class="btn btn-sm" style="background: #ff9800; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer;">
                        Adjust
                    </button>
                </td>
            '''.format(
                member_name_js_adj,
                cost,
                0, 0, 0, 0,  # placeholders for format positions 2-5
                member.PeopleId,
                org_id
            )

        # Individual goal reminder button
        member_email = getattr(member, 'EmailAddress', '') or ''
        # Use _escape_js_string for JavaScript onclick handlers
        member_name_js = _escape_js_string(member.Name or '')
        email_js = _escape_js_string(member_email)
        # HTML-escaped name for title attribute (display purposes)
        member_name_html = _escape_html(member.Name or '')

        if member_email:
            # Ensure we pass valid numeric values for cost and outstanding
            # Use the actual float values, not formatted strings (which may contain "-")
            numeric_cost = float(cost) if cost else 0.0
            numeric_outstanding = float(outstanding) if outstanding else 0.0
            remind_cell = '''
                <td class="text-center">
                    <button onclick="MissionsEmail.openIndividualGoalReminder({0}, '{1}', '{2}', '{3}', {4}, {5}, {6})"
                            class="btn btn-sm" style="background: #2196F3; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer;" title="Send goal reminder to {7}">
                        &#128276;
                    </button>
                </td>
            '''.format(
                member.PeopleId,
                email_js,
                member_name_js,
                trip_name_js,
                numeric_cost,  # Use actual numeric value
                numeric_outstanding,  # Use actual numeric value
                org_id,  # Pass org_id for Advanced Options link
                member_name_html  # For title attribute (HTML context)
            )
        else:
            remind_cell = '<td class="text-center text-muted">No email</td>'

        # Format currency displays - show "Not Charged" indicator for trip cost if applicable
        cost_display = '<span style="color: #999;">Not Set</span>' if not_charged else format_currency(cost)
        # For "Not Charged" members, show "-" consistently for paid and outstanding
        paid_display = '-' if not_charged else format_currency(paid)
        outstanding_display = '<span style="color: #999;">--</span>' if not_charged else format_currency(outstanding)

        html.append('''
            <tr{8}>
                <td>{0}</td>
                <td class="text-right">{1}</td>
                <td class="text-right">{2}</td>
                <td class="text-right">{3}</td>
                <td>
                    <div class="progress-container" style="min-width: 100px;">
                        <div class="progress-bar {4}" style="width: {5:.0f}%;"></div>
                    </div>
                    <div class="progress-text">{9}</div>
                </td>
                {6}
                {7}
            </tr>
        '''.format(
            _escape_html(member.Name),
            cost_display,
            paid_display,  # Use paid_display which handles "Not Charged" members
            outstanding_display,
            bar_class,
            pct if not not_charged else 0.0,
            remind_cell,
            actions_cell,
            ' style="background-color: #fff3cd;"' if not_charged else '',
            status_text
        ))

    html.append('</tbody></table>')
    html.append('</div>')  # card-body
    html.append('</div>')  # card

    # Note for leaders about privacy
    if not is_admin:
        html.append('''
            <div class="alert alert-info mt-3">
                <strong>Note:</strong> Individual donor information is not shown.
                For detailed financial records, please contact the church office.
            </div>
        ''')

    return ''.join(html)


def render_trip_documents(org_id, user_role):
    """Render document checklist for a trip."""
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    html = []

    # Section header
    html.append('''
        <div class="section-header">
            <div>
                <div class="breadcrumb">
                    <a href="?">Dashboard</a> &rsaquo;
                    <a href="?trip={0}">{1}</a> &rsaquo; Documents
                </div>
                <h2>Documents</h2>
            </div>
        </div>
    '''.format(org_id, _escape_html(trip.OrganizationName)))

    # Get document status for members - passport info is in RecReg table, background check in Volunteer table
    docs_sql = '''
    SELECT
        p.PeopleId,
        p.Name2 as Name,
        CASE WHEN rr.passportnumber IS NOT NULL AND rr.passportnumber != '' THEN 1 ELSE 0 END as HasPassportNumber,
        rr.passportexpires as PassportExpires,
        vs.Description as BGCheckStatus
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    LEFT JOIN RecReg rr WITH (NOLOCK) ON rr.PeopleId = p.PeopleId
    LEFT JOIN Volunteer v WITH (NOLOCK) ON v.PeopleId = p.PeopleId
    LEFT JOIN lookup.VolApplicationStatus vs WITH (NOLOCK) ON vs.Id = v.StatusId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
    ORDER BY p.Name2
    '''.format(org_id)

    members = list(q.QuerySql(docs_sql))

    html.append('<div class="card">')
    html.append('<div class="card-header"><h4>Document Checklist</h4></div>')
    html.append('<div class="card-body">')

    html.append('<table class="mission-table">')
    html.append('''
        <thead>
            <tr>
                <th>Member</th>
                <th class="text-center">Passport Number</th>
                <th class="text-center">Passport Expires</th>
                <th class="text-center">Background Check</th>
            </tr>
        </thead>
        <tbody>
    ''')

    for member in members:
        # Passport number
        number_icon = '&#10003;' if member.HasPassportNumber else '&#10007;'
        number_class = 'text-success' if member.HasPassportNumber else 'text-danger'

        # Passport expiration
        exp_text = ''
        exp_class = 'text-muted'
        if member.PassportExpires:
            exp_date = member.PassportExpires
            if hasattr(exp_date, 'strftime'):
                exp_text = exp_date.strftime('%b %Y')
                # Check if expiring within 6 months
                today = datetime.date.today()
                if hasattr(exp_date, 'date'):
                    exp_date = exp_date.date()
                months_until = (exp_date.year - today.year) * 12 + (exp_date.month - today.month)
                if months_until < 6:
                    exp_class = 'text-warning'
                if exp_date < today:
                    exp_class = 'text-danger'
                    exp_text += ' (EXPIRED)'

        # Background check (from Volunteer table via VolApplicationStatus lookup)
        bg_status = member.BGCheckStatus
        if bg_status and bg_status.lower() == 'complete':
            bg_icon = '&#10003;'
            bg_class = 'text-success'
        elif bg_status:
            bg_icon = '&#10007;'
            bg_class = 'text-warning'
        else:
            bg_icon = '&#8211;'
            bg_class = 'text-muted'

        html.append('''
            <tr>
                <td><a href="/Person2/{0}" target="_blank">{1}</a></td>
                <td class="text-center {2}">{3}</td>
                <td class="text-center {4}">{5}</td>
                <td class="text-center {6}">{7}</td>
            </tr>
        '''.format(
            member.PeopleId,
            _escape_html(member.Name),
            number_class, number_icon,
            exp_class, exp_text or '&#8211;',
            bg_class, bg_icon
        ))

    html.append('</tbody></table>')
    html.append('</div>')  # card-body
    html.append('</div>')  # card

    return ''.join(html)


def render_trip_messages(org_id, user_role):
    """Render message history for a trip - shows emails received by trip members."""
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    html = []

    # Get filter parameters
    filter_person_id = None
    if hasattr(model.Data, 'filterPerson') and model.Data.filterPerson:
        try:
            filter_person_id = int(model.Data.filterPerson)
        except:
            pass

    # Get search query
    search_query = ''
    if hasattr(model.Data, 'search') and model.Data.search:
        search_query = str(model.Data.search).strip()

    # Get pagination parameters
    page_size = 25  # Default page size
    if hasattr(model.Data, 'pageSize') and model.Data.pageSize:
        try:
            page_size = int(model.Data.pageSize)
            if page_size not in [25, 50, 100]:
                page_size = 25
        except:
            pass

    current_page = 1
    if hasattr(model.Data, 'page') and model.Data.page:
        try:
            current_page = max(1, int(model.Data.page))
        except:
            pass

    # Section header
    html.append('''
        <div class="section-header">
            <div>
                <div class="breadcrumb">
                    <a href="?">Dashboard</a> &rsaquo;
                    <a href="?trip={0}">{1}</a> &rsaquo; Messages
                </div>
                <h2>Messages</h2>
            </div>
        </div>
    '''.format(org_id, _escape_html(trip.OrganizationName)))

    # Get trip members for the filter dropdown
    members_sql = '''
    SELECT DISTINCT p.PeopleId, p.Name2
    FROM OrganizationMembers om WITH (NOLOCK)
    INNER JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    WHERE om.OrganizationId = {0}
      AND om.MemberTypeId <> {1}
    ORDER BY p.Name2
    '''.format(org_id, config.MEMBER_TYPE_LEADER)

    try:
        members = list(q.QuerySql(members_sql))
    except:
        members = []

    # Filter controls
    html.append('<div class="card mb-3">')
    html.append('<div class="card-body">')

    # Row 1: Search and Person filter
    html.append('<div class="row align-items-end mb-2">')

    # Search box
    html.append('<div class="col-md-4">')
    html.append('<label><strong>Search:</strong></label>')
    html.append('<input type="text" id="searchQuery" class="form-control" placeholder="Search subject or sender..." value="{0}" onkeypress="if(event.keyCode==13) filterMessages();">'.format(_escape_html(search_query)))
    html.append('</div>')

    # Person filter dropdown
    html.append('<div class="col-md-3">')
    html.append('<label><strong>Filter by Person:</strong></label>')
    html.append('<select id="personFilter" class="form-control" onchange="filterMessages()">')
    html.append('<option value="">All Team Members</option>')
    for member in members:
        selected = ' selected' if filter_person_id and member.PeopleId == filter_person_id else ''
        html.append('<option value="{0}"{2}>{1}</option>'.format(
            member.PeopleId,
            _escape_html(member.Name2),
            selected
        ))
    html.append('</select>')
    html.append('</div>')

    # Page size selector
    html.append('<div class="col-md-2">')
    html.append('<label><strong>Show:</strong></label>')
    html.append('<select id="pageSize" class="form-control" onchange="filterMessages()">')
    for size in [25, 50, 100]:
        selected = ' selected' if page_size == size else ''
        html.append('<option value="{0}"{1}>{0} per page</option>'.format(size, selected))
    html.append('</select>')
    html.append('</div>')

    # Search button and Send new message
    html.append('''
        <div class="col-md-3 text-right">
            <button class="btn btn-secondary" onclick="filterMessages()">Search</button>
            <a href="/OrgMembers/{0}/Email" class="btn btn-primary" target="_blank">
                Send Message
            </a>
        </div>
    '''.format(org_id))

    html.append('</div>')  # row
    html.append('</div>')  # card-body
    html.append('</div>')  # card

    # Build person filter clause for SQL
    person_filter = ''
    if filter_person_id:
        person_filter = 'AND eqt.PeopleId = {0}'.format(filter_person_id)

    # Get emails sent to trip members
    # This query finds all emails received by people who are members of this trip
    messages_sql = '''
    SELECT
        eq.Id,
        eq.Subject,
        eq.Queued,
        eq.FromAddr,
        eq.FromName,
        p.PeopleId,
        p.Name2 as RecipientName,
        eqt.Sent,
        CASE WHEN eqt.Sent IS NOT NULL THEN 'Sent' ELSE 'Pending' END as Status
    FROM EmailQueueTo eqt WITH (NOLOCK)
    INNER JOIN EmailQueue eq WITH (NOLOCK) ON eq.Id = eqt.Id
    INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = eqt.PeopleId
    WHERE eqt.PeopleId IN (
        SELECT om.PeopleId
        FROM OrganizationMembers om WITH (NOLOCK)
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId <> {1}
    )
      AND eq.Queued >= DATEADD(year, -2, GETDATE())
      {2}
    ORDER BY eq.Queued DESC, p.Name2
    '''.format(org_id, config.MEMBER_TYPE_LEADER, person_filter)

    try:
        messages = list(q.QuerySql(messages_sql))
    except:
        messages = []

    # Group messages by email ID for display
    email_groups = {}
    for msg in messages:
        if msg.Id not in email_groups:
            email_groups[msg.Id] = {
                'id': msg.Id,
                'subject': msg.Subject,
                'queued': msg.Queued,
                'fromAddr': msg.FromAddr,
                'fromName': msg.FromName,
                'recipients': []
            }
        email_groups[msg.Id]['recipients'].append({
            'peopleId': msg.PeopleId,
            'name': msg.RecipientName,
            'sent': msg.Sent,
            'status': msg.Status
        })

    html.append('<div class="card">')
    html.append('<div class="card-header"><h4>Email History for Team Members</h4></div>')
    html.append('<div class="card-body">')

    if email_groups:
        html.append('<table class="mission-table" id="messagesTable">')
        html.append('''
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Subject</th>
                    <th>From</th>
                    <th>Recipients</th>
                    <th class="text-center">Actions</th>
                </tr>
            </thead>
            <tbody>
        ''')

        # Sort by date descending
        sorted_emails = sorted(email_groups.values(), key=lambda x: x['queued'] if x['queued'] else '', reverse=True)

        # Apply search filter
        if search_query:
            search_lower = search_query.lower()
            filtered_emails = []
            for email in sorted_emails:
                subject_match = email['subject'] and search_lower in email['subject'].lower()
                from_match = (email['fromName'] and search_lower in email['fromName'].lower()) or \
                             (email['fromAddr'] and search_lower in email['fromAddr'].lower())
                if subject_match or from_match:
                    filtered_emails.append(email)
            sorted_emails = filtered_emails

        # Calculate pagination
        total_emails = len(sorted_emails)
        total_pages = max(1, (total_emails + page_size - 1) // page_size)  # Ceiling division
        current_page = min(current_page, total_pages)  # Don't go past last page
        start_idx = (current_page - 1) * page_size
        end_idx = start_idx + page_size
        paginated_emails = sorted_emails[start_idx:end_idx]

        # Show result count
        if search_query:
            html.append('<p class="text-muted mb-2">Found {0} message(s) matching "{1}"</p>'.format(total_emails, _escape_html(search_query)))

        for email in paginated_emails:
            msg_date = email['queued'].strftime('%b %d, %Y') if hasattr(email['queued'], 'strftime') else str(email['queued'])

            # Build recipients display
            recipient_count = len(email['recipients'])
            if recipient_count <= 3:
                recipient_names = ', '.join([r['name'] for r in email['recipients']])
            else:
                recipient_names = ', '.join([r['name'] for r in email['recipients'][:2]]) + ' +{0} more'.format(recipient_count - 2)

            from_display = email['fromName'] or email['fromAddr'] or 'Unknown'

            html.append('''
                <tr>
                    <td>{0}</td>
                    <td>
                        <a href="#" onclick="showMessageDetail({1}, {5}); return false;" style="cursor: pointer;">
                            {2}
                        </a>
                    </td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td class="text-center">
                        <a href="/Manage/Emails/Details/{1}" target="_blank" class="btn btn-sm btn-secondary" title="View in TouchPoint">
                            View
                        </a>
                    </td>
                </tr>
            '''.format(
                msg_date,
                email['id'],
                _escape_html(email['subject'] or '(No Subject)'),
                _escape_html(from_display),
                _escape_html(recipient_names),
                org_id
            ))

        html.append('</tbody></table>')

        # Pagination controls
        if total_pages > 1:
            html.append('<div class="d-flex justify-content-between align-items-center mt-3">')

            # Showing X-Y of Z
            html.append('<div class="text-muted">Showing {0}-{1} of {2} messages</div>'.format(
                start_idx + 1,
                min(end_idx, total_emails),
                total_emails
            ))

            # Pagination buttons
            html.append('<div class="btn-group">')

            # Previous button
            if current_page > 1:
                html.append('<button class="btn btn-sm btn-secondary" onclick="goToPage({0})">&laquo; Previous</button>'.format(current_page - 1))
            else:
                html.append('<button class="btn btn-sm btn-secondary" disabled>&laquo; Previous</button>')

            # Page numbers (show up to 5 pages)
            start_page = max(1, current_page - 2)
            end_page = min(total_pages, start_page + 4)
            if end_page - start_page < 4:
                start_page = max(1, end_page - 4)

            for page_num in range(start_page, end_page + 1):
                if page_num == current_page:
                    html.append('<button class="btn btn-sm btn-primary">{0}</button>'.format(page_num))
                else:
                    html.append('<button class="btn btn-sm btn-secondary" onclick="goToPage({0})">{0}</button>'.format(page_num))

            # Next button
            if current_page < total_pages:
                html.append('<button class="btn btn-sm btn-secondary" onclick="goToPage({0})">Next &raquo;</button>'.format(current_page + 1))
            else:
                html.append('<button class="btn btn-sm btn-secondary" disabled>Next &raquo;</button>')

            html.append('</div>')  # btn-group
            html.append('</div>')  # d-flex

    else:
        if filter_person_id:
            html.append('<p class="text-muted">No messages found for this person.</p>')
        elif search_query:
            html.append('<p class="text-muted">No messages found matching "{0}".</p>'.format(_escape_html(search_query)))
        else:
            html.append('<p class="text-muted">No messages found for team members.</p>')

    html.append('</div>')  # card-body
    html.append('</div>')  # card

    # Message detail modal
    html.append('''
    <div id="messageDetailModal" class="modal" style="display: none;">
        <div class="modal-content" style="max-width: 800px;">
            <div class="modal-header">
                <h4 id="messageDetailTitle">Message Details</h4>
                <span class="close" onclick="closeMessageDetailModal()">&times;</span>
            </div>
            <div class="modal-body" id="messageDetailContent">
                <div class="text-center"><p>Loading...</p></div>
            </div>
        </div>
    </div>
    ''')

    # JavaScript for filtering and message details
    html.append('''
    <script>
    function filterMessages() {
        var personId = document.getElementById('personFilter').value;
        var searchQuery = document.getElementById('searchQuery').value;
        var pageSize = document.getElementById('pageSize').value;
        var url = '?trip=''' + str(org_id) + '''&section=messages';
        if (personId) {
            url += '&filterPerson=' + personId;
        }
        if (searchQuery) {
            url += '&search=' + encodeURIComponent(searchQuery);
        }
        if (pageSize) {
            url += '&pageSize=' + pageSize;
        }
        // Reset to page 1 when filtering
        url += '&page=1';
        window.location.href = url;
    }

    function goToPage(pageNum) {
        var personId = document.getElementById('personFilter').value;
        var searchQuery = document.getElementById('searchQuery').value;
        var pageSize = document.getElementById('pageSize').value;
        var url = '?trip=''' + str(org_id) + '''&section=messages';
        if (personId) {
            url += '&filterPerson=' + personId;
        }
        if (searchQuery) {
            url += '&search=' + encodeURIComponent(searchQuery);
        }
        if (pageSize) {
            url += '&pageSize=' + pageSize;
        }
        url += '&page=' + pageNum;
        window.location.href = url;
    }

    function showMessageDetail(emailId, orgId) {
        document.getElementById('messageDetailModal').style.display = 'block';
        document.getElementById('messageDetailContent').innerHTML = '<div class="text-center"><p>Loading message details...</p></div>';

        // Load message details via AJAX
        var xhr = new XMLHttpRequest();
        xhr.open('POST', window.location.pathname, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var response = JSON.parse(xhr.responseText);
                        if (response.success) {
                            displayMessageDetail(response.data, orgId);
                        } else {
                            document.getElementById('messageDetailContent').innerHTML = '<div class="alert alert-danger">' + (response.error || 'Failed to load message') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('messageDetailContent').innerHTML = '<div class="alert alert-danger">Error parsing response</div>';
                    }
                } else {
                    document.getElementById('messageDetailContent').innerHTML = '<div class="alert alert-danger">Failed to load message details</div>';
                }
            }
        };
        xhr.send('ajax=1&action=get_message_detail&emailId=' + emailId + '&orgId=' + orgId);
    }

    function displayMessageDetail(data, orgId) {
        var html = '<div class="mb-3">';
        html += '<p><strong>Subject:</strong> ' + escapeHtml(data.subject || '(No Subject)') + '</p>';
        html += '<p><strong>From:</strong> ' + escapeHtml(data.fromName || '') + ' &lt;' + escapeHtml(data.fromAddr || '') + '&gt;</p>';
        html += '<p><strong>Sent:</strong> ' + escapeHtml(data.queued || '') + '</p>';
        html += '</div>';

        // Recipients section
        html += '<div class="mb-3">';
        html += '<h5>Recipients on this trip:</h5>';
        html += '<table class="mission-table">';
        html += '<thead><tr><th>Name</th><th>Status</th><th>Action</th></tr></thead>';
        html += '<tbody>';

        for (var i = 0; i < data.recipients.length; i++) {
            var r = data.recipients[i];
            var statusBadge = r.sent ? '<span class="status-badge status-complete">Sent</span>' : '<span class="status-badge status-pending">Pending</span>';
            html += '<tr>';
            html += '<td><a href="/Person2/' + r.peopleId + '" target="_blank">' + escapeHtml(r.name) + '</a></td>';
            html += '<td>' + statusBadge + '</td>';
            html += '<td><button class="btn btn-sm btn-primary" onclick="resendMessage(' + data.emailId + ', ' + r.peopleId + ', \\'' + escapeHtml(r.name).replace(/'/g, "\\\\'") + '\\')">Resend</button></td>';
            html += '</tr>';
        }

        html += '</tbody></table>';
        html += '</div>';

        // Message body preview
        if (data.body) {
            html += '<div class="mb-3">';
            html += '<h5>Message Preview:</h5>';
            html += '<div style="border: 1px solid #ddd; padding: 15px; max-height: 300px; overflow-y: auto; background: #f9f9f9;">';
            html += data.body;
            html += '</div>';
            html += '</div>';
        }

        html += '<div class="mt-3">';
        html += '<a href="/Manage/Emails/Details/' + data.emailId + '" target="_blank" class="btn btn-secondary">View Full Email in TouchPoint</a>';
        html += '</div>';

        document.getElementById('messageDetailContent').innerHTML = html;
    }

    function resendMessage(emailId, peopleId, personName) {
        if (!confirm('Resend this email to ' + personName + '?')) {
            return;
        }

        var xhr = new XMLHttpRequest();
        xhr.open('POST', window.location.pathname, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var response = JSON.parse(xhr.responseText);
                        if (response.success) {
                            alert('Email queued for resend to ' + personName);
                        } else {
                            alert('Failed to resend: ' + (response.error || 'Unknown error'));
                        }
                    } catch (e) {
                        alert('Error processing response');
                    }
                } else {
                    alert('Failed to resend email');
                }
            }
        };
        xhr.send('ajax=1&action=resend_message&emailId=' + emailId + '&peopleId=' + peopleId);
    }

    function closeMessageDetailModal() {
        document.getElementById('messageDetailModal').style.display = 'none';
    }

    function escapeHtml(text) {
        if (!text) return '';
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Close modal on outside click
    window.onclick = function(event) {
        var modal = document.getElementById('messageDetailModal');
        if (event.target === modal) {
            modal.style.display = 'none';
        }
    };
    </script>
    ''')

    return ''.join(html)


def render_trip_tasks(org_id, user_role):
    """Render tasks and milestones for a trip."""
    trip = _get_trip_info(org_id)
    if not trip:
        return '<div class="alert alert-danger">Trip not found.</div>'

    html = []

    # Section header
    html.append('''
        <div class="section-header">
            <div>
                <div class="breadcrumb">
                    <a href="?">Dashboard</a> &rsaquo;
                    <a href="?trip={0}">{1}</a> &rsaquo; Tasks
                </div>
                <h2>Tasks & Milestones</h2>
            </div>
        </div>
    '''.format(org_id, _escape_html(trip.OrganizationName)))

    # Calculate milestones based on trip dates
    milestones = _calculate_trip_milestones(trip)

    html.append('<div class="card">')
    html.append('<div class="card-header"><h4>Trip Milestones</h4></div>')
    html.append('<div class="card-body">')

    if milestones:
        html.append('<div class="timeline">')

        for milestone in milestones:
            status_class = 'status-' + milestone['status']
            html.append('''
                <div class="timeline-item">
                    <div class="timeline-date">{0}</div>
                    <div class="timeline-content">
                        <h4>{1}</h4>
                        <p><span class="status-badge {2}">{3}</span></p>
                    </div>
                </div>
            '''.format(
                milestone['date'],
                _escape_html(milestone['title']),
                status_class,
                milestone['status'].title()
            ))

        html.append('</div>')  # timeline
    else:
        html.append('<p class="text-muted">No trip dates set. Set trip dates to see milestones.</p>')

    html.append('</div>')  # card-body
    html.append('</div>')  # card

    # Quick checklist
    html.append('<div class="card mt-4">')
    html.append('<div class="card-header"><h4>Pre-Trip Checklist</h4></div>')
    html.append('<div class="card-body">')

    checklist_items = [
        ('Confirm all team members registered', True),
        ('Collect passport information', True),
        ('Complete background checks', True),
        ('Schedule team meetings', True),
        ('Send fundraising letters', False),
        ('Book flights/transportation', False),
        ('Arrange lodging', False),
        ('Plan ministry activities', False),
        ('Pack supplies', False),
        ('Final team meeting', False)
    ]

    html.append('<ul class="list-group">')
    for item, is_automated in checklist_items:
        icon = '&#9744;'  # unchecked box
        if is_automated:
            icon = '&#128269;'  # magnifying glass - tracked automatically
        html.append('''
            <li class="list-group-item d-flex justify-content-between align-items-center">
                <span>{0} {1}</span>
                {2}
            </li>
        '''.format(
            icon,
            item,
            '<small class="text-muted">(Auto-tracked)</small>' if is_automated else ''
        ))
    html.append('</ul>')

    html.append('</div>')  # card-body
    html.append('</div>')  # card

    return ''.join(html)


def _calculate_trip_milestones(trip):
    """Calculate trip milestones based on trip dates."""
    milestones = []
    today = datetime.date.today()

    if not trip.TripBegin:
        return milestones

    trip_start = trip.TripBegin
    if hasattr(trip_start, 'date'):
        trip_start = trip_start.date()

    trip_end = trip.TripEnd
    if trip_end and hasattr(trip_end, 'date'):
        trip_end = trip_end.date()

    # Registration deadline (8 weeks before)
    reg_deadline = trip_start - datetime.timedelta(weeks=8)
    milestones.append({
        'date': reg_deadline.strftime('%b %d, %Y'),
        'title': 'Registration Deadline',
        'status': 'completed' if today > reg_deadline else 'pending'
    })

    # Passport deadline (6 weeks before)
    passport_deadline = trip_start - datetime.timedelta(weeks=6)
    milestones.append({
        'date': passport_deadline.strftime('%b %d, %Y'),
        'title': 'Passport Information Due',
        'status': 'completed' if today > passport_deadline else 'pending'
    })

    # 50% payment (4 weeks before)
    payment_50 = trip_start - datetime.timedelta(weeks=4)
    milestones.append({
        'date': payment_50.strftime('%b %d, %Y'),
        'title': '50% Payment Due',
        'status': 'completed' if today > payment_50 else 'pending'
    })

    # Full payment (2 weeks before)
    payment_full = trip_start - datetime.timedelta(weeks=2)
    milestones.append({
        'date': payment_full.strftime('%b %d, %Y'),
        'title': 'Full Payment Due',
        'status': 'completed' if today > payment_full else 'pending'
    })

    # Trip start
    milestones.append({
        'date': trip_start.strftime('%b %d, %Y'),
        'title': 'Trip Begins',
        'status': 'active' if today == trip_start else ('completed' if today > trip_start else 'pending')
    })

    # Trip end
    if trip_end:
        milestones.append({
            'date': trip_end.strftime('%b %d, %Y'),
            'title': 'Trip Ends',
            'status': 'completed' if today > trip_end else 'pending'
        })

    return milestones


# ::END:: Trip Section Views

#####################################################################
# STYLE DEFINITIONS
#####################################################################

# ::START:: Styles
def get_modern_styles():
    """Return modernized CSS styles with improved mobile support"""
    return '''
    <style>
    /* Modern CSS Variables for theming */
    :root {
        --primary-color: #003366;
        --secondary-color: #4A90E2;
        --success-color: #4CAF50;
        --warning-color: #ff9800;
        --danger-color: #f44336;
        --light-bg: #f8f9fa;
        --border-color: #dee2e6;
        --text-muted: #6c757d;
        --shadow: 0 2px 4px rgba(0,0,0,0.1);
        --radius: 8px;
    }
    
    /* Compact mobile-first design */
    * {
        box-sizing: border-box;
    }
    
    /* Navigation */
    .mission-nav {
        display: flex;
        gap: 7px;
        padding: 10px;
        background: var(--light-bg);
        border-radius: var(--radius);
        margin-bottom: 15px;
        flex-wrap: wrap;
    }
    
    .mission-nav a {
        padding: 8px 15px;
        color: var(--text-muted);
        text-decoration: none;
        border-radius: var(--radius);
        transition: all 0.3s ease;
        font-weight: 500;
        font-size: 14px;
    }
    
    .mission-nav a:hover {
        background: white;
        color: var(--secondary-color);
        box-shadow: var(--shadow);
    }
    
    .mission-nav a.active {
        background: var(--secondary-color);
        color: white;
    }
    
    /* Compact KPI Cards */
    .kpi-container {
        display: flex;
        gap: 10px;
        margin-bottom: 20px;
        flex-wrap: wrap;
    }
    
    .kpi-card {
        background: var(--light-bg);
        padding: 15px;
        border-radius: var(--radius);
        text-align: center;
        flex: 1;
        min-width: 100px;
    }
    
    .kpi-card .value {
        font-size: 1.5em;
        font-weight: bold;
        color: var(--primary-color);
    }
    
    .kpi-card .label {
        color: var(--text-muted);
        font-size: 0.85em;
        margin-top: 2px;
    }
    
    /* Urgent deadlines section */
    .urgent-deadlines {
        background: #fff3e0;
        border-left: 4px solid #ff9800;
        padding: 15px;
        margin-bottom: 20px;
        border-radius: var(--radius);
    }
    
    /* Main content layout */
    .dashboard-container {
        display: flex;
        gap: 30px;
        align-items: flex-start;
        margin: 0 auto;
        max-width: 1400px;
    }
    
    .main-content {
        flex: 1;
        min-width: 300px;
    }
    
    .sidebar {
        width: 320px;
        min-width: 280px;
        flex-shrink: 0;
        padding: 0 10px;
        position: sticky;
        top: 20px;
        max-height: 90vh;
        overflow-y: auto;
        min-height: 200px;
        z-index: 10;
    }
    
    /* Timeline styles for upcoming meetings */
    .timeline {
        border-left: 4px solid #003366;
        padding-left: 20px;
        margin-left: 10px;
    }
    
    .timeline-item {
        margin-bottom: 20px;
        position: relative;
    }
    
    .timeline-item::before {
        content: '';
        position: absolute;
        left: -24px;
        top: 0;
        width: 12px;
        height: 12px;
        border-radius: 50%;
        background: #003366;
        border: 2px solid white;
    }
    
    .timeline-date {
        font-weight: bold;
        color: #003366;
        font-size: 13px;
    }
    
    .timeline-date a {
        color: #003366;
        text-decoration: none;
    }
    
    .timeline-date a:hover {
        text-decoration: underline;
    }
    
    .timeline-content {
        background: #f8f9fa;
        padding: 10px;
        border-radius: 5px;
        margin-top: 5px;
    }
    
    .timeline-content h4 {
        margin: 0;
        padding: 0;
        font-size: 16px;
        color: #333;
    }
    
    .timeline-content p {
        margin: 5px 0;
        font-size: 13px;
        color: #666;
    }
    
    /* Upcoming meetings cards */
    .meeting-card {
        background: white;
        padding: 10px;
        margin-bottom: 8px;
        border-radius: 5px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        border-left: 3px solid var(--secondary-color);
    }
    
    .meeting-card:hover {
        box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }
    
    /* Compact table styles */
    .mission-table {
        width: 100%;
        background: white;
        border-radius: var(--radius);
        overflow: hidden;
        box-shadow: var(--shadow);
        font-size: 14px;
        table-layout: auto;
        margin-bottom: 20px;
    }
    
    .mission-table th {
        background: var(--primary-color);
        color: white;
        padding: 8px;
        text-align: left;
        font-weight: 600;
        font-size: 13px;
        position: relative;
        vertical-align: middle;
    }
    
    .mission-table th.text-center {
        text-align: center;
    }
    
    .mission-table td {
        padding: 6px 8px;
        border-bottom: 1px solid var(--border-color);
        vertical-align: middle;
    }
    
    .mission-table td.text-center {
        text-align: center;
    }
    
    .mission-table tr:hover {
        background: var(--light-bg);
    }

    /* Family grouping bands - subtle alternating backgrounds */
    .mission-table tr.family-band-0 {
        background-color: rgba(0, 123, 255, 0.06);
    }
    .mission-table tr.family-band-1 {
        background-color: transparent;
    }
    .mission-table tr.family-band-0:hover,
    .mission-table tr.family-band-1:hover {
        background: var(--light-bg);
    }

    /* Compact mission row */
    .mission-row {
        font-size: 14px;
    }
    
    .mission-row .title {
        font-weight: bold;
        color: var(--primary-color);
    }
    
    .mission-row .dates {
        font-size: 12px;
        color: var(--text-muted);
        font-style: italic;
    }
    
    /* Icon styles with tooltips */
    .icon-header {
        font-size: 24px;
        text-align: center;
        cursor: pointer;
        position: relative;
        display: inline-block;
        padding: 5px;
        line-height: 1;
        vertical-align: middle;
    }
    
    .icon-header:hover {
        background: rgba(0,0,0,0.05);
        border-radius: 4px;
    }
    
    /* Tooltip styles for icon headers */
    .icon-tooltip {
        display: none;
        position: fixed;
        background-color: #333;
        color: white;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 12px;
        white-space: nowrap;
        z-index: 1000;
        box-shadow: 0 2px 5px rgba(0,0,0,0.3);
    }
    
    .icon-tooltip.show {
        display: block;
    }
    
    .text-center {
        text-align: center;
        vertical-align: middle;
    }
    
    /* Status badges */
    .status-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 0.75em;
        font-weight: 500;
    }
    
    .status-urgent { background: #ffebee; color: #c62828; }
    .status-high { background: #fff3e0; color: #f57c00; }
    .status-medium { background: #e3f2fd; color: #1565c0; }
    .status-normal { background: #e8f5e9; color: #2e7d32; }
    .status-pending { background: #f3e5f5; color: #6a1b9a; }
    .status-active { background: #e8f5e9; color: #2e7d32; }
    .status-closed { background: #ffebee; color: #c62828; }
    
    /* Progress bars - more compact */
    .progress-container {
        width: 100%;
        max-width: 200px;
        height: 6px;
        background: #e0e0e0;
        border-radius: 3px;
        overflow: hidden;
        margin: 5px 0;
    }
    
    .progress-bar {
        height: 100%;
        background: var(--success-color);
        transition: width 0.3s ease;
    }
    
    .progress-text {
        font-size: 11px;
        color: var(--text-muted);
    }
    
    /* Leader highlighting */
    .no-leader-warning {
        background-color: yellow;
        padding: 2px 6px;
        font-style: italic;
        font-size: 12px;
    }
    
    /* Popup styles */
    .popup-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.5);
        z-index: 1000;
    }
    
    .popup-content {
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        padding: 20px;
        border-radius: var(--radius);
        max-width: 500px;
        max-height: 70vh;
        overflow-y: auto;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }
    
    .popup-close {
        float: right;
        font-size: 24px;
        cursor: pointer;
        color: var(--danger-color);
    }
    
    /* Performance indicator */
    .performance-indicator {
        position: fixed;
        top: 10px;
        right: 10px;
        background: rgba(0,0,0,0.7);
        color: white;
        padding: 5px 10px;
        border-radius: 4px;
        font-size: 11px;
        display: none;
    }
    
    /* Mobile optimizations - very compact */
    @media (max-width: 1024px) {
        .sidebar {
            width: 280px;
            min-width: 250px;
        }
    }
    
    @media (max-width: 768px) {
        .dashboard-container {
            flex-direction: column;
        }
        
        .sidebar {
            width: 100%;
            position: static;
            max-height: none;
        }
        
        .mission-table {
            font-size: 12px;
        }
        
        .mission-table th {
            display: none;
        }
        
        .mission-table td {
            padding: 6px 12px;
            display: inline-block;
            text-align: left;
            border-bottom: none;
            vertical-align: middle;
        }

        .mission-table td[data-label]:before {
            content: attr(data-label) ": ";
            font-weight: 600;
            color: #555;
            font-size: 11px;
        }

        /* Name cell - full width, prominent header */
        .mission-table td[data-label="Name"] {
            display: block;
            width: 100%;
            background: #f8f9fa;
            padding: 10px 12px;
            border-bottom: 2px solid var(--primary-color);
        }
        .mission-table td[data-label="Name"]:before {
            display: none;
        }

        /* Role and Age on same row */
        .mission-table td[data-label="Role"],
        .mission-table td[data-label="Age"] {
            width: 48%;
            box-sizing: border-box;
        }

        /* Contact - full width */
        .mission-table td[data-label="Contact"] {
            display: block;
            width: 100%;
            padding: 8px 12px;
            border-top: 1px solid #eee;
        }

        /* Status indicators in a row: Photo, Passport, Background */
        .mission-table td[data-label="Photo"],
        .mission-table td[data-label="Passport"],
        .mission-table td[data-label="Background"] {
            width: 32%;
            text-align: center;
            box-sizing: border-box;
            padding: 8px 4px;
        }
        .mission-table td[data-label="Photo"]:before,
        .mission-table td[data-label="Passport"]:before,
        .mission-table td[data-label="Background"]:before {
            display: block;
            text-align: center;
            min-width: auto;
        }

        /* Actions cell - full width */
        .mission-table td[data-label="Actions"] {
            display: block;
            width: 100%;
            text-align: center;
            padding: 10px 12px;
            background: #f8f9fa;
            border-top: 1px solid #eee;
        }
        .mission-table td[data-label="Actions"]:before {
            display: none;
        }
        .mission-table td[data-label="Actions"] .actions-dropdown-btn {
            width: 100%;
            padding: 10px 16px;
        }

        .mission-table tr {
            border: 1px solid var(--border-color);
            margin-bottom: 12px;
            border-radius: 8px;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            display: block;
        }
        
        .kpi-card {
            padding: 10px;
        }
        
        .kpi-card .value {
            font-size: 1.2em;
        }
        
        .icon-header {
            font-size: 16px;
        }
        
        .progress-container {
            max-width: 150px;
        }
        
        /* Make popup full screen on mobile */
        .popup-content {
            width: 90%;
            height: 90%;
            max-width: none;
            max-height: none;
        }
    }
    
    /* Status row separator */
    .status-separator td {
        background: var(--light-bg);
        font-weight: bold;
        text-align: center;
        padding: 5px;
        font-size: 13px;
    }
    
    /* Loading states */
    .loading-spinner {
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid var(--light-bg);
        border-top-color: var(--primary-color);
        border-radius: 50%;
        animation: spin 1s linear infinite;
    }
    
    @keyframes spin {
        to { transform: rotate(360deg); }
    }
    
    /* Clickable badges */
    .clickable-badge {
        cursor: pointer;
        text-decoration: underline;
        color: var(--secondary-color);
    }
    
    .clickable-badge:hover {
        color: var(--primary-color);
    }
    
    /* Organization separator */
    .org-separator td {
        background: var(--light-bg);
        font-weight: bold;
        padding: 10px 8px;
    }
    
    /* Member card styles */
    .member-card {
        background: white;
        padding: 15px;
        margin-bottom: 15px;
        border-radius: var(--radius);
        box-shadow: var(--shadow);
    }
    
    .badge {
        display: inline-block;
        padding: 4px 8px;
        border-radius: 4px;
        font-size: 12px;
    }
    
    /* Stat card styles */
    .stat-card {
        background: white;
        padding: 20px;
        border-radius: var(--radius);
        box-shadow: var(--shadow);
    }
    
    .stat-table {
        width: 100%;
    }
    
    .stat-table td {
        padding: 5px 0;
    }
    
    /* Alert styles */
    .alert {
        padding: 15px;
        border-radius: var(--radius);
        margin-bottom: 20px;
    }
    
    .alert-danger {
        background: #ffebee;
        color: #c62828;
        border-left: 4px solid var(--danger-color);
    }
    
    /* Modal styles */
    .modal {
        display: none;
        position: fixed;
        z-index: 1000;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0,0,0,0.4);
    }
    
    .modal-content {
        background-color: #fefefe;
        margin: 15% auto;
        padding: 20px;
        border: 1px solid #888;
        width: 80%;
        max-width: 600px;
        border-radius: var(--radius);
    }
    
    .close {
        color: #aaa;
        float: right;
        font-size: 28px;
        font-weight: bold;
        cursor: pointer;
    }
    
    .close:hover,
    .close:focus {
        color: black;
    }
    
    /* Chart containers for stats */
    .chart-container {
        height: 300px;
        margin: 20px 0;
    }
    
    /* Timeline styles for trends */
    .timeline-chart {
        background: white;
        padding: 20px;
        border-radius: var(--radius);
        box-shadow: var(--shadow);
        margin-bottom: 20px;
    }

    /* ================================================================
       SIDEBAR NAVIGATION STYLES - ManagedMissions-like navigation
       ================================================================ */

    /* Sidebar CSS Variables */
    :root {
        --sidebar-width: 280px;
        --sidebar-collapsed-width: 60px;
        --sidebar-bg: #1a1a2e;
        --sidebar-text: #e8e8e8;
        --sidebar-hover: #16213e;
        --sidebar-active: #0f3460;
        --sidebar-border: rgba(255,255,255,0.1);
        --sidebar-accent: #4A90E2;
        --transition-speed: 0.3s;
        --tp-header-height: 100px;  /* TouchPoint page header height - includes margin */
    }

    /* App container - sidebar + main content */
    .app-container {
        display: flex;
        width: 100%;
        position: relative;
        margin-top: 0;
    }

    /* Navigation Sidebar - contained within .box-content, no independent scroll */
    .nav-sidebar {
        width: var(--sidebar-width);
        background: var(--sidebar-bg);
        color: var(--sidebar-text);
        flex-shrink: 0;
        overflow: visible;  /* No independent scrolling - flows with page */
        transition: width var(--transition-speed) ease;
        z-index: 100;
        box-shadow: 2px 0 10px rgba(0,0,0,0.2);
    }

    .nav-sidebar.collapsed {
        width: var(--sidebar-collapsed-width);
    }

    .nav-sidebar.collapsed .sidebar-text,
    .nav-sidebar.collapsed .trip-sections,
    .nav-sidebar.collapsed .sidebar-header h3,
    .nav-sidebar.collapsed .trips-header h4 {
        display: none;
    }

    .nav-sidebar.collapsed .trip-header {
        justify-content: center;
        padding: 12px 5px;
    }

    .nav-sidebar.collapsed .expand-icon {
        display: none;
    }

    /* Sidebar Header */
    .sidebar-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 15px 20px;
        background: rgba(0,0,0,0.2);
        border-bottom: 1px solid var(--sidebar-border);
    }

    .sidebar-header h3 {
        margin: 0;
        font-size: 18px;
        font-weight: 600;
        color: white;
    }

    .sidebar-toggle {
        background: transparent;
        border: none;
        color: var(--sidebar-text);
        cursor: pointer;
        padding: 5px 10px;
        border-radius: 4px;
        transition: background var(--transition-speed);
        font-size: 16px;
    }

    .sidebar-toggle:hover {
        background: var(--sidebar-hover);
    }

    /* Admin Navigation Links */
    .admin-nav {
        padding: 10px 0;
        border-bottom: 1px solid var(--sidebar-border);
    }

    .admin-nav .nav-link {
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 12px 20px;
        color: var(--sidebar-text);
        text-decoration: none;
        transition: background var(--transition-speed), border-left var(--transition-speed);
        border-left: 3px solid transparent;
    }

    .admin-nav .nav-link:hover {
        background: var(--sidebar-hover);
    }

    .admin-nav .nav-link.active {
        background: var(--sidebar-active);
        border-left-color: var(--sidebar-accent);
    }

    .admin-nav .nav-link .icon {
        font-size: 18px;
        width: 24px;
        text-align: center;
    }

    /* Trips Header */
    .trips-header {
        padding: 15px 20px 10px;
    }

    .trips-header h4 {
        margin: 0;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: rgba(255,255,255,0.5);
    }

    /* Trip List */
    .trip-list {
        list-style: none;
        padding: 0;
        margin: 0;
    }

    .trip-item {
        border-bottom: 1px solid var(--sidebar-border);
    }

    .trip-header {
        display: flex;
        align-items: center;
        padding: 12px 15px;
        cursor: pointer;
        transition: background var(--transition-speed);
        gap: 10px;
        position: relative;
    }

    .trip-header:hover {
        background: var(--sidebar-hover);
    }

    .trip-header.active {
        background: var(--sidebar-active);
        border-left: 3px solid var(--sidebar-accent);
    }

    .expand-icon {
        font-size: 10px;
        transition: transform var(--transition-speed);
        color: rgba(255,255,255,0.5);
        flex-shrink: 0;
    }

    .trip-item.expanded .expand-icon {
        transform: rotate(90deg);
    }

    .trip-link {
        color: var(--sidebar-text);
        text-decoration: none;
        font-size: 14px;
        flex: 1;
        word-wrap: break-word;
        overflow-wrap: break-word;
        line-height: 1.3;
    }

    .trip-link:hover {
        color: white;
    }

    .trip-info {
        display: flex;
        flex-direction: column;
        flex: 1;
        min-width: 0;
        overflow: hidden;
    }

    .trip-dates {
        font-size: 11px;
        color: rgba(255, 255, 255, 0.5);
        margin-top: 2px;
        font-weight: normal;
    }

    .nav-sidebar.collapsed .trip-dates {
        display: none;
    }

    /* Tooltip for full trip name on hover */
    .trip-header[title]:hover::after {
        content: attr(title);
        position: absolute;
        left: 100%;
        top: 50%;
        transform: translateY(-50%);
        background: #333;
        color: white;
        padding: 8px 12px;
        border-radius: 4px;
        white-space: nowrap;
        z-index: 1000;
        font-size: 13px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        margin-left: 5px;
    }

    .trip-member-badge {
        font-size: 11px;
        padding: 2px 8px;
        border-radius: 10px;
        flex-shrink: 0;
        background: rgba(255, 255, 255, 0.15);
        color: rgba(255, 255, 255, 0.9);
        white-space: nowrap;
    }

    .nav-sidebar.collapsed .trip-member-badge {
        display: none;
    }

    /* Expandable Trip Sections */
    .trip-sections {
        display: none;
        background: rgba(0,0,0,0.3);
    }

    .trip-item.expanded .trip-sections {
        display: block;
    }

    .section-link {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 15px 10px 35px;
        color: rgba(255,255,255,0.7);
        text-decoration: none;
        font-size: 13px;
        transition: background var(--transition-speed), color var(--transition-speed);
        border-left: 3px solid transparent;
    }

    .section-link:hover {
        background: var(--sidebar-hover);
        color: white;
    }

    .section-link.active {
        background: var(--sidebar-active);
        color: white;
        border-left-color: var(--sidebar-accent);
    }

    .section-link .section-icon {
        font-size: 14px;
        width: 20px;
        text-align: center;
    }

    /* Mission Trips Section (collapsible) */
    .trips-section {
        margin-top: 0;
    }

    .trips-header {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px 15px;
        color: rgba(255,255,255,0.6);
        cursor: pointer;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        transition: background 0.2s;
        border-bottom: 1px solid rgba(255,255,255,0.1);
    }

    .trips-header:hover {
        background: rgba(255,255,255,0.05);
        color: rgba(255,255,255,0.8);
    }

    .trips-header .expand-icon {
        font-size: 10px;
        transition: transform 0.2s;
    }

    .trips-header .expand-icon.expanded {
        transform: rotate(90deg);
    }

    .trips-count {
        background: rgba(255,255,255,0.15);
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 11px;
        margin-left: auto;
    }

    .nav-sidebar.collapsed .trips-section .trips-header {
        display: none;
    }

    /* Application Orgs Section (collapsible) */
    .app-orgs-section {
        border-top: 1px solid rgba(255,255,255,0.1);
        margin-top: 10px;
        padding-top: 5px;
    }

    .app-orgs-header {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px 15px;
        color: rgba(255,255,255,0.6);
        cursor: pointer;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        transition: background 0.2s;
    }

    .app-orgs-header:hover {
        background: rgba(255,255,255,0.05);
        color: rgba(255,255,255,0.8);
    }

    .app-orgs-header .expand-icon {
        font-size: 10px;
        transition: transform 0.2s;
    }

    .app-orgs-header .expand-icon.expanded {
        transform: rotate(90deg);
    }

    .app-org-count {
        background: rgba(255,255,255,0.15);
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 11px;
        margin-left: auto;
    }

    .app-orgs-list {
        list-style: none;
        padding: 0;
        margin: 0;
    }

    .app-org-item {
        margin: 0;
    }

    .app-org-link {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px 15px 8px 25px;
        color: rgba(255,255,255,0.6);
        text-decoration: none;
        font-size: 13px;
        transition: all 0.2s;
        border-left: 3px solid transparent;
    }

    .app-org-link:hover {
        background: var(--sidebar-hover);
        color: white;
    }

    .app-org-item.active .app-org-link {
        background: var(--sidebar-active);
        color: white;
        border-left-color: var(--sidebar-accent);
    }

    .app-org-icon {
        font-size: 14px;
        opacity: 0.7;
    }

    .app-org-name {
        flex: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    .app-org-badge {
        font-size: 10px;
        padding: 2px 6px;
        border-radius: 8px;
        background: rgba(255,255,255,0.1);
    }

    .nav-sidebar.collapsed .app-orgs-section {
        display: none;
    }

    /* Main Panel - flex grows to fill remaining space */
    .main-panel {
        flex: 1;
        padding: 20px;
        background: #f5f5f5;
        min-width: 0;  /* Allow flex item to shrink below content size */
    }

    /* Mobile Toggle Button */
    .mobile-toggle {
        display: none;
        position: fixed;
        bottom: 20px;
        right: 20px;
        left: auto;
        top: auto;
        z-index: 1050;
        background: var(--primary-color);
        color: white;
        border: 3px solid white;
        padding: 10px 16px;
        border-radius: 50px;
        cursor: pointer;
        font-size: 20px;
        line-height: 1;
        box-shadow: 0 4px 16px rgba(0,0,0,0.35);
        -webkit-tap-highlight-color: transparent;
        flex-direction: row;
        align-items: center;
        justify-content: center;
        gap: 8px;
    }

    .mobile-toggle-label {
        display: inline;
        font-size: 14px;
        font-weight: 700;
        letter-spacing: 0.5px;
    }

    .mobile-toggle:hover,
    .mobile-toggle:active {
        background: var(--secondary-color);
        transform: scale(1.05);
    }

    /* Mobile Overlay - clicking closes sidebar */
    .mobile-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        z-index: 999;
        cursor: pointer;
        -webkit-tap-highlight-color: transparent;
    }

    /* Mobile close button inside sidebar */
    .mobile-close-btn {
        display: none;
        position: sticky;
        bottom: 0;
        left: 0;
        right: 0;
        padding: 16px;
        background: linear-gradient(to top, var(--sidebar-bg) 80%, transparent);
        text-align: center;
        z-index: 11;
    }

    .mobile-close-btn button {
        width: 100%;
        padding: 16px;
        background: rgba(255,255,255,0.15);
        border: 1px solid rgba(255,255,255,0.3);
        border-radius: 8px;
        color: white;
        font-size: 16px;
        font-weight: 600;
        cursor: pointer;
        -webkit-tap-highlight-color: transparent;
    }

    .mobile-close-btn button:active {
        background: rgba(255,255,255,0.25);
    }

    @media (max-width: 768px) {
        .mobile-close-btn {
            display: block;
        }
    }

    /* Floating Action Button - Email Team */
    .fab-container {
        position: fixed;
        bottom: 30px;
        right: 30px;
        z-index: 200;
    }

    .fab-button {
        width: 56px;
        height: 56px;
        border-radius: 50%;
        background: var(--primary-color);
        color: white;
        border: none;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 24px;
        transition: all 0.3s ease;
    }

    .fab-button:hover {
        transform: scale(1.1);
        box-shadow: 0 6px 16px rgba(0,0,0,0.4);
    }

    .fab-tooltip {
        position: absolute;
        right: 70px;
        top: 50%;
        transform: translateY(-50%);
        background: #333;
        color: white;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 14px;
        white-space: nowrap;
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.3s ease;
    }

    .fab-container:hover .fab-tooltip {
        opacity: 1;
    }

    /* Action Buttons */
    .action-btn {
        padding: 4px 8px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        transition: all 0.2s ease;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }

    .action-btn-email {
        background: #17a2b8;
        color: white;
    }

    .action-btn-email:hover {
        background: #138496;
    }

    .action-btn-view {
        background: #6c757d;
        color: white;
    }

    .action-btn-view:hover {
        background: #545b62;
    }

    /* Actions dropdown menu */
    .actions-dropdown {
        position: relative;
        display: inline-block;
    }

    .actions-dropdown-btn {
        padding: 6px 12px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 12px;
        background: #6c757d;
        color: white;
    }

    .actions-dropdown-btn:hover {
        background: #545b62;
    }

    .actions-dropdown-content {
        display: none;
        position: fixed;
        background-color: white;
        min-width: 180px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
        z-index: 10000;
        border-radius: 4px;
        overflow: visible;
    }

    .actions-dropdown-content.show {
        display: block;
    }

    .actions-dropdown-item {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px 12px;
        text-decoration: none;
        color: #333;
        cursor: pointer;
        border: none;
        background: none;
        width: 100%;
        text-align: left;
        font-size: 13px;
    }

    .actions-dropdown-item:hover {
        background-color: #f1f1f1;
    }

    .actions-dropdown-item .icon {
        width: 20px;
        text-align: center;
    }

    .actions-dropdown-item .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        margin-left: auto;
    }

    .actions-dropdown-item .status-dot.complete {
        background: #28a745;
    }

    .actions-dropdown-item .status-dot.incomplete {
        background: #dc3545;
    }

    .actions-dropdown-item .status-dot.warning {
        background: #fd7e14;
    }

    /* Email Popup Modal */
    .email-modal-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.6);
        z-index: 1000;
        align-items: center;
        justify-content: center;
    }

    .email-modal-overlay.active {
        display: flex;
    }

    .email-modal {
        background: white;
        border-radius: 8px;
        width: 90%;
        max-width: 600px;
        max-height: 90vh;
        overflow: hidden;
        box-shadow: 0 10px 40px rgba(0,0,0,0.3);
    }

    .email-modal-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 15px 20px;
        background: var(--primary-color);
        color: white;
    }

    .email-modal-header h4 {
        margin: 0;
        font-size: 18px;
    }

    .email-modal-close {
        background: none;
        border: none;
        color: white;
        font-size: 24px;
        cursor: pointer;
        padding: 0;
        line-height: 1;
    }

    .email-modal-body {
        padding: 20px;
        max-height: 60vh;
        overflow-y: auto;
    }

    .email-form-group {
        margin-bottom: 15px;
    }

    .email-form-group label {
        display: block;
        margin-bottom: 5px;
        font-weight: 600;
        color: #333;
    }

    .email-form-group input,
    .email-form-group textarea {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 14px;
        box-sizing: border-box;
    }

    .email-form-group input:focus,
    .email-form-group textarea:focus {
        outline: none;
        border-color: var(--primary-color);
        box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.2);
    }

    .email-form-group textarea {
        min-height: 150px;
        resize: vertical;
    }

    /* Email Formatting Toolbar */
    .email-formatting-toolbar {
        display: flex;
        flex-wrap: wrap;
        gap: 4px;
        padding: 8px;
        background: #f5f5f5;
        border: 1px solid #ddd;
        border-bottom: none;
        border-radius: 4px 4px 0 0;
    }

    .email-formatting-toolbar button {
        padding: 6px 10px;
        background: white;
        border: 1px solid #d1d5db;
        border-radius: 3px;
        cursor: pointer;
        font-size: 13px;
        color: #333;
        transition: background 0.2s, border-color 0.2s;
    }

    .email-formatting-toolbar button:hover {
        background: #e5e7eb;
        border-color: #9ca3af;
    }

    .email-formatting-toolbar button:active {
        background: #d1d5db;
    }

    .email-body-editor {
        width: 100%;
        min-height: 180px;
        max-height: 300px;
        overflow-y: auto;
        padding: 12px;
        border: 1px solid #ddd;
        border-radius: 0 0 4px 4px;
        font-size: 14px;
        font-family: Arial, sans-serif;
        line-height: 1.6;
        box-sizing: border-box;
        background: white;
        outline: none;
    }

    .email-body-editor:focus {
        border-color: var(--primary-color);
        box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.2);
    }

    .email-body-editor:empty:before {
        content: attr(placeholder);
        color: #999;
        pointer-events: none;
    }

    .email-body-editor ul,
    .email-body-editor ol {
        margin: 0 0 10px 20px;
        padding: 0;
    }

    .email-body-editor a {
        color: #1a73e8;
        text-decoration: underline;
    }

    .email-modal-footer {
        padding: 15px 20px;
        background: #f5f5f5;
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
    }

    .email-footer-buttons {
        display: flex;
        gap: 10px;
    }

    .email-advanced-link {
        color: #666;
        font-size: 13px;
        text-decoration: none;
        padding: 8px 12px;
        border-radius: 4px;
        transition: background 0.2s, color 0.2s;
    }

    .email-advanced-link:hover {
        background: #e0e0e0;
        color: #333;
        text-decoration: none;
    }

    .email-btn {
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
    }

    .email-btn-cancel {
        background: #e0e0e0;
        color: #333;
    }

    .email-btn-send {
        background: var(--primary-color);
        color: white;
    }

    .email-btn-send:hover {
        background: #1557b0;
    }

    /* Section Action Bar */
    .section-actions {
        display: flex;
        gap: 10px;
        margin-bottom: 15px;
    }

    .section-action-btn {
        padding: 8px 16px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        transition: all 0.2s ease;
    }

    .section-action-btn-primary {
        background: var(--primary-color);
        color: white;
    }

    .section-action-btn-primary:hover {
        background: #1557b0;
    }

    .section-action-btn-secondary {
        background: #6c757d;
        color: white;
    }

    .section-action-btn-secondary:hover {
        background: #545b62;
    }

    /* Mobile Responsive */
    @media (max-width: 768px) {
        .nav-sidebar {
            position: fixed;
            top: 0;
            left: 0;
            height: 100vh;
            transform: translateX(-100%);
            z-index: 1000;
            overflow-y: auto;
            overflow-x: hidden;
            -webkit-overflow-scrolling: touch;
            overscroll-behavior: contain;
        }

        .nav-sidebar.mobile-open {
            transform: translateX(0);
        }

        /* Larger touch targets for mobile */
        .nav-sidebar .nav-link {
            padding: 16px 20px;
            min-height: 52px;
        }

        .nav-sidebar .trip-header {
            padding: 16px 15px;
            min-height: 56px;
        }

        .nav-sidebar .section-link {
            padding: 14px 20px 14px 45px;
            min-height: 48px;
        }

        .nav-sidebar .expand-icon {
            font-size: 14px;
            padding: 8px;
            margin: -8px;
            margin-right: 0;
        }

        /* Make trip list scrollable within sidebar */
        .nav-sidebar .trip-list {
            padding-bottom: 100px;
        }

        /* Close button at bottom of mobile sidebar for easy access */
        .nav-sidebar .sidebar-header {
            position: sticky;
            top: 0;
            z-index: 10;
        }

        .main-panel {
            padding: 60px 15px 15px;
            width: 100%;
        }

        .mobile-toggle {
            display: flex;
        }

        .fab-container {
            display: none !important;
        }

        .nav-sidebar.mobile-open ~ .mobile-overlay {
            display: block;
        }

        /* Adjust KPI cards for mobile */
        .kpi-container {
            flex-direction: column;
        }

        .kpi-card {
            min-width: 100%;
        }
    }

    /* Tablet adjustments */
    @media (min-width: 769px) and (max-width: 1024px) {
        :root {
            --sidebar-width: 240px;
        }

        .section-link {
            padding-left: 25px;
            font-size: 12px;
        }
    }

    /* Section content header */
    .section-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 20px;
        padding-bottom: 15px;
        border-bottom: 2px solid var(--primary-color);
    }

    .section-header h2 {
        margin: 0;
        color: var(--primary-color);
        font-size: 24px;
    }

    .section-header .breadcrumb {
        font-size: 13px;
        color: var(--text-muted);
    }

    .section-header .breadcrumb a {
        color: var(--secondary-color);
        text-decoration: none;
    }

    .section-header .breadcrumb a:hover {
        text-decoration: underline;
    }

    /* Content cards for sections */
    .content-card {
        background: white;
        border-radius: var(--radius);
        box-shadow: var(--shadow);
        padding: 20px;
        margin-bottom: 20px;
    }

    .content-card h3 {
        margin-top: 0;
        color: var(--primary-color);
        border-bottom: 1px solid var(--border-color);
        padding-bottom: 10px;
    }

    /* Trip info grid */
    .trip-info-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 15px;
    }

    .trip-info-item {
        padding: 15px;
        background: var(--light-bg);
        border-radius: var(--radius);
    }

    .trip-info-item .label {
        font-size: 12px;
        color: var(--text-muted);
        text-transform: uppercase;
        margin-bottom: 5px;
    }

    .trip-info-item .value {
        font-size: 18px;
        font-weight: bold;
        color: var(--primary-color);
    }

    /* End Sidebar Navigation Styles */

    </style>
    '''

# ::END:: Styles

#####################################################################
# PERFORMANCE MONITORING
#####################################################################

# ::START:: Performance
def start_timer():
    """Start performance timer"""
    return datetime.datetime.now()

def end_timer(start_time, label="Query"):
    """End timer and optionally display result"""
    elapsed = (datetime.datetime.now() - start_time).total_seconds()
    if config.ENABLE_DEVELOPER_VISUALIZATION and model.UserIsInRole("SuperAdmin"):
        return '<span class="performance-indicator">{0}: {1:.2f}s</span>'.format(label, elapsed)
    return ""

# ::END:: Performance

#####################################################################
# DATABASE QUERIES - OPTIMIZED FOR PERFORMANCE
#####################################################################

# ::START:: Database Queries
def get_total_stats_query():
    """Get overall mission statistics - using optimized CTE that runs in ~1 second"""
    # Build the exclusion list for application orgs
    app_org_list = ','.join(str(x) for x in config.APPLICATION_ORG_IDS)
    
    # Get the mission trip totals CTE for ALL trips (including closed)
    mission_trip_cte_all = get_mission_trip_totals_cte(include_closed=True).replace('WITH ', '').strip()
    
    # Extract just the CTE definitions from the mission trip CTE (before the final SELECT)
    # Find where MissionTripTotals AS ends
    cte_end_pos = mission_trip_cte_all.rfind('MissionTripTotals AS')
    if cte_end_pos > 0:
        # Find the closing parenthesis after MissionTripTotals definition
        paren_count = 0
        found_start = False
        for i in range(cte_end_pos, len(mission_trip_cte_all)):
            if mission_trip_cte_all[i] == '(':
                paren_count += 1
                found_start = True
            elif mission_trip_cte_all[i] == ')' and found_start:
                paren_count -= 1
                if paren_count == 0:
                    # Include everything up to and including this closing paren
                    mission_trip_cte_defs = mission_trip_cte_all[:i+1]
                    break
    else:
        # Fallback - use the whole thing
        mission_trip_cte_defs = mission_trip_cte_all
    
    # Build the complete query with all CTEs merged
    query = '''
    WITH ''' + mission_trip_cte_defs + ''',
    ActiveMissions AS (
        SELECT o.OrganizationId, o.MemberCount
        FROM Organizations o WITH (NOLOCK)
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND NOT EXISTS (
              SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
              WHERE oe.OrganizationId = o.OrganizationId
                AND oe.Field = 'Close'
                AND oe.DateValue IS NOT NULL
                AND oe.DateValue <= GETDATE()
          )
          AND o.OrganizationId NOT IN ({2})
    ),
    ApplicationOrgs AS (
        SELECT SUM(MemberCount) as TotalApplications
        FROM Organizations WITH (NOLOCK)
        WHERE OrganizationId IN ({2})
    ),
    OutstandingPaymentsOpen AS (
        SELECT
            SUM(mtt.Due) AS TotalDue
        FROM MissionTripTotals mtt
        INNER JOIN Organizations o WITH (NOLOCK) ON mtt.InvolvementId = o.OrganizationId
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND mtt.Due > 0
          AND mtt.Name <> 'total'
          AND mtt.SortOrder <> 'ZZZZZ'
          AND NOT EXISTS (
              SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
              WHERE oe.OrganizationId = o.OrganizationId
                AND oe.Field = 'Close'
                AND oe.DateValue IS NOT NULL
                AND oe.DateValue <= GETDATE()
          )
    ),
    OutstandingPaymentsClosed AS (
        SELECT 
            SUM(mtt.Due) AS TotalDue
        FROM MissionTripTotals mtt
        INNER JOIN Organizations o WITH (NOLOCK) ON mtt.InvolvementId = o.OrganizationId
        WHERE o.IsMissionTrip = {0} 
          AND o.OrganizationStatusId = {1} 
          AND mtt.Due > 0
          AND mtt.Name <> 'total'
          AND mtt.SortOrder <> 'ZZZZZ'
          AND EXISTS (
              SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
              WHERE oe.OrganizationId = o.OrganizationId 
                AND oe.Field = 'Close'
                AND oe.DateValue <= GETDATE()
          )
    ),
    OutstandingPaymentsAll AS (
        SELECT 
            SUM(mtt.Due) AS TotalDue
        FROM MissionTripTotals mtt
        INNER JOIN Organizations o WITH (NOLOCK) ON mtt.InvolvementId = o.OrganizationId
        WHERE o.IsMissionTrip = {0} 
          AND o.OrganizationStatusId = {1} 
          AND mtt.Due > 0
          AND mtt.Name <> 'total'
          AND mtt.SortOrder <> 'ZZZZZ'
    )
    SELECT 
        ISNULL(SUM(am.MemberCount), 0) AS TotalMembers,
        ISNULL((SELECT TotalApplications FROM ApplicationOrgs), 0) AS TotalApplications,
        ISNULL((SELECT TotalDue FROM OutstandingPaymentsOpen), 0) AS TotalOutstandingOpen,
        ISNULL((SELECT TotalDue FROM OutstandingPaymentsClosed), 0) AS TotalOutstandingClosed,
        ISNULL((SELECT TotalDue FROM OutstandingPaymentsAll), 0) AS TotalOutstandingAll
    FROM ActiveMissions am
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, app_org_list)
    
    return query


def get_active_missions_query_optimized(show_closed=False):
    """Highly optimized query for active missions"""
    
    # Build closed filter
    closed_filter = ""
    if not show_closed:
        closed_filter = "AND EventStatus <> 'Closed'"
    
    # Get the mission trip totals CTE without the WITH keyword
    mission_trip_cte = get_mission_trip_totals_cte(include_closed=show_closed).replace('WITH ', '').strip()
    
    # Extract just the CTE definitions (everything before the final MissionTripTotals result)
    cte_end_pos = mission_trip_cte.rfind('MissionTripTotals AS')
    if cte_end_pos > 0:
        # Find the closing parenthesis after MissionTripTotals definition
        paren_count = 0
        found_start = False
        for i in range(cte_end_pos, len(mission_trip_cte)):
            if mission_trip_cte[i] == '(':
                paren_count += 1
                found_start = True
            elif mission_trip_cte[i] == ')' and found_start:
                paren_count -= 1
                if paren_count == 0:
                    # Include everything up to and including this closing paren
                    mission_trip_cte_defs = mission_trip_cte[:i+1]
                    break
    else:
        # Fallback - use the whole thing
        mission_trip_cte_defs = mission_trip_cte
    
    # Use single CTE for better performance
    return '''
    WITH ''' + mission_trip_cte_defs + ''',
    MissionSummary AS (
        SELECT 
            o.OrganizationId,
            o.OrganizationName,
            o.MemberCount,
            o.ImageUrl,
            o.RegStart,
            o.RegEnd,
            
            -- Financial summary
            ISNULL(mtt.Outstanding, 0) AS Outstanding,
            ISNULL(mtt.TotalDue, 0) AS TotalDue,
            ISNULL(mtt.TotalFee, 0) AS TotalFee,
            
            -- Event dates and status calculation
            oe.StartDate,
            oe.EndDate,
            CASE 
                WHEN oe.StartDate > GETDATE() THEN 'Pre'
                WHEN oe.StartDate <= GETDATE() AND oe.EndDate >= GETDATE() THEN 'Event'
                WHEN oe.CloseDate < GETDATE() THEN 'Closed'
                WHEN oe.EndDate < GETDATE() THEN 'Post'
                ELSE 'Unknown'
            END AS EventStatus,
            
            -- Registration status
            CASE 
                WHEN o.RegStart > GETDATE() THEN 'Pending'
                WHEN o.RegEnd < GETDATE() THEN 'Closed'
                ELSE 'Open'
            END AS RegistrationStatus,
            
            -- Pre-calculate meeting counts
            ISNULL(m.PastMeetings, 0) AS PastMeetings,
            ISNULL(m.FutureMeetings, 0) AS FutureMeetings,
            
            -- Pre-calculate check counts
            ISNULL(bc.BGGood, 0) AS BackgroundCheckGood,
            ISNULL(bc.BGBad, 0) AS BackgroundCheckBad,
            ISNULL(bc.BGMissing, 0) AS BackgroundCheckMissing,
            
            -- Passport count
            ISNULL(pc.PassportCount, 0) AS PeopleWithPassports
            
        FROM Organizations o WITH (NOLOCK)
        
        -- Get financial totals from our CTE
        LEFT JOIN (
            SELECT 
                InvolvementId,
                SUM(Due) AS Outstanding,
                SUM(Raised) AS TotalDue,
                SUM(TripCost) AS TotalFee
            FROM MissionTripTotals
            WHERE SortOrder <> 'ZZZZZ'
            GROUP BY InvolvementId
        ) mtt ON mtt.InvolvementId = o.OrganizationId
        
        -- Get event dates
        LEFT JOIN (
            SELECT 
                OrganizationId,
                MAX(CASE WHEN Field = 'Main Event Start' THEN DateValue END) AS StartDate,
                MAX(CASE WHEN Field = 'Main Event End' THEN DateValue END) AS EndDate,
                MAX(CASE WHEN Field = 'Close' THEN DateValue END) AS CloseDate
            FROM OrganizationExtra WITH (NOLOCK)
            WHERE Field IN ('Main Event Start', 'Main Event End', 'Close')
            GROUP BY OrganizationId
        ) oe ON oe.OrganizationId = o.OrganizationId
        
        -- Get meeting counts
        LEFT JOIN (
            SELECT 
                OrganizationId,
                SUM(CASE WHEN MeetingDate < GETDATE() THEN 1 ELSE 0 END) AS PastMeetings,
                SUM(CASE WHEN MeetingDate >= GETDATE() THEN 1 ELSE 0 END) AS FutureMeetings
            FROM Meetings WITH (NOLOCK)
            GROUP BY OrganizationId
        ) m ON m.OrganizationId = o.OrganizationId
        
        -- Get background check summary
        LEFT JOIN (
            SELECT 
                om.OrganizationId,
                SUM(CASE WHEN vs.Description = 'Complete' AND v.ProcessedDate >= DATEADD(YEAR, -3, GETDATE()) THEN 1 ELSE 0 END) AS BGGood,
                SUM(CASE WHEN vs.Description <> 'Complete' AND v.ProcessedDate >= DATEADD(YEAR, -3, GETDATE()) THEN 1 ELSE 0 END) AS BGBad,
                SUM(CASE WHEN v.ProcessedDate IS NULL OR v.ProcessedDate < DATEADD(YEAR, -3, GETDATE()) THEN 1 ELSE 0 END) AS BGMissing
            FROM OrganizationMembers om WITH (NOLOCK)
            LEFT JOIN Volunteer v WITH (NOLOCK) ON v.PeopleId = om.PeopleId
            LEFT JOIN lookup.VolApplicationStatus vs WITH (NOLOCK) ON vs.Id = v.StatusId
            WHERE om.MemberTypeId <> {0}
              AND om.InactiveDate IS NULL
            GROUP BY om.OrganizationId
        ) bc ON bc.OrganizationId = o.OrganizationId
        
        -- Get passport count
        LEFT JOIN (
            SELECT 
                om.OrganizationId,
                COUNT(DISTINCT om.PeopleId) AS PassportCount
            FROM OrganizationMembers om WITH (NOLOCK)
            INNER JOIN RecReg r WITH (NOLOCK) ON om.PeopleId = r.PeopleId
            WHERE om.MemberTypeId <> {0}
              AND om.InactiveDate IS NULL
              AND r.passportnumber IS NOT NULL 
              AND r.passportexpires IS NOT NULL
            GROUP BY om.OrganizationId
        ) pc ON pc.OrganizationId = o.OrganizationId
        
        WHERE o.IsMissionTrip = {1}
          AND o.OrganizationStatusId = {2}
    )
    SELECT * FROM MissionSummary
    WHERE 1=1 {3}
    ORDER BY 
        CASE EventStatus
            WHEN 'Event' THEN 1
            WHEN 'Pre' THEN 2
            WHEN 'Post' THEN 3
            WHEN 'Closed' THEN 4
            ELSE 5
        END,
        StartDate,
        OrganizationName
    '''.format(config.MEMBER_TYPE_LEADER, config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, closed_filter)

def get_active_missions_query(show_closed=False):
    """Get active missions - use optimized version"""
    # Check if caching is enabled
    if config.USE_CACHING:
        cache_key = "missions_data_{0}".format("all" if show_closed else "active")
        cached_data = model.Cache(cache_key) if hasattr(model, 'Cache') else None
        if cached_data:
            return cached_data
    
    return get_active_missions_query_optimized(show_closed)

def get_active_mission_leaders_query(org_id):
    """Get leaders for a specific mission - optimized"""
    return '''
    SELECT p.PeopleId,
        p.Name,
        p.Age,
        om.OrganizationId,
        mt.Description AS Leader
    FROM OrganizationMembers om WITH (NOLOCK)
    INNER JOIN lookup.MemberType mt WITH (NOLOCK) ON mt.Id = om.MemberTypeId
    INNER JOIN lookup.AttendType at WITH (NOLOCK) ON at.Id = mt.AttendanceTypeId
    INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
    WHERE om.OrganizationId = {0} 
      AND at.id = {1}
      AND om.InactiveDate IS NULL
    ORDER BY p.Name
    '''.format(org_id, config.ATTENDANCE_TYPE_LEADER)

def get_popup_data_query(org_id, list_type):
    """Get data for popup displays based on list type"""
    
    if list_type == 'meetings':
        return '''
        SELECT
            COALESCE(me_desc.Data, m.Description) as Description,
            COALESCE(me_loc.Data, m.Location) as Location,
            m.MeetingDate,
            CASE WHEN m.MeetingDate < GETDATE() THEN 'Past' ELSE 'Upcoming' END AS Status
        FROM Meetings m WITH (NOLOCK)
        LEFT JOIN MeetingExtra me_desc WITH (NOLOCK) ON m.MeetingId = me_desc.MeetingId AND me_desc.Field = 'Description'
        LEFT JOIN MeetingExtra me_loc WITH (NOLOCK) ON m.MeetingId = me_loc.MeetingId AND me_loc.Field = 'Location'
        WHERE m.OrganizationId = {0}
        ORDER BY m.MeetingDate DESC
        '''.format(org_id)
    
    elif list_type == 'passports':
        return '''
        SELECT 
            p.Name,
            p.Age,
            p.PeopleId
        FROM OrganizationMembers om WITH (NOLOCK)
        INNER JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
        LEFT JOIN RecReg r WITH (NOLOCK) ON p.PeopleId = r.PeopleId
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId <> {1}
          AND (r.passportnumber IS NULL OR r.passportexpires IS NULL)
          AND om.InactiveDate IS NULL
        ORDER BY p.Name2
        '''.format(org_id, config.MEMBER_TYPE_LEADER)
    
    elif list_type == 'bgchecks':
        return '''
        SELECT 
            p.Name,
            p.Age,
            p.PeopleId,
            CASE 
                WHEN vs.Description <> 'Complete' THEN 'Failed: ' + vs.Description
                WHEN v.ProcessedDate < DATEADD(YEAR, -3, GETDATE()) THEN 'Expired'
                WHEN v.ProcessedDate IS NULL THEN 'Missing'
                ELSE 'Unknown Issue'
            END AS Issue
        FROM OrganizationMembers om WITH (NOLOCK)
        INNER JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
        LEFT JOIN Volunteer v WITH (NOLOCK) ON v.PeopleId = p.PeopleId
        LEFT JOIN lookup.VolApplicationStatus vs WITH (NOLOCK) ON vs.Id = v.StatusId
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId <> {1}
          AND om.InactiveDate IS NULL
          AND (vs.Description IS NULL 
               OR vs.Description <> 'Complete' 
               OR v.ProcessedDate IS NULL 
               OR v.ProcessedDate < DATEADD(YEAR, -3, GETDATE()))
        ORDER BY p.Name2
        '''.format(org_id, config.MEMBER_TYPE_LEADER)
    
    elif list_type == 'people':
        return '''
        SELECT 
            p.Name,
            p.Age,
            p.PeopleId,
            mt.Description AS Role,
            p.CellPhone,
            p.EmailAddress
        FROM OrganizationMembers om WITH (NOLOCK)
        INNER JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
        LEFT JOIN lookup.MemberType mt WITH (NOLOCK) ON mt.Id = om.MemberTypeId
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId <> {1}
          AND om.InactiveDate IS NULL
        ORDER BY mt.Description, p.Name2
        '''.format(org_id, config.MEMBER_TYPE_LEADER)
    
    else:
        return None

def get_enhanced_stats_queries():
    """Get enhanced statistics queries for missions pastors"""
    
    # Get the mission trip totals CTE for ALL trips (including closed)
    mission_trip_cte_all = get_mission_trip_totals_cte(include_closed=True).replace('WITH ', '').strip()
    
    # Extract just the CTE definitions from the mission trip CTE
    cte_end_pos = mission_trip_cte_all.rfind('MissionTripTotals AS')
    if cte_end_pos > 0:
        paren_count = 0
        found_start = False
        for i in range(cte_end_pos, len(mission_trip_cte_all)):
            if mission_trip_cte_all[i] == '(':
                paren_count += 1
                found_start = True
            elif mission_trip_cte_all[i] == ')' and found_start:
                paren_count -= 1
                if paren_count == 0:
                    mission_trip_cte_defs = mission_trip_cte_all[:i+1]
                    break
    else:
        mission_trip_cte_defs = mission_trip_cte_all
    
    queries = {
        'financial_summary': '''
            WITH ''' + mission_trip_cte_defs + ''',
            FinancialBreakdown AS (
                SELECT 
                    'Open' AS TripStatus,
                    COUNT(DISTINCT o.OrganizationId) AS TotalTrips,
                    SUM(mtt.TripCost) AS TotalGoal,
                    SUM(mtt.Raised) AS TotalRaised,
                    SUM(mtt.Due) AS TotalOutstanding
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN MissionTripTotals mtt ON mtt.InvolvementId = o.OrganizationId
                WHERE o.IsMissionTrip = ''' + str(config.MISSION_TRIP_FLAG) + '''
                  AND o.OrganizationStatusId = ''' + str(config.ACTIVE_ORG_STATUS_ID) + '''
                  AND (mtt.Name <> 'total' OR mtt.Name IS NULL)
                  AND (mtt.SortOrder <> 'ZZZZZ' OR mtt.SortOrder IS NULL)
                  AND NOT EXISTS (
                      SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
                      WHERE oe.OrganizationId = o.OrganizationId
                        AND oe.Field = 'Close'
                        AND oe.DateValue IS NOT NULL
                        AND oe.DateValue <= GETDATE()
                  )

                UNION ALL

                SELECT
                    'Closed' AS TripStatus,
                    COUNT(DISTINCT o.OrganizationId) AS TotalTrips,
                    SUM(mtt.TripCost) AS TotalGoal,
                    SUM(mtt.Raised) AS TotalRaised,
                    SUM(mtt.Due) AS TotalOutstanding
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN MissionTripTotals mtt ON mtt.InvolvementId = o.OrganizationId
                WHERE o.IsMissionTrip = ''' + str(config.MISSION_TRIP_FLAG) + '''
                  AND o.OrganizationStatusId = ''' + str(config.ACTIVE_ORG_STATUS_ID) + '''
                  AND (mtt.Name <> 'total' OR mtt.Name IS NULL)
                  AND (mtt.SortOrder <> 'ZZZZZ' OR mtt.SortOrder IS NULL)
                  AND EXISTS (
                      SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
                      WHERE oe.OrganizationId = o.OrganizationId 
                        AND oe.Field = 'Close'
                        AND oe.DateValue <= GETDATE()
                  )
                
                UNION ALL
                
                SELECT 
                    'All' AS TripStatus,
                    COUNT(DISTINCT o.OrganizationId) AS TotalTrips,
                    SUM(mtt.TripCost) AS TotalGoal,
                    SUM(mtt.Raised) AS TotalRaised,
                    SUM(mtt.Due) AS TotalOutstanding
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN MissionTripTotals mtt ON mtt.InvolvementId = o.OrganizationId
                WHERE o.IsMissionTrip = ''' + str(config.MISSION_TRIP_FLAG) + '''
                  AND o.OrganizationStatusId = ''' + str(config.ACTIVE_ORG_STATUS_ID) + '''
                  AND (mtt.Name <> 'total' OR mtt.Name IS NULL)
                  AND (mtt.SortOrder <> 'ZZZZZ' OR mtt.SortOrder IS NULL)
            )
            SELECT 
                TripStatus,
                TotalTrips,
                TotalGoal,
                TotalRaised,
                TotalOutstanding,
                CASE 
                    WHEN TotalGoal > 0 
                    THEN ROUND((TotalRaised * 100.0) / TotalGoal, 0)
                    ELSE 0 
                END AS PercentRaised
            FROM FinancialBreakdown
        ''',
        
        'trip_trends': '''
            SELECT 
                YEAR(oe.DateValue) AS TripYear,
                COUNT(DISTINCT o.OrganizationId) AS TripCount,
                SUM(o.MemberCount) AS TotalParticipants
            FROM Organizations o WITH (NOLOCK)
            INNER JOIN OrganizationExtra oe WITH (NOLOCK) ON o.OrganizationId = oe.OrganizationId
            WHERE o.IsMissionTrip = {0}
              AND o.OrganizationStatusId = {1}
              AND oe.Field = 'Main Event Start'
              AND oe.DateValue >= DATEADD(YEAR, -5, GETDATE())
            GROUP BY YEAR(oe.DateValue)
            ORDER BY TripYear DESC
        '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID),
        
        'participation_by_age': '''
            SELECT 
                CASE 
                    WHEN p.Age < 18 THEN 'Youth (Under 18)'
                    WHEN p.Age BETWEEN 18 AND 25 THEN 'Young Adults (18-25)'
                    WHEN p.Age BETWEEN 26 AND 35 THEN 'Adults (26-35)'
                    WHEN p.Age BETWEEN 36 AND 50 THEN 'Middle Age (36-50)'
                    WHEN p.Age BETWEEN 51 AND 65 THEN 'Seniors (51-65)'
                    WHEN p.Age > 65 THEN 'Senior Plus (65+)'
                    ELSE 'Unknown'
                END AS AgeGroup,
                COUNT(DISTINCT p.PeopleId) AS ParticipantCount,
                ROUND(AVG(CAST(p.Age AS FLOAT)), 1) AS AvgAge,
                CASE 
                    WHEN MIN(p.Age) < 18 THEN 1
                    WHEN MIN(p.Age) BETWEEN 18 AND 25 THEN 2
                    WHEN MIN(p.Age) BETWEEN 26 AND 35 THEN 3
                    WHEN MIN(p.Age) BETWEEN 36 AND 50 THEN 4
                    WHEN MIN(p.Age) BETWEEN 51 AND 65 THEN 5
                    WHEN MIN(p.Age) > 65 THEN 6
                    ELSE 7
                END AS SortOrder
            FROM Organizations o WITH (NOLOCK)
            INNER JOIN OrganizationMembers om WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
            INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
            WHERE o.IsMissionTrip = {0}
              AND o.OrganizationStatusId = {1}
              AND om.InactiveDate IS NULL
              AND p.Age IS NOT NULL
            GROUP BY 
                CASE 
                    WHEN p.Age < 18 THEN 'Youth (Under 18)'
                    WHEN p.Age BETWEEN 18 AND 25 THEN 'Young Adults (18-25)'
                    WHEN p.Age BETWEEN 26 AND 35 THEN 'Adults (26-35)'
                    WHEN p.Age BETWEEN 36 AND 50 THEN 'Middle Age (36-50)'
                    WHEN p.Age BETWEEN 51 AND 65 THEN 'Seniors (51-65)'
                    WHEN p.Age > 65 THEN 'Senior Plus (65+)'
                    ELSE 'Unknown'
                END
            ORDER BY 
                MIN(CASE 
                    WHEN p.Age < 18 THEN 1
                    WHEN p.Age BETWEEN 18 AND 25 THEN 2
                    WHEN p.Age BETWEEN 26 AND 35 THEN 3
                    WHEN p.Age BETWEEN 36 AND 50 THEN 4
                    WHEN p.Age BETWEEN 51 AND 65 THEN 5
                    WHEN p.Age > 65 THEN 6
                    ELSE 7
                END)
        '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID),
        
        'repeat_participants': '''
            SELECT 
                TripCount,
                COUNT(*) AS PeopleCount
            FROM (
                SELECT 
                    om.PeopleId,
                    COUNT(DISTINCT om.OrganizationId) AS TripCount
                FROM OrganizationMembers om WITH (NOLOCK)
                INNER JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
                WHERE o.IsMissionTrip = {0}
                  AND o.OrganizationStatusId IN (30, 40) -- Active and Inactive
                  AND om.MemberTypeId <> {2}
                GROUP BY om.PeopleId
            ) tc
            GROUP BY TripCount
            ORDER BY TripCount
        '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, config.MEMBER_TYPE_LEADER),
        
        'financial_by_year': '''
            {0}
            SELECT 
                YEAR(GETDATE()) AS Year,
                COUNT(DISTINCT mtt.InvolvementId) AS TripCount,
                SUM(mtt.TripCost) AS TotalGoal,
                SUM(mtt.Raised) AS TotalRaised,
                SUM(mtt.Due) AS TotalOutstanding,
                CASE 
                    WHEN SUM(mtt.TripCost) > 0 
                    THEN ROUND((SUM(mtt.Raised) * 100.0) / SUM(mtt.TripCost), 0)
                    ELSE 0 
                END AS PercentRaised
            FROM MissionTripTotals mtt
            INNER JOIN Organizations o WITH (NOLOCK) ON mtt.InvolvementId = o.OrganizationId
            LEFT JOIN OrganizationExtra oe WITH (NOLOCK) ON o.OrganizationId = oe.OrganizationId AND oe.Field = 'Main Event Start'
            WHERE YEAR(ISNULL(oe.DateValue, GETDATE())) = YEAR(GETDATE())
              AND mtt.SortOrder <> 'ZZZZZ'
              AND mtt.Name <> 'total'
            
            UNION ALL
            
            SELECT 
                YEAR(GETDATE()) - 1 AS Year,
                COUNT(DISTINCT mtt.InvolvementId) AS TripCount,
                SUM(mtt.TripCost) AS TotalGoal,
                SUM(mtt.Raised) AS TotalRaised,
                SUM(mtt.Due) AS TotalOutstanding,
                CASE 
                    WHEN SUM(mtt.TripCost) > 0 
                    THEN ROUND((SUM(mtt.Raised) * 100.0) / SUM(mtt.TripCost), 0)
                    ELSE 0 
                END AS PercentRaised
            FROM MissionTripTotals mtt
            INNER JOIN Organizations o WITH (NOLOCK) ON mtt.InvolvementId = o.OrganizationId
            LEFT JOIN OrganizationExtra oe WITH (NOLOCK) ON o.OrganizationId = oe.OrganizationId AND oe.Field = 'Main Event Start'
            WHERE YEAR(ISNULL(oe.DateValue, YEAR(GETDATE()) - 1)) = YEAR(GETDATE()) - 1
              AND mtt.SortOrder <> 'ZZZZZ'
              AND mtt.Name <> 'total'
            
            ORDER BY Year DESC
        '''.format(get_mission_trip_totals_cte(include_closed=True)),
        
        'upcoming_deadlines': '''
            {2}
            SELECT TOP 10
                o.OrganizationName,
                o.OrganizationId,
                oe.DateValue AS DeadlineDate,
                oe.Field AS DeadlineType,
                DATEDIFF(DAY, GETDATE(), oe.DateValue) AS DaysUntil,
                ISNULL((SELECT SUM(Due) FROM MissionTripTotals WHERE InvolvementId = o.OrganizationId AND SortOrder <> 'ZZZZZ' AND Name <> 'total'), 0) AS AmountRemaining
            FROM Organizations o WITH (NOLOCK)
            INNER JOIN OrganizationExtra oe WITH (NOLOCK) ON o.OrganizationId = oe.OrganizationId
            WHERE o.IsMissionTrip = {0}
              AND o.OrganizationStatusId = {1}
              AND oe.DateValue > GETDATE()
              AND oe.Field IN ('Close', 'Main Event Start', 'Registration Deadline', 'Final Payment Due')
            ORDER BY oe.DateValue
        '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, get_mission_trip_totals_cte(include_closed=True))
    }
    
    return queries

# ::END:: Database Queries

#####################################################################
# HELPER FUNCTIONS
#####################################################################

# ::START:: Helper Functions
def get_navigation_html(active_tab):
    """Generate responsive navigation with active state"""
    tabs = [
        {'id': 'dashboard', 'label': 'Dashboard', 'icon': '📊'},
        {'id': 'due', 'label': 'Finance', 'icon': '💰', 'enabled': config.ENABLE_FINANCE_TAB},
        {'id': 'messages', 'label': 'Messages', 'icon': '✉️', 'enabled': config.ENABLE_MESSAGES_TAB},
        {'id': 'stats', 'label': 'Stats', 'icon': '📈', 'enabled': config.ENABLE_STATS_TAB}
    ]
    
    nav_html = '<nav class="mission-nav">'
    for tab in tabs:
        if tab.get('enabled', True):
            active_class = 'active' if tab['id'] == active_tab else ''
            nav_html += '<a href="?simplenav={0}" class="{1}">{2} {3}</a>'.format(
                tab['id'], active_class, tab['icon'], tab['label']
            )
    nav_html += '</nav>'
    
    return nav_html

def render_kpi_cards(total_members, total_applications, total_outstanding_open, total_outstanding_closed, total_outstanding_all):
    """Render KPI cards with modern styling"""
    return '''
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="value">{0}</div>
            <div class="label">Active Members</div>
        </div>
        <div class="kpi-card">
            <div class="value">{1}</div>
            <div class="label">In the Queue</div>
        </div>
        <div class="kpi-card">
            <div class="value">{2}</div>
            <div class="label">Total Due (Open)</div>
            <div style="font-size: 0.75em; color: var(--text-muted); margin-top: 5px;">
                Active trips only<br>
                <em>From people with balances</em>
            </div>
        </div>
        <div class="kpi-card">
            <div class="value">{3}</div>
            <div class="label">Total Due (Closed)</div>
            <div style="font-size: 0.75em; color: var(--text-muted); margin-top: 5px;">
                Past trips<br>
                <em>From people with balances</em>
            </div>
        </div>
        <div class="kpi-card">
            <div class="value">{4}</div>
            <div class="label">Total Due (All)</div>
            <div style="font-size: 0.75em; color: var(--text-muted); margin-top: 5px;">
                Combined total<br>
                <em>From people with balances</em>
            </div>
        </div>
    </div>
    <div style="margin: 10px 0; padding: 10px; background: #f5f5f5; border-radius: 5px; font-size: 0.85em; color: #666;">
        <strong>Note:</strong> Dashboard totals only include outstanding amounts from participants who still owe money. 
        This differs from the Stats view which shows comprehensive financial totals for all participants.
    </div>
    '''.format(total_members, total_applications, format_currency(total_outstanding_open), 
               format_currency(total_outstanding_closed), format_currency(total_outstanding_all))

def get_popup_script():
    """JavaScript for popup functionality with AJAX loading"""
    return '''
    <script>
    // Performance monitoring
    var queryStartTime = new Date();
    
    // Get the PyScriptForm address dynamically
    function getPyScriptFormAddress() {
        let path = window.location.pathname;
        return path.replace("/PyScript/", "/PyScriptForm/");
    }
    
    function showPopup(title, content) {
        var popup = document.getElementById('dataPopup');
        var popupTitle = document.getElementById('popupTitle');
        var popupBody = document.getElementById('popupBody');
        
        popupTitle.innerHTML = title;
        popupBody.innerHTML = content;
        popup.style.display = 'block';
    }
    
    function closePopup() {
        document.getElementById('dataPopup').style.display = 'none';
    }
    
    // Close popup when clicking outside
    window.onclick = function(event) {
        var popup = document.getElementById('dataPopup');
        if (event.target == popup) {
            popup.style.display = 'none';
        }
    }
    
    // Icon header tooltips
    function showIconTooltip(event, text) {
        event.stopPropagation();
        
        // Hide any existing tooltips
        const existing = document.querySelectorAll('.icon-tooltip.show');
        existing.forEach(t => t.classList.remove('show'));
        
        // Create or get tooltip
        let tooltip = document.getElementById('iconTooltip');
        if (!tooltip) {
            tooltip = document.createElement('div');
            tooltip.id = 'iconTooltip';
            tooltip.className = 'icon-tooltip';
            document.body.appendChild(tooltip);
        }
        
        // Set content and position
        tooltip.textContent = text;
        tooltip.style.left = (event.pageX - 50) + 'px';
        tooltip.style.top = (event.pageY + 10) + 'px';
        tooltip.classList.add('show');
        
        // Hide on next click anywhere
        setTimeout(() => {
            document.addEventListener('click', function hideTooltip() {
                tooltip.classList.remove('show');
                document.removeEventListener('click', hideTooltip);
            });
        }, 100);
    }
    
    // AJAX function to load popup data on demand
    function loadAndShowPopup(orgId, listType, title) {
        // Show loading message
        showPopup(title, '<div class="loading-spinner"></div> Loading...');
        
        // Prepare form data
        var formData = new FormData();
        formData.append('action', 'get_popup_data');
        formData.append('org_id', orgId);
        formData.append('list_type', listType);
        
        // Make AJAX request
        fetch(getPyScriptFormAddress(), {
            method: 'POST',
            body: formData
        })
        .then(response => response.text())
        .then(data => {
            // Update popup with received data
            showPopup(title, data);
        })
        .catch(error => {
            showPopup(title, '<p>Error loading data: ' + error + '</p>');
        });
    }
    
    // Show performance indicator
    document.addEventListener('DOMContentLoaded', function() {
        const loadTime = (new Date() - queryStartTime) / 1000;
        if (loadTime > 5) {
            const perfIndicator = document.createElement('div');
            perfIndicator.className = 'performance-indicator';
            perfIndicator.style.display = 'block';
            perfIndicator.textContent = 'Page load: ' + loadTime.toFixed(2) + 's';
            document.body.appendChild(perfIndicator);
            
            setTimeout(() => {
                perfIndicator.style.display = 'none';
            }, 5000);
        }
    });
    </script>
    '''


def get_sidebar_javascript():
    """JavaScript for sidebar navigation interactions"""
    return '''
    <script>
    // Missions Dashboard Sidebar Controller
    var MissionsDashboard = {
        // State
        state: {
            sidebarCollapsed: false,
            mobileOpen: false,
            expandedTrips: [],
            appOrgsExpanded: false,
            tripsExpanded: true
        },

        // Initialize
        init: function() {
            this.setupContainer();
            this.loadState();
            this.applyState();
            this.bindKeyboardShortcuts();
        },

        // Set up the app container inside .box-content
        setupContainer: function() {
            var boxContent = document.querySelector('.box-content');
            var appContainer = document.querySelector('.app-container');

            if (boxContent && appContainer) {
                // Style box-content to contain our layout properly
                boxContent.style.padding = '0';
                boxContent.style.margin = '0';
                boxContent.style.position = 'relative';
                boxContent.style.overflow = 'visible';

                // Ensure app-container fills the space properly
                appContainer.style.width = '100%';

                console.log('Dashboard container set up inside .box-content');
            } else {
                console.log('Could not find .box-content or .app-container');
            }
        },

        // Load state from localStorage
        loadState: function() {
            try {
                var saved = localStorage.getItem('missionsDashboardState');
                if (saved) {
                    var parsed = JSON.parse(saved);
                    this.state.sidebarCollapsed = parsed.sidebarCollapsed || false;
                    this.state.expandedTrips = parsed.expandedTrips || [];
                    this.state.appOrgsExpanded = parsed.appOrgsExpanded !== undefined ? parsed.appOrgsExpanded : false;
                    this.state.tripsExpanded = parsed.tripsExpanded !== undefined ? parsed.tripsExpanded : true;
                }
            } catch (e) {
                console.log('Could not load sidebar state:', e);
            }
        },

        // Save state to localStorage
        saveState: function() {
            try {
                localStorage.setItem('missionsDashboardState', JSON.stringify({
                    sidebarCollapsed: this.state.sidebarCollapsed,
                    expandedTrips: this.state.expandedTrips
                }));
            } catch (e) {
                console.log('Could not save sidebar state:', e);
            }
        },

        // Apply saved state to DOM
        applyState: function() {
            var sidebar = document.getElementById('navSidebar');
            if (!sidebar) return;

            // Apply collapsed state
            if (this.state.sidebarCollapsed) {
                sidebar.classList.add('collapsed');
            }

            // Apply expanded trips
            var self = this;
            this.state.expandedTrips.forEach(function(tripId) {
                var tripItem = document.querySelector('.trip-item[data-trip-id="' + tripId + '"]');
                if (tripItem) {
                    tripItem.classList.add('expanded');
                }
            });

            // Apply app orgs state (only if not viewing an app org)
            var appList = document.querySelector('.app-orgs-list');
            var appIcon = document.querySelector('.app-orgs-header .expand-icon');
            if (appList && appIcon) {
                // Check if an app org is currently selected (has active class on li)
                var hasActiveAppOrg = appList.querySelector('.app-org-item.active');
                if (hasActiveAppOrg) {
                    // Keep expanded if viewing an app org
                    appList.style.display = 'block';
                    appIcon.classList.add('expanded');
                } else if (this.state.appOrgsExpanded) {
                    appList.style.display = 'block';
                    appIcon.classList.add('expanded');
                } else {
                    appList.style.display = 'none';
                    appIcon.classList.remove('expanded');
                }
            }

            // Apply trips section state (only if not viewing a trip)
            var tripsList = document.querySelector('.trips-section .trip-list');
            var tripsIcon = document.querySelector('.trips-header .expand-icon');
            if (tripsList && tripsIcon) {
                // Check if a trip is currently selected (has active class)
                var hasActiveTrip = tripsList.querySelector('.trip-item.expanded');
                if (hasActiveTrip) {
                    // Keep expanded if viewing a trip
                    tripsList.style.display = 'block';
                    tripsIcon.classList.add('expanded');
                } else if (this.state.tripsExpanded) {
                    tripsList.style.display = 'block';
                    tripsIcon.classList.add('expanded');
                } else {
                    tripsList.style.display = 'none';
                    tripsIcon.classList.remove('expanded');
                }
            }
        },

        // Toggle sidebar collapsed state
        toggleSidebar: function() {
            var sidebar = document.getElementById('navSidebar');
            if (!sidebar) return;

            this.state.sidebarCollapsed = !this.state.sidebarCollapsed;
            sidebar.classList.toggle('collapsed');

            // Update toggle button icon
            var toggleBtn = sidebar.querySelector('.sidebar-toggle');
            if (toggleBtn) {
                toggleBtn.innerHTML = this.state.sidebarCollapsed ? '&#9654;' : '&#9664;';
            }

            this.saveState();
        },

        // Toggle mobile sidebar
        toggleMobile: function() {
            var sidebar = document.getElementById('navSidebar');
            if (!sidebar) return;

            this.state.mobileOpen = !this.state.mobileOpen;
            sidebar.classList.toggle('mobile-open');
        },

        // Close mobile sidebar
        closeMobile: function() {
            var sidebar = document.getElementById('navSidebar');
            if (!sidebar) return;

            this.state.mobileOpen = false;
            sidebar.classList.remove('mobile-open');
        },

        // Toggle trip expansion
        toggleTrip: function(tripId) {
            var tripItem = document.querySelector('.trip-item[data-trip-id="' + tripId + '"]');
            if (!tripItem) return;

            var isExpanded = tripItem.classList.contains('expanded');

            if (isExpanded) {
                // Collapse this trip
                tripItem.classList.remove('expanded');
                var idx = this.state.expandedTrips.indexOf(tripId);
                if (idx > -1) {
                    this.state.expandedTrips.splice(idx, 1);
                }
            } else {
                // Expand this trip
                tripItem.classList.add('expanded');
                if (this.state.expandedTrips.indexOf(tripId) === -1) {
                    this.state.expandedTrips.push(tripId);
                }
            }

            this.saveState();
        },

        // Load a trip section via AJAX or navigation
        loadSection: function(tripId, section, event) {
            // For now, use standard navigation
            // In phase 2, this can be converted to AJAX loading

            // Update URL without preventing default navigation
            // The href on the link handles the actual navigation
        },

        // Keyboard shortcuts
        bindKeyboardShortcuts: function() {
            var self = this;
            document.addEventListener('keydown', function(e) {
                // Escape key closes mobile sidebar
                if (e.key === 'Escape' && self.state.mobileOpen) {
                    self.closeMobile();
                }
                // Ctrl+B toggles sidebar (like VS Code)
                if (e.ctrlKey && e.key === 'b') {
                    e.preventDefault();
                    self.toggleSidebar();
                }
            });
        },

        // Navigate to a trip section
        navigateToSection: function(tripId, section) {
            window.location.href = '?trip=' + tripId + '&section=' + section;
        },

        // Navigate to admin view
        navigateToView: function(view) {
            window.location.href = '?view=' + view;
        },

        // Toggle application orgs section collapse/expand
        toggleAppOrgs: function() {
            var list = document.querySelector('.app-orgs-list');
            var icon = document.querySelector('.app-orgs-header .expand-icon');

            if (!list || !icon) return;

            var isExpanded = list.style.display !== 'none';

            if (isExpanded) {
                // Collapse
                list.style.display = 'none';
                icon.classList.remove('expanded');
            } else {
                // Expand
                list.style.display = 'block';
                icon.classList.add('expanded');
            }

            // Save state to localStorage
            try {
                var saved = JSON.parse(localStorage.getItem('missionsDashboardState') || '{}');
                saved.appOrgsExpanded = !isExpanded;
                localStorage.setItem('missionsDashboardState', JSON.stringify(saved));
            } catch (e) {
                console.log('Could not save app orgs state:', e);
            }
        },

        // Toggle mission trips section collapse/expand
        toggleTripsSection: function() {
            var list = document.querySelector('.trips-section .trip-list');
            var icon = document.querySelector('.trips-header .expand-icon');

            if (!list || !icon) return;

            var isExpanded = list.style.display !== 'none';

            if (isExpanded) {
                // Collapse
                list.style.display = 'none';
                icon.classList.remove('expanded');
            } else {
                // Expand
                list.style.display = 'block';
                icon.classList.add('expanded');
            }

            // Save state to localStorage
            try {
                var saved = JSON.parse(localStorage.getItem('missionsDashboardState') || '{}');
                saved.tripsExpanded = !isExpanded;
                localStorage.setItem('missionsDashboardState', JSON.stringify(saved));
            } catch (e) {
                console.log('Could not save trips section state:', e);
            }
        }
    };

    // Actions Dropdown Controller
    var ActionsDropdown = {
        toggle: function(btn) {
            var dropdown = btn.nextElementSibling;
            var isOpen = dropdown.classList.contains('show');

            // Close all other dropdowns first
            this.closeAll();

            // Toggle this dropdown
            if (!isOpen) {
                dropdown.classList.add('show');

                // Position the dropdown using fixed positioning
                var btnRect = btn.getBoundingClientRect();
                var dropdownHeight = dropdown.offsetHeight || 200; // Estimate if not yet rendered
                var viewportHeight = window.innerHeight;
                var spaceBelow = viewportHeight - btnRect.bottom;
                var spaceAbove = btnRect.top;

                // Position horizontally - align right edge with button right edge
                dropdown.style.right = (window.innerWidth - btnRect.right) + 'px';
                dropdown.style.left = 'auto';

                // Position vertically - prefer below, but go above if not enough space
                if (spaceBelow >= dropdownHeight || spaceBelow >= spaceAbove) {
                    // Position below the button
                    dropdown.style.top = btnRect.bottom + 'px';
                    dropdown.style.bottom = 'auto';
                } else {
                    // Position above the button (dropup)
                    dropdown.style.bottom = (viewportHeight - btnRect.top) + 'px';
                    dropdown.style.top = 'auto';
                }
            }
        },

        closeAll: function() {
            var dropdowns = document.querySelectorAll('.actions-dropdown-content');
            for (var i = 0; i < dropdowns.length; i++) {
                dropdowns[i].classList.remove('show');
                // Reset positioning
                dropdowns[i].style.top = '';
                dropdowns[i].style.bottom = '';
                dropdowns[i].style.left = '';
                dropdowns[i].style.right = '';
            }
        }
    };

    // Member Role Controller - Set member type (Leader/Member)
    var MemberRole = {
        set: function(peopleId, orgId, memberType) {
            if (!confirm('Set this person as ' + memberType + '?')) return;

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'set_member_type',
                    people_id: peopleId,
                    org_id: orgId,
                    mem_role: memberType
                },
                success: function(response) {
                    try {
                        var data = typeof response === 'string' ? JSON.parse(response) : response;
                        if (data.success) {
                            // Reload to reflect the change
                            window.location.reload();
                        } else {
                            alert('Error: ' + (data.message || 'Unknown error'));
                        }
                    } catch(e) {
                        alert('Error processing response');
                    }
                },
                error: function(xhr, status, error) {
                    alert('Failed to update member type: ' + error);
                }
            });
        }
    };

    // Close dropdown when clicking outside
    document.addEventListener('click', function(e) {
        if (!e.target.closest('.actions-dropdown')) {
            ActionsDropdown.closeAll();
        }
    });

    // Initialize on DOM ready
    document.addEventListener('DOMContentLoaded', function() {
        MissionsDashboard.init();
    });
    </script>
    '''


def get_email_javascript(is_admin=False):
    """JavaScript for email popup functionality

    Args:
        is_admin: Boolean indicating if user is an admin (controls which templates are visible)
    """
    admin_value = 'true' if is_admin else 'false'
    return '''
    <script>
    // Missions Dashboard Email Controller
    var MissionsEmail = {
        currentPeopleId: null,
        currentEmail: null,
        currentEmails: [],  // For team emails
        currentName: null,
        isTeamEmail: false,
        currentContext: 'general',  // Context for template filtering: general, team, budget, meetings, passport
        currentTripName: '',
        currentMemberData: {},  // For individual template placeholders
        isAdmin: ''' + admin_value + ''',  // Set by Python based on user role

        // =====================================================================
        // EMAIL TEMPLATES - Managed via Settings > Quick Email Templates
        // =====================================================================
        // Templates are loaded dynamically from the server.
        // To manage templates, go to Settings tab and use the Email Templates section.
        //
        // PLACEHOLDERS available in templates:
        //   {{TripName}}     - Name of the trip
        //   {{PersonName}}   - Recipient's first name
        //   {{TripCost}}     - Individual's trip cost (individual emails only)
        //   {{Outstanding}}  - Amount still needed (individual emails only)
        //   {{MyGivingLink}} - Link to their giving page (replaced by TouchPoint)
        //   {{SupportLink}}  - Support link (replaced by TouchPoint)
        // =====================================================================
        templates: [],
        templatesLoaded: false,

        // Load templates from server (called automatically when needed)
        loadTemplatesFromServer: function(callback) {
            var self = this;

            // If already loaded, just call callback
            if (this.templatesLoaded && this.templates.length > 0) {
                console.log('Templates already loaded, using cache:', this.templates.length);
                if (callback) callback();
                return;
            }

            // Use PyScriptForm URL for AJAX
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            console.log('Loading templates from:', ajaxUrl);

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'get_dropdown_templates'
                },
                success: function(response) {
                    console.log('Template AJAX response:', response);
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success && result.templates) {
                            self.templates = result.templates;
                            self.templatesLoaded = true;
                            console.log('Templates loaded successfully:', self.templates.length);
                        } else {
                            console.error('Template response not successful:', result);
                        }
                    } catch (e) {
                        console.error('Error parsing templates:', e, response);
                    }
                    if (callback) callback();
                },
                error: function(xhr, status, error) {
                    console.error('Error loading templates:', error, xhr.responseText);
                    if (callback) callback();
                }
            });
        },

        // Load templates into dropdown based on current context and user role
        loadTemplates: function() {
            var self = this;

            // If templates haven't been loaded from server yet, load them first
            if (!this.templatesLoaded) {
                this.loadTemplatesFromServer(function() {
                    self.populateTemplateDropdown();
                });
            } else {
                this.populateTemplateDropdown();
            }
        },

        // Populate the template dropdown (called after templates are loaded)
        populateTemplateDropdown: function() {
            var select = document.getElementById('emailTemplate');
            if (!select) return;

            // Clear existing options
            select.innerHTML = '<option value="">-- Select a template (optional) --</option>';

            var self = this;
            var emailType = this.isTeamEmail ? 'team' : 'individual';

            console.log('Populating dropdown. Templates:', this.templates.length, 'Context:', this.currentContext, 'EmailType:', emailType, 'IsAdmin:', this.isAdmin);

            // Filter templates by context, type, and role
            this.templates.forEach(function(template) {
                // Default missing fields for backwards compatibility
                var tContext = template.context || 'all';
                var tType = template.type || 'both';
                var tRole = template.role || 'all';

                // Check if template matches current context
                var contextMatch = tContext === 'all' || tContext === self.currentContext;
                // Check if template matches email type (team or individual)
                var typeMatch = tType === 'both' || tType === emailType;
                // Check if user has permission based on role
                // role: 'all' = everyone, 'admin' = admins only, 'leader' = leaders only
                var roleMatch = tRole === 'all' ||
                               (tRole === 'admin' && self.isAdmin) ||
                               (tRole === 'leader' && !self.isAdmin);

                console.log('Template:', template.name, '| context:', tContext, contextMatch, '| type:', tType, typeMatch, '| role:', tRole, roleMatch);

                if (contextMatch && typeMatch && roleMatch) {
                    var option = document.createElement('option');
                    option.value = template.id;
                    option.textContent = template.name;
                    select.appendChild(option);
                }
            });
        },

        // Apply selected template to subject and body
        applyTemplate: function() {
            var select = document.getElementById('emailTemplate');
            if (!select || !select.value) return;

            var templateId = select.value;
            var template = null;

            // Find the template
            for (var i = 0; i < this.templates.length; i++) {
                if (this.templates[i].id === templateId) {
                    template = this.templates[i];
                    break;
                }
            }

            if (!template) return;

            // Replace placeholders in subject and body
            var subject = this.replacePlaceholders(template.subject);
            var body = this.replacePlaceholders(template.body);

            // Set subject and body
            document.getElementById('emailSubject').value = subject;
            this.setBodyContent(body, false);
        },

        // Replace placeholders with actual values
        replacePlaceholders: function(text) {
            var result = text;

            // Church URL and Trip name
            result = result.replace(/\\{\\{ChurchUrl\\}\\}/g, window.location.origin);
            result = result.replace(/\\{\\{TripName\\}\\}/g, this.currentTripName || 'the mission trip');

            // Person name (first name if individual, 'Team' if team email)
            if (this.isTeamEmail) {
                result = result.replace(/\\{\\{PersonName\\}\\}/g, '{{PersonName}}');  // Keep placeholder for server-side replacement
            } else {
                var firstName = this.currentName || 'there';
                if (firstName.indexOf(',') !== -1) {
                    // Handle "Last, First" format
                    firstName = firstName.split(',')[1].trim().split(' ')[0];
                }
                result = result.replace(/\\{\\{PersonName\\}\\}/g, firstName);
            }

            // Financial placeholders
            if (this.currentMemberData.tripCost) {
                result = result.replace(/\\{\\{TripCost\\}\\}/g, this.currentMemberData.tripCost);
            }
            if (this.currentMemberData.outstanding) {
                result = result.replace(/\\{\\{Outstanding\\}\\}/g, this.currentMemberData.outstanding);
            }

            // Keep server-side placeholders intact
            // {{MyGivingLink}} and {{SupportLink}} are replaced by TouchPoint when sending

            return result;
        },

        // Get the PyScriptForm address for AJAX calls
        getPyScriptFormAddress: function() {
            var pathname = window.location.pathname;
            // Replace /PyScript/ with /PyScriptForm/ if present
            if (pathname.indexOf('/PyScript/') !== -1) {
                return pathname.replace('/PyScript/', '/PyScriptForm/');
            }
            return pathname;
        },

        // Format text in the email body editor
        formatText: function(command) {
            document.execCommand(command, false, null);
            document.getElementById('emailBody').focus();
        },

        // Insert a link in the email body
        insertLink: function() {
            var url = prompt('Enter URL:');
            if (url) {
                // Ensure URL has protocol
                if (url && !url.match(/^https?:\\/\\//i)) {
                    url = 'https://' + url;
                }
                document.execCommand('createLink', false, url);
                document.getElementById('emailBody').focus();
            }
        },

        // Set the email body content (handles both text and HTML)
        setBodyContent: function(content, isHtml) {
            var body = document.getElementById('emailBody');
            if (isHtml) {
                body.innerHTML = content;
            } else {
                // Convert plain text to HTML (preserve line breaks)
                // Handle both actual newlines and literal \\n strings from stored templates
                var html = content
                    .replace(/\\\\n/g, '<br>')  // Literal \\n (escaped in storage)
                    .replace(/\\n/g, '<br>');   // Actual newline characters
                body.innerHTML = html;
            }
        },

        // Get the email body content as HTML
        getBodyContent: function() {
            return document.getElementById('emailBody').innerHTML;
        },

        // Update visibility of the advanced link based on whether we have an org ID
        updateAdvancedLink: function() {
            var advancedLink = document.getElementById('emailAdvancedLink');
            if (advancedLink) {
                advancedLink.style.display = this.currentOrgId ? 'inline-block' : 'none';
            }
        },

        // Open TouchPoint's built-in involvement email tools
        openAdvanced: function() {
            if (this.currentOrgId) {
                window.open('/Org/' + this.currentOrgId + '#tab-Members-tab', '_blank');
            } else {
                alert('No organization context available for advanced options.');
            }
        },

        // Open email modal for an individual
        openIndividual: function(peopleId, email, name, orgId, tripName, context) {
            this.currentPeopleId = peopleId;
            this.currentEmail = email;
            this.currentEmails = [email];
            this.currentName = name;
            this.isTeamEmail = false;
            this.currentOrgId = orgId || null;  // Optional org context
            this.currentTripName = tripName || '';
            this.currentContext = context || 'general';
            this.currentMemberData = {};

            // Update modal fields
            document.getElementById('emailModalTitle').textContent = 'Email ' + name;
            document.getElementById('emailTo').value = email || 'No email on file';
            document.getElementById('emailSubject').value = '';
            document.getElementById('emailBody').innerHTML = '';

            // Reset and load templates
            var templateSelect = document.getElementById('emailTemplate');
            if (templateSelect) templateSelect.value = '';
            this.loadTemplates();

            // Show/hide advanced link based on org context
            this.updateAdvancedLink();

            // Show modal
            document.getElementById('emailModal').classList.add('active');

            // Focus subject field
            setTimeout(function() {
                document.getElementById('emailSubject').focus();
            }, 100);
        },

        // Open email modal for the whole team
        openTeam: function(orgId, emailsString, tripName, context) {
            var emails = emailsString ? emailsString.split(',') : [];
            this.currentPeopleId = null;
            this.currentEmail = null;
            this.currentEmails = emails;
            this.currentName = 'Team';
            this.isTeamEmail = true;
            this.currentOrgId = orgId;
            this.currentTripName = tripName || '';
            this.currentContext = context || 'general';
            this.currentMemberData = {};

            if (emails.length === 0) {
                alert('No team members have email addresses on file.');
                return;
            }

            // Update modal fields
            document.getElementById('emailModalTitle').textContent = 'Email Team (' + emails.length + ' recipients)';
            document.getElementById('emailTo').value = emails.length + ' team members';
            document.getElementById('emailSubject').value = '';
            document.getElementById('emailBody').innerHTML = '';

            // Reset and load templates
            var templateSelect = document.getElementById('emailTemplate');
            if (templateSelect) templateSelect.value = '';
            this.loadTemplates();

            // Show/hide advanced link
            this.updateAdvancedLink();

            // Show modal
            document.getElementById('emailModal').classList.add('active');

            // Focus subject field
            setTimeout(function() {
                document.getElementById('emailSubject').focus();
            }, 100);
        },

        // Open email modal with goal reminder template (for all team)
        openGoalReminder: function(orgId, emailsString, tripName) {
            var emails = emailsString ? emailsString.split(',') : [];
            this.currentPeopleId = null;
            this.currentEmail = null;
            this.currentEmails = emails;
            this.currentName = 'Team';
            this.isTeamEmail = true;
            this.currentOrgId = orgId;
            this.currentTripName = tripName || '';
            this.currentContext = 'budget';
            this.currentMemberData = {};

            if (emails.length === 0) {
                alert('No team members have email addresses on file.');
                return;
            }

            // Pre-fill with goal reminder template
            // Uses personalized placeholders: {{PersonName}}, {{MyGivingLink}}, {{SupportLink}}
            var subject = tripName + ' - Fundraising Goal Reminder';
            var body = 'Hi {{PersonName}},\\n\\n' +
                'This is a friendly reminder about your fundraising goals for ' + tripName + '.\\n\\n' +
                'Please check your current fundraising status and reach out if you need any help meeting your goal.\\n\\n' +
                'View your payment status here:\\n{{MyGivingLink}}\\n\\n' +
                'Share this link with friends and family who want to support your trip:\\n{{SupportLink}}\\n\\n' +
                'Thank you for your commitment to this mission trip!\\n\\n' +
                'Blessings';

            // Update modal fields
            document.getElementById('emailModalTitle').textContent = 'Send Goal Reminder (' + emails.length + ' recipients)';
            document.getElementById('emailTo').value = emails.length + ' team members';
            document.getElementById('emailSubject').value = subject;
            this.setBodyContent(body, false);

            // Load templates (allows user to switch to different template)
            var templateSelect = document.getElementById('emailTemplate');
            if (templateSelect) templateSelect.value = '';
            this.loadTemplates();

            // Show/hide advanced link
            this.updateAdvancedLink();

            // Show modal
            document.getElementById('emailModal').classList.add('active');

            // Focus body field so they can customize
            setTimeout(function() {
                document.getElementById('emailBody').focus();
            }, 100);
        },

        // Open email modal with individual goal reminder (for one person)
        openIndividualGoalReminder: function(peopleId, email, memberName, tripName, tripCost, outstanding, orgId) {
            this.currentPeopleId = peopleId;
            this.currentEmail = email;
            this.currentEmails = [email];
            this.currentName = memberName;
            this.isTeamEmail = false;
            this.currentOrgId = orgId || null;  // Optional org context
            this.currentTripName = tripName || '';
            this.currentContext = 'budget';

            // Format currency values for display
            var costFormatted = '$' + parseFloat(tripCost).toFixed(2).replace(/\\d(?=(\\d{3})+\\.)/g, '$&,');
            var outstandingFormatted = '$' + parseFloat(outstanding).toFixed(2).replace(/\\d(?=(\\d{3})+\\.)/g, '$&,');
            this.currentMemberData = {
                tripCost: costFormatted,
                outstanding: outstandingFormatted
            };

            if (!email) {
                alert('No email address on file for ' + memberName + '.');
                return;
            }

            // Pre-fill with personalized goal reminder template
            // Uses personalized placeholders: {{PersonName}}, {{MyGivingLink}}, {{SupportLink}}
            var subject = tripName + ' - Fundraising Goal Reminder';
            var body = 'Hi ' + memberName.split(',')[0] + ',\\n\\n' +
                'This is a friendly reminder about your fundraising goal for ' + tripName + '.\\n\\n' +
                'Your trip cost: ' + costFormatted + '\\n' +
                'Amount still needed: ' + outstandingFormatted + '\\n\\n' +
                'View your payment status here:\\n{{MyGivingLink}}\\n\\n' +
                'Share this link with friends and family who want to support your trip:\\n{{SupportLink}}\\n\\n' +
                'Please reach out if you need any help meeting your goal.\\n\\n' +
                'Thank you for your commitment to this mission trip!\\n\\n' +
                'Blessings';

            // Update modal fields
            document.getElementById('emailModalTitle').textContent = 'Send Goal Reminder to ' + memberName;
            document.getElementById('emailTo').value = email;
            document.getElementById('emailSubject').value = subject;
            this.setBodyContent(body, false);

            // Load templates (allows user to switch to different template)
            var templateSelect = document.getElementById('emailTemplate');
            if (templateSelect) templateSelect.value = '';
            this.loadTemplates();

            // Update advanced link visibility (hidden for individual emails)
            this.updateAdvancedLink();

            // Show modal
            document.getElementById('emailModal').classList.add('active');

            // Focus body field so they can customize
            setTimeout(function() {
                document.getElementById('emailBody').focus();
            }, 100);
        },

        // Open email modal with passport request for an individual
        openPassportRequest: function(peopleId, email, memberName, tripName, orgId) {
            this.currentPeopleId = peopleId;
            this.currentEmail = email;
            this.currentEmails = [email];
            this.currentName = memberName;
            this.isTeamEmail = false;
            this.isPassportRequest = true;
            this.currentOrgId = orgId;
            this.currentTripName = tripName || '';
            this.currentContext = 'team';
            this.currentMemberData = {};

            if (!email) {
                alert('No email address on file for ' + memberName + '.');
                return;
            }

            // Get first name from "Last, First" format
            var firstName = memberName;
            if (memberName.indexOf(',') !== -1) {
                firstName = memberName.split(',')[1].trim().split(' ')[0];
            }

            // Pre-fill with passport request template
            var subject = tripName + ' - Passport Information Needed';
            var body = 'Hi ' + firstName + ',\\n\\n' +
                'We need your passport information for the upcoming ' + tripName + ' mission trip.\\n\\n' +
                'Please complete the passport form at the link below:\\n\\n' +
                CHURCH_URL + '/OnlineReg/3421\\n\\n' +
                'This form will collect your passport number, expiration date, and other travel details we need for trip planning.\\n\\n' +
                'Please complete this as soon as possible so we can finalize travel arrangements.\\n\\n' +
                'Thank you!\\n\\n' +
                'Blessings';

            // Update modal fields
            document.getElementById('emailModalTitle').textContent = 'Request Passport Info from ' + memberName;
            document.getElementById('emailTo').value = email;
            document.getElementById('emailSubject').value = subject;
            this.setBodyContent(body, false);

            // Load templates (allows user to switch to different template)
            var templateSelect = document.getElementById('emailTemplate');
            if (templateSelect) templateSelect.value = '';
            this.loadTemplates();

            // Show advanced link (has org context)
            this.updateAdvancedLink();

            // Show modal
            document.getElementById('emailModal').classList.add('active');

            // Focus body field so they can customize
            setTimeout(function() {
                document.getElementById('emailBody').focus();
            }, 100);
        },

        // Open email modal with meeting reminder template
        openMeetingReminder: function(orgId, emailsString, tripName, meetingDate, meetingTime, meetingDescription, meetingLocation) {
            var emails = emailsString ? emailsString.split(',') : [];
            this.currentPeopleId = null;
            this.currentEmail = null;
            this.currentEmails = emails;
            this.currentName = 'Team';
            this.isTeamEmail = true;
            this.currentOrgId = orgId;
            this.currentTripName = tripName || '';
            this.currentContext = 'meetings';
            this.currentMemberData = {};

            if (emails.length === 0) {
                alert('No team members have email addresses on file.');
                return;
            }

            // Format meeting details
            var description = meetingDescription || 'Team Meeting';
            var location = meetingLocation ? '\\n\\nLocation: ' + meetingLocation : '';

            // Pre-fill with meeting reminder template
            var subject = tripName + ' - Meeting Reminder: ' + description;
            var body = 'Hi Team,\\n\\n' +
                'This is a reminder about our upcoming meeting:\\n\\n' +
                '📅 ' + description + '\\n' +
                '🕐 ' + meetingDate + ' at ' + meetingTime + location + '\\n\\n' +
                'Please make sure to attend. If you cannot make it, please let us know as soon as possible.\\n\\n' +
                'See you there!\\n\\n' +
                'Blessings';

            // Update modal fields
            document.getElementById('emailModalTitle').textContent = 'Meeting Reminder (' + emails.length + ' recipients)';
            document.getElementById('emailTo').value = emails.length + ' team members';
            document.getElementById('emailSubject').value = subject;
            this.setBodyContent(body, false);

            // Load templates (allows user to switch to different template)
            var templateSelect = document.getElementById('emailTemplate');
            if (templateSelect) templateSelect.value = '';
            this.loadTemplates();

            // Show/hide advanced link
            this.updateAdvancedLink();

            // Show modal
            document.getElementById('emailModal').classList.add('active');

            // Focus body field so they can customize
            setTimeout(function() {
                document.getElementById('emailBody').focus();
            }, 100);
        },

        // Close email modal
        close: function() {
            document.getElementById('emailModal').classList.remove('active');
            this.currentPeopleId = null;
            this.currentEmail = null;
            this.currentEmails = [];
            this.currentName = null;
            this.isTeamEmail = false;
            this.isPassportRequest = false;
            this.currentOrgId = null;
            this.currentTripName = '';
            this.currentContext = 'general';
            this.currentMemberData = {};
        },

        // Send email - opens email client with pre-filled data
        send: function() {
            var subject = document.getElementById('emailSubject').value;
            var body = this.getBodyContent();

            if (this.currentEmails.length === 0) {
                alert('No email addresses available.');
                return;
            }

            if (!subject.trim()) {
                alert('Please enter a subject.');
                document.getElementById('emailSubject').focus();
                return;
            }

            // Check if body is empty (strip HTML tags and whitespace)
            var bodyText = body.replace(/<[^>]*>/g, '').trim();
            if (!bodyText) {
                alert('Please enter a message.');
                document.getElementById('emailBody').focus();
                return;
            }

            // Disable the send button and show loading state
            var sendBtn = document.querySelector('.email-send-btn');
            var originalText = sendBtn ? sendBtn.innerHTML : 'Send';
            if (sendBtn) {
                sendBtn.disabled = true;
                sendBtn.innerHTML = 'Sending...';
            }

            // Send via TouchPoint's email system using AJAX
            var self = this;
            var formData = new FormData();
            formData.append('action', 'send_email');
            formData.append('to_emails', this.currentEmails.join(','));
            formData.append('subject', subject);
            // IMPORTANT: Encode HTML to avoid ASP.NET request validation blocking
            // HTML tags in POST data are blocked as "potentially dangerous"
            var encodedBody = body.replace(/</g, '&lt;').replace(/>/g, '&gt;')
                                  .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            formData.append('body', encodedBody);
            if (this.currentPeopleId) {
                formData.append('people_id', this.currentPeopleId);
            }
            // Always pass org_id for personalized link placeholders
            if (this.currentOrgId) {
                formData.append('org_id', this.currentOrgId);
            }
            if (this.isPassportRequest) {
                formData.append('is_passport_request', 'true');
            }

            fetch(this.getPyScriptFormAddress(), {
                method: 'POST',
                body: formData
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                if (data.success) {
                    self.close();
                } else {
                    alert('Error: ' + (data.message || 'Failed to send email.'));
                }
            })
            .catch(function(error) {
                console.error('Email error:', error);
                alert('Error sending email. Please try again.');
            })
            .finally(function() {
                // Re-enable the send button
                if (sendBtn) {
                    sendBtn.disabled = false;
                    sendBtn.innerHTML = originalText;
                }
            });
        },

        // Handle escape key to close modal
        handleKeydown: function(e) {
            if (e.key === 'Escape') {
                var modal = document.getElementById('emailModal');
                if (modal && modal.classList.contains('active')) {
                    MissionsEmail.close();
                }
            }
        }
    };

    // Initialize email handlers
    document.addEventListener('keydown', MissionsEmail.handleKeydown);

    // Close modal when clicking outside
    document.getElementById('emailModal').addEventListener('click', function(e) {
        if (e.target === this) {
            MissionsEmail.close();
        }
    });

    </script>
    '''


def get_fee_adjustment_javascript():
    """JavaScript for fee adjustment functionality"""
    return '''
    <script>
    // Missions Dashboard Fee Adjustment Controller
    var MissionsFee = {
        currentPeopleId: null,
        currentOrgId: null,
        currentName: null,
        currentCost: 0,

        // Open fee adjustment modal for a member
        openAdjust: function(peopleId, orgId, memberName, currentCost) {
            this.currentPeopleId = peopleId;
            this.currentOrgId = orgId;
            this.currentName = memberName;
            this.currentCost = currentCost || 0;

            // Update modal fields
            document.getElementById('feeModalTitle').textContent = 'Adjust Fee - ' + memberName;
            document.getElementById('feeCurrentCost').textContent = '$' + this.formatMoney(currentCost);
            document.getElementById('feeAdjustAmount').value = '';
            document.getElementById('feeDescription').value = '';

            // Show modal
            document.getElementById('feeModal').classList.add('active');

            // Focus amount field
            setTimeout(function() {
                document.getElementById('feeAdjustAmount').focus();
            }, 100);
        },

        // Close fee adjustment modal
        close: function() {
            document.getElementById('feeModal').classList.remove('active');
            this.currentPeopleId = null;
            this.currentOrgId = null;
            this.currentName = null;
            this.currentCost = 0;
        },

        // Format number as money
        formatMoney: function(amount) {
            return parseFloat(amount).toFixed(2).replace(/\\d(?=(\\d{3})+\\.)/g, '$&,');
        },

        // Submit fee adjustment via AJAX
        submit: function() {
            var amount = document.getElementById('feeAdjustAmount').value;
            var description = document.getElementById('feeDescription').value;

            if (!amount || isNaN(parseFloat(amount))) {
                alert('Please enter a valid adjustment amount.');
                document.getElementById('feeAdjustAmount').focus();
                return;
            }

            var amountValue = parseFloat(amount);
            if (amountValue === 0) {
                alert('Adjustment amount cannot be zero.');
                document.getElementById('feeAdjustAmount').focus();
                return;
            }

            if (!description.trim()) {
                description = 'Fee adjustment from Mission Dashboard';
            }

            // Show loading state
            var submitBtn = document.querySelector('#feeModal .fee-btn-submit');
            var originalText = submitBtn.textContent;
            submitBtn.textContent = 'Processing...';
            submitBtn.disabled = true;

            // Get script name from URL
            var pathname = window.location.pathname;
            var scriptName = 'Mission_Dashboard';
            var parts = pathname.split('/');
            for (var i = 0; i < parts.length; i++) {
                if (parts[i] === 'PyScriptForm' || parts[i] === 'PyScript') {
                    if (i + 1 < parts.length) {
                        scriptName = parts[i + 1].split('?')[0];
                        break;
                    }
                }
            }

            // Send AJAX request
            var self = this;
            var xhr = new XMLHttpRequest();
            xhr.open('POST', '/PyScriptForm/' + scriptName, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    submitBtn.textContent = originalText;
                    submitBtn.disabled = false;

                    if (xhr.status === 200) {
                        try {
                            var response = JSON.parse(xhr.responseText);
                            if (response.success) {
                                self.close();
                                // Reload the current section to show updated values
                                if (typeof MissionsDashboardSidebar !== 'undefined') {
                                    MissionsDashboardSidebar.loadSection(self.currentOrgId, 'budget');
                                } else {
                                    // Fallback: reload page
                                    window.location.reload();
                                }
                            } else {
                                alert('Error: ' + (response.message || 'Unknown error'));
                            }
                        } catch (e) {
                            alert('Error processing response. Please try again.');
                            console.error('Parse error:', e, xhr.responseText);
                        }
                    } else {
                        alert('Request failed. Please try again.');
                    }
                }
            };

            var data = 'ajax=true&action=adjust_fee' +
                       '&people_id=' + encodeURIComponent(self.currentPeopleId) +
                       '&org_id=' + encodeURIComponent(self.currentOrgId) +
                       '&amount=' + encodeURIComponent(amountValue) +
                       '&description=' + encodeURIComponent(description);

            xhr.send(data);
        },

        // Handle escape key to close modal
        handleKeydown: function(e) {
            if (e.key === 'Escape') {
                var modal = document.getElementById('feeModal');
                if (modal && modal.classList.contains('active')) {
                    MissionsFee.close();
                }
            }
        }
    };

    // Initialize fee adjustment handlers when DOM is ready
    document.addEventListener('DOMContentLoaded', function() {
        var feeModal = document.getElementById('feeModal');
        if (feeModal) {
            document.addEventListener('keydown', MissionsFee.handleKeydown);
            feeModal.addEventListener('click', function(e) {
                if (e.target === this) {
                    MissionsFee.close();
                }
            });
        }
    });

    </script>
    '''


def get_dates_javascript():
    """JavaScript for trip dates editing functionality"""
    return '''
    <script>
    // Missions Dashboard Dates Controller
    var MissionsDates = {
        currentOrgId: null,
        tripName: null,

        // Open date editing modal
        openEdit: function(orgId, tripName, startDate, endDate, closeDate) {
            this.currentOrgId = orgId;
            this.tripName = tripName;

            // Update modal title
            document.getElementById('datesModalTitle').textContent = 'Edit Trip Dates - ' + tripName;

            // Set current values (format for date input: YYYY-MM-DD)
            document.getElementById('dateEventStart').value = this.formatDateForInput(startDate);
            document.getElementById('dateEventEnd').value = this.formatDateForInput(endDate);
            document.getElementById('dateClose').value = this.formatDateForInput(closeDate);

            // Show modal
            document.getElementById('datesModal').classList.add('active');

            // Focus first date field
            setTimeout(function() {
                document.getElementById('dateEventStart').focus();
            }, 100);
        },

        // Format date string for input element (YYYY-MM-DD)
        formatDateForInput: function(dateStr) {
            if (!dateStr || dateStr === '' || dateStr === 'None' || dateStr === 'N/A') {
                return '';
            }
            // Try to parse various date formats
            try {
                var d = new Date(dateStr);
                if (isNaN(d.getTime())) return '';
                var year = d.getFullYear();
                var month = ('0' + (d.getMonth() + 1)).slice(-2);
                var day = ('0' + d.getDate()).slice(-2);
                return year + '-' + month + '-' + day;
            } catch (e) {
                return '';
            }
        },

        // Close modal
        close: function() {
            document.getElementById('datesModal').classList.remove('active');
            this.currentOrgId = null;
            this.tripName = null;
        },

        // Submit date changes via AJAX
        submit: function() {
            var startDate = document.getElementById('dateEventStart').value;
            var endDate = document.getElementById('dateEventEnd').value;
            var closeDate = document.getElementById('dateClose').value;

            // Validate at least one date is set
            if (!startDate && !endDate && !closeDate) {
                alert('Please enter at least one date.');
                return;
            }

            // Show loading state
            var submitBtn = document.querySelector('#datesModal .dates-btn-submit');
            var originalText = submitBtn.textContent;
            submitBtn.textContent = 'Saving...';
            submitBtn.disabled = true;

            // Get script name from URL
            var pathname = window.location.pathname;
            var scriptName = 'Mission_Dashboard';
            var parts = pathname.split('/');
            for (var i = 0; i < parts.length; i++) {
                if (parts[i] === 'PyScriptForm' || parts[i] === 'PyScript') {
                    if (i + 1 < parts.length) {
                        scriptName = parts[i + 1].split('?')[0];
                        break;
                    }
                }
            }

            // Send AJAX request
            var self = this;
            var xhr = new XMLHttpRequest();
            xhr.open('POST', '/PyScriptForm/' + scriptName, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    submitBtn.textContent = originalText;
                    submitBtn.disabled = false;

                    if (xhr.status === 200) {
                        try {
                            // Extract JSON from response (may have extra content from TouchPoint API)
                            var responseText = xhr.responseText;
                            var jsonStart = responseText.indexOf('{');
                            var jsonEnd = responseText.lastIndexOf('}');
                            var jsonStr = (jsonStart >= 0 && jsonEnd > jsonStart)
                                ? responseText.substring(jsonStart, jsonEnd + 1)
                                : responseText;
                            var response = JSON.parse(jsonStr);
                            if (response.success) {
                                self.close();
                                // Reload the overview section
                                if (typeof MissionsDashboardSidebar !== 'undefined') {
                                    MissionsDashboardSidebar.loadSection(self.currentOrgId, 'overview');
                                } else {
                                    window.location.reload();
                                }
                            } else {
                                alert('Error: ' + (response.message || 'Unknown error'));
                            }
                        } catch (e) {
                            // Check if dates might have been updated anyway
                            if (xhr.responseText.indexOf('"success": true') >= 0 || xhr.responseText.indexOf('"success":true') >= 0) {
                                self.close();
                                if (typeof MissionsDashboardSidebar !== 'undefined') {
                                    MissionsDashboardSidebar.loadSection(self.currentOrgId, 'overview');
                                } else {
                                    window.location.reload();
                                }
                            } else {
                                alert('Error processing response. Please try again.');
                                console.error('Parse error:', e, xhr.responseText);
                            }
                        }
                    } else {
                        alert('Request failed. Please try again.');
                    }
                }
            };

            var data = 'ajax=true&action=update_dates' +
                       '&org_id=' + encodeURIComponent(self.currentOrgId) +
                       '&start_date=' + encodeURIComponent(startDate) +
                       '&end_date=' + encodeURIComponent(endDate) +
                       '&close_date=' + encodeURIComponent(closeDate);

            xhr.send(data);
        },

        // Handle escape key
        handleKeydown: function(e) {
            if (e.key === 'Escape') {
                var modal = document.getElementById('datesModal');
                if (modal && modal.classList.contains('active')) {
                    MissionsDates.close();
                }
            }
        }
    };

    // Initialize date editing handlers when DOM is ready
    document.addEventListener('DOMContentLoaded', function() {
        var datesModal = document.getElementById('datesModal');
        if (datesModal) {
            document.addEventListener('keydown', MissionsDates.handleKeydown);
            datesModal.addEventListener('click', function(e) {
                if (e.target === this) {
                    MissionsDates.close();
                }
            });
        }
    });

    </script>
    '''


def get_meeting_javascript():
    """JavaScript for creating meetings functionality"""
    return '''
    <script>
    // Missions Dashboard Meeting Controller
    var MissionsMeeting = {
        currentOrgId: null,
        tripName: null,
        editMode: false,
        currentMeetingId: null,

        // Open create meeting modal
        openCreate: function(orgId, tripName) {
            this.currentOrgId = orgId;
            this.tripName = tripName;
            this.editMode = false;
            this.currentMeetingId = null;

            // Update modal title and button
            document.getElementById('meetingModalTitle').textContent = 'Create Meeting - ' + tripName;
            document.querySelector('#meetingModal .meeting-btn-submit').textContent = 'Create Meeting';

            // Clear/reset all fields
            document.getElementById('meetingDescription').value = '';
            document.getElementById('meetingLocation').value = '';

            // Set default date to tomorrow
            var tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            var dateStr = tomorrow.toISOString().split('T')[0];
            document.getElementById('meetingDate').value = dateStr;

            // Set default time to 6:00 PM
            document.getElementById('meetingTime').value = '18:00';

            // Show modal
            document.getElementById('meetingModal').classList.add('active');

            // Focus description field
            setTimeout(function() {
                document.getElementById('meetingDescription').focus();
            }, 100);
        },

        // Open edit meeting modal with existing data
        openEdit: function(meetingId, orgId, tripName, dateStr, timeStr, description, location) {
            this.currentOrgId = orgId;
            this.tripName = tripName;
            this.editMode = true;
            this.currentMeetingId = meetingId;

            // Update modal title and button
            document.getElementById('meetingModalTitle').textContent = 'Edit Meeting - ' + tripName;
            document.querySelector('#meetingModal .meeting-btn-submit').textContent = 'Update Meeting';

            // Populate fields with existing data
            document.getElementById('meetingDescription').value = description || '';
            document.getElementById('meetingLocation').value = location || '';
            document.getElementById('meetingDate').value = dateStr || '';
            document.getElementById('meetingTime').value = timeStr || '18:00';

            // Show modal
            document.getElementById('meetingModal').classList.add('active');

            // Focus description field
            setTimeout(function() {
                document.getElementById('meetingDescription').focus();
            }, 100);
        },

        // Close modal
        close: function() {
            document.getElementById('meetingModal').classList.remove('active');
            this.currentOrgId = null;
            this.tripName = null;
            this.editMode = false;
            this.currentMeetingId = null;
        },

        // Submit meeting create/update via AJAX
        submit: function() {
            var meetingDescription = document.getElementById('meetingDescription').value;
            var meetingDate = document.getElementById('meetingDate').value;
            var meetingTime = document.getElementById('meetingTime').value;
            var meetingLocation = document.getElementById('meetingLocation').value;

            if (!meetingDate) {
                alert('Please select a date for the meeting.');
                return;
            }

            // Show loading state
            var submitBtn = document.querySelector('#meetingModal .meeting-btn-submit');
            var originalText = submitBtn.textContent;
            submitBtn.textContent = this.editMode ? 'Updating...' : 'Creating...';
            submitBtn.disabled = true;

            // Get script name from URL
            var pathname = window.location.pathname;
            var scriptName = 'Mission_Dashboard';
            var parts = pathname.split('/');
            for (var i = 0; i < parts.length; i++) {
                if (parts[i] === 'PyScriptForm' || parts[i] === 'PyScript') {
                    if (i + 1 < parts.length) {
                        scriptName = parts[i + 1].split('?')[0];
                        break;
                    }
                }
            }

            // Send AJAX request
            var self = this;
            var xhr = new XMLHttpRequest();
            xhr.open('POST', '/PyScriptForm/' + scriptName, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    submitBtn.textContent = originalText;
                    submitBtn.disabled = false;

                    if (xhr.status === 200) {
                        try {
                            // Extract JSON from response
                            var responseText = xhr.responseText;
                            var jsonStart = responseText.indexOf('{');
                            var jsonEnd = responseText.lastIndexOf('}');
                            var jsonStr = (jsonStart >= 0 && jsonEnd > jsonStart)
                                ? responseText.substring(jsonStart, jsonEnd + 1)
                                : responseText;
                            var response = JSON.parse(jsonStr);
                            if (response.success) {
                                self.close();
                                // Reload the meetings section
                                if (typeof MissionsDashboardSidebar !== 'undefined' && self.currentOrgId) {
                                    MissionsDashboardSidebar.loadSection(self.currentOrgId, 'meetings');
                                } else {
                                    window.location.reload();
                                }
                            } else {
                                alert('Error: ' + (response.message || 'Unknown error'));
                            }
                        } catch (e) {
                            console.error('Parse error:', e, xhr.responseText);
                            alert('Error processing response. Please try again.');
                        }
                    } else {
                        alert('Network error. Please try again.');
                    }
                }
            };

            // Build POST data based on mode
            var postData;
            if (self.editMode) {
                postData = 'action=update_meeting&meeting_id=' + encodeURIComponent(self.currentMeetingId) +
                           '&org_id=' + encodeURIComponent(self.currentOrgId) +
                           '&meeting_date=' + encodeURIComponent(meetingDate) +
                           '&meeting_time=' + encodeURIComponent(meetingTime || '18:00') +
                           '&description=' + encodeURIComponent(meetingDescription || 'Team Meeting') +
                           '&location=' + encodeURIComponent(meetingLocation || '');
            } else {
                postData = 'action=create_meeting&org_id=' + encodeURIComponent(self.currentOrgId) +
                           '&meeting_date=' + encodeURIComponent(meetingDate) +
                           '&meeting_time=' + encodeURIComponent(meetingTime || '18:00') +
                           '&description=' + encodeURIComponent(meetingDescription || 'Team Meeting') +
                           '&location=' + encodeURIComponent(meetingLocation || '');
            }
            xhr.send(postData);
        },

        // Handle escape key
        handleKeydown: function(e) {
            if (e.key === 'Escape') {
                MissionsMeeting.close();
            }
        }
    };

    // Setup event listeners when DOM is ready
    document.addEventListener('DOMContentLoaded', function() {
        var meetingModal = document.getElementById('meetingModal');
        if (meetingModal) {
            // Close on escape key
            document.addEventListener('keydown', MissionsMeeting.handleKeydown);

            // Close on overlay click
            meetingModal.addEventListener('click', function(e) {
                if (e.target === this) {
                    MissionsMeeting.close();
                }
            });
        }
    });

    // Toggle meeting attendance visibility
    function toggleAttendance(meetingId) {
        var attendanceList = document.getElementById('attendance-' + meetingId);
        var arrow = document.getElementById('arrow-' + meetingId);

        if (attendanceList) {
            if (attendanceList.classList.contains('expanded')) {
                attendanceList.classList.remove('expanded');
                if (arrow) arrow.innerHTML = '&#9660;';  // Down arrow
            } else {
                attendanceList.classList.add('expanded');
                if (arrow) arrow.innerHTML = '&#9650;';  // Up arrow
            }
        }
    }

    </script>
    '''


def get_person_details_javascript():
    """JavaScript for person details popup modal with larger photo and more info"""
    return '''
    <style>
    /* Person Details Modal Styles */
    .person-modal-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.6);
        z-index: 10000;
        opacity: 0;
        transition: opacity 0.3s ease;
    }
    .person-modal-overlay.active {
        display: flex;
        align-items: center;
        justify-content: center;
        opacity: 1;
    }
    .person-modal {
        background: white;
        border-radius: 12px;
        width: 90%;
        max-width: 500px;
        max-height: 90vh;
        overflow-y: auto;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        transform: scale(0.9);
        transition: transform 0.3s ease;
    }
    .person-modal-overlay.active .person-modal {
        transform: scale(1);
    }
    .person-modal-header {
        position: relative;
        background: linear-gradient(135deg, #003366 0%, #4A90E2 100%);
        color: white;
        padding: 30px 20px;
        text-align: center;
        border-radius: 12px 12px 0 0;
    }
    .person-modal-close {
        position: absolute;
        top: 10px;
        right: 15px;
        font-size: 24px;
        cursor: pointer;
        opacity: 0.8;
        background: none;
        border: none;
        color: white;
    }
    .person-modal-close:hover {
        opacity: 1;
    }
    .person-modal-photo {
        width: 120px;
        height: 120px;
        border-radius: 50%;
        object-fit: cover;
        border: 4px solid white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        margin-bottom: 15px;
    }
    .person-modal-photo-placeholder {
        width: 120px;
        height: 120px;
        border-radius: 50%;
        background: rgba(255,255,255,0.2);
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 48px;
        font-weight: bold;
        border: 4px solid white;
        margin: 0 auto 15px;
    }
    .person-modal-name {
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 5px;
    }
    .person-modal-role {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    .person-modal-body {
        padding: 20px;
    }
    .person-detail-section {
        margin-bottom: 20px;
    }
    .person-detail-section h4 {
        font-size: 0.8rem;
        text-transform: uppercase;
        color: #666;
        letter-spacing: 1px;
        margin-bottom: 10px;
        border-bottom: 1px solid #eee;
        padding-bottom: 5px;
    }
    .person-detail-row {
        display: flex;
        justify-content: space-between;
        padding: 8px 0;
        border-bottom: 1px solid #f0f0f0;
    }
    .person-detail-row:last-child {
        border-bottom: none;
    }
    .person-detail-label {
        color: #666;
        font-size: 0.9rem;
    }
    .person-detail-value {
        font-weight: 500;
        color: #333;
    }
    .person-detail-value a {
        color: #4A90E2;
        text-decoration: none;
    }
    .person-detail-value a:hover {
        text-decoration: underline;
    }
    .person-trips-list {
        list-style: none;
        padding: 0;
        margin: 0;
    }
    .person-trips-list li {
        padding: 10px;
        background: #f8f9fa;
        border-radius: 6px;
        margin-bottom: 8px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .person-trips-list .trip-name {
        font-weight: 500;
    }
    .person-trips-list .trip-date {
        font-size: 0.85rem;
        color: #666;
    }
    .person-modal-actions {
        display: flex;
        gap: 10px;
        padding: 15px 20px;
        border-top: 1px solid #eee;
        background: #f8f9fa;
        border-radius: 0 0 12px 12px;
    }
    .person-modal-btn {
        flex: 1;
        padding: 10px;
        border: none;
        border-radius: 6px;
        cursor: pointer;
        font-weight: 500;
        text-align: center;
        text-decoration: none;
        display: inline-block;
    }
    .person-modal-btn-primary {
        background: #4A90E2;
        color: white;
    }
    .person-modal-btn-secondary {
        background: #e9ecef;
        color: #333;
    }
    .person-modal-btn:hover {
        opacity: 0.9;
    }
    .person-modal-loading {
        text-align: center;
        padding: 40px;
        color: #666;
    }
    .person-modal-loading .spinner {
        width: 40px;
        height: 40px;
        border: 3px solid #e9ecef;
        border-top: 3px solid #4A90E2;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin: 0 auto 15px;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    </style>

    <!-- Person Details Modal -->
    <div id="personModal" class="person-modal-overlay">
        <div class="person-modal">
            <div id="personModalContent">
                <!-- Content loaded dynamically -->
            </div>
        </div>
    </div>

    <script>
    // Person Details Popup Controller
    var PersonDetails = {
        // Get the PyScriptForm address dynamically
        getPyScriptFormAddress: function() {
            var path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        },

        // Open modal and load person details
        open: function(peopleId, orgId) {
            var modal = document.getElementById('personModal');
            var content = document.getElementById('personModalContent');

            // Show loading state
            content.innerHTML = '<div class="person-modal-loading"><div class="spinner"></div><p>Loading details...</p></div>';
            modal.classList.add('active');

            // Fetch person details via AJAX
            var formData = new FormData();
            formData.append('action', 'get_person_details');
            formData.append('people_id', peopleId);
            formData.append('org_id', orgId || '');

            fetch(this.getPyScriptFormAddress(), {
                method: 'POST',
                body: formData
            })
            .then(function(response) { return response.text(); })
            .then(function(html) {
                content.innerHTML = html;
            })
            .catch(function(error) {
                content.innerHTML = '<div class="person-modal-loading"><p style="color: red;">Error loading details: ' + error + '</p></div>';
            });
        },

        // Close modal
        close: function() {
            var modal = document.getElementById('personModal');
            modal.classList.remove('active');
        },

        // Handle escape key
        handleKeydown: function(e) {
            if (e.key === 'Escape') {
                PersonDetails.close();
            }
        }
    };

    // Initialize person modal handlers when DOM is ready
    document.addEventListener('DOMContentLoaded', function() {
        var personModal = document.getElementById('personModal');
        if (personModal) {
            document.addEventListener('keydown', PersonDetails.handleKeydown);
            personModal.addEventListener('click', function(e) {
                if (e.target === this) {
                    PersonDetails.close();
                }
            });
        }
    });
    </script>
    '''


def get_approval_workflow_javascript():
    """JavaScript for trip registration approval workflow modal"""
    return '''
    <style>
    /* Approval Workflow Styles */
    .approval-filter-tabs {
        display: flex;
        gap: 8px;
        margin-bottom: 15px;
        flex-wrap: wrap;
    }
    .approval-filter-tab {
        padding: 8px 16px;
        border: 1px solid #dee2e6;
        border-radius: 20px;
        background: white;
        cursor: pointer;
        font-size: 0.9rem;
        transition: all 0.2s ease;
    }
    .approval-filter-tab:hover {
        border-color: #667eea;
        background: #f8f9ff;
    }
    .approval-filter-tab.active {
        background: #667eea;
        color: white;
        border-color: #667eea;
    }
    .approval-filter-tab .filter-count {
        font-weight: 600;
    }
    .approval-filter-tab[data-filter="pending"] .filter-count {
        color: #fd7e14;
    }
    .approval-filter-tab.active[data-filter="pending"] .filter-count {
        color: white;
    }
    .approval-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 500;
    }
    .approval-status-badge.approved {
        background: #d4edda;
        color: #155724;
    }
    .approval-status-badge.denied {
        background: #f8d7da;
        color: #721c24;
    }
    .approval-status-badge.pending {
        background: #fff3cd;
        color: #856404;
    }
    .btn-approve-small, .btn-deny-small, .btn-revoke-small {
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.75rem;
        cursor: pointer;
        border: none;
        transition: all 0.2s ease;
        margin-left: 5px;
    }
    .btn-approve-small {
        background: #667eea;
        color: white;
    }
    .btn-approve-small:hover {
        background: #5a6fd6;
    }
    .btn-deny-small {
        background: #dc3545;
        color: white;
    }
    .btn-deny-small:hover {
        background: #c82333;
    }
    .btn-revoke-small {
        background: #6c757d;
        color: white;
    }
    .btn-revoke-small:hover {
        background: #5a6268;
    }
    /* Approval Modal Styles */
    .approval-modal-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.6);
        z-index: 10001;
        opacity: 0;
        transition: opacity 0.3s ease;
    }
    .approval-modal-overlay.active {
        display: flex;
        align-items: center;
        justify-content: center;
        opacity: 1;
    }
    .approval-modal {
        background: white;
        border-radius: 12px;
        width: 95%;
        max-width: 600px;
        max-height: 90vh;
        overflow-y: auto;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        transform: scale(0.9);
        transition: transform 0.3s ease;
    }
    .approval-modal-overlay.active .approval-modal {
        transform: scale(1);
    }
    .approval-modal-header {
        background: linear-gradient(135deg, #003366 0%, #4A90E2 100%);
        color: white;
        padding: 20px;
        border-radius: 12px 12px 0 0;
        position: relative;
    }
    .approval-modal-header h3 {
        margin: 0;
        font-size: 1.5rem;
    }
    .approval-modal-header .modal-subtitle {
        opacity: 0.9;
        font-size: 1rem;
        margin-top: 5px;
    }
    .approval-modal-close {
        position: absolute;
        top: 15px;
        right: 15px;
        font-size: 24px;
        cursor: pointer;
        opacity: 0.8;
        background: none;
        border: none;
        color: white;
    }
    .approval-modal-close:hover {
        opacity: 1;
    }
    .approval-modal-body {
        padding: 20px;
        font-size: 12px;
    }
    .approval-section {
        margin-bottom: 20px;
    }
    .approval-section h4 {
        font-size: 12px;
        color: #495057;
        margin-bottom: 10px;
        text-transform: uppercase;
        font-weight: 600;
        border-bottom: 1px solid #e9ecef;
        padding-bottom: 5px;
    }
    .person-info-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 12px;
    }
    .person-info-item {
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        text-align: left;
    }
    .person-info-item .label {
        font-size: 10px;
        color: #6c757d;
        text-transform: uppercase;
        text-align: left;
        margin-bottom: 2px;
    }
    .person-info-item .value {
        font-size: 12px;
        text-align: left;
        color: #212529;
    }
    .registration-qa-list {
        max-height: 350px;
        overflow-y: auto;
        background: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
    }
    .registration-qa-item {
        margin-bottom: 15px;
        padding-bottom: 15px;
        border-bottom: 1px solid #e9ecef;
    }
    .registration-qa-item:last-child {
        margin-bottom: 0;
        padding-bottom: 0;
        border-bottom: none;
    }
    .registration-qa-item .question {
        font-weight: 600;
        color: #495057;
        font-size: 12px;
        margin-bottom: 5px;
    }
    .registration-qa-item .answer {
        color: #212529;
        font-size: 12px;
        padding-left: 10px;
        border-left: 3px solid #667eea;
        line-height: 1.5;
    }
    .denial-reason-input {
        width: 100%;
        min-height: 80px;
        padding: 10px;
        border: 1px solid #ced4da;
        border-radius: 6px;
        font-size: 12px;
        resize: vertical;
    }
    .denial-reason-input:focus {
        outline: none;
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102,126,234,0.15);
    }
    .approval-modal-footer {
        display: flex;
        gap: 10px;
        padding: 20px;
        background: #f8f9fa;
        border-radius: 0 0 12px 12px;
        justify-content: flex-end;
    }
    .btn-approve, .btn-deny, .btn-cancel {
        padding: 10px 20px;
        border-radius: 6px;
        font-size: 0.95rem;
        font-weight: 500;
        cursor: pointer;
        border: none;
        transition: all 0.2s ease;
    }
    .btn-approve {
        background: #28a745;
        color: white;
    }
    .btn-approve:hover {
        background: #218838;
    }
    .btn-deny {
        background: #dc3545;
        color: white;
    }
    .btn-deny:hover {
        background: #c82333;
    }
    .btn-cancel {
        background: #6c757d;
        color: white;
    }
    .btn-cancel:hover {
        background: #5a6268;
    }
    .approval-loading {
        text-align: center;
        padding: 40px;
        color: #6c757d;
    }
    .current-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 5px 12px;
        border-radius: 15px;
        font-size: 0.85rem;
        font-weight: 500;
        margin-top: 5px;
    }
    .denial-history {
        background: #fff3cd;
        border: 1px solid #ffc107;
        border-radius: 6px;
        padding: 10px;
        margin-top: 10px;
    }
    .denial-history .denial-label {
        font-weight: 600;
        color: #856404;
        font-size: 0.85rem;
    }
    .denial-history .denial-reason-text {
        color: #856404;
        font-size: 0.9rem;
        margin-top: 5px;
    }
    .denial-history .denial-meta {
        color: #856404;
        font-size: 0.75rem;
        margin-top: 5px;
        opacity: 0.8;
    }
    </style>

    <!-- Approval Modal HTML -->
    <div id="approval-modal-overlay" class="approval-modal-overlay">
        <div class="approval-modal">
            <div class="approval-modal-header">
                <button class="approval-modal-close" onclick="ApprovalWorkflow.closeModal()">&times;</button>
                <h3 id="approval-modal-title">Review Registration</h3>
                <div class="modal-subtitle" id="approval-modal-subtitle"></div>
            </div>
            <div class="approval-modal-body">
                <div id="approval-modal-content">
                    <div class="approval-loading">
                        Loading registration details...
                    </div>
                </div>
            </div>
            <div class="approval-modal-footer" id="approval-modal-footer">
                <button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>
                <button class="btn-deny" id="btn-deny-registration" onclick="ApprovalWorkflow.showDenyForm()">Deny</button>
                <button class="btn-approve" id="btn-approve-registration" onclick="ApprovalWorkflow.approve()">Approve</button>
            </div>
        </div>
    </div>

    <script>
    // Approval Workflow Controller
    var ApprovalWorkflow = {
        currentPeopleId: null,
        currentOrgId: null,
        currentPersonName: null,
        currentStatus: null,
        denialReason: '',
        showApprovalButtons: true,

        // Initialize the workflow
        init: function(orgId) {
            this.currentOrgId = orgId;
        },

        // Filter table rows by approval status
        filterTable: function(filter) {
            // Update active tab
            var tabs = document.querySelectorAll('.approval-filter-tab');
            for (var i = 0; i < tabs.length; i++) {
                tabs[i].classList.remove('active');
                if (tabs[i].getAttribute('data-filter') === filter) {
                    tabs[i].classList.add('active');
                }
            }

            // Filter rows
            var rows = document.querySelectorAll('tr[data-status]');
            for (var j = 0; j < rows.length; j++) {
                var row = rows[j];
                var status = row.getAttribute('data-status');
                if (filter === 'all' || status === filter) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            }
        },

        // Show the approval modal for a person
        // showApprovalBtns parameter controls whether Approve/Deny buttons are shown
        showModal: function(peopleId, personName, orgId, showApprovalBtns) {
            this.currentPeopleId = peopleId;
            this.currentOrgId = orgId;
            this.currentPersonName = personName;
            this.denialReason = '';
            this.showApprovalButtons = (showApprovalBtns !== false); // Default to true

            // Update modal title
            document.getElementById('approval-modal-title').textContent = 'Review Registration';
            document.getElementById('approval-modal-subtitle').textContent = personName;

            // Show loading state
            document.getElementById('approval-modal-content').innerHTML = '<div class="approval-loading">Loading registration details...</div>';

            // Reset footer buttons - conditionally show Approve/Deny based on showApprovalButtons
            if (this.showApprovalButtons) {
                document.getElementById('approval-modal-footer').innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>' +
                    '<button class="btn-deny" id="btn-deny-registration" onclick="ApprovalWorkflow.showDenyForm()">Deny</button>' +
                    '<button class="btn-approve" id="btn-approve-registration" onclick="ApprovalWorkflow.approve()">Approve</button>';
            } else {
                document.getElementById('approval-modal-footer').innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Close</button>';
            }

            // Show modal
            var overlay = document.getElementById('approval-modal-overlay');
            overlay.style.display = 'flex';
            setTimeout(function() {
                overlay.classList.add('active');
            }, 10);

            // Load registration data
            this.loadRegistrationData(peopleId, orgId);
        },

        // Load registration data via AJAX
        loadRegistrationData: function(peopleId, orgId) {
            var self = this;
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'get_registration_answers',
                    people_id: peopleId,
                    org_id: orgId
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.currentStatus = result.status;
                            self.renderRegistrationData(result);
                        } else {
                            document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error loading registration data: ' + (result.error || 'Unknown error') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error parsing response</div>';
                        console.error('Parse error:', e);
                    }
                },
                error: function(xhr, status, error) {
                    document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error loading registration data: ' + error + '</div>';
                }
            });
        },

        // Render the registration data in the modal
        renderRegistrationData: function(data) {
            var html = '';
            var person = data.person || {};

            // Person info section
            html += '<div class="approval-section">';
            html += '<h4>Person Information</h4>';
            html += '<div class="person-info-grid">';
            html += '<div class="person-info-item"><span class="label">Name</span><span class="value">' + this.escapeHtml(person.name || 'N/A') + '</span></div>';
            // Email - clickable to open email popup
            if (person.email) {
                var escapedEmail = this.escapeJsString(person.email);
                var escapedName = this.escapeJsString(person.name || '');
                html += '<div class="person-info-item"><span class="label">Email</span><span class="value"><a href="#" onclick="ApprovalWorkflow.closeModal(); MissionsEmail.openIndividual(' + person.people_id + ', \\'' + escapedEmail + '\\', \\'' + escapedName + '\\', ' + this.currentOrgId + '); return false;" style="color: #007bff; text-decoration: none;">' + this.escapeHtml(person.email) + '</a></span></div>';
            } else {
                html += '<div class="person-info-item"><span class="label">Email</span><span class="value">N/A</span></div>';
            }
            // Phone - clickable tel: link
            if (person.phone) {
                var phoneDigits = person.phone.replace(/[^0-9]/g, '');
                html += '<div class="person-info-item"><span class="label">Phone</span><span class="value"><a href="tel:' + phoneDigits + '" style="color: #007bff; text-decoration: none;">' + this.escapeHtml(person.phone) + '</a></span></div>';
            } else {
                html += '<div class="person-info-item"><span class="label">Phone</span><span class="value">N/A</span></div>';
            }
            html += '<div class="person-info-item"><span class="label">Age</span><span class="value">' + (person.age || 'N/A') + '</span></div>';
            html += '</div>';

            // Current status - only show if approval workflow is enabled
            if (this.showApprovalButtons) {
                html += '<div style="margin-top: 10px;">';
                html += '<span class="label">Current Status: </span>';
                if (data.status === 'approved') {
                    html += '<span class="current-status-badge approved" style="background:#d4edda;color:#155724;">&#10003; Approved</span>';
                } else if (data.status === 'denied') {
                    html += '<span class="current-status-badge denied" style="background:#f8d7da;color:#721c24;">&#10007; Denied</span>';
                } else {
                    html += '<span class="current-status-badge pending" style="background:#fff3cd;color:#856404;">&#8987; Pending</span>';
                }
                html += '</div>';

                // Show denial reason if exists
                if (data.status === 'denied' && data.denial_reason) {
                    html += '<div class="denial-history">';
                    html += '<div class="denial-label">Previous Denial Reason:</div>';
                    html += '<div class="denial-reason-text">' + this.escapeHtml(data.denial_reason) + '</div>';
                    if (data.denied_by_name || data.denied_date) {
                        html += '<div class="denial-meta">';
                        if (data.denied_by_name) html += 'By: ' + this.escapeHtml(data.denied_by_name);
                        if (data.denied_date) html += ' on ' + data.denied_date;
                        html += '</div>';
                    }
                    html += '</div>';
                }
            }

            html += '</div>';

            // Registration Q&A section
            html += '<div class="approval-section">';
            html += '<h4>Registration Answers</h4>';
            if (data.answers && data.answers.length > 0) {
                html += '<div class="registration-qa-list">';
                for (var i = 0; i < data.answers.length; i++) {
                    var qa = data.answers[i];
                    html += '<div class="registration-qa-item">';
                    html += '<div class="question">' + this.escapeHtml(qa.question || 'Question ' + (i + 1)) + '</div>';
                    html += '<div class="answer">' + this.escapeHtml(qa.answer || 'No answer provided') + '</div>';
                    html += '</div>';
                }
                html += '</div>';
            } else {
                html += '<div style="color: #6c757d; font-style: italic;">No registration questions found for this trip.</div>';
            }
            html += '</div>';

            document.getElementById('approval-modal-content').innerHTML = html;

            // Update buttons based on status
            this.updateFooterButtons(data.status);
        },

        // Update footer buttons based on current status
        updateFooterButtons: function(status) {
            var footer = document.getElementById('approval-modal-footer');

            // If approval buttons not enabled, just show Close button
            if (!this.showApprovalButtons) {
                footer.innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Close</button>';
                return;
            }

            var html = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>';

            if (status === 'approved') {
                html += '<button class="btn-deny" onclick="ApprovalWorkflow.revoke()">Revoke Approval</button>';
            } else if (status === 'denied') {
                html += '<button class="btn-approve" onclick="ApprovalWorkflow.approve()">Approve</button>';
                html += '<button class="btn-deny" onclick="ApprovalWorkflow.revoke()">Reset to Pending</button>';
            } else {
                html += '<button class="btn-deny" id="btn-deny-registration" onclick="ApprovalWorkflow.showDenyForm()">Deny</button>';
                html += '<button class="btn-approve" id="btn-approve-registration" onclick="ApprovalWorkflow.approve()">Approve</button>';
            }

            footer.innerHTML = html;
        },

        // Show the denial reason form
        showDenyForm: function() {
            var content = document.getElementById('approval-modal-content');
            var existingHtml = content.innerHTML;

            // Add denial reason textarea
            var denyFormHtml = '<div class="approval-section" id="denial-reason-section">' +
                '<h4>Denial Reason</h4>' +
                '<textarea id="denial-reason-textarea" class="denial-reason-input" placeholder="Please provide a reason for denying this registration...">' + this.escapeHtml(this.denialReason) + '</textarea>' +
                '</div>';

            content.innerHTML = existingHtml + denyFormHtml;

            // Update buttons
            var footer = document.getElementById('approval-modal-footer');
            footer.innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>' +
                '<button class="btn-approve" onclick="ApprovalWorkflow.showModal(' + this.currentPeopleId + ', \\'' + this.escapeJsString(this.currentPersonName) + '\\', ' + this.currentOrgId + ', ' + this.showApprovalButtons + ')">Back</button>' +
                '<button class="btn-deny" onclick="ApprovalWorkflow.submitDenial()">Confirm Denial</button>';

            // Focus textarea
            document.getElementById('denial-reason-textarea').focus();
        },

        // Submit approval
        approve: function() {
            var self = this;
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            // Show loading
            document.getElementById('approval-modal-content').innerHTML = '<div class="approval-loading">Processing approval...</div>';

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'approve_member',
                    people_id: self.currentPeopleId,
                    org_id: self.currentOrgId
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            // Check if we should prompt to send approval notification email
                            if (result.prompt_email && result.person && result.person.email) {
                                ApprovalWorkflow.showApprovalEmailPrompt(result);
                            } else {
                                // Reload the page to show updated status
                                window.location.reload();
                            }
                        } else {
                            document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + (result.error || 'Unknown error') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error processing response</div>';
                    }
                },
                error: function(xhr, status, error) {
                    document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + error + '</div>';
                }
            });
        },

        // Show approval email prompt with editable email form
        showApprovalEmailPrompt: function(approvalData) {
            var self = this;
            var person = approvalData.person;
            var trip = approvalData.trip;
            var depositStr = trip.deposit > 0 ? '$' + trip.deposit.toLocaleString() : '[Deposit Amount]';
            var costStr = trip.cost > 0 ? '$' + trip.cost.toLocaleString() : '[Trip Cost]';

            // Store the approval data for the email
            ApprovalWorkflow.pendingApprovalEmail = {
                person: person,
                trip: trip
            };

            // Load the approval notification template from MissionsEmail templates
            var defaultSubject = trip.name + ' - Registration Approved!';
            var defaultBody = 'Dear ' + person.first_name + ',\\n\\n' +
                'Great news! Your registration for ' + trip.name + ' has been approved!\\n\\n' +
                'TRIP DETAILS:\\n' +
                '- Trip Cost: ' + costStr + '\\n' +
                '- Required Deposit: ' + depositStr + '\\n\\n' +
                'NEXT STEPS:\\n' +
                '1. Pay your deposit as soon as possible to secure your spot\\n' +
                '2. Start fundraising for the remaining balance\\n' +
                '3. Share your personal fundraising page with friends and family\\n\\n' +
                'PAYMENT OPTIONS:\\n' +
                'View your payment status and make payments here:\\n' +
                '{{MyGivingLink}}\\n\\n' +
                'FUNDRAISING:\\n' +
                'Share this link with supporters who want to help fund your trip:\\n' +
                '{{SupportLink}}\\n\\n' +
                'If you have any questions, please do not hesitate to reach out.\\n\\n' +
                'We are excited to have you on this mission trip!\\n\\n' +
                'Blessings';

            // Try to load the approval_notification template
            if (typeof MissionsEmail !== 'undefined' && MissionsEmail.templates && MissionsEmail.templates.length > 0) {
                for (var i = 0; i < MissionsEmail.templates.length; i++) {
                    if (MissionsEmail.templates[i].id === 'approval_notification') {
                        var tpl = MissionsEmail.templates[i];
                        // Replace placeholders
                        defaultSubject = (tpl.subject || defaultSubject)
                            .replace(/\\{\\{TripName\\}\\}/g, trip.name)
                            .replace(/\\{\\{PersonName\\}\\}/g, person.first_name);
                        defaultBody = (tpl.body || defaultBody)
                            .replace(/\\{\\{TripName\\}\\}/g, trip.name)
                            .replace(/\\{\\{PersonName\\}\\}/g, person.first_name)
                            .replace(/\\{\\{TripCost\\}\\}/g, costStr)
                            .replace(/\\{\\{DepositAmount\\}\\}/g, depositStr);
                        break;
                    }
                }
            }

            // Convert \\n to actual newlines for textarea display
            var bodyForTextarea = defaultBody.replace(/\\\\n/g, '\\n');

            // Create the approval email modal with editable form
            var modalHtml = '<div id="approval-email-prompt-modal" class="email-modal-overlay active">' +
                '<div class="email-modal" style="max-width: 650px;">' +
                    '<div class="email-modal-header" style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white;">' +
                        '<h3 style="margin: 0;">&#10003; Member Approved - Send Notification</h3>' +
                        '<button class="email-close-btn" onclick="ApprovalWorkflow.closeApprovalEmailPrompt()" style="color: white;">&times;</button>' +
                    '</div>' +
                    '<div class="email-modal-body" style="padding: 20px;">' +
                        '<div style="background: #d4edda; border: 1px solid #c3e6cb; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px;">' +
                            '<strong style="color: #155724;">' + person.name + '</strong> has been approved for <strong style="color: #155724;">' + trip.name + '</strong>' +
                        '</div>' +
                        '<div style="background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px;">' +
                            '<div style="display: flex; align-items: center; gap: 10px;">' +
                                '<span style="font-size: 20px;">&#128176;</span>' +
                                '<div>' +
                                    '<strong style="color: #856404;">Trip Fee to be Set:</strong> ' +
                                    '<span style="font-size: 18px; font-weight: bold; color: #856404;">' + costStr + '</span>' +
                                    (trip.deposit > 0 ? '<br><small style="color: #856404;">Deposit required: ' + depositStr + '</small>' : '') +
                                '</div>' +
                            '</div>' +
                        '</div>' +
                        '<div style="margin-bottom: 12px;">' +
                            '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">To:</label>' +
                            '<input type="text" id="approval-email-to" value="' + person.email + '" readonly ' +
                                'style="width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; background: #f8f9fa; box-sizing: border-box;">' +
                        '</div>' +
                        '<div style="margin-bottom: 12px;">' +
                            '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">Subject:</label>' +
                            '<input type="text" id="approval-email-subject" value="' + defaultSubject.replace(/"/g, '&quot;') + '" ' +
                                'style="width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; box-sizing: border-box;">' +
                        '</div>' +
                        '<div style="margin-bottom: 12px;">' +
                            '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">Message:</label>' +
                            '<textarea id="approval-email-body" rows="12" ' +
                                'style="width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-family: inherit; font-size: 14px; resize: vertical; box-sizing: border-box;">' + bodyForTextarea + '</textarea>' +
                            '<p style="margin: 6px 0 0 0; font-size: 12px; color: #888;">Tip: Edit the default template in Settings > Quick Email Templates > "Registration Approved"</p>' +
                        '</div>' +
                    '</div>' +
                    '<div class="email-modal-footer" style="display: flex; justify-content: flex-end; gap: 10px; padding: 15px 20px; background: #f8f9fa;">' +
                        '<button onclick="ApprovalWorkflow.closeApprovalEmailPrompt()" style="padding: 10px 20px; border: 1px solid #ddd; border-radius: 6px; background: white; cursor: pointer;">Skip (No Fee Set)</button>' +
                        '<button id="approval-send-btn" onclick="ApprovalWorkflow.sendApprovalEmail()" style="padding: 10px 20px; border: none; border-radius: 6px; background: #28a745; color: white; cursor: pointer;">&#128176; Set Fee &amp; Send Email</button>' +
                    '</div>' +
                '</div>' +
            '</div>';

            document.body.insertAdjacentHTML('beforeend', modalHtml);
            document.getElementById('approval-email-prompt-modal').addEventListener('click', function(e) {
                if (e.target === this) ApprovalWorkflow.closeApprovalEmailPrompt();
            });
        },

        closeApprovalEmailPrompt: function() {
            var modal = document.getElementById('approval-email-prompt-modal');
            if (modal) modal.remove();
            ApprovalWorkflow.pendingApprovalEmail = null;
            // Reload to show updated approval status
            window.location.reload();
        },

        sendApprovalEmail: function() {
            var data = ApprovalWorkflow.pendingApprovalEmail;
            if (!data) {
                window.location.reload();
                return;
            }

            var toEmail = document.getElementById('approval-email-to').value;
            var subject = document.getElementById('approval-email-subject').value;
            var body = document.getElementById('approval-email-body').value;

            if (!subject.trim()) {
                alert('Please enter a subject.');
                document.getElementById('approval-email-subject').focus();
                return;
            }

            if (!body.trim()) {
                alert('Please enter a message.');
                document.getElementById('approval-email-body').focus();
                return;
            }

            // Disable the send button and show loading
            var sendBtn = document.getElementById('approval-send-btn');
            if (sendBtn) {
                sendBtn.disabled = true;
                sendBtn.innerHTML = 'Setting fee...';
            }

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            // Encode body to avoid ASP.NET request validation blocking HTML
            var encodedBody = body.replace(/</g, '&lt;').replace(/>/g, '&gt;')
                                  .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

            // Step 1: First adjust the fee (trip cost)
            var tripCost = data.trip.cost || 0;

            function sendEmailAfterFee() {
                // Step 2: Now send the email
                if (sendBtn) sendBtn.innerHTML = 'Sending email...';

                $.ajax({
                    url: ajaxUrl,
                    type: 'POST',
                    data: {
                        action: 'send_email',
                        to_emails: toEmail,
                        subject: subject,
                        body: encodedBody,
                        people_id: data.person.people_id,
                        org_id: data.trip.org_id
                    },
                    success: function(response) {
                        try {
                            var result = typeof response === 'string' ? JSON.parse(response) : response;
                            if (result.success) {
                                // Close modal and reload
                                var modal = document.getElementById('approval-email-prompt-modal');
                                if (modal) modal.remove();
                                ApprovalWorkflow.pendingApprovalEmail = null;
                                window.location.reload();
                            } else {
                                alert('Fee was set but email failed: ' + (result.message || 'Failed to send email.'));
                                // Still reload since fee was set
                                window.location.reload();
                            }
                        } catch (e) {
                            alert('Fee was set. Email status unclear - please check sent emails.');
                            window.location.reload();
                        }
                    },
                    error: function(xhr, status, error) {
                        alert('Fee was set but email failed: ' + error);
                        // Still reload since fee was set
                        window.location.reload();
                    }
                });
            }

            // Only adjust fee if trip cost is greater than 0
            if (tripCost > 0) {
                $.ajax({
                    url: ajaxUrl,
                    type: 'POST',
                    data: {
                        ajax: 'true',
                        action: 'adjust_fee',
                        people_id: data.person.people_id,
                        org_id: data.trip.org_id,
                        amount: -tripCost,  // Negative to increase what they owe
                        description: 'Trip fee set upon approval'
                    },
                    success: function(response) {
                        try {
                            var result = typeof response === 'string' ? JSON.parse(response) : response;
                            if (result.success) {
                                // Fee set successfully, now send email
                                sendEmailAfterFee();
                            } else {
                                alert('Error setting fee: ' + (result.message || 'Unknown error'));
                                if (sendBtn) {
                                    sendBtn.disabled = false;
                                    sendBtn.innerHTML = '&#128176; Set Fee &amp; Send Email';
                                }
                            }
                        } catch (e) {
                            alert('Error processing fee response');
                            if (sendBtn) {
                                sendBtn.disabled = false;
                                sendBtn.innerHTML = '&#128176; Set Fee &amp; Send Email';
                            }
                        }
                    },
                    error: function(xhr, status, error) {
                        alert('Error setting fee: ' + error);
                        if (sendBtn) {
                            sendBtn.disabled = false;
                            sendBtn.innerHTML = '&#128176; Set Fee &amp; Send Email';
                        }
                    }
                });
            } else {
                // No trip cost, just send the email
                sendEmailAfterFee();
            }
        },

        pendingApprovalEmail: null,

        // Submit denial
        submitDenial: function() {
            var self = this;
            var reason = document.getElementById('denial-reason-textarea').value.trim();

            if (!reason) {
                alert('Please provide a reason for denial.');
                document.getElementById('denial-reason-textarea').focus();
                return;
            }

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            // Show loading
            document.getElementById('approval-modal-content').innerHTML = '<div class="approval-loading">Processing denial...</div>';

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'deny_member',
                    people_id: self.currentPeopleId,
                    org_id: self.currentOrgId,
                    reason: reason
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            window.location.reload();
                        } else {
                            document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + (result.error || 'Unknown error') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error processing response</div>';
                    }
                },
                error: function(xhr, status, error) {
                    document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + error + '</div>';
                }
            });
        },

        // Revoke approval (return to pending)
        revoke: function(peopleId, orgId) {
            var self = this;
            var targetPeopleId = peopleId || this.currentPeopleId;
            var targetOrgId = orgId || this.currentOrgId;

            if (!confirm('Are you sure you want to revoke this status and return to pending?')) {
                return;
            }

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'revoke_approval',
                    people_id: targetPeopleId,
                    org_id: targetOrgId
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            window.location.reload();
                        } else {
                            alert('Error: ' + (result.error || 'Unknown error'));
                        }
                    } catch (e) {
                        alert('Error processing response');
                    }
                },
                error: function(xhr, status, error) {
                    alert('Error: ' + error);
                }
            });
        },

        // Close the modal
        closeModal: function() {
            var overlay = document.getElementById('approval-modal-overlay');
            overlay.classList.remove('active');
            setTimeout(function() {
                overlay.style.display = 'none';
            }, 300);
        },

        // Helper: Escape HTML
        escapeHtml: function(text) {
            if (!text) return '';
            var div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        },

        // Helper: Escape JS string (uses fromCharCode to avoid Python/JS escaping issues)
        escapeJsString: function(str) {
            if (!str) return '';
            var bs = String.fromCharCode(92);  // backslash
            var sq = String.fromCharCode(39);  // single quote
            var dq = String.fromCharCode(34);  // double quote
            var result = str.split(bs).join(bs + bs);
            result = result.split(sq).join(bs + sq);
            result = result.split(dq).join(bs + dq);
            return result;
        }
    };

    // Trip Approval Settings Controller - Manages per-trip approval overrides
    var TripApprovalSettings = {
        currentOrgId: null,
        currentSetting: 'inherit',

        // Show the settings modal for a trip
        showModal: function(orgId, currentSetting) {
            this.currentOrgId = orgId;
            this.currentSetting = currentSetting || 'inherit';

            // Create modal if it doesn't exist
            if (!document.getElementById('trip-approval-settings-modal')) {
                this.createModal();
            }

            // Update radio selection
            var radios = document.querySelectorAll('input[name="tripApprovalSetting"]');
            radios.forEach(function(radio) {
                radio.checked = (radio.value === currentSetting);
            });

            // Show modal
            document.getElementById('trip-approval-settings-modal').classList.add('active');
        },

        // Create the modal HTML
        createModal: function() {
            var modalHtml = '<div id="trip-approval-settings-modal" class="email-modal-overlay">' +
                '<div class="email-modal" style="max-width: 450px;">' +
                    '<div class="email-modal-header" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">' +
                        '<h3 style="margin: 0;">&#9881; Trip Approval Settings</h3>' +
                        '<button class="email-close-btn" onclick="TripApprovalSettings.closeModal()" style="color: white;">&times;</button>' +
                    '</div>' +
                    '<div class="email-modal-body" style="padding: 20px;">' +
                        '<p style="margin-bottom: 15px; color: #666;">Choose how approvals should work for this trip:</p>' +
                        '<div style="display: flex; flex-direction: column; gap: 12px;">' +
                            '<label style="display: flex; align-items: flex-start; gap: 10px; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; cursor: pointer;">' +
                                '<input type="radio" name="tripApprovalSetting" value="inherit" style="margin-top: 3px;">' +
                                '<div>' +
                                    '<strong style="color: #333;">Use Global Setting</strong>' +
                                    '<p style="margin: 5px 0 0 0; font-size: 13px; color: #666;">Follow the global approval setting from Dashboard Settings</p>' +
                                '</div>' +
                            '</label>' +
                            '<label style="display: flex; align-items: flex-start; gap: 10px; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; cursor: pointer;">' +
                                '<input type="radio" name="tripApprovalSetting" value="enabled" style="margin-top: 3px;">' +
                                '<div>' +
                                    '<strong style="color: #4caf50;">&#10003; Enable Approvals</strong>' +
                                    '<p style="margin: 5px 0 0 0; font-size: 13px; color: #666;">Require approval for registrations (overrides global)</p>' +
                                '</div>' +
                            '</label>' +
                            '<label style="display: flex; align-items: flex-start; gap: 10px; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; cursor: pointer;">' +
                                '<input type="radio" name="tripApprovalSetting" value="disabled" style="margin-top: 3px;">' +
                                '<div>' +
                                    '<strong style="color: #f44336;">&#10007; Disable Approvals</strong>' +
                                    '<p style="margin: 5px 0 0 0; font-size: 13px; color: #666;">No approval required (overrides global)</p>' +
                                '</div>' +
                            '</label>' +
                        '</div>' +
                    '</div>' +
                    '<div class="email-modal-footer" style="display: flex; justify-content: flex-end; gap: 10px; padding: 15px 20px; background: #f8f9fa;">' +
                        '<button onclick="TripApprovalSettings.closeModal()" style="padding: 10px 20px; border: 1px solid #ddd; border-radius: 6px; background: white; cursor: pointer;">Cancel</button>' +
                        '<button onclick="TripApprovalSettings.saveSetting()" style="padding: 10px 20px; border: none; border-radius: 6px; background: #667eea; color: white; cursor: pointer;">Save Changes</button>' +
                    '</div>' +
                '</div>' +
            '</div>';
            document.body.insertAdjacentHTML('beforeend', modalHtml);
            document.getElementById('trip-approval-settings-modal').addEventListener('click', function(e) {
                if (e.target === this) TripApprovalSettings.closeModal();
            });
        },

        closeModal: function() {
            var modal = document.getElementById('trip-approval-settings-modal');
            if (modal) modal.classList.remove('active');
        },

        saveSetting: function() {
            var self = this;
            var selectedRadio = document.querySelector('input[name="tripApprovalSetting"]:checked');
            var setting = selectedRadio ? selectedRadio.value : 'inherit';
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: { ajax: 'true', action: 'save_trip_approval_setting', org_id: self.currentOrgId, setting: setting },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            window.location.reload();
                        } else {
                            alert('Error: ' + (result.message || 'Unknown error'));
                        }
                    } catch (e) { alert('Error processing response'); }
                },
                error: function(xhr, status, error) { alert('Error: ' + error); }
            });
        }
    };
    window.TripApprovalSettings = TripApprovalSettings;

    // Trip Fee Settings Controller - Manages trip cost and deposit amount settings
    var TripFeeSettings = {
        currentOrgId: null,
        currentTripCost: 0,
        currentDepositAmount: 0,

        // Show the settings modal for a trip
        showModal: function(orgId, tripCost, depositAmount) {
            this.currentOrgId = orgId;
            this.currentTripCost = tripCost || 0;
            this.currentDepositAmount = depositAmount || 0;

            // Create modal if it doesn't exist
            if (!document.getElementById('trip-fee-settings-modal')) {
                this.createModal();
            }

            // Set input values
            document.getElementById('tripCostInput').value = this.currentTripCost > 0 ? this.currentTripCost : '';
            document.getElementById('tripDepositInput').value = this.currentDepositAmount > 0 ? this.currentDepositAmount : '';

            // Show modal
            document.getElementById('trip-fee-settings-modal').classList.add('active');
        },

        // Create the modal HTML
        createModal: function() {
            var modalHtml = '<div id="trip-fee-settings-modal" class="email-modal-overlay">' +
                '<div class="email-modal" style="max-width: 480px;">' +
                    '<div class="email-modal-header" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">' +
                        '<h3 style="margin: 0;">&#128176; Trip Fee Settings</h3>' +
                        '<button class="email-close-btn" onclick="TripFeeSettings.closeModal()" style="color: white;">&times;</button>' +
                    '</div>' +
                    '<div class="email-modal-body" style="padding: 20px;">' +
                        '<p style="margin-bottom: 15px; color: #666;">Set trip cost and deposit amount for approval notifications:</p>' +
                        '<div style="display: flex; flex-direction: column; gap: 16px;">' +
                            '<div>' +
                                '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">Trip Cost ($)</label>' +
                                '<input type="number" id="tripCostInput" placeholder="Enter trip cost (e.g., 1500)" ' +
                                    'style="width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; box-sizing: border-box;">' +
                                '<p style="margin: 6px 0 0 0; font-size: 12px; color: #888;">Leave empty to use TouchPoint built-in trip cost</p>' +
                            '</div>' +
                            '<div>' +
                                '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">Required Deposit ($)</label>' +
                                '<input type="number" id="tripDepositInput" placeholder="Enter deposit amount (e.g., 300)" ' +
                                    'style="width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; box-sizing: border-box;">' +
                                '<p style="margin: 6px 0 0 0; font-size: 12px; color: #888;">Deposit amount required after approval</p>' +
                            '</div>' +
                        '</div>' +
                    '</div>' +
                    '<div class="email-modal-footer" style="display: flex; justify-content: flex-end; gap: 10px; padding: 15px 20px; background: #f8f9fa;">' +
                        '<button onclick="TripFeeSettings.closeModal()" style="padding: 10px 20px; border: 1px solid #ddd; border-radius: 6px; background: white; cursor: pointer;">Cancel</button>' +
                        '<button onclick="TripFeeSettings.saveSetting()" style="padding: 10px 20px; border: none; border-radius: 6px; background: #667eea; color: white; cursor: pointer;">Save Changes</button>' +
                    '</div>' +
                '</div>' +
            '</div>';
            document.body.insertAdjacentHTML('beforeend', modalHtml);
            document.getElementById('trip-fee-settings-modal').addEventListener('click', function(e) {
                if (e.target === this) TripFeeSettings.closeModal();
            });
        },

        closeModal: function() {
            var modal = document.getElementById('trip-fee-settings-modal');
            if (modal) modal.classList.remove('active');
        },

        saveSetting: function() {
            var self = this;
            var tripCost = document.getElementById('tripCostInput').value || 0;
            var depositAmount = document.getElementById('tripDepositInput').value || 0;
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'save_trip_fee_settings',
                    org_id: self.currentOrgId,
                    trip_cost: tripCost,
                    deposit_amount: depositAmount
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            window.location.reload();
                        } else {
                            alert('Error: ' + (result.message || 'Unknown error'));
                        }
                    } catch (e) { alert('Error processing response'); }
                },
                error: function(xhr, status, error) { alert('Error: ' + error); }
            });
        }
    };
    window.TripFeeSettings = TripFeeSettings;

    // Initialize event listeners when DOM is ready
    (function() {
        function initApprovalWorkflowEvents() {
            // Close modal when clicking outside
            var overlay = document.getElementById('approval-modal-overlay');
            if (overlay) {
                overlay.addEventListener('click', function(e) {
                    if (e.target === this) {
                        ApprovalWorkflow.closeModal();
                    }
                });
            }

            // Close modal with Escape key
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') {
                    var modalOverlay = document.getElementById('approval-modal-overlay');
                    if (modalOverlay && modalOverlay.classList.contains('active')) {
                        ApprovalWorkflow.closeModal();
                    }
                }
            });
        }

        // Run immediately if DOM is ready, otherwise wait
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initApprovalWorkflowEvents);
        } else {
            initApprovalWorkflowEvents();
        }
    })();

    // Make ApprovalWorkflow globally available
    window.ApprovalWorkflow = ApprovalWorkflow;
    </script>
    '''


def create_visualization_diagram():
    """Create Mermaid diagram for code structure visualization"""
    if not config.ENABLE_DEVELOPER_VISUALIZATION or not model.UserIsInRole("SuperAdmin"):
        return ""
    
    return '''
    <details class="code-visualization" style="margin: 20px 0;">
        <summary>📊 Developer: View Code Structure</summary>
        <div class="mermaid" style="background: white; padding: 20px; margin-top: 10px;">
        graph TD
            A[main] -->|Check Auth| B[User Authentication]
            A -->|Route| C{Active Tab?}
            
            C -->|dashboard| D[Dashboard View]
            C -->|due| E[Finance View]
            C -->|messages| F[Messages View]
            C -->|stats| G[Stats View]
            
            D --> H[Load Total Stats]
            D --> I[Load Active Missions]
            D --> J[Load Recent Signups]
            D --> K[Load Upcoming Meetings]
            
            E --> L[Load Outstanding Payments]
            E --> M[Load Transactions]
            
            F --> N[Load Email History]
            
            G --> O[Load Demographics]
            G --> P[Financial Summary]
            G --> Q[Trip Trends]
            G --> R[Participation Analysis]
            
            style A fill:#f9f,stroke:#333,stroke-width:4px
            style B fill:#bbf,stroke:#333,stroke-width:2px
            style C fill:#ffd,stroke:#333,stroke-width:2px
        </div>
        <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
        <script>mermaid.initialize({ startOnLoad: true });</script>
    </details>
    '''

def render_progress_bar(current, total, label=""):
    """Render a progress bar with label"""
    if total == 0:
        percentage = 100
    else:
        percentage = (float(current) / float(total)) * 100
    
    return '''
    <div style="margin-top: 10px;">
        <div class="progress-container">
            <div class="progress-bar" style="width: {0:.1f}%;"></div>
        </div>
        <div class="progress-text">{1}</div>
    </div>
    '''.format(percentage, label)

def format_popup_data(data, list_type):
    """Format data for popup display"""
    if not data:
        return '<p>No data found.</p>'
    
    html = '<div style="max-height: 400px; overflow-y: auto;">'
    
    if list_type == 'meetings':
        html += '<table class="mission-table">'
        html += '<tbody>'
        for row in data:
            status_class = 'text-muted' if row.Status == 'Past' else 'text-primary'
            html += '''
            <tr>
                <td class="{0}">
                    <strong>{1}</strong><br>
                    📅 {2}<br>
                    📍 {3}
                </td>
            </tr>
            '''.format(
                status_class,
                row.Description or "No description",
                format_date(str(row.MeetingDate)),
                row.Location or "TBD"
            )
        html += '</tbody></table>'
    
    elif list_type == 'passports':
        html += '<table class="mission-table">'
        html += '<tbody>'
        for row in data:
            html += '''
            <tr>
                <td>
                    <a href="/Person2/{0}" target="_blank">{1}</a> (Age: {2})
                </td>
            </tr>
            '''.format(row.PeopleId, row.Name, row.Age)
        html += '</tbody></table>'
    
    elif list_type == 'bgchecks':
        html += '<table class="mission-table">'
        html += '<tbody>'
        for row in data:
            html += '''
            <tr>
                <td>
                    <a href="/Person2/{0}" target="_blank">{1}</a> (Age: {2})<br>
                    <span class="status-badge status-urgent">{3}</span>
                </td>
            </tr>
            '''.format(row.PeopleId, row.Name, row.Age, row.Issue)
        html += '</tbody></table>'
    
    elif list_type == 'people':
        html += '<table class="mission-table">'
        html += '<tbody>'
        current_role = None
        for row in data:
            if row.Role != current_role:
                html += '<tr class="org-separator"><td><strong>{0}</strong></td></tr>'.format(row.Role)
                current_role = row.Role
            
            html += '''
            <tr>
                <td>
                    <a href="/Person2/{0}" target="_blank">{1}</a> (Age: {2})<br>
                    📱 {3} | ✉️ {4}
                </td>
            </tr>
            '''.format(
                row.PeopleId, 
                row.Name, 
                row.Age,
                model.FmtPhone(row.CellPhone) if row.CellPhone else 'No phone',
                row.EmailAddress or 'No email'
            )
        html += '</tbody></table>'
    
    html += '</div>'
    return html

def handle_ajax_request():
    """Handle AJAX requests for popup data and section loading"""
    import json

    # Check if this is a POST request with action parameter
    if not hasattr(model.Data, 'action'):
        # Also check for ajax=true parameter for GET requests
        if hasattr(model.Data, 'ajax') and str(model.Data.ajax).lower() == 'true':
            return handle_ajax_section_load()
        return False

    action = str(model.Data.action) if model.Data.action else ""

    if action == 'get_popup_data':
        # Get parameters
        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') else ""
        list_type = str(model.Data.list_type) if hasattr(model.Data, 'list_type') else ""

        if org_id and list_type:
            # Get the query
            query = get_popup_data_query(org_id, list_type)

            if query:
                # Execute query
                timer = start_timer()
                data = execute_query_with_debug(query, "Popup Data: " + list_type, "sql")

                # Format and return the data
                html = format_popup_data(data, list_type)
                print html
            else:
                print '<p>Invalid list type.</p>'
        else:
            print '<p>Missing parameters.</p>'

        # Return True to indicate AJAX was handled
        return True

    elif action == 'load_section':
        # AJAX request to load a trip section
        return handle_ajax_section_load()

    elif action == 'adjust_fee':
        # AJAX request to adjust member fee (admin only)
        return handle_fee_adjustment()

    elif action == 'update_dates':
        # AJAX request to update trip dates (admin only)
        return handle_date_update()

    elif action == 'get_person_details':
        # AJAX request to get person details for popup
        return handle_person_details()

    elif action == 'send_email':
        # AJAX request to send email via TouchPoint's email system
        return handle_send_email()

    elif action == 'create_meeting':
        # AJAX request to create a new meeting
        return handle_create_meeting()

    elif action == 'update_meeting':
        # AJAX request to update an existing meeting
        return handle_update_meeting()

    elif action == 'get_message_detail':
        # AJAX request to get message details for modal
        return handle_get_message_detail()

    elif action == 'resend_message':
        # AJAX request to resend a message to a person
        return handle_resend_message()

    elif action == 'get_templates':
        # AJAX request to get all email templates
        return handle_get_templates()

    elif action == 'save_template':
        # AJAX request to save an email template
        return handle_save_template()

    elif action == 'reset_template':
        # AJAX request to reset a template to default
        return handle_reset_template()

    elif action == 'get_dropdown_templates':
        # AJAX request to get email dropdown templates
        return handle_get_dropdown_templates()

    elif action == 'save_dropdown_template':
        # AJAX request to save an email dropdown template
        return handle_save_dropdown_template()

    elif action == 'delete_dropdown_template':
        # AJAX request to delete an email dropdown template
        return handle_delete_dropdown_template()

    elif action == 'get_config':
        # AJAX request to get dashboard configuration
        return handle_get_config()

    elif action == 'save_config':
        # AJAX request to save dashboard configuration
        return handle_save_config()

    elif action == 'get_global_settings':
        # AJAX request to get global dashboard settings
        return handle_get_global_settings()

    elif action == 'save_global_settings':
        # AJAX request to save global dashboard settings
        return handle_save_global_settings()

    elif action == 'get_trip_approval_setting':
        # AJAX request to get approval setting for a specific trip
        return handle_get_trip_approval_setting()

    elif action == 'save_trip_approval_setting':
        # AJAX request to save approval setting for a specific trip
        return handle_save_trip_approval_setting()

    elif action == 'get_trip_fee_settings':
        # AJAX request to get trip fee settings (cost and deposit)
        return handle_get_trip_fee_settings()

    elif action == 'save_trip_fee_settings':
        # AJAX request to save trip fee settings (cost and deposit)
        return handle_save_trip_fee_settings()

    # =====================================================
    # TRIP REGISTRATION APPROVAL WORKFLOW AJAX HANDLERS
    # =====================================================

    elif action == 'get_approval_statuses':
        # AJAX request to get approval statuses for all trip members
        return handle_get_approval_statuses()

    elif action == 'get_registration_answers':
        # AJAX request to get registration form Q&A for a person
        return handle_get_registration_answers()

    elif action == 'approve_member':
        # AJAX request to approve a trip member
        return handle_approve_member()

    elif action == 'deny_member':
        # AJAX request to deny a trip member
        return handle_deny_member()

    elif action == 'revoke_approval':
        # AJAX request to revoke approval/denial (return to pending)
        return handle_revoke_approval()

    elif action == 'set_member_type':
        # AJAX request to set member type (Leader/Member)
        return handle_set_member_type()

    # DEBUG: If we got here with an ajax=1 parameter, log what action we received
    if hasattr(model.Data, 'ajax') and str(model.Data.ajax) == '1':
        print '{"debug_unknown_action": true, "action_received": "' + action + '"}'
        return True

    return False  # Not an AJAX request


def handle_person_details():
    """Handle AJAX request to get person details for popup modal."""
    people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None
    org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None

    if not people_id:
        print '<div class="person-modal-loading"><p style="color: red;">Missing person ID.</p></div>'
        return True

    try:
        # Query person details with picture from Picture table
        person_sql = '''
        SELECT
            p.PeopleId,
            p.Name2,
            p.FirstName,
            p.LastName,
            p.NickName,
            p.Age,
            p.GenderId,
            g.Description as Gender,
            p.EmailAddress,
            p.CellPhone,
            p.HomePhone,
            p.WorkPhone,
            p.AddressLineOne,
            p.AddressLineTwo,
            p.CityName,
            p.StateCode,
            p.ZipCode,
            ms.Description as MemberStatus,
            p.BaptismDate,
            pic.LargeUrl as PictureUrl
        FROM People p WITH (NOLOCK)
        LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
        LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
        LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
        WHERE p.PeopleId = {0}
        '''.format(people_id)

        person_result = list(q.QuerySql(person_sql))
        if not person_result:
            print '<div class="person-modal-loading"><p style="color: red;">Person not found.</p></div>'
            return True

        person = person_result[0]

        # Get member type for this organization if org_id provided
        member_type = "Team Member"
        if org_id:
            member_type_sql = '''
            SELECT mt.Description
            FROM OrganizationMembers om WITH (NOLOCK)
            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
            WHERE om.OrganizationId = {0} AND om.PeopleId = {1}
            '''.format(org_id, people_id)
            mt_result = list(q.QuerySql(member_type_sql))
            if mt_result and mt_result[0].Description:
                member_type = mt_result[0].Description

        # Get all mission trips this person has been on
        trips_sql = '''
        SELECT
            o.OrganizationId,
            o.OrganizationName,
            o.Location,
            (SELECT TOP 1 oe.DateValue FROM OrganizationExtra oe
             WHERE oe.OrganizationId = o.OrganizationId AND oe.Field = 'Main Event Start') as TripDate,
            om.EnrollmentDate,
            CASE WHEN o.OrganizationId = {1} THEN 1 ELSE 0 END as IsCurrentTrip
        FROM OrganizationMembers om WITH (NOLOCK)
        JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
        WHERE om.PeopleId = {0}
          AND o.IsMissionTrip = 1
          AND om.MemberTypeId NOT IN (230, 311)
        ORDER BY
            CASE WHEN o.OrganizationId = {1} THEN 0 ELSE 1 END,
            om.EnrollmentDate DESC
        '''.format(people_id, org_id or 0)

        trips = list(q.QuerySql(trips_sql))

        # Format display name
        display_name = person.Name2 or 'Unknown'
        if ',' in display_name:
            name_parts = display_name.split(',')
            display_name = '{0} {1}'.format(name_parts[1].strip(), name_parts[0].strip())

        # Build the modal HTML
        html = []

        # Header with photo
        html.append('<div class="person-modal-header">')
        html.append('<button class="person-modal-close" onclick="PersonDetails.close()">&times;</button>')

        if person.PictureUrl:
            html.append('<img src="{0}" alt="{1}" class="person-modal-photo" onerror="this.style.display=\'none\'; this.nextElementSibling.style.display=\'flex\';">'.format(
                person.PictureUrl, _escape_html(display_name)))
            html.append('<div class="person-modal-photo-placeholder" style="display: none;">{0}</div>'.format(
                display_name[0].upper() if display_name else '?'))
        else:
            html.append('<div class="person-modal-photo-placeholder">{0}</div>'.format(
                display_name[0].upper() if display_name else '?'))

        html.append('<div class="person-modal-name">{0}</div>'.format(_escape_html(display_name)))
        html.append('<div class="person-modal-role">{0}</div>'.format(_escape_html(member_type)))
        html.append('</div>')

        # Body with details
        html.append('<div class="person-modal-body">')

        # Basic Info Section
        html.append('<div class="person-detail-section">')
        html.append('<h4>Basic Information</h4>')

        if person.Age:
            html.append('<div class="person-detail-row"><span class="person-detail-label">Age</span><span class="person-detail-value">{0}</span></div>'.format(person.Age))

        if person.Gender:
            html.append('<div class="person-detail-row"><span class="person-detail-label">Gender</span><span class="person-detail-value">{0}</span></div>'.format(_escape_html(person.Gender)))

        if person.MemberStatus:
            html.append('<div class="person-detail-row"><span class="person-detail-label">Status</span><span class="person-detail-value">{0}</span></div>'.format(_escape_html(person.MemberStatus)))

        if person.BaptismDate:
            html.append('<div class="person-detail-row"><span class="person-detail-label">Baptized</span><span class="person-detail-value">{0}</span></div>'.format(format_date(str(person.BaptismDate))))

        html.append('</div>')

        # Contact Section
        html.append('<div class="person-detail-section">')
        html.append('<h4>Contact Information</h4>')

        if person.EmailAddress:
            # Make email clickable to open MissionsEmail popup (close person modal first)
            escaped_email_contact = _escape_html(person.EmailAddress).replace("'", "\\'")
            escaped_name_contact = _escape_html(display_name.split()[0] if display_name else 'Team Member').replace("'", "\\'")
            org_id_js = org_id if org_id else 'null'
            html.append('<div class="person-detail-row"><span class="person-detail-label">Email</span><span class="person-detail-value"><a href="#" onclick="PersonDetails.close(); MissionsEmail.openIndividual({0}, \'{1}\', \'{2}\', {4}); return false;" style="color: #007bff; cursor: pointer;">{3}</a></span></div>'.format(
                people_id, escaped_email_contact, escaped_name_contact, _escape_html(person.EmailAddress), org_id_js))

        if person.CellPhone:
            formatted_phone = model.FmtPhone(person.CellPhone, "") if person.CellPhone else ""
            html.append('<div class="person-detail-row"><span class="person-detail-label">Cell</span><span class="person-detail-value"><a href="tel:{0}">{1}</a></span></div>'.format(
                person.CellPhone.replace('-', '').replace(' ', '').replace('(', '').replace(')', ''),
                formatted_phone or person.CellPhone))

        if person.HomePhone:
            formatted_phone = model.FmtPhone(person.HomePhone, "") if person.HomePhone else ""
            html.append('<div class="person-detail-row"><span class="person-detail-label">Home</span><span class="person-detail-value">{0}</span></div>'.format(formatted_phone or person.HomePhone))

        # Address
        address_parts = []
        if person.AddressLineOne:
            address_parts.append(_escape_html(person.AddressLineOne))
        if person.AddressLineTwo:
            address_parts.append(_escape_html(person.AddressLineTwo))
        if person.CityName or person.StateCode or person.ZipCode:
            city_state_zip = []
            if person.CityName:
                city_state_zip.append(_escape_html(person.CityName))
            if person.StateCode:
                city_state_zip.append(_escape_html(person.StateCode))
            if person.ZipCode:
                city_state_zip.append(_escape_html(person.ZipCode))
            address_parts.append(', '.join(city_state_zip[:2]) + (' ' + city_state_zip[2] if len(city_state_zip) > 2 else ''))

        if address_parts:
            html.append('<div class="person-detail-row"><span class="person-detail-label">Address</span><span class="person-detail-value">{0}</span></div>'.format('<br>'.join(address_parts)))

        html.append('</div>')

        # Mission Trips Section
        if trips:
            html.append('<div class="person-detail-section">')
            html.append('<h4>Mission Trip History ({0} trip{1})</h4>'.format(len(trips), 's' if len(trips) != 1 else ''))
            html.append('<ul class="person-trips-list">')

            for trip in trips:
                trip_date_str = ''
                if trip.TripDate:
                    try:
                        trip_date_str = format_date(str(trip.TripDate))
                    except:
                        trip_date_str = str(trip.TripDate)[:10] if trip.TripDate else ''

                is_current = trip.IsCurrentTrip == 1
                trip_name = _escape_html(trip.OrganizationName or 'Unknown Trip')
                location = _escape_html(trip.Location or '')

                html.append('''
                    <li{0}>
                        <div>
                            <span class="trip-name">{1}</span>
                            {2}
                        </div>
                        <span class="trip-date">{3}</span>
                    </li>
                '''.format(
                    ' style="border: 2px solid #4A90E2; background: #e8f4fd;"' if is_current else '',
                    trip_name,
                    '<br><small style="color: #666;">{0}</small>'.format(location) if location else '',
                    trip_date_str or 'Date TBD'
                ))

            html.append('</ul>')
            html.append('</div>')
        else:
            html.append('<div class="person-detail-section">')
            html.append('<h4>Mission Trip History</h4>')
            html.append('<p style="color: #666; font-style: italic;">No mission trip history found.</p>')
            html.append('</div>')

        html.append('</div>')  # person-modal-body

        # Get first name for email modal
        first_name = display_name.split()[0] if display_name else 'Team Member'

        # Action buttons
        html.append('<div class="person-modal-actions">')
        html.append('<a href="/Person2/{0}" target="_blank" class="person-modal-btn person-modal-btn-primary">View Full Profile</a>'.format(people_id))
        if person.EmailAddress:
            # Use MissionsEmail popup instead of mailto: - close person modal first
            escaped_email = _escape_html(person.EmailAddress).replace("'", "\\'")
            escaped_name = _escape_html(first_name).replace("'", "\\'")
            org_id_btn = org_id if org_id else 'null'
            html.append('<button type="button" class="person-modal-btn person-modal-btn-secondary" onclick="PersonDetails.close(); MissionsEmail.openIndividual({0}, \'{1}\', \'{2}\', {3})">Send Email</button>'.format(
                people_id, escaped_email, escaped_name, org_id_btn))
        html.append('</div>')

        print ''.join(html)

    except Exception as e:
        print '<div class="person-modal-loading"><p style="color: red;">Error loading details: {0}</p></div>'.format(str(e))

    return True


def _record_passport_request(org_id, people_id, person_name):
    """
    Record a passport request in the organization's extra values.
    Stores JSON data with person ID, name, and request timestamp.
    """
    import json
    from datetime import datetime

    try:
        # Get existing passport requests or create new list
        extra_value_name = "PassportRequests"
        existing_value = model.ExtraValueTextOrg(org_id, extra_value_name)

        if existing_value:
            try:
                requests = json.loads(existing_value)
            except:
                requests = []
        else:
            requests = []

        # Add new request
        new_request = {
            'people_id': people_id,
            'name': person_name,
            'requested_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'requested_by': model.UserPeopleId
        }

        # Check if person already has a request - update if so
        found = False
        for i, req in enumerate(requests):
            if req.get('people_id') == people_id:
                requests[i] = new_request
                found = True
                break

        if not found:
            requests.append(new_request)

        # Save updated list
        model.AddExtraValueTextOrg(org_id, extra_value_name, json.dumps(requests))

    except Exception as e:
        # Don't fail the email send if tracking fails
        pass


def _personalize_email_body(body, people_id, org_id):
    """
    Replace personalized link placeholders in email body.

    Available placeholders:
    - {{MyGivingLink}} - Link to recipient's own giving/registrations page
    - {{SupportLink}} - Shareable link for others to support this person
    - {{PersonName}} - Recipient's name
    """
    if not people_id:
        return body

    # Get person info for name placeholder
    person = model.GetPerson(int(people_id))
    person_name = person.Name if person else ""

    # Build personalized links
    my_giving_link = "{0}/Person2/{1}#tab-registrations".format(CHURCH_URL, people_id)

    # Support link requires org_id
    if org_id:
        support_link = "{0}/OnlineReg/{1}?goerid={2}".format(CHURCH_URL, org_id, people_id)
    else:
        support_link = ""

    # Replace placeholders (case-insensitive)
    result = body
    result = result.replace('{{MyGivingLink}}', my_giving_link)
    result = result.replace('{{mygivinglink}}', my_giving_link)
    result = result.replace('{{SupportLink}}', support_link)
    result = result.replace('{{supportlink}}', support_link)
    result = result.replace('{{PersonName}}', person_name)
    result = result.replace('{{personname}}', person_name)
    result = result.replace('{{ChurchUrl}}', CHURCH_URL)
    result = result.replace('{{churchurl}}', CHURCH_URL)

    return result


def handle_send_email():
    """Handle AJAX request to send email via TouchPoint's email system."""
    import json

    try:
        # Get email parameters
        to_emails = str(model.Data.to_emails) if hasattr(model.Data, 'to_emails') and model.Data.to_emails else ""
        subject = str(model.Data.subject) if hasattr(model.Data, 'subject') and model.Data.subject else ""
        body = str(model.Data.body) if hasattr(model.Data, 'body') and model.Data.body else ""
        people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None

        # Check if this is a passport request
        is_passport_request = str(model.Data.is_passport_request) if hasattr(model.Data, 'is_passport_request') and model.Data.is_passport_request else ""
        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None

        if not to_emails:
            print json.dumps({'success': False, 'message': 'No recipients specified.'})
            return True

        if not subject.strip():
            print json.dumps({'success': False, 'message': 'Subject is required.'})
            return True

        if not body.strip():
            print json.dumps({'success': False, 'message': 'Message body is required.'})
            return True

        # Decode HTML entities that were encoded client-side to avoid ASP.NET request validation
        # This reverses the encoding done in JavaScript before POST
        body = body.replace('&lt;', '<').replace('&gt;', '>')
        body = body.replace('&quot;', '"').replace('&#39;', "'")

        # Get current user info for the from address
        current_user_id = model.UserPeopleId
        current_user = model.GetPerson(current_user_id)

        from_email = current_user.EmailAddress if current_user and current_user.EmailAddress else ""
        from_name = current_user.Name if current_user else "Missions Dashboard"

        # Parse recipients (comma-separated)
        recipient_list = [e.strip() for e in to_emails.split(',') if e.strip()]

        # Format body as HTML (convert newlines to <br>)
        html_body = body.replace('\n', '<br>')

        # Send to each recipient using TouchPoint's email system
        # model.Email(query, queuedBy, fromEmail, fromName, subject, body, ccList)
        if people_id and len(recipient_list) == 1:
            # Single recipient with people_id - personalize the email
            person = model.GetPerson(int(people_id))
            if person:
                # Personalize the email body with links for this recipient
                personalized_body = _personalize_email_body(html_body, people_id, org_id)

                # Use PeopleId query to send to specific person (no spaces around =)
                query = "PeopleId={0}".format(int(people_id))
                model.Email(query, current_user_id, from_email, from_name, subject, personalized_body, None)

                # Record passport request if applicable
                if is_passport_request == 'true' and org_id:
                    _record_passport_request(int(org_id), int(people_id), person.Name)

                print json.dumps({'success': True, 'message': 'Email sent successfully to {0}.'.format(person.Name)})
            else:
                print json.dumps({'success': False, 'message': 'Person not found.'})
        else:
            # Multiple recipients - find people by email using SQL and send individually
            # Each recipient gets personalized links
            emails_sent = 0
            errors = []

            for email in recipient_list:
                try:
                    # Find person by email using SQL query
                    sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}' AND IsDeceased = 0".format(
                        email.replace("'", "''"))
                    result = q.QuerySql(sql)
                    pid = None
                    if result:
                        for row in result:
                            if row.PeopleId:
                                pid = row.PeopleId
                                break

                    if pid:
                        # Personalize the email body for this recipient
                        personalized_body = _personalize_email_body(html_body, pid, org_id)

                        # Send email to this person
                        query = "PeopleId={0}".format(pid)
                        model.Email(query, current_user_id, from_email, from_name, subject, personalized_body, None)
                        emails_sent += 1
                    else:
                        errors.append("Could not find person for: " + email)
                except Exception as e:
                    errors.append("Error sending to {0}: {1}".format(email, str(e)))

            if emails_sent > 0:
                msg = 'Email sent successfully to {0} recipient(s).'.format(emails_sent)
                if errors:
                    msg += ' Some errors occurred: ' + '; '.join(errors)
                print json.dumps({'success': True, 'message': msg})
            else:
                print json.dumps({'success': False, 'message': 'Failed to send emails. ' + '; '.join(errors)})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_fee_adjustment():
    """Handle AJAX request to adjust a member's fee."""
    import json

    # Check finance permission - requires Finance/FinanceAdmin/ManageTransactions role
    user_role = get_user_role_and_trips()
    if not user_role.get('can_manage_finance', False):
        print json.dumps({'success': False, 'message': 'Access denied. Finance role required.'})
        return True

    people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None
    org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
    amount = str(model.Data.amount) if hasattr(model.Data, 'amount') and model.Data.amount else None
    description = str(model.Data.description) if hasattr(model.Data, 'description') and model.Data.description else 'Fee adjustment from Mission Dashboard'

    if not people_id or not org_id:
        print json.dumps({'success': False, 'message': 'Missing people ID or organization ID.'})
        return True

    if not amount:
        print json.dumps({'success': False, 'message': 'Missing adjustment amount.'})
        return True

    try:
        # Parse amount - can be positive (add to fee) or negative (reduce fee)
        amount_float = float(amount.replace(',', '').replace('$', ''))

        # Use TouchPoint's AdjustFee method
        # This adjusts the member's balance in the organization
        model.AdjustFee(int(people_id), int(org_id), amount_float, description)

        print json.dumps({
            'success': True,
            'message': 'Fee adjusted successfully.',
            'amount': amount_float
        })
    except ValueError:
        print json.dumps({'success': False, 'message': 'Invalid amount format.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_date_update():
    """Handle AJAX request to update trip dates (admin only)."""
    import json

    # Check admin permission
    user_role = get_user_role_and_trips()
    if not user_role.get('is_admin', False):
        print json.dumps({'success': False, 'message': 'Access denied. Admin only.'})
        return True

    org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
    start_date = str(model.Data.start_date) if hasattr(model.Data, 'start_date') and model.Data.start_date else None
    end_date = str(model.Data.end_date) if hasattr(model.Data, 'end_date') and model.Data.end_date else None
    close_date = str(model.Data.close_date) if hasattr(model.Data, 'close_date') and model.Data.close_date else None

    if not org_id:
        print json.dumps({'success': False, 'message': 'Missing organization ID.'})
        return True

    try:
        updates_made = []

        # Update Main Event Start date if provided
        if start_date and start_date.strip():
            model.AddExtraValueDateOrg(int(org_id), "Main Event Start", start_date)
            updates_made.append("Main Event Start")

        # Update Main Event End date if provided
        if end_date and end_date.strip():
            model.AddExtraValueDateOrg(int(org_id), "Main Event End", end_date)
            updates_made.append("Main Event End")

        # Update Close date if provided
        if close_date and close_date.strip():
            model.AddExtraValueDateOrg(int(org_id), "Close", close_date)
            updates_made.append("Close")

        if updates_made:
            print json.dumps({
                'success': True,
                'message': 'Updated: ' + ', '.join(updates_made),
                'updated_fields': updates_made
            })
        else:
            print json.dumps({
                'success': False,
                'message': 'No dates were provided to update.'
            })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_create_meeting():
    """Handle AJAX request to create a new meeting for a trip."""
    import json
    import datetime

    # Get parameters
    org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
    meeting_date = str(model.Data.meeting_date) if hasattr(model.Data, 'meeting_date') and model.Data.meeting_date else None
    meeting_time = str(model.Data.meeting_time) if hasattr(model.Data, 'meeting_time') and model.Data.meeting_time else None
    description = str(model.Data.description) if hasattr(model.Data, 'description') and model.Data.description else 'Team Meeting'
    location = str(model.Data.location) if hasattr(model.Data, 'location') and model.Data.location else ''

    if not org_id:
        print json.dumps({'success': False, 'message': 'Missing organization ID.'})
        return True

    if not meeting_date:
        print json.dumps({'success': False, 'message': 'Meeting date is required.'})
        return True

    try:
        # Parse the date
        date_parts = meeting_date.split('-')
        year = int(date_parts[0])
        month = int(date_parts[1])
        day = int(date_parts[2])

        # Parse time (default to 6:00 PM if not provided)
        if meeting_time:
            time_parts = meeting_time.split(':')
            hour = int(time_parts[0])
            minute = int(time_parts[1]) if len(time_parts) > 1 else 0
        else:
            hour = 18
            minute = 0

        # Create datetime object
        meeting_datetime = datetime.datetime(year, month, day, hour, minute, 0)

        # Create the meeting using TouchPoint API
        # GetMeetingIdByDateTime will create if not exists
        meeting_id = model.GetMeetingIdByDateTime(int(org_id), meeting_datetime, createIfNotExists=True)

        if meeting_id:
            # Set the Description and Location using AddExtraValueMeeting
            # These are standard meeting fields in TouchPoint
            if description and description.strip():
                try:
                    model.AddExtraValueMeeting(meeting_id, 'Description', description.strip())
                except:
                    pass  # Description may already be set or not supported as extra value

            if location and location.strip():
                try:
                    model.AddExtraValueMeeting(meeting_id, 'Location', location.strip())
                except:
                    pass  # Location may already be set or not supported as extra value

            # Format date for display
            formatted_date = meeting_datetime.strftime('%B %d, %Y at %I:%M %p')
            print json.dumps({
                'success': True,
                'message': 'Meeting created successfully!',
                'meeting_id': meeting_id,
                'formatted_date': formatted_date,
                'description': description,
                'location': location
            })
        else:
            print json.dumps({'success': False, 'message': 'Failed to create meeting.'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_update_meeting():
    """Handle AJAX request to update an existing meeting's date/time, description, and location."""
    import json
    import datetime

    # Set content type for JSON response
    print 'Content-Type: application/json'
    print ''

    try:
        meeting_id = model.Data.meeting_id
        org_id = model.Data.org_id
        meeting_date = model.Data.meeting_date  # Format: YYYY-MM-DD
        meeting_time = model.Data.meeting_time  # Format: HH:MM
        description = model.Data.description if hasattr(model.Data, 'description') else ''
        location = model.Data.location if hasattr(model.Data, 'location') else ''

        if not meeting_id:
            print json.dumps({'success': False, 'message': 'Meeting ID is required.'})
            return True

        if not meeting_date or not meeting_time:
            print json.dumps({'success': False, 'message': 'Meeting date and time are required.'})
            return True

        # Parse and combine date and time
        try:
            date_parts = meeting_date.split('-')
            time_parts = meeting_time.split(':')

            year = int(date_parts[0])
            month = int(date_parts[1])
            day = int(date_parts[2])
            hour = int(time_parts[0])
            minute = int(time_parts[1])

            meeting_datetime = datetime.datetime(year, month, day, hour, minute)
        except Exception as parse_error:
            print json.dumps({'success': False, 'message': 'Invalid date/time format: ' + str(parse_error)})
            return True

        # Update the meeting date/time using model.UpdateMeetingDate
        try:
            model.UpdateMeetingDate(int(meeting_id), meeting_datetime)
        except Exception as update_error:
            print json.dumps({'success': False, 'message': 'Error updating meeting date/time: ' + str(update_error)})
            return True

        # Update Description using AddExtraValueMeeting
        if description is not None:
            try:
                model.AddExtraValueMeeting(int(meeting_id), 'Description', description.strip() if description else '')
            except:
                pass  # Description may not be supported as extra value

        # Update Location using AddExtraValueMeeting
        if location is not None:
            try:
                model.AddExtraValueMeeting(int(meeting_id), 'Location', location.strip() if location else '')
            except:
                pass  # Location may not be supported as extra value

        # Format date for display
        formatted_date = meeting_datetime.strftime('%B %d, %Y at %I:%M %p')
        print json.dumps({
            'success': True,
            'message': 'Meeting updated successfully!',
            'meeting_id': meeting_id,
            'formatted_date': formatted_date,
            'description': description,
            'location': location
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_get_message_detail():
    """Handle AJAX request to get message details for the modal."""
    import json

    # Set content type for JSON response
    print 'Content-Type: application/json'
    print ''

    try:
        email_id = model.Data.emailId if hasattr(model.Data, 'emailId') else None
        org_id = model.Data.orgId if hasattr(model.Data, 'orgId') else None

        if not email_id:
            print json.dumps({'success': False, 'error': 'Missing email ID'})
            return True

        # Get email details
        email_sql = '''
        SELECT
            eq.Id,
            eq.Subject,
            eq.FromAddr,
            eq.FromName,
            eq.Queued,
            eq.Body
        FROM EmailQueue eq WITH (NOLOCK)
        WHERE eq.Id = {0}
        '''.format(int(email_id))

        email_result = list(q.QuerySql(email_sql))
        if not email_result:
            print json.dumps({'success': False, 'error': 'Email not found'})
            return True

        email = email_result[0]

        # Get recipients who are trip members
        recipients_sql = '''
        SELECT
            p.PeopleId,
            p.Name2 as Name,
            eqt.Sent
        FROM EmailQueueTo eqt WITH (NOLOCK)
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = eqt.PeopleId
        WHERE eqt.Id = {0}
        '''.format(int(email_id))

        # If org_id is provided, filter to only trip members
        if org_id:
            recipients_sql = '''
            SELECT
                p.PeopleId,
                p.Name2 as Name,
                eqt.Sent
            FROM EmailQueueTo eqt WITH (NOLOCK)
            INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = eqt.PeopleId
            INNER JOIN OrganizationMembers om WITH (NOLOCK) ON om.PeopleId = p.PeopleId
            WHERE eqt.Id = {0}
              AND om.OrganizationId = {1}
              AND om.MemberTypeId <> {2}
            ORDER BY p.Name2
            '''.format(int(email_id), int(org_id), config.MEMBER_TYPE_LEADER)

        recipients = list(q.QuerySql(recipients_sql))

        # Format the response
        recipients_list = []
        for r in recipients:
            recipients_list.append({
                'peopleId': r.PeopleId,
                'name': r.Name,
                'sent': r.Sent.strftime('%Y-%m-%d %H:%M') if r.Sent and hasattr(r.Sent, 'strftime') else None
            })

        # Format queued date
        queued_str = ''
        if email.Queued:
            if hasattr(email.Queued, 'strftime'):
                queued_str = email.Queued.strftime('%B %d, %Y at %I:%M %p')
            else:
                queued_str = str(email.Queued)

        response = {
            'success': True,
            'data': {
                'emailId': email.Id,
                'subject': email.Subject or '',
                'fromAddr': email.FromAddr or '',
                'fromName': email.FromName or '',
                'queued': queued_str,
                'body': email.Body or '',
                'recipients': recipients_list
            }
        }

        print json.dumps(response)

    except Exception as e:
        print json.dumps({'success': False, 'error': str(e)})

    return True


def handle_resend_message():
    """Handle AJAX request to resend an email to a specific person."""
    import json

    # Set content type for JSON response
    print 'Content-Type: application/json'
    print ''

    try:
        email_id = model.Data.emailId if hasattr(model.Data, 'emailId') else None
        people_id = model.Data.peopleId if hasattr(model.Data, 'peopleId') else None

        if not email_id or not people_id:
            print json.dumps({'success': False, 'error': 'Missing email ID or person ID'})
            return True

        # Get the original email
        email_sql = '''
        SELECT
            eq.Subject,
            eq.Body,
            eq.FromAddr,
            eq.FromName
        FROM EmailQueue eq WITH (NOLOCK)
        WHERE eq.Id = {0}
        '''.format(int(email_id))

        email_result = list(q.QuerySql(email_sql))
        if not email_result:
            print json.dumps({'success': False, 'error': 'Original email not found'})
            return True

        email = email_result[0]

        # Get the person's email
        person = model.GetPerson(int(people_id))
        if not person:
            print json.dumps({'success': False, 'error': 'Person not found'})
            return True

        if not person.EmailAddress:
            print json.dumps({'success': False, 'error': 'Person has no email address'})
            return True

        # Send the email using TouchPoint's email system
        # Use model.Email to send to a single person
        from_email = email.FromAddr or 'noreply@touchpointsoftware.com'
        from_name = email.FromName or 'TouchPoint'
        subject = email.Subject or '(No Subject)'
        body = email.Body or ''

        # Create a query that returns just this person
        person_query = 'PeopleId = {0}'.format(int(people_id))

        try:
            model.Email(person_query, model.UserPeopleId, from_email, from_name, subject, body, None)
            print json.dumps({'success': True, 'message': 'Email queued for resend'})
        except Exception as send_error:
            print json.dumps({'success': False, 'error': 'Failed to send email: ' + str(send_error)})

    except Exception as e:
        print json.dumps({'success': False, 'error': str(e)})

    return True


def get_template(template_id):
    """
    Get an email template by ID.
    First checks for custom template in Special Content, then falls back to default.

    Args:
        template_id: Template identifier (e.g., 'goal_reminder_team')

    Returns:
        Dict with 'name', 'description', 'subject', 'body', 'is_custom'
    """
    # Try to load custom template from Special Content
    content_name = config.TEMPLATE_CONTENT_PREFIX + template_id
    try:
        custom_content = model.TextContent(content_name)
        if custom_content:
            import json
            template_data = json.loads(custom_content)
            template_data['is_custom'] = True
            template_data['id'] = template_id
            return template_data
    except:
        pass

    # Fall back to default template
    default = config.DEFAULT_TEMPLATES.get(template_id)
    if default:
        return {
            'id': template_id,
            'name': default['name'],
            'description': default['description'],
            'subject': default['subject'],
            'body': default['body'],
            'is_custom': False
        }

    return None


def save_template(template_id, subject, body):
    """
    Save a custom email template to Special Content.

    Args:
        template_id: Template identifier
        subject: Email subject line
        body: Email body content

    Returns:
        True on success, False on failure
    """
    import json

    # Get default template info for name/description
    default = config.DEFAULT_TEMPLATES.get(template_id)
    if not default:
        return False

    content_name = config.TEMPLATE_CONTENT_PREFIX + template_id
    template_data = {
        'name': default['name'],
        'description': default['description'],
        'subject': subject,
        'body': body
    }

    try:
        # Decode HTML entities that were encoded for safe transmission
        decoded_subject = subject.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&#39;', "'").replace('&amp;', '&')
        decoded_body = body.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&#39;', "'").replace('&amp;', '&')

        template_data['subject'] = decoded_subject
        template_data['body'] = decoded_body

        model.WriteContentText(content_name, json.dumps(template_data), "")
        return True
    except:
        return False


def delete_template(template_id):
    """
    Delete a custom template, reverting to default.

    Args:
        template_id: Template identifier

    Returns:
        True on success, False on failure
    """
    content_name = config.TEMPLATE_CONTENT_PREFIX + template_id
    try:
        # Write empty content to effectively delete
        model.WriteContentText(content_name, "", "")
        return True
    except:
        return False


def handle_get_templates():
    """Handle AJAX request to get all email templates."""
    import json

    try:
        templates = []
        # Use list() to ensure we get the keys properly in IronPython
        template_ids = list(config.DEFAULT_TEMPLATES.keys())

        for template_id in template_ids:
            template = get_template(template_id)
            if template:
                templates.append(template)

        result = {
            'success': True,
            'templates': templates,
            'count': len(templates)
        }
        print json.dumps(result)
    except Exception as e:
        import traceback
        result = {
            'success': False,
            'message': 'Error loading templates: ' + str(e),
            'traceback': traceback.format_exc()
        }
        print json.dumps(result)

    return True


def handle_save_template():
    """Handle AJAX request to save an email template."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        template_id = str(model.Data.template_id) if hasattr(model.Data, 'template_id') and model.Data.template_id else None
        subject = str(model.Data.subject) if hasattr(model.Data, 'subject') and model.Data.subject else ''
        body = str(model.Data.body) if hasattr(model.Data, 'body') and model.Data.body else ''

        if not template_id:
            print json.dumps({'success': False, 'message': 'Missing template ID'})
            return True

        if save_template(template_id, subject, body):
            print json.dumps({'success': True, 'message': 'Template saved successfully'})
        else:
            print json.dumps({'success': False, 'message': 'Failed to save template'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_reset_template():
    """Handle AJAX request to reset a template to default."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        template_id = str(model.Data.template_id) if hasattr(model.Data, 'template_id') and model.Data.template_id else None

        if not template_id:
            print json.dumps({'success': False, 'message': 'Missing template ID'})
            return True

        if delete_template(template_id):
            # Return the default template
            default = config.DEFAULT_TEMPLATES.get(template_id)
            if default:
                print json.dumps({
                    'success': True,
                    'message': 'Template reset to default',
                    'template': {
                        'id': template_id,
                        'name': default['name'],
                        'description': default['description'],
                        'subject': default['subject'],
                        'body': default['body'],
                        'is_custom': False
                    }
                })
            else:
                print json.dumps({'success': True, 'message': 'Template reset'})
        else:
            print json.dumps({'success': False, 'message': 'Failed to reset template'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


# Storage key for dropdown templates
DROPDOWN_TEMPLATES_STORAGE_KEY = "MissionsDashboard_DropdownTemplates"

# Storage key for global dashboard settings
GLOBAL_SETTINGS_STORAGE_KEY = "MissionsDashboard_GlobalSettings"

# Default global settings
DEFAULT_GLOBAL_SETTINGS = {
    'approvals_enabled': True,  # Global toggle for approval workflow
    'require_approval_for_new_trips': True,  # Default for new trips
}


def get_global_settings():
    """Get global dashboard settings from storage or return defaults."""
    import json
    try:
        stored = model.TextContent(GLOBAL_SETTINGS_STORAGE_KEY)
        if stored:
            settings = json.loads(stored)
            # Merge with defaults to ensure all keys exist
            merged = DEFAULT_GLOBAL_SETTINGS.copy()
            merged.update(settings)
            return merged
    except:
        pass
    return DEFAULT_GLOBAL_SETTINGS.copy()


def save_global_settings(settings):
    """Save global dashboard settings to storage."""
    import json
    try:
        content_str = json.dumps(settings)
        model.WriteContentText(GLOBAL_SETTINGS_STORAGE_KEY, content_str, "")
        # Verify the save worked by reading back
        stored = model.TextContent(GLOBAL_SETTINGS_STORAGE_KEY)
        if stored:
            return True
        return False
    except Exception as e:
        # Log the error for debugging
        return False


def get_trip_approval_setting(org_id):
    """Get the approval setting for a specific trip.

    Returns: 'enabled', 'disabled', or 'inherit' (use global setting)
    """
    try:
        setting = model.ExtraValueOrg(int(org_id), 'TripApprovalsEnabled')
        if setting in ['enabled', 'disabled', 'inherit']:
            return setting
    except:
        pass
    return 'inherit'


def set_trip_approval_setting(org_id, setting):
    """Set the approval setting for a specific trip.

    Args:
        org_id: Organization ID
        setting: 'enabled', 'disabled', or 'inherit'
    """
    try:
        # Always save the setting value (including 'inherit')
        # TouchPoint may not have DeleteExtraValueOrg, so we store 'inherit' as a value
        model.AddExtraValueCodeOrg(int(org_id), 'TripApprovalsEnabled', setting)
        return True
    except:
        return False


def are_approvals_enabled_for_trip(org_id):
    """Check if approvals are enabled for a specific trip.

    Logic:
    1. Check trip-specific override first
    2. If 'inherit' or not set, use global setting
    3. Default to True (enabled) for backward compatibility

    Returns: True if approvals should be shown, False otherwise
    """
    # Check trip-specific setting
    trip_setting = get_trip_approval_setting(org_id)

    if trip_setting == 'enabled':
        return True
    elif trip_setting == 'disabled':
        return False

    # Inherit from global setting
    global_settings = get_global_settings()
    return global_settings.get('approvals_enabled', True)


# =============================================================================
# TRIP FEE SETTINGS - Trip Cost and Deposit Amount for Approval Workflow
# =============================================================================

def get_trip_fee_settings(org_id):
    """Get the trip cost and deposit amount settings for a specific trip.

    These values are used when TouchPoint's built-in trip cost is not set,
    or as an override for the approval workflow email notifications.

    Returns: dict with 'trip_cost' and 'deposit_amount' (both integers, 0 if not set)
    """
    try:
        # Use ExtraValueOrg to get string values, then parse as integers
        trip_cost_str = model.ExtraValueOrg(int(org_id), 'TripCostOverride')
        deposit_amount_str = model.ExtraValueOrg(int(org_id), 'TripDepositAmount')

        trip_cost = 0
        deposit_amount = 0

        if trip_cost_str:
            try:
                trip_cost = int(trip_cost_str)
            except:
                pass

        if deposit_amount_str:
            try:
                deposit_amount = int(deposit_amount_str)
            except:
                pass

        return {
            'trip_cost': trip_cost,
            'deposit_amount': deposit_amount
        }
    except:
        return {'trip_cost': 0, 'deposit_amount': 0}


def set_trip_fee_settings(org_id, trip_cost, deposit_amount):
    """Set the trip cost and deposit amount settings for a specific trip.

    Args:
        org_id: Organization ID
        trip_cost: Trip cost in dollars (integer)
        deposit_amount: Required deposit amount in dollars (integer)

    Returns: True if successful, False otherwise
    """
    try:
        org_id_int = int(org_id)
        # Store as code values (strings) using AddExtraValueCodeOrg
        trip_cost_str = str(int(trip_cost)) if trip_cost and int(trip_cost) > 0 else '0'
        deposit_amount_str = str(int(deposit_amount)) if deposit_amount and int(deposit_amount) > 0 else '0'

        model.AddExtraValueCodeOrg(org_id_int, 'TripCostOverride', trip_cost_str)
        model.AddExtraValueCodeOrg(org_id_int, 'TripDepositAmount', deposit_amount_str)

        return True
    except Exception as e:
        return False


def get_effective_trip_cost(org_id):
    """Get the effective trip cost for a trip.

    Priority:
    1. TripCostOverride extra value (if set and > 0)
    2. TouchPoint's built-in IndAmt from TransactionSummary

    Returns: Trip cost in dollars (integer)
    """
    # First check for override
    fee_settings = get_trip_fee_settings(org_id)
    if fee_settings['trip_cost'] > 0:
        return fee_settings['trip_cost']

    # Fall back to TouchPoint's built-in trip cost
    try:
        sql = '''
        SELECT TOP 1 ts.IndAmt
        FROM TransactionSummary ts WITH (NOLOCK)
        JOIN OrganizationMembers om WITH (NOLOCK) ON ts.OrganizationId = om.OrganizationId
            AND ts.PeopleId = om.PeopleId AND ts.RegId = om.TranId
        WHERE ts.OrganizationId = {0} AND ts.IndAmt > 0
        '''.format(int(org_id))
        result = list(q.QuerySql(sql))
        if result and result[0].IndAmt:
            return int(result[0].IndAmt)
    except:
        pass

    return 0


# Default dropdown templates (used if none saved)
DEFAULT_DROPDOWN_TEMPLATES = [
    {
        'id': 'goal_reminder_team',
        'name': 'Fundraising Goal Reminder (Team)',
        'context': 'all',
        'type': 'team',
        'role': 'all',
        'subject': '{{TripName}} - Fundraising Goal Reminder',
        'body': 'Hi {{PersonName}},\\n\\nThis is a friendly reminder about your fundraising goals for {{TripName}}.\\n\\nPlease check your current fundraising status and reach out if you need any help meeting your goal.\\n\\nView your payment status here:\\n{{MyGivingLink}}\\n\\nShare this link with friends and family who want to support your trip:\\n{{SupportLink}}\\n\\nThank you for your commitment to this mission trip!\\n\\nBlessings'
    },
    {
        'id': 'goal_reminder_individual',
        'name': 'Fundraising Goal Reminder (Individual)',
        'context': 'all',
        'type': 'individual',
        'role': 'all',
        'subject': '{{TripName}} - Fundraising Goal Reminder',
        'body': 'Hi {{PersonName}},\\n\\nThis is a friendly reminder about your fundraising goal for {{TripName}}.\\n\\nYour trip cost: {{TripCost}}\\nAmount still needed: {{Outstanding}}\\n\\nView your payment status here:\\n{{MyGivingLink}}\\n\\nShare this link with friends and family who want to support your trip:\\n{{SupportLink}}\\n\\nPlease reach out if you need any help meeting your goal.\\n\\nThank you for your commitment to this mission trip!\\n\\nBlessings'
    },
    {
        'id': 'passport_request',
        'name': 'Passport Information Request',
        'context': 'all',
        'type': 'individual',
        'role': 'admin',
        'subject': '{{TripName}} - Passport Information Needed',
        'body': 'Hi {{PersonName}},\\n\\nWe need your passport information for the upcoming {{TripName}} mission trip.\\n\\nPlease complete the passport form at the link below:\\n\\n{{ChurchUrl}}/OnlineReg/3421\\n\\nThis form will collect your passport number, expiration date, and other travel details we need for trip planning.\\n\\nPlease complete this as soon as possible so we can finalize travel arrangements.\\n\\nThank you!\\n\\nBlessings'
    },
    {
        'id': 'meeting_reminder',
        'name': 'Meeting Reminder',
        'context': 'meetings',
        'type': 'team',
        'role': 'all',
        'subject': '{{TripName}} - Meeting Reminder',
        'body': 'Hi Team,\\n\\nThis is a reminder about our upcoming meeting for {{TripName}}.\\n\\nPlease make sure to attend. If you cannot make it, please let us know as soon as possible.\\n\\nSee you there!\\n\\nBlessings'
    },
    {
        'id': 'welcome_team',
        'name': 'Welcome to the Team',
        'context': 'team',
        'type': 'individual',
        'role': 'admin',
        'subject': '{{TripName}} - Welcome to the Team!',
        'body': 'Hi {{PersonName}},\\n\\nWelcome to the {{TripName}} mission team! We are excited to have you join us.\\n\\nOver the coming weeks, you will receive information about meetings, fundraising goals, and preparation for the trip.\\n\\nIf you have any questions, please do not hesitate to reach out.\\n\\nBlessings'
    },
    {
        'id': 'approval_notification',
        'name': 'Registration Approved',
        'context': 'team',
        'type': 'individual',
        'role': 'admin',
        'subject': '{{TripName}} - Registration Approved!',
        'body': 'Dear {{PersonName}},\\n\\nGreat news! Your registration for {{TripName}} has been approved!\\n\\nTRIP DETAILS:\\n- Trip Cost: {{TripCost}}\\n- Required Deposit: {{DepositAmount}}\\n\\nNEXT STEPS:\\n1. Pay your deposit as soon as possible to secure your spot\\n2. Start fundraising for the remaining balance\\n3. Share your personal fundraising page with friends and family\\n\\nPAYMENT OPTIONS:\\nView your payment status and make payments here:\\n{{MyGivingLink}}\\n\\nFUNDRAISING:\\nShare this link with supporters who want to help fund your trip:\\n{{SupportLink}}\\n\\nIf you have any questions, please do not hesitate to reach out.\\n\\nWe are excited to have you on this mission trip!\\n\\nBlessings'
    },
    {
        'id': 'payment_deadline',
        'name': 'Payment Deadline Reminder',
        'context': 'budget',
        'type': 'both',
        'role': 'admin',
        'subject': '{{TripName}} - Payment Deadline Approaching',
        'body': 'Hi {{PersonName}},\\n\\nThis is a reminder that a payment deadline is approaching for {{TripName}}.\\n\\nPlease check your payment status and make sure you are on track with your fundraising goal.\\n\\nView your payment status here:\\n{{MyGivingLink}}\\n\\nIf you have any questions about your account or need assistance, please let us know.\\n\\nBlessings'
    },
    {
        'id': 'general_update',
        'name': 'General Team Update',
        'context': 'all',
        'type': 'team',
        'role': 'all',
        'subject': '{{TripName}} - Team Update',
        'body': 'Hi Team,\\n\\nWe wanted to share a quick update about {{TripName}}.\\n\\n[Add your update here]\\n\\nIf you have any questions, please let us know.\\n\\nBlessings'
    },
    {
        'id': 'document_request',
        'name': 'Document Request',
        'context': 'all',
        'type': 'individual',
        'role': 'admin',
        'subject': '{{TripName}} - Document Needed',
        'body': 'Hi {{PersonName}},\\n\\nWe need you to submit the following document(s) for {{TripName}}:\\n\\n[List documents needed]\\n\\nPlease submit these as soon as possible to help us finalize trip preparations.\\n\\nThank you!\\n\\nBlessings'
    },
    {
        'id': 'encouragement',
        'name': 'Team Encouragement',
        'context': 'all',
        'type': 'both',
        'role': 'leader',
        'subject': '{{TripName}} - A Word of Encouragement',
        'body': 'Hi {{PersonName}},\\n\\nI wanted to reach out and encourage you as we prepare for {{TripName}}.\\n\\n[Add your personal message here]\\n\\nThank you for being part of this team!\\n\\nBlessings'
    }
]


def get_dropdown_templates():
    """Get all dropdown templates - merges defaults with user-saved templates.

    Built-in default templates are always included and can be customized.
    User-added templates are preserved separately.
    """
    import json

    # Start with a copy of defaults (these will always be present)
    default_ids = set(t['id'] for t in DEFAULT_DROPDOWN_TEMPLATES)
    result = list(DEFAULT_DROPDOWN_TEMPLATES)  # Copy of defaults

    try:
        stored = model.TextContent(DROPDOWN_TEMPLATES_STORAGE_KEY)
        if stored:
            stored_templates = json.loads(stored)
            stored_by_id = {t['id']: t for t in stored_templates}

            # Update defaults with any customizations from stored
            for i, t in enumerate(result):
                if t['id'] in stored_by_id:
                    # User customized this default template - use their version
                    result[i] = stored_by_id[t['id']]

            # Add any user-created templates (not in defaults)
            for t in stored_templates:
                if t['id'] not in default_ids:
                    result.append(t)
    except:
        pass

    return result


def save_dropdown_templates(templates):
    """Save dropdown templates to storage."""
    import json
    try:
        model.WriteContentText(DROPDOWN_TEMPLATES_STORAGE_KEY, json.dumps(templates), "")
        return True
    except:
        return False


def handle_get_dropdown_templates():
    """Handle AJAX request to get email dropdown templates."""
    import json

    try:
        templates = get_dropdown_templates()
        result = {
            'success': True,
            'templates': templates,
            'count': len(templates)
        }
        print json.dumps(result)
    except Exception as e:
        import traceback
        result = {
            'success': False,
            'message': 'Error loading dropdown templates: ' + str(e),
            'traceback': traceback.format_exc()
        }
        print json.dumps(result)

    return True


def handle_save_dropdown_template():
    """Handle AJAX request to save an email dropdown template.

    NOTE: All parameter names use 'tpl_' prefix to avoid ASP.NET reserved words
    like 'type' and 'role' which cause blank responses.
    """
    import json
    import base64

    def base64_decode(s):
        """Decode Base64 string back to UTF-8 text."""
        if not s:
            return ''
        try:
            # Base64 decode - in Python 2.7, this returns a string directly
            # The JS encoding is: btoa(unescape(encodeURIComponent(str)))
            # which produces UTF-8 bytes encoded as Base64
            decoded = base64.b64decode(s)
            # In Python 2.7, decoded is already a string
            # but we need to handle UTF-8 properly
            if isinstance(decoded, bytes):
                return decoded.decode('utf-8')
            return decoded
        except Exception as e:
            # Return original if decode fails
            return s

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        # Get template data from request (using tpl_ prefix to avoid ASP.NET reserved words)
        template_id = str(model.Data.tpl_id) if hasattr(model.Data, 'tpl_id') and model.Data.tpl_id else None
        name = str(model.Data.tpl_name) if hasattr(model.Data, 'tpl_name') and model.Data.tpl_name else None
        context = str(model.Data.tpl_context) if hasattr(model.Data, 'tpl_context') and model.Data.tpl_context else 'trip'
        template_type = str(model.Data.tpl_type) if hasattr(model.Data, 'tpl_type') and model.Data.tpl_type else 'general'
        role = str(model.Data.tpl_role) if hasattr(model.Data, 'tpl_role') and model.Data.tpl_role else 'all'

        # Get Base64-encoded subject and body (tpl_ prefix)
        b64_subject = str(model.Data.tpl_subject) if hasattr(model.Data, 'tpl_subject') and model.Data.tpl_subject else ''
        b64_body = str(model.Data.tpl_body) if hasattr(model.Data, 'tpl_body') and model.Data.tpl_body else ''

        # Decode Base64 (JavaScript encodes these to bypass ASP.NET request validation)
        subject = base64_decode(b64_subject)
        body = base64_decode(b64_body)

        # Validate required fields
        if not template_id or not name or not subject or not body:
            print json.dumps({'success': False, 'message': 'Missing required fields (id, name, subject, body)'})
            return True

        # Convert escaped newlines back to actual newlines
        body = body.replace('\\n', '\n')

        # Get existing templates
        templates = get_dropdown_templates()

        # Check if this is an update or new template
        existing_index = None
        for i, t in enumerate(templates):
            if t['id'] == template_id:
                existing_index = i
                break

        # Create template object
        template = {
            'id': template_id,
            'name': name,
            'context': context,
            'type': template_type,
            'role': role,
            'subject': subject,
            'body': body
        }

        if existing_index is not None:
            # Update existing
            templates[existing_index] = template
        else:
            # Add new
            templates.append(template)

        # Save templates
        if save_dropdown_templates(templates):
            print json.dumps({'success': True, 'message': 'Template saved successfully'})
        else:
            print json.dumps({'success': False, 'message': 'Failed to save template'})

    except Exception as e:
        import traceback
        print json.dumps({
            'success': False,
            'message': 'Error saving template: ' + str(e),
            'traceback': traceback.format_exc()
        })

    return True


def handle_delete_dropdown_template():
    """Handle AJAX request to delete an email dropdown template."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        template_id = str(model.Data.template_id) if hasattr(model.Data, 'template_id') and model.Data.template_id else None

        if not template_id:
            print json.dumps({'success': False, 'message': 'Missing template ID'})
            return True

        # Get existing templates
        templates = get_dropdown_templates()

        # Remove the template
        new_templates = [t for t in templates if t['id'] != template_id]

        if len(new_templates) == len(templates):
            print json.dumps({'success': False, 'message': 'Template not found'})
            return True

        # Save templates
        if save_dropdown_templates(new_templates):
            print json.dumps({'success': True, 'message': 'Template deleted successfully'})
        else:
            print json.dumps({'success': False, 'message': 'Failed to delete template'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


# =====================================================
# GLOBAL SETTINGS HANDLERS
# =====================================================

def handle_get_config():
    """Handle AJAX request to get dashboard configuration."""
    import json
    try:
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        data = {
            'ACTIVE_ORG_STATUS_ID': config.ACTIVE_ORG_STATUS_ID,
            'MISSION_TRIP_FLAG': config.MISSION_TRIP_FLAG,
            'MEMBER_TYPE_LEADER': config.MEMBER_TYPE_LEADER,
            'ATTENDANCE_TYPE_LEADER': config.ATTENDANCE_TYPE_LEADER,
            'ITEMS_PER_PAGE': config.ITEMS_PER_PAGE,
            'SHOW_CLOSED_BY_DEFAULT': config.SHOW_CLOSED_BY_DEFAULT,
            'ENABLE_SQL_DEBUG': config.ENABLE_SQL_DEBUG,
            'MY_MISSIONS_LINK': config.MY_MISSIONS_LINK,
            'CURRENCY_SYMBOL': config.CURRENCY_SYMBOL,
            'ADMIN_ROLES': ','.join(config.ADMIN_ROLES) if isinstance(config.ADMIN_ROLES, list) else config.ADMIN_ROLES,
            'FINANCE_ROLES': ','.join(config.FINANCE_ROLES) if isinstance(config.FINANCE_ROLES, list) else config.FINANCE_ROLES,
            'LEADER_MEMBER_TYPES': ','.join(str(x) for x in config.LEADER_MEMBER_TYPES) if isinstance(config.LEADER_MEMBER_TYPES, list) else config.LEADER_MEMBER_TYPES,
            'SIDEBAR_BG_COLOR': config.SIDEBAR_BG_COLOR,
            'SIDEBAR_ACTIVE_COLOR': config.SIDEBAR_ACTIVE_COLOR,
            'APPLICATION_ORG_IDS': ','.join(str(x) for x in config.APPLICATION_ORG_IDS) if isinstance(config.APPLICATION_ORG_IDS, list) else config.APPLICATION_ORG_IDS,
            'ENABLE_STATS_TAB': config.ENABLE_STATS_TAB,
            'ENABLE_FINANCE_TAB': config.ENABLE_FINANCE_TAB,
            'ENABLE_MESSAGES_TAB': config.ENABLE_MESSAGES_TAB,
        }
        print json.dumps({'success': True, 'config': data})
    except Exception as e:
        print json.dumps({'success': False, 'message': str(e)})
    return True


def handle_save_config():
    """Handle AJAX request to save dashboard configuration."""
    import json
    try:
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        # Read each setting from the request
        fields = [
            ('ACTIVE_ORG_STATUS_ID', 'int'), ('MISSION_TRIP_FLAG', 'int'),
            ('MEMBER_TYPE_LEADER', 'int'), ('ATTENDANCE_TYPE_LEADER', 'int'),
            ('ITEMS_PER_PAGE', 'int'), ('SHOW_CLOSED_BY_DEFAULT', 'bool'),
            ('ENABLE_SQL_DEBUG', 'bool'), ('MY_MISSIONS_LINK', 'str'),
            ('CURRENCY_SYMBOL', 'str'), ('ADMIN_ROLES', 'list'),
            ('FINANCE_ROLES', 'list'), ('LEADER_MEMBER_TYPES', 'intlist'),
            ('SIDEBAR_BG_COLOR', 'str'), ('SIDEBAR_ACTIVE_COLOR', 'str'),
            ('APPLICATION_ORG_IDS', 'intlist'),
            ('ENABLE_STATS_TAB', 'bool'), ('ENABLE_FINANCE_TAB', 'bool'),
            ('ENABLE_MESSAGES_TAB', 'bool'),
        ]

        for field_name, field_type in fields:
            if hasattr(Data, field_name):
                raw = str(getattr(Data, field_name)).strip()
                # Skip empty values for non-string/non-bool types to preserve defaults
                if not raw and field_type in ('int', 'list', 'intlist'):
                    continue
                if field_type == 'int':
                    try:
                        setattr(config, field_name, int(raw))
                    except:
                        continue  # Keep default if parse fails
                elif field_type == 'bool':
                    setattr(config, field_name, raw.lower() in ('true', '1', 'yes'))
                elif field_type == 'list':
                    parsed = [x.strip() for x in raw.split(',') if x.strip()]
                    if parsed:
                        setattr(config, field_name, parsed)
                elif field_type == 'intlist':
                    try:
                        parsed = [int(x.strip()) for x in raw.split(',') if x.strip()]
                        if parsed:
                            setattr(config, field_name, parsed)
                    except:
                        continue  # Keep default if parse fails
                else:
                    setattr(config, field_name, raw)

        save_config(config)
        print json.dumps({'success': True, 'message': 'Configuration saved'})
    except Exception as e:
        print json.dumps({'success': False, 'message': str(e)})
    return True


def handle_get_global_settings():
    """Handle AJAX request to get global dashboard settings."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        settings = get_global_settings()
        print json.dumps({
            'success': True,
            'settings': settings
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_save_global_settings():
    """Handle AJAX request to save global dashboard settings."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        # Get settings from request
        approvals_enabled = str(model.Data.approvals_enabled) if hasattr(model.Data, 'approvals_enabled') else 'true'
        new_approval_value = approvals_enabled.lower() == 'true'

        settings = get_global_settings()
        settings['approvals_enabled'] = new_approval_value

        if save_global_settings(settings):
            # Verify by reading back
            verified_settings = get_global_settings()
            print json.dumps({
                'success': True,
                'message': 'Settings saved successfully',
                'settings': verified_settings,
                'approvals_enabled': verified_settings.get('approvals_enabled', True)
            })
        else:
            print json.dumps({'success': False, 'message': 'Failed to save settings to storage'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error saving settings: ' + str(e)})

    return True


def handle_get_trip_approval_setting():
    """Handle AJAX request to get approval setting for a specific trip."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None

        if not org_id:
            print json.dumps({'success': False, 'message': 'Missing organization ID'})
            return True

        trip_setting = get_trip_approval_setting(org_id)
        is_enabled = are_approvals_enabled_for_trip(org_id)
        global_settings = get_global_settings()

        print json.dumps({
            'success': True,
            'trip_setting': trip_setting,
            'is_enabled': is_enabled,
            'global_enabled': global_settings.get('approvals_enabled', True)
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_save_trip_approval_setting():
    """Handle AJAX request to save approval setting for a specific trip."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        setting = str(model.Data.setting) if hasattr(model.Data, 'setting') and model.Data.setting else 'inherit'

        if not org_id:
            print json.dumps({'success': False, 'message': 'Missing organization ID'})
            return True

        if setting not in ['enabled', 'disabled', 'inherit']:
            print json.dumps({'success': False, 'message': 'Invalid setting value'})
            return True

        if set_trip_approval_setting(org_id, setting):
            is_enabled = are_approvals_enabled_for_trip(org_id)
            print json.dumps({
                'success': True,
                'message': 'Trip approval setting saved',
                'trip_setting': setting,
                'is_enabled': is_enabled
            })
        else:
            print json.dumps({'success': False, 'message': 'Failed to save setting'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_get_trip_fee_settings():
    """Handle AJAX request to get trip fee settings (cost and deposit)."""
    import json

    try:
        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None

        if not org_id:
            print json.dumps({'success': False, 'message': 'Missing organization ID'})
            return True

        fee_settings = get_trip_fee_settings(org_id)
        effective_cost = get_effective_trip_cost(org_id)

        print json.dumps({
            'success': True,
            'trip_cost': fee_settings.get('trip_cost', 0),
            'deposit_amount': fee_settings.get('deposit_amount', 0),
            'effective_trip_cost': effective_cost
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_save_trip_fee_settings():
    """Handle AJAX request to save trip fee settings (cost and deposit)."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        trip_cost = str(model.Data.trip_cost) if hasattr(model.Data, 'trip_cost') and model.Data.trip_cost else '0'
        deposit_amount = str(model.Data.deposit_amount) if hasattr(model.Data, 'deposit_amount') and model.Data.deposit_amount else '0'

        if not org_id:
            print json.dumps({'success': False, 'message': 'Missing organization ID'})
            return True

        # Convert to integers
        try:
            trip_cost_int = int(float(trip_cost)) if trip_cost else 0
            deposit_amount_int = int(float(deposit_amount)) if deposit_amount else 0
        except ValueError:
            print json.dumps({'success': False, 'message': 'Invalid cost or deposit value'})
            return True

        if set_trip_fee_settings(org_id, trip_cost_int, deposit_amount_int):
            print json.dumps({
                'success': True,
                'message': 'Trip fee settings saved',
                'trip_cost': trip_cost_int,
                'deposit_amount': deposit_amount_int
            })
        else:
            print json.dumps({'success': False, 'message': 'Failed to save settings'})

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


# =====================================================
# TRIP REGISTRATION APPROVAL WORKFLOW HANDLERS
# =====================================================

def handle_get_approval_statuses():
    """Handle AJAX request to get approval statuses for all trip members."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None

        if not org_id:
            print json.dumps({'success': False, 'message': 'Missing organization ID'})
            return True

        # Get all approval statuses
        statuses = _get_all_approval_statuses(org_id)

        # Convert to JSON-serializable format (peopleId as string key)
        status_dict = {}
        for pid, status_info in statuses.items():
            status_dict[str(pid)] = status_info

        print json.dumps({
            'success': True,
            'statuses': status_dict
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_get_registration_answers():
    """Handle AJAX request to get registration form Q&A for a person."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None

        if not org_id or not people_id:
            print json.dumps({'success': False, 'message': 'Missing organization or person ID'})
            return True

        # Get registration answers
        answers = _get_registration_answers(org_id, people_id)

        # Get person info
        person_sql = '''
        SELECT p.Name2, p.EmailAddress, p.CellPhone, p.Age,
               om.EnrollmentDate
        FROM People p WITH (NOLOCK)
        LEFT JOIN OrganizationMembers om ON om.PeopleId = p.PeopleId AND om.OrganizationId = {0}
        WHERE p.PeopleId = {1}
        '''.format(org_id, people_id)

        person_result = list(q.QuerySql(person_sql))
        person_info = None
        if person_result:
            p = person_result[0]
            person_info = {
                'people_id': int(people_id),
                'name': p.Name2 or '',
                'email': p.EmailAddress or '',
                'phone': p.CellPhone or '',
                'age': p.Age if p.Age else None,
                'enrollment_date': str(p.EnrollmentDate)[:10] if p.EnrollmentDate else None
            }

        # Get current approval status
        status = _get_approval_status(org_id, people_id)

        print json.dumps({
            'success': True,
            'person': person_info,
            'answers': answers,
            'status': status
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_approve_member():
    """Handle AJAX request to approve a trip member."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None

        if not org_id or not people_id:
            print json.dumps({'success': False, 'message': 'Missing organization or person ID'})
            return True

        org_id_int = int(org_id)
        people_id_int = int(people_id)

        # Remove from trip-denied if present
        if model.InSubGroup(people_id_int, org_id_int, 'trip-denied'):
            model.RemoveSubGroup(people_id_int, org_id_int, 'trip-denied')

        # Remove denial reason if any
        _remove_denial_reason(org_id_int, people_id_int)

        # Add to trip-approved subgroup
        model.AddSubGroup(people_id_int, org_id_int, 'trip-approved')

        # Get person details for email prompt
        person = model.GetPerson(people_id_int)
        person_name = person.Name2 if person else 'Member'
        person_first = person.NickName or person.FirstName if person else 'Member'
        person_email = person.EmailAddress if person else ''

        # Get trip info
        trip = _get_trip_info(org_id_int)
        trip_name = trip.OrganizationName if trip else 'Mission Trip'

        # Get fee settings
        fee_settings = get_trip_fee_settings(org_id_int)
        effective_cost = get_effective_trip_cost(org_id_int)
        deposit_amount = fee_settings.get('deposit_amount', 0)

        print json.dumps({
            'success': True,
            'message': 'Member approved successfully',
            'status': 'approved',
            'prompt_email': True,
            'person': {
                'people_id': people_id_int,
                'name': person_name,
                'first_name': person_first,
                'email': person_email
            },
            'trip': {
                'org_id': org_id_int,
                'name': trip_name,
                'cost': effective_cost,
                'deposit': deposit_amount
            }
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_deny_member():
    """Handle AJAX request to deny a trip member."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None
        reason = str(model.Data.reason) if hasattr(model.Data, 'reason') and model.Data.reason else ''

        if not org_id or not people_id:
            print json.dumps({'success': False, 'message': 'Missing organization or person ID'})
            return True

        if not reason:
            print json.dumps({'success': False, 'message': 'Denial reason is required'})
            return True

        org_id_int = int(org_id)
        people_id_int = int(people_id)

        # Get person name and current user info for denial record
        person = model.GetPerson(people_id_int)
        person_name = person.Name2 if person else 'Unknown'

        current_user_id = model.UserPeopleId
        current_user = model.GetPerson(current_user_id)
        current_user_name = current_user.Name2 if current_user else 'Admin'

        # Remove from trip-approved if present
        if model.InSubGroup(people_id_int, org_id_int, 'trip-approved'):
            model.RemoveSubGroup(people_id_int, org_id_int, 'trip-approved')

        # Add to trip-denied subgroup
        model.AddSubGroup(people_id_int, org_id_int, 'trip-denied')

        # Store denial reason
        _add_denial_reason(org_id_int, people_id_int, person_name, reason, current_user_id, current_user_name)

        print json.dumps({
            'success': True,
            'message': 'Member denied',
            'status': 'denied'
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_revoke_approval():
    """Handle AJAX request to revoke approval/denial (return to pending)."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None

        if not org_id or not people_id:
            print json.dumps({'success': False, 'message': 'Missing organization or person ID'})
            return True

        org_id_int = int(org_id)
        people_id_int = int(people_id)

        # Remove from both subgroups
        if model.InSubGroup(people_id_int, org_id_int, 'trip-approved'):
            model.RemoveSubGroup(people_id_int, org_id_int, 'trip-approved')

        if model.InSubGroup(people_id_int, org_id_int, 'trip-denied'):
            model.RemoveSubGroup(people_id_int, org_id_int, 'trip-denied')

        # Remove denial reason if any
        _remove_denial_reason(org_id_int, people_id_int)

        print json.dumps({
            'success': True,
            'message': 'Status revoked, member returned to pending',
            'status': 'pending'
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_set_member_type():
    """Handle AJAX request to set a member's type (Leader or Member) in an organization."""
    import json

    try:
        # Check admin permission
        user_role = get_user_role_and_trips()
        if not user_role.get('is_admin', False):
            print json.dumps({'success': False, 'message': 'Access denied. Admin role required.'})
            return True

        org_id = str(model.Data.org_id) if hasattr(model.Data, 'org_id') and model.Data.org_id else None
        people_id = str(model.Data.people_id) if hasattr(model.Data, 'people_id') and model.Data.people_id else None
        member_type = str(model.Data.mem_role) if hasattr(model.Data, 'mem_role') and model.Data.mem_role else None

        if not org_id or not people_id or not member_type:
            print json.dumps({'success': False, 'message': 'Missing required parameters'})
            return True

        if member_type not in ('Leader', 'Member'):
            print json.dumps({'success': False, 'message': 'Invalid member type. Must be Leader or Member.'})
            return True

        org_id_int = int(org_id)
        people_id_int = int(people_id)

        model.SetMemberType(people_id_int, org_id_int, member_type)

        print json.dumps({
            'success': True,
            'message': 'Member type set to ' + member_type,
            'member_type': member_type
        })

    except Exception as e:
        print json.dumps({'success': False, 'message': 'Error: ' + str(e)})

    return True


def handle_ajax_section_load():
    """Handle AJAX requests to load trip sections dynamically."""
    trip_id = str(model.Data.trip) if hasattr(model.Data, 'trip') and model.Data.trip else None
    section = str(model.Data.section) if hasattr(model.Data, 'section') and model.Data.section else 'overview'

    if not trip_id:
        print '<div class="alert alert-danger">Missing trip ID.</div>'
        return True

    # Get user role for access control
    user_role = get_user_role_and_trips()

    # Check access
    if not has_trip_access(user_role, trip_id):
        print '<div class="alert alert-danger">Access denied.</div>'
        return True

    # Render the requested section
    print render_trip_section(trip_id, section, user_role)
    return True

def get_queue_statistics():
    """Get statistics about people in the application queue"""
    
    app_org_list = ','.join(str(x) for x in config.APPLICATION_ORG_IDS)
    
    queue_stats_sql = '''
    WITH QueuePeople AS (
        SELECT DISTINCT 
            p.PeopleId,
            p.Name2 AS Name,
            p.Age,
            p.GenderId,
            p.MemberStatusId,
            om.EnrollmentDate
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN OrganizationMembers om WITH (NOLOCK) ON o.OrganizationId = om.OrganizationId
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
        WHERE o.OrganizationId IN ({0})
          AND om.InactiveDate IS NULL
    ),
    PreviousMissions AS (
        SELECT 
            qp.PeopleId,
            COUNT(DISTINCT o.OrganizationId) AS PreviousTripCount
        FROM QueuePeople qp
        INNER JOIN OrganizationMembers om WITH (NOLOCK) ON qp.PeopleId = om.PeopleId
        INNER JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
        WHERE o.IsMissionTrip = {1}
          AND o.OrganizationStatusId = {2}
          AND o.OrganizationId NOT IN ({0})
        GROUP BY qp.PeopleId
    )
    SELECT 
        COUNT(DISTINCT qp.PeopleId) AS TotalInQueue,
        COUNT(DISTINCT pm.PeopleId) AS HasPreviousTrips,
        COUNT(DISTINCT CASE WHEN pm.PeopleId IS NULL THEN qp.PeopleId END) AS FirstTimers,
        AVG(qp.Age) AS AvgAge,
        MIN(qp.Age) AS MinAge,
        MAX(qp.Age) AS MaxAge,
        SUM(CASE WHEN g.Description = 'Male' THEN 1 ELSE 0 END) AS MaleCount,
        SUM(CASE WHEN g.Description = 'Female' THEN 1 ELSE 0 END) AS FemaleCount,
        SUM(CASE WHEN ms.Description = 'Member' THEN 1 ELSE 0 END) AS MemberCount,
        SUM(CASE WHEN ms.Description != 'Member' OR ms.Description IS NULL THEN 1 ELSE 0 END) AS NonMemberCount,
        DATEDIFF(DAY, MIN(qp.EnrollmentDate), GETDATE()) AS OldestApplicationDays,
        DATEDIFF(DAY, MAX(qp.EnrollmentDate), GETDATE()) AS NewestApplicationDays
    FROM QueuePeople qp
    LEFT JOIN PreviousMissions pm ON qp.PeopleId = pm.PeopleId
    LEFT JOIN lookup.Gender g WITH (NOLOCK) ON qp.GenderId = g.Id
    LEFT JOIN lookup.MemberStatus ms WITH (NOLOCK) ON qp.MemberStatusId = ms.Id
    '''.format(app_org_list, config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)
    
    return execute_query_with_debug(queue_stats_sql, "Queue Statistics Query", "top1")

def render_queue_statistics():
    """Render queue statistics section on dashboard"""
    
    queue_stats = get_queue_statistics()
    
    if not queue_stats or queue_stats.TotalInQueue == 0:
        return  # Don't show if no one in queue
    
    print '''
    <div style="margin: 20px 0; padding: 15px; background: #fff3e0; border-radius: 8px; border-left: 4px solid #ff9800;">
        <h4 style="margin-top: 0; color: #e65100;">
            <span style="font-size: 1.2em;">📋</span> Application Queue Statistics
        </h4>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
            <div>
                <div style="font-size: 20px; font-weight: bold; color: #ff6f00;">{0}</div>
                <div style="color: #666; font-size: 0.9em;">Total Applicants</div>
                <div style="margin-top: 5px; font-size: 0.85em;">
                    <span style="color: #388e3c;"><strong>{1}</strong> have been on trips</span><br>
                    <span style="color: #1976d2;"><strong>{2}</strong> first-timers</span>
                </div>
            </div>
            <div>
                <div style="font-size: 20px; font-weight: bold; color: #ff6f00;">{3} yrs</div>
                <div style="color: #666; font-size: 0.9em;">Average Age</div>
                <div style="margin-top: 5px; font-size: 0.85em;">
                    Range: {4} - {5} years
                </div>
            </div>
            <div>
                <div style="font-size: 20px; font-weight: bold; color: #ff6f00;">{6}/{7}</div>
                <div style="color: #666; font-size: 0.9em;">Gender Split</div>
                <div style="margin-top: 5px; font-size: 0.85em;">
                    Male / Female
                </div>
            </div>
            <div>
                <div style="font-size: 20px; font-weight: bold; color: #ff6f00;">{8}/{9}</div>
                <div style="color: #666; font-size: 0.9em;">Membership</div>
                <div style="margin-top: 5px; font-size: 0.85em;">
                    Members / Non-members
                </div>
            </div>
            <div>
                <div style="font-size: 20px; font-weight: bold; color: #ff6f00;">{10} days</div>
                <div style="color: #666; font-size: 0.9em;">Oldest Application</div>
                <div style="margin-top: 5px; font-size: 0.85em;">
                    Newest: {11} days ago
                </div>
            </div>
        </div>
    </div>
    '''.format(
        queue_stats.TotalInQueue,
        queue_stats.HasPreviousTrips or 0,
        queue_stats.FirstTimers or 0,
        int(queue_stats.AvgAge or 0),
        queue_stats.MinAge or 0,
        queue_stats.MaxAge or 0,
        queue_stats.MaleCount or 0,
        queue_stats.FemaleCount or 0,
        queue_stats.MemberCount or 0,
        queue_stats.NonMemberCount or 0,
        queue_stats.OldestApplicationDays or 0,
        queue_stats.NewestApplicationDays or 0
    )

def render_dashboard_deadlines():
    """Render upcoming deadlines prominently on dashboard"""
    
    # Get the mission trip totals CTE
    mission_trip_cte = get_mission_trip_totals_cte(include_closed=True)
    
    deadlines_sql = '''
        {2}
        SELECT TOP 5
            o.OrganizationName,
            o.OrganizationId,
            oe.DateValue AS DeadlineDate,
            oe.Field AS DeadlineType,
            DATEDIFF(DAY, GETDATE(), oe.DateValue) AS DaysUntil,
            ISNULL((SELECT SUM(Due) FROM MissionTripTotals WHERE InvolvementId = o.OrganizationId), 0) AS AmountRemaining
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN OrganizationExtra oe WITH (NOLOCK) ON o.OrganizationId = oe.OrganizationId
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND oe.DateValue > GETDATE()
          AND oe.DateValue <= DATEADD(DAY, 30, GETDATE())  -- Next 30 days only
          AND oe.Field IN ('Close', 'Main Event Start', 'Registration Deadline', 'Final Payment Due')
        ORDER BY oe.DateValue
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, mission_trip_cte)
    
    deadlines = execute_query_with_debug(deadlines_sql, "Dashboard Urgent Deadlines", "sql")
    
    if deadlines and len(deadlines) > 0:
        print '''
        <div class="urgent-deadlines">
            <h4 style="margin-top: 0; color: #e65100;">⚠️ Upcoming Deadlines (Next 30 Days)</h4>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 10px;">
        '''
        
        for deadline in deadlines:
            urgency_color = '#d32f2f' if deadline.DaysUntil <= 7 else '#f57c00' if deadline.DaysUntil <= 14 else '#388e3c'
            icon = '🚨' if deadline.DaysUntil <= 7 else '⚠️'
            
            # Format the amount remaining text
            if deadline.AmountRemaining > 0:
                amount_text = '<strong style="color: #d32f2f;">{0} remaining</strong>'.format(format_currency(deadline.AmountRemaining))
            else:
                amount_text = '<strong style="color: #2e7d32;">Fully paid</strong>'
            
            print '''
            <div style="background: white; padding: 10px; border-radius: 5px; border-left: 3px solid {5};">
                <div style="font-weight: bold; font-size: 12px;">
                    {6} <a href="/PyScript/Mission_Dashboard?OrgView={0}" style="color: {5};">{1}</a>
                </div>
                <div style="font-size: 11px; color: #666; margin-top: 3px;">
                    {2}: {3}<br>
                    <strong style="color: {5};">{4} days left</strong><br>
                    {7}
                </div>
            </div>
            '''.format(
                deadline.OrganizationId,
                deadline.OrganizationName[:30] + '...' if len(deadline.OrganizationName) > 30 else deadline.OrganizationName,
                deadline.DeadlineType.replace('_', ' '),
                format_date(str(deadline.DeadlineDate)),
                deadline.DaysUntil,
                urgency_color,
                icon,
                amount_text
            )
        
        print '''
            </div>
        </div>
        '''

# ::END:: Helper Functions

#####################################################################
# VIEW CONTROLLERS
#####################################################################

# ::START:: Leader Dashboard View
def render_leader_dashboard_view(user_role):
    """Dashboard view for trip leaders (non-admins).

    Shows a welcome page with only their assigned trips in a clean card layout.
    """
    accessible_trips = user_role.get('accessible_trips', [])
    user_id = user_role.get('user_id', 0)

    # Get user's name for welcome message
    user_name = "Leader"
    try:
        user_person = model.GetPerson(user_id)
        if user_person and user_person.PreferredName:
            user_name = user_person.PreferredName
        elif user_person and user_person.FirstName:
            user_name = user_person.FirstName
    except:
        pass

    print '''
    <div style="max-width: 1000px; margin: 0 auto; padding: 20px;">
        <!-- Welcome Header -->
        <div style="text-align: center; margin-bottom: 40px; padding: 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; color: white;">
            <h1 style="margin: 0 0 10px 0; font-size: 28px;">&#127758; Welcome, {0}!</h1>
            <p style="margin: 0; font-size: 16px; opacity: 0.9;">You are a leader for {1} mission trip{2}. Select a trip below to manage.</p>
        </div>
    '''.format(user_name, len(accessible_trips), 's' if len(accessible_trips) != 1 else '')

    if not accessible_trips:
        print '''
        <div style="text-align: center; padding: 60px 20px; background: #f8f9fa; border-radius: 12px;">
            <div style="font-size: 48px; margin-bottom: 20px;">&#128566;</div>
            <h2 style="color: #333; margin-bottom: 10px;">No Trips Assigned</h2>
            <p style="color: #666;">You are not currently assigned as a leader for any mission trips.</p>
            <p style="color: #666;">If you believe this is an error, please contact your administrator.</p>
        </div>
        '''
    else:
        # Group trips by status
        upcoming_trips = [t for t in accessible_trips if t.TripStatus == 'upcoming']
        active_trips = [t for t in accessible_trips if t.TripStatus == 'active']
        completed_trips = [t for t in accessible_trips if t.TripStatus == 'completed']

        print '<div style="display: grid; gap: 20px;">'

        # Active trips first (most important)
        if active_trips:
            print '<h3 style="color: #28a745; margin: 20px 0 10px 0;">&#128994; Active Trips</h3>'
            for trip in active_trips:
                print_leader_trip_card(trip, 'active')

        # Upcoming trips
        if upcoming_trips:
            print '<h3 style="color: #007bff; margin: 20px 0 10px 0;">&#128197; Upcoming Trips</h3>'
            for trip in upcoming_trips:
                print_leader_trip_card(trip, 'upcoming')

        # Completed trips (collapsed by default)
        if completed_trips:
            print '''
            <details style="margin-top: 20px;">
                <summary style="cursor: pointer; color: #6c757d; font-size: 16px; font-weight: 600; padding: 10px 0;">
                    &#9989; Completed Trips ({0})
                </summary>
                <div style="margin-top: 10px;">
            '''.format(len(completed_trips))
            for trip in completed_trips:
                print_leader_trip_card(trip, 'completed')
            print '</div></details>'

        print '</div>'

    print '</div>'


def print_leader_trip_card(trip, status):
    """Print a trip card for the leader dashboard."""
    import datetime

    # Status colors
    status_colors = {
        'active': {'bg': '#d4edda', 'border': '#28a745', 'badge': '#28a745'},
        'upcoming': {'bg': '#cce5ff', 'border': '#007bff', 'badge': '#007bff'},
        'completed': {'bg': '#f8f9fa', 'border': '#6c757d', 'badge': '#6c757d'}
    }
    colors = status_colors.get(status, status_colors['upcoming'])

    # Format dates
    start_date_str = ''
    if trip.StartDate:
        try:
            start_date_str = trip.StartDate.strftime('%b %d, %Y')
        except:
            start_date_str = str(trip.StartDate)[:10]

    end_date_str = ''
    if trip.EndDate:
        try:
            end_date_str = trip.EndDate.strftime('%b %d, %Y')
        except:
            end_date_str = str(trip.EndDate)[:10]

    date_range = ''
    if start_date_str and end_date_str:
        date_range = '{0} - {1}'.format(start_date_str, end_date_str)
    elif start_date_str:
        date_range = 'Starts {0}'.format(start_date_str)

    # Days until/since trip
    days_info = ''
    if trip.StartDate and status == 'upcoming':
        try:
            today = datetime.datetime.now()
            delta = (trip.StartDate - today).days
            if delta > 0:
                days_info = '<span style="color: {0}; font-weight: 600;">{1} days away</span>'.format(colors['badge'], delta)
            elif delta == 0:
                days_info = '<span style="color: #dc3545; font-weight: 600;">TODAY!</span>'
        except:
            pass

    print '''
    <a href="?trip={0}&section=overview" style="text-decoration: none; color: inherit; display: block;">
        <div style="background: {1}; border: 2px solid {2}; border-radius: 10px; padding: 20px;
                    transition: transform 0.2s, box-shadow 0.2s; cursor: pointer;"
             onmouseover="this.style.transform='translateY(-3px)'; this.style.boxShadow='0 8px 20px rgba(0,0,0,0.15)';"
             onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='none';">
            <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 10px;">
                <div>
                    <h3 style="margin: 0 0 8px 0; color: #333; font-size: 20px;">{3}</h3>
                    <div style="display: flex; gap: 15px; flex-wrap: wrap; align-items: center;">
                        <span style="color: #666;">&#128197; {4}</span>
                        <span style="color: #666;">&#128101; {5} members</span>
                        {6}
                    </div>
                </div>
                <div style="background: {7}; color: white; padding: 6px 14px; border-radius: 20px; font-size: 13px; font-weight: 600; text-transform: uppercase;">
                    {8}
                </div>
            </div>
            <div style="margin-top: 15px; color: #007bff; font-weight: 600;">
                Click to manage this trip &rarr;
            </div>
        </div>
    </a>
    '''.format(
        trip.OrganizationId,                          # 0
        colors['bg'],                                  # 1
        colors['border'],                              # 2
        trip.OrganizationName or 'Unknown Trip',      # 3
        date_range or 'Dates TBD',                    # 4
        trip.MemberCount or 0,                        # 5
        days_info,                                    # 6
        colors['badge'],                              # 7
        status.title()                                # 8
    )


# ::END:: Leader Dashboard View

# ::START:: Member Dashboard View (My Missions Portal)
def render_member_dashboard_view(user_role):
    """Dashboard view for trip members (non-leaders).

    Shows a personal "My Missions" portal with their trip info, payments, meetings, and documents.
    """
    import datetime

    member_trips = user_role.get('member_trips', [])
    leader_trips = user_role.get('accessible_trips', [])  # In case they're also a leader of other trips
    user_id = user_role.get('user_id', 0)

    # Get user's name for welcome message
    user_name = "Missionary"
    try:
        user_person = model.GetPerson(user_id)
        if user_person and user_person.PreferredName:
            user_name = user_person.PreferredName
        elif user_person and user_person.FirstName:
            user_name = user_person.FirstName
    except:
        pass

    # Combine all trips for the member
    all_trips = list(member_trips)
    # Add leader trips too (they can also see their leader trips here)
    for lt in leader_trips:
        if not any(t.OrganizationId == lt.OrganizationId for t in all_trips):
            all_trips.append(lt)

    total_trips = len(all_trips)

    print '''
    <div style="max-width: 900px; margin: 0 auto; padding: 20px;">
        <!-- Welcome Header -->
        <div style="text-align: center; margin-bottom: 30px; padding: 30px; background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); border-radius: 12px; color: white;">
            <h1 style="margin: 0 0 10px 0; font-size: 28px;">&#127758; My Missions</h1>
            <p style="margin: 0; font-size: 18px;">Welcome, {0}!</p>
            <p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.9;">You are registered for {1} mission trip{2}.</p>
        </div>
    '''.format(user_name, total_trips, 's' if total_trips != 1 else '')

    if not all_trips:
        print '''
        <div style="text-align: center; padding: 60px 20px; background: #f8f9fa; border-radius: 12px;">
            <div style="font-size: 48px; margin-bottom: 20px;">&#127758;</div>
            <h2 style="color: #333; margin-bottom: 10px;">No Mission Trips Found</h2>
            <p style="color: #666;">You are not currently registered for any mission trips.</p>
            <p style="color: #666;">Contact your missions director if you believe this is an error.</p>
        </div>
        '''
    else:
        # Show each trip with detailed info
        for trip in all_trips:
            render_member_trip_card(trip, user_id, user_role)

    # Bookmark reminder
    print '''
        <div style="margin-top: 30px; padding: 16px; background: #e7f3ff; border: 1px solid #b8daff; border-radius: 8px; text-align: center;">
            <strong>&#128278; Tip:</strong> Bookmark this page (<strong>{my_missions_display}</strong>) for quick access to your trip information!
        </div>
    </div>
    '''.format(my_missions_display=MY_MISSIONS_DISPLAY)


def render_member_trip_card(trip, user_id, user_role):
    """Render a detailed trip card for a member showing their personal trip info."""
    import datetime

    org_id = trip.OrganizationId
    trip_name = trip.OrganizationName or 'Mission Trip'
    status = trip.TripStatus or 'upcoming'

    # Status styling
    status_colors = {
        'active': {'bg': '#d4edda', 'border': '#28a745', 'badge_bg': '#28a745', 'text': 'Active'},
        'upcoming': {'bg': '#e7f3ff', 'border': '#007bff', 'badge_bg': '#007bff', 'text': 'Upcoming'},
        'completed': {'bg': '#f8f9fa', 'border': '#6c757d', 'badge_bg': '#6c757d', 'text': 'Completed'}
    }
    colors = status_colors.get(status, status_colors['upcoming'])

    # Format dates
    start_date_str = ''
    days_until = None
    if trip.StartDate:
        try:
            start_date_str = trip.StartDate.strftime('%B %d, %Y')
            if status == 'upcoming':
                days_until = (trip.StartDate - datetime.datetime.now()).days
        except:
            start_date_str = str(trip.StartDate)[:10]

    end_date_str = ''
    if trip.EndDate:
        try:
            end_date_str = trip.EndDate.strftime('%B %d, %Y')
        except:
            end_date_str = str(trip.EndDate)[:10]

    # Get effective trip cost (checks TripCostOverride extra value first, then TouchPoint default)
    trip_cost = get_effective_trip_cost(org_id)

    # Get member's payment info from TransactionSummary (most accurate source)
    financial_sql = '''
    SELECT
        ISNULL(SUM(ts.IndPaid), 0) AS AmountPaid,
        ISNULL(SUM(ts.IndDue), 0) AS AmountDue
    FROM TransactionSummary ts WITH (NOLOCK)
    WHERE ts.OrganizationId = {0} AND ts.PeopleId = {1}
    '''.format(org_id, user_id)

    amount_paid = 0
    amount_due = 0
    try:
        financial_result = list(q.QuerySql(financial_sql))
        if financial_result and financial_result[0]:
            amount_paid = financial_result[0].AmountPaid or 0
            amount_due = financial_result[0].AmountDue or 0
    except:
        pass

    # If we have a trip cost but no amount_due, use trip cost minus paid
    if trip_cost > 0 and amount_due == 0:
        amount_due = max(0, trip_cost - amount_paid)

    # Get upcoming meetings for this trip
    meetings_sql = '''
    SELECT TOP 3
        m.MeetingDate,
        m.Description,
        m.Location
    FROM Meetings m
    WHERE m.OrganizationId = {0}
      AND m.MeetingDate >= GETDATE()
    ORDER BY m.MeetingDate
    '''.format(org_id)

    upcoming_meetings = []
    try:
        upcoming_meetings = list(q.QuerySql(meetings_sql))
    except:
        pass

    # Check if user is a leader of this trip
    is_leader = any(lt.OrganizationId == org_id for lt in user_role.get('accessible_trips', []))

    # Start card
    print '''
    <div style="background: {0}; border: 2px solid {1}; border-radius: 12px; margin-bottom: 20px; overflow: hidden;">
        <!-- Trip Header -->
        <div style="background: {1}; color: white; padding: 16px 20px;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div>
                    <h2 style="margin: 0; font-size: 22px;">{2}</h2>
                    <p style="margin: 5px 0 0 0; opacity: 0.9;">{3}{4}</p>
                </div>
                <div style="display: flex; gap: 8px; align-items: center;">
                    {5}
                    <span style="background: rgba(255,255,255,0.2); padding: 6px 14px; border-radius: 20px; font-size: 13px;">
                        {6}
                    </span>
                </div>
            </div>
        </div>
    '''.format(
        colors['bg'],                                           # 0 - card bg
        colors['border'],                                       # 1 - border/header bg
        trip_name,                                              # 2 - trip name
        start_date_str + (' - ' + end_date_str if end_date_str else ''),  # 3 - dates
        ' ({0} days away)'.format(days_until) if days_until and days_until > 0 else '',  # 4 - countdown
        '<span style="background: #ffc107; color: #333; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600;">LEADER</span>' if is_leader else '',  # 5 - leader badge
        colors['text']                                          # 6 - status text
    )

    # Card body with sections
    print '<div style="padding: 20px;">'

    # === PAYMENT STATUS SECTION ===
    print '''
        <div style="margin-bottom: 20px;">
            <h4 style="margin: 0 0 12px 0; color: #333; font-size: 16px;">&#128176; Payment Status</h4>
    '''

    if trip_cost > 0:
        payment_percent = int((amount_paid / trip_cost) * 100) if trip_cost > 0 else 0
        progress_color = '#28a745' if payment_percent >= 100 else ('#ffc107' if payment_percent >= 50 else '#dc3545')

        print '''
            <div style="background: #f8f9fa; border-radius: 8px; padding: 16px;">
                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                    <span style="color: #666;">Trip Cost:</span>
                    <strong>${0:,.0f}</strong>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                    <span style="color: #28a745;">Amount Paid:</span>
                    <strong style="color: #28a745;">${1:,.0f}</strong>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
                    <span style="color: #dc3545;">Balance Due:</span>
                    <strong style="color: #dc3545;">${2:,.0f}</strong>
                </div>
                <!-- Progress bar -->
                <div style="background: #e9ecef; border-radius: 10px; height: 20px; overflow: hidden;">
                    <div style="background: {3}; height: 100%; width: {4}%; transition: width 0.3s;"></div>
                </div>
                <p style="text-align: center; margin: 8px 0 0 0; font-size: 13px; color: #666;">{4}% funded</p>
            </div>
        '''.format(float(trip_cost), float(amount_paid), float(amount_due), progress_color, payment_percent)
    else:
        print '''
            <div style="background: #f8f9fa; border-radius: 8px; padding: 16px; text-align: center; color: #666;">
                <p style="margin: 0;">Trip cost not yet set. Check back later for payment details.</p>
            </div>
        '''

    print '</div>'

    # === UPCOMING MEETINGS SECTION ===
    print '''
        <div style="margin-bottom: 20px;">
            <h4 style="margin: 0 0 12px 0; color: #333; font-size: 16px;">&#128197; Upcoming Meetings</h4>
    '''

    if upcoming_meetings:
        print '<div style="background: #f8f9fa; border-radius: 8px; padding: 12px;">'
        for meeting in upcoming_meetings:
            meeting_date_str = ''
            try:
                meeting_date_str = meeting.MeetingDate.strftime('%a, %b %d at %I:%M %p')
            except:
                meeting_date_str = str(meeting.MeetingDate)[:16]

            location = meeting.Location or ''
            description = meeting.Description or ''

            print '''
                <div style="padding: 10px 0; border-bottom: 1px solid #e9ecef;">
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <span style="font-size: 20px;">&#128197;</span>
                        <div>
                            <strong style="color: #333;">{0}</strong>
                            {1}
                            {2}
                        </div>
                    </div>
                </div>
            '''.format(
                meeting_date_str,
                '<br><span style="color: #666; font-size: 13px;">&#128205; {0}</span>'.format(location) if location else '',
                '<br><span style="color: #888; font-size: 12px;">{0}</span>'.format(description) if description else ''
            )
        print '</div>'
    else:
        print '''
            <div style="background: #f8f9fa; border-radius: 8px; padding: 16px; text-align: center; color: #666;">
                <p style="margin: 0;">No upcoming meetings scheduled.</p>
            </div>
        '''

    print '</div>'

    # === ACTION BUTTON ===
    # Single link to Person2 registrations tab which has funding page and email supporter links
    print '''
        <div style="display: flex; gap: 10px; flex-wrap: wrap;">
            <a href="{church_url}/Person2/{0}#tab-registrations" target="_blank"
               style="flex: 1; min-width: 150px; padding: 12px 16px; background: #007bff; color: white; text-decoration: none;
                      border-radius: 8px; text-align: center; font-weight: 600;">
                &#128179; View My Trip Details
            </a>
        </div>
    '''.format(user_id, church_url=CHURCH_URL)

    # Leader link if they're a leader
    if is_leader:
        print '''
        <div style="margin-top: 12px;">
            <a href="?trip={0}&section=overview"
               style="display: block; padding: 12px 16px; background: #6c757d; color: white; text-decoration: none;
                      border-radius: 8px; text-align: center; font-weight: 600;">
                &#128736; Manage Trip (Leader View)
            </a>
        </div>
        '''.format(org_id)

    print '</div></div>'  # Close card body and card


# ::END:: Member Dashboard View

# ::START:: Dashboard View
def render_dashboard_view():
    """Main dashboard view with mission overview (ADMIN ONLY)"""
    
    timer = start_timer()
    
    # Get total statistics with debug
    stats_query = get_total_stats_query()
    stats = execute_query_with_debug(stats_query, "Total Statistics Query", "top1")
    
    # Render KPI cards with all total due values
    print render_kpi_cards(stats.TotalMembers, stats.TotalApplications, stats.TotalOutstandingOpen, 
                          stats.TotalOutstandingClosed, stats.TotalOutstandingAll)
    
    # Add urgent deadlines section
    render_dashboard_deadlines()
    
    # Add queue statistics section
    render_queue_statistics()

    # Add pending approvals section
    render_pending_approvals_section()

    # Check for show closed parameter
    show_closed = model.Data.ShowClosed == '1' if hasattr(model.Data, 'ShowClosed') else config.SHOW_CLOSED_BY_DEFAULT
    
    # Toggle links
    toggle_text = "Hide Closed" if show_closed else "Show Closed"
    toggle_value = "0" if show_closed else "1"
    print '<p><a href="?ShowClosed={0}">{1}</a>'.format(toggle_value, toggle_text)
    
    # Add debug toggle for developers
    if model.UserIsInRole("SuperAdmin"):
        debug_text = "Disable SQL Debug" if SQL_DEBUG else "Enable SQL Debug"
        debug_value = "0" if SQL_DEBUG else "1"
        current_params = "ShowClosed=" + (model.Data.ShowClosed if hasattr(model.Data, 'ShowClosed') else '0')
        print ' | <a href="?{0}&debug={1}" style="color: #666;">{2}</a>'.format(current_params, debug_value, debug_text)
    
    print '</p>'
    
    # Add popup HTML
    print '''
    <div id="dataPopup" class="popup-overlay">
        <div class="popup-content">
            <span class="popup-close" onclick="closePopup()">&times;</span>
            <h3 id="popupTitle"></h3>
            <div id="popupBody"></div>
        </div>
    </div>
    
    <!-- Tooltip container for icon headers -->
    <div id="iconTooltip" class="icon-tooltip"></div>
    '''
    
    # Add popup script
    print get_popup_script()
    
    # Main content container
    print '<div class="dashboard-container">'
    print '<div class="main-content">'
    
    # Active missions table with icon headers
    print '<table class="mission-table">'
    print '''
    <thead>
        <tr>
            <th style="width: 40%;">Mission</th>
            <th class="text-center" style="width: 15%;">
                <div class="icon-header" onclick="showIconTooltip(event, 'Meetings: Upcoming-Completed')">
                    📅
                </div>
            </th>
            <th class="text-center" style="width: 15%;">
                <div class="icon-header" onclick="showIconTooltip(event, 'Passports: Active')">
                    🆔
                </div>
            </th>
            <th class="text-center" style="width: 15%;">
                <div class="icon-header" onclick="showIconTooltip(event, 'Background Checks: Approved-Failed-None')">
                    🛡️
                </div>
            </th>
            <th class="text-center" style="width: 15%;">
                <div class="icon-header" onclick="showIconTooltip(event, 'Total Members in Organization')">
                    👥
                </div>
            </th>
        </tr>
    </thead>
    <tbody>
    '''
    
    # Get active missions data
    missions_timer = start_timer()
    missions = execute_query_with_debug(get_active_missions_query(show_closed), "Active Missions Query", "sql")
    print end_timer(missions_timer, "Missions Query")
    
    last_status = None
    for mission in missions:
        # Status separator
        if mission.EventStatus != last_status:
            print '''
            <tr class="status-separator">
                <td colspan="5" style="text-align: center;">{0}</td>
            </tr>
            '''.format(mission.EventStatus)
            last_status = mission.EventStatus
        
        # Prepare popup data
        org_id = str(mission.OrganizationId)
        
        # Mission row
        print '''
        <tr class="mission-row">
            <td data-label="Mission">
                <div class="title">🏔️<a href="?OrgView={0}">{1}</a></div>
                <div class="dates">Trip: {2} - {3}</div>
            </td>
            <td data-label="Meetings" class="text-center">
                <a href="javascript:void(0)" onclick="loadAndShowPopup({0}, 'meetings', 'Meetings')" class="clickable-badge">{4}-{5}</a>
            </td>
            <td data-label="Passports" class="text-center">
                <a href="javascript:void(0)" onclick="loadAndShowPopup({0}, 'passports', 'Missing Passports')" class="clickable-badge">{6}</a>
            </td>
            <td data-label="Background" class="text-center">
                <a href="javascript:void(0)" onclick="loadAndShowPopup({0}, 'bgchecks', 'Background Check Issues')" class="clickable-badge">{7}-{8}-{9}</a>
            </td>
            <td data-label="Members" class="text-center">
                <a href="javascript:void(0)" onclick="loadAndShowPopup({0}, 'people', 'Team Members')" class="clickable-badge">{10}</a>
            </td>
        </tr>
        '''.format(
            mission.OrganizationId,
            mission.OrganizationName,
            format_date(str(mission.StartDate)),
            format_date(str(mission.EndDate)),
            mission.FutureMeetings,
            mission.PastMeetings,
            mission.PeopleWithPassports,
            mission.BackgroundCheckGood,
            mission.BackgroundCheckBad,
            mission.BackgroundCheckMissing,
            mission.MemberCount
        )
        
        # Check for leaders
        leaders = execute_query_with_debug(get_active_mission_leaders_query(org_id), "Leaders for Org " + org_id, "sql")
        if not leaders:
            print '''
            <tr>
                <td colspan="5" style="padding: 2px 8px;">
                    <span class="no-leader-warning">-- No leader(s) --</span>
                </td>
            </tr>
            '''
        else:
            for leader in leaders:
                print '''
                <tr>
                    <td colspan="5" style="padding: 2px 8px; font-size: 12px;">
                        <a href="/Person2/{0}" target="_blank">ℹ️</a> <i>{1}: {2} ({3})</i>
                    </td>
                </tr>
                '''.format(leader.PeopleId, leader.Leader, leader.Name, leader.Age)
        
        # Progress bar for payments (more compact)
        if mission.TotalFee and mission.TotalFee > 0:
            percentage = (float(mission.TotalDue) / float(mission.TotalFee)) * 100
            if mission.TotalDue >= mission.TotalFee:
                progress_text = "${:,.0f} of ${:,.0f} <strong>goal met!</strong>".format(mission.TotalDue, mission.TotalFee)
            else:
                progress_text = "<strong>${:,.0f} remaining</strong> of ${:,.0f} goal".format(mission.Outstanding, mission.TotalFee)
            
            print '''
            <tr>
                <td colspan="5" style="padding: 4px 8px;">
                    <div class="progress-container">
                        <div class="progress-bar" style="width: {0:.1f}%;"></div>
                    </div>
                    <div class="progress-text">{1}</div>
                </td>
            </tr>
            '''.format(percentage, progress_text)
    
    print '</tbody></table>'

    print '</div>'  # End main-content
    
    # Sidebar with Upcoming Meetings
    print '<div class="sidebar">'
    print '<div style="background: white; padding: 20px; border-radius: var(--radius); box-shadow: var(--shadow);">'
    print '<h3 style="margin-top: 0; color: var(--primary-color);">📅 Upcoming Meetings</h3>'
    
    upcoming_meetings_sql = '''
        SELECT TOP 15
            o.OrganizationName,
            COALESCE(me_desc.Data, m.Description) as Description,
            COALESCE(me_loc.Data, m.Location) as Location,
            m.MeetingDate,
            m.OrganizationId
        FROM Meetings m WITH (NOLOCK)
        INNER JOIN Organizations o WITH (NOLOCK) ON m.OrganizationId = o.OrganizationId
        LEFT JOIN MeetingExtra me_desc WITH (NOLOCK) ON m.MeetingId = me_desc.MeetingId AND me_desc.Field = 'Description'
        LEFT JOIN MeetingExtra me_loc WITH (NOLOCK) ON m.MeetingId = me_loc.MeetingId AND me_loc.Field = 'Location'
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND m.MeetingDate >= CAST(GETDATE() AS DATE)
        ORDER BY m.MeetingDate
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)
    
    upcoming_meetings = execute_query_with_debug(upcoming_meetings_sql, "Upcoming Meetings Query", "sql")
    
    if not upcoming_meetings:
        print '<p style="text-align: center; color: var(--text-muted);">Looks like the ministry calendars are as clear as a sunny day! 🌞</p>'
        print '<p style="text-align: center; color: var(--text-muted);">Why not schedule some meetings and make it a bit more exciting?</p>'
    else:
        print '<div class="timeline">'
        for meeting in upcoming_meetings:
            # Use TouchPoint's date formatting if available
            meeting_date = format_date(str(meeting.MeetingDate))
            if hasattr(model, 'FmtDate'):
                try:
                    meeting_date = model.FmtDate(meeting.MeetingDate)
                except:
                    pass
            
            print '''<div class="timeline-item">
                    <div class="timeline-date"><a href="/PyScript/Mission_Dashboard?OrgView={4}">{0}</a></div>
                    <div class="timeline-content">
                        <h4>{1}</h4>
                        <p>⏰ {2} 📍 {3}</p>
                    </div>
                </div>
            '''.format(
                meeting.OrganizationName,
                meeting.Description or "No description",
                meeting_date,
                meeting.Location or "TBD",
                meeting.OrganizationId
            )
        print '</div>'  # End timeline
    
    print '</div>'  # End sidebar content
    print '</div>'  # End sidebar
    print '</div>'  # End dashboard-container
    
    print end_timer(timer, "Dashboard Load")

# ::END:: Dashboard View

# ::START:: Review View
def render_pending_approvals_section():
    """Render pending approvals section for dashboard home."""
    # Get all pending approvals across all active trips
    pending_sql = '''
    WITH TripSubgroups AS (
        SELECT
            omt.OrgId,
            omt.PeopleId,
            MAX(CASE WHEN sg.Name = 'trip-approved' THEN 1 ELSE 0 END) as IsApproved,
            MAX(CASE WHEN sg.Name = 'trip-denied' THEN 1 ELSE 0 END) as IsDenied
        FROM OrgMemMemTags omt WITH (NOLOCK)
        JOIN MemberTags sg WITH (NOLOCK) ON omt.MemberTagId = sg.Id
        WHERE sg.Name IN ('trip-approved', 'trip-denied')
        GROUP BY omt.OrgId, omt.PeopleId
    )
    SELECT TOP 10
        om.OrganizationId,
        o.OrganizationName,
        om.PeopleId,
        p.Name,
        p.EmailAddress,
        om.EnrollmentDate
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    LEFT JOIN TripSubgroups ts ON om.OrganizationId = ts.OrgId AND om.PeopleId = ts.PeopleId
    WHERE o.IsMissionTrip = 1
      AND o.OrganizationStatusId = 30
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
      AND (ts.IsApproved IS NULL OR ts.IsApproved = 0)
      AND (ts.IsDenied IS NULL OR ts.IsDenied = 0)
    ORDER BY om.EnrollmentDate DESC
    '''

    try:
        pending = list(q.QuerySql(pending_sql))
        pending_count = len(pending)

        if pending_count == 0:
            return  # Don't show section if no pending approvals

        print '''
        <div class="kpi-section">
            <div class="overview-card" style="margin-top: 20px;">
                <div class="overview-card-header">
                    <h3>&#9998; Pending Trip Approvals ({0})</h3>
                    <a href="?view=review" class="btn btn-sm btn-primary">View All</a>
                </div>
                <div class="overview-card-body" style="padding: 0;">
                    <table class="signups-table">
                        <thead>
                            <tr>
                                <th>Name</th>
                                <th>Trip</th>
                                <th>Enrolled</th>
                                <th style="text-align: center;">Action</th>
                            </tr>
                        </thead>
                        <tbody>
        '''.format(pending_count)

        for member in pending:
            enrollment_date = ''
            if member.EnrollmentDate:
                try:
                    enrollment_date = member.EnrollmentDate.strftime('%b %d, %Y')
                except:
                    enrollment_date = str(member.EnrollmentDate)[:10]

            # Check if approvals are enabled for this specific trip
            show_approval_btns = are_approvals_enabled_for_trip(member.OrganizationId)

            print '''
                            <tr>
                                <td>
                                    <a href="/Person2/{0}" target="_blank" class="signup-name">{1}</a>
                                </td>
                                <td>
                                    <a href="?OrgView={2}" class="text-muted">{3}</a>
                                </td>
                                <td>{4}</td>
                                <td style="text-align: center;">
                                    <button class="btn btn-sm btn-outline-primary"
                                            onclick="ApprovalWorkflow.showModal({0}, '{1}', {2}, {5})"
                                            title="Review Application">
                                        Review
                                    </button>
                                </td>
                            </tr>
            '''.format(
                member.PeopleId,
                _escape_html(member.Name or ''),
                member.OrganizationId,
                _escape_html(member.OrganizationName or ''),
                enrollment_date,
                'true' if show_approval_btns else 'false'
            )

        print '''
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        '''
    except Exception as e:
        pass  # Silently fail if query errors


def render_review_view():
    """Full Review page showing all registrations across trips, with approval UI when enabled."""
    timer = start_timer()

    # Get global approval setting
    global_settings = get_global_settings()
    global_approvals_enabled = global_settings.get('approvals_enabled', True)

    # Cache for trip approval settings
    trip_approval_cache = {}

    print '''
    <style>
        .review-container {
            padding: 20px;
        }
        .review-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }
        .review-header h2 {
            margin: 0;
            color: var(--primary-color);
        }
        .review-filters {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
        }
        .review-filters .filter-btn {
            padding: 8px 16px;
            border: 1px solid #ddd;
            background: white;
            border-radius: 20px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .review-filters .filter-btn:hover {
            border-color: var(--primary-color);
        }
        .review-filters .filter-btn.active {
            background: var(--primary-color);
            color: white;
            border-color: var(--primary-color);
        }
        .review-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: var(--shadow);
        }
        .review-table th {
            text-align: left;
            padding: 14px 16px;
            background: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            font-weight: 600;
            font-size: 0.85rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .review-table td {
            padding: 14px 16px;
            border-bottom: 1px solid #f0f0f0;
            vertical-align: middle;
        }
        .review-table tr:last-child td { border-bottom: none; }
        .review-table tr:hover { background: #f8f9fa; }
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.8rem;
            font-weight: 500;
        }
        .status-pending {
            background: #fff3cd;
            color: #856404;
        }
        .status-approved {
            background: #d4edda;
            color: #155724;
        }
        .status-denied {
            background: #f8d7da;
            color: #721c24;
        }
        .review-actions {
            display: flex;
            gap: 6px;
        }
        .review-actions button {
            padding: 6px 12px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.85rem;
            transition: all 0.2s;
        }
        .btn-review {
            background: #667eea;
            color: white;
        }
        .btn-review:hover {
            background: #5a6fd6;
        }
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #6c757d;
        }
        .empty-state-icon {
            font-size: 3rem;
            margin-bottom: 10px;
        }
    </style>
    '''

    # Get filter from URL
    filter_type = str(model.Data.filter) if hasattr(model.Data, 'filter') else 'pending'

    # Build query based on filter
    if filter_type == 'approved':
        status_filter = '''
            AND ts.IsApproved = 1
            AND (ts.IsDenied IS NULL OR ts.IsDenied = 0)
        '''
        filter_title = 'Approved'
    elif filter_type == 'denied':
        status_filter = '''
            AND ts.IsDenied = 1
        '''
        filter_title = 'Denied'
    else:  # pending (default)
        status_filter = '''
            AND (ts.IsApproved IS NULL OR ts.IsApproved = 0)
            AND (ts.IsDenied IS NULL OR ts.IsDenied = 0)
        '''
        filter_type = 'pending'
        filter_title = 'Pending'

    review_sql = '''
    WITH TripSubgroups AS (
        SELECT
            omt.OrgId,
            omt.PeopleId,
            MAX(CASE WHEN sg.Name = 'trip-approved' THEN 1 ELSE 0 END) as IsApproved,
            MAX(CASE WHEN sg.Name = 'trip-denied' THEN 1 ELSE 0 END) as IsDenied
        FROM OrgMemMemTags omt WITH (NOLOCK)
        JOIN MemberTags sg WITH (NOLOCK) ON omt.MemberTagId = sg.Id
        WHERE sg.Name IN ('trip-approved', 'trip-denied')
        GROUP BY omt.OrgId, omt.PeopleId
    )
    SELECT
        om.OrganizationId,
        o.OrganizationName,
        om.PeopleId,
        p.Name,
        p.EmailAddress,
        p.CellPhone,
        om.EnrollmentDate,
        COALESCE(ts.IsApproved, 0) as IsApproved,
        COALESCE(ts.IsDenied, 0) as IsDenied
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
    LEFT JOIN TripSubgroups ts ON om.OrganizationId = ts.OrgId AND om.PeopleId = ts.PeopleId
    WHERE o.IsMissionTrip = 1
      AND o.OrganizationStatusId = 30
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
      {0}
    ORDER BY om.EnrollmentDate DESC
    '''.format(status_filter)

    # Get counts for filter tabs
    counts_sql = '''
    WITH TripSubgroups AS (
        SELECT
            omt.OrgId,
            omt.PeopleId,
            MAX(CASE WHEN sg.Name = 'trip-approved' THEN 1 ELSE 0 END) as IsApproved,
            MAX(CASE WHEN sg.Name = 'trip-denied' THEN 1 ELSE 0 END) as IsDenied
        FROM OrgMemMemTags omt WITH (NOLOCK)
        JOIN MemberTags sg WITH (NOLOCK) ON omt.MemberTagId = sg.Id
        WHERE sg.Name IN ('trip-approved', 'trip-denied')
        GROUP BY omt.OrgId, omt.PeopleId
    )
    SELECT
        SUM(CASE WHEN (COALESCE(ts.IsApproved, 0) = 0 AND COALESCE(ts.IsDenied, 0) = 0) THEN 1 ELSE 0 END) as PendingCount,
        SUM(CASE WHEN ts.IsApproved = 1 AND COALESCE(ts.IsDenied, 0) = 0 THEN 1 ELSE 0 END) as ApprovedCount,
        SUM(CASE WHEN ts.IsDenied = 1 THEN 1 ELSE 0 END) as DeniedCount
    FROM OrganizationMembers om WITH (NOLOCK)
    JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
    LEFT JOIN TripSubgroups ts ON om.OrganizationId = ts.OrgId AND om.PeopleId = ts.PeopleId
    WHERE o.IsMissionTrip = 1
      AND o.OrganizationStatusId = 30
      AND om.MemberTypeId NOT IN (230, 311)
      AND om.InactiveDate IS NULL
    '''

    try:
        counts = list(q.QuerySql(counts_sql))[0]
        pending_count = counts.PendingCount or 0
        approved_count = counts.ApprovedCount or 0
        denied_count = counts.DeniedCount or 0
    except:
        pending_count = 0
        approved_count = 0
        denied_count = 0

    print '''
    <div class="review-container">
        <div class="review-header">
            <h2>&#9998; Trip Registration Review</h2>
        </div>
    '''

    # Only show approval filter tabs if global approvals are enabled
    if global_approvals_enabled:
        print '''
        <div class="review-filters">
            <button class="filter-btn {0}" onclick="window.location.href='?view=review&filter=pending'">
                Pending ({1})
            </button>
            <button class="filter-btn {2}" onclick="window.location.href='?view=review&filter=approved'">
                Approved ({3})
            </button>
            <button class="filter-btn {4}" onclick="window.location.href='?view=review&filter=denied'">
                Denied ({5})
            </button>
        </div>
        '''.format(
            'active' if filter_type == 'pending' else '',
            pending_count,
            'active' if filter_type == 'approved' else '',
            approved_count,
            'active' if filter_type == 'denied' else '',
            denied_count
        )
    else:
        print '''
        <div style="padding: 10px 16px; background: #f0f0f0; border-radius: 8px; margin-bottom: 20px; color: #666;">
            &#128712; Approval workflow is currently disabled globally. Showing all registrations.
        </div>
        '''

    try:
        members = list(q.QuerySql(review_sql))

        if not members:
            empty_title = filter_title if global_approvals_enabled else 'Trip'
            print '''
            <div class="empty-state">
                <div class="empty-state-icon">&#128203;</div>
                <h3>No {0} Registrations</h3>
                <p>There are no {1} registrations to review.</p>
            </div>
            '''.format(empty_title, empty_title.lower())
        else:
            # Build table header - Status column only when approvals enabled, Action always shown
            if global_approvals_enabled:
                print '''
                <table class="review-table">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th>Trip</th>
                            <th>Contact</th>
                            <th>Enrolled</th>
                            <th>Status</th>
                            <th style="text-align: center;">Action</th>
                        </tr>
                    </thead>
                    <tbody>
                '''
            else:
                print '''
                <table class="review-table">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th>Trip</th>
                            <th>Contact</th>
                            <th>Enrolled</th>
                            <th style="text-align: center;">Action</th>
                        </tr>
                    </thead>
                    <tbody>
                '''

            for member in members:
                enrollment_date = ''
                if member.EnrollmentDate:
                    try:
                        enrollment_date = member.EnrollmentDate.strftime('%b %d, %Y')
                    except:
                        enrollment_date = str(member.EnrollmentDate)[:10]

                # Check if approvals are enabled for this specific trip (use cache)
                org_id = member.OrganizationId
                if org_id not in trip_approval_cache:
                    trip_approval_cache[org_id] = are_approvals_enabled_for_trip(org_id)
                trip_approvals_on = trip_approval_cache[org_id]

                # Build clickable contact info
                contact_parts = []
                if member.EmailAddress:
                    escaped_email = _escape_html(member.EmailAddress).replace("'", "\\'")
                    escaped_name = _escape_html(member.Name or '').replace("'", "\\'")
                    contact_parts.append('<a href="#" onclick="MissionsEmail.openIndividual({0}, \'{1}\', \'{2}\', {3}); return false;" style="color: #007bff; text-decoration: none;">{4}</a>'.format(
                        member.PeopleId, escaped_email, escaped_name, member.OrganizationId, _escape_html(member.EmailAddress)))
                if member.CellPhone:
                    phone_digits = ''.join(c for c in member.CellPhone if c.isdigit())
                    formatted_phone = model.FmtPhone(member.CellPhone, "") if member.CellPhone else member.CellPhone
                    contact_parts.append('<a href="tel:{0}" style="color: #007bff; text-decoration: none;">{1}</a>'.format(
                        phone_digits, _escape_html(formatted_phone or member.CellPhone)))
                contact_info = '<br>'.join(contact_parts) if contact_parts else ''

                # Check if approvals are enabled for this trip
                show_approval_status = global_approvals_enabled and trip_approvals_on

                # Review button is always shown (to view registration questions)
                review_button = '''
                    <button class="btn-review"
                            onclick="ApprovalWorkflow.showModal({0}, '{1}', {2}, {3})"
                            title="Review Registration">
                        Review
                    </button>
                '''.format(
                    member.PeopleId,
                    _escape_html(member.Name or '').replace("'", "\\'"),
                    member.OrganizationId,
                    'true' if show_approval_status else 'false'
                )

                # Render row based on whether approvals are enabled
                if show_approval_status:
                    # Determine status HTML for trips with approvals enabled
                    if member.IsDenied == 1:
                        status_html = '<span class="status-badge status-denied">&#10006; Denied</span>'
                        # Show both Review and Revoke for denied
                        action_html = review_button + '''
                            <button class="btn-review" style="margin-left: 5px;"
                                    onclick="ApprovalWorkflow.revokeStatus({0}, {1})"
                                    title="Move back to pending">
                                Revoke
                            </button>
                        '''.format(member.PeopleId, member.OrganizationId)
                    elif member.IsApproved == 1:
                        status_html = '<span class="status-badge status-approved">&#10004; Approved</span>'
                        # Show both Review and Revoke for approved
                        action_html = review_button + '''
                            <button class="btn-review" style="margin-left: 5px;"
                                    onclick="ApprovalWorkflow.revokeStatus({0}, {1})"
                                    title="Move back to pending">
                                Revoke
                            </button>
                        '''.format(member.PeopleId, member.OrganizationId)
                    else:
                        status_html = '<span class="status-badge status-pending">&#9203; Pending</span>'
                        action_html = review_button

                    print '''
                        <tr data-people-id="{0}" data-org-id="{1}">
                            <td>
                                <a href="/Person2/{0}" target="_blank" class="signup-name">{2}</a>
                            </td>
                            <td>
                                <a href="?trip={1}" class="text-muted">{3}</a>
                            </td>
                            <td style="font-size: 0.85rem;">{4}</td>
                            <td>{5}</td>
                            <td>{6}</td>
                            <td style="text-align: center;">
                                <div class="review-actions">
                                    {7}
                                </div>
                            </td>
                        </tr>
                    '''.format(
                        member.PeopleId,
                        member.OrganizationId,
                        _escape_html(member.Name or ''),
                        _escape_html(member.OrganizationName or ''),
                        contact_info,
                        enrollment_date,
                        status_html,
                        action_html
                    )
                else:
                    # Approvals disabled - show Review button but no Status column
                    print '''
                        <tr data-people-id="{0}" data-org-id="{1}">
                            <td>
                                <a href="/Person2/{0}" target="_blank" class="signup-name">{2}</a>
                            </td>
                            <td>
                                <a href="?trip={1}" class="text-muted">{3}</a>
                            </td>
                            <td style="font-size: 0.85rem;">{4}</td>
                            <td>{5}</td>
                            <td style="text-align: center;">
                                <div class="review-actions">
                                    {6}
                                </div>
                            </td>
                        </tr>
                    '''.format(
                        member.PeopleId,
                        member.OrganizationId,
                        _escape_html(member.Name or ''),
                        _escape_html(member.OrganizationName or ''),
                        contact_info,
                        enrollment_date,
                        review_button
                    )

            print '''
                </tbody>
            </table>
            '''
    except Exception as e:
        print '<div class="alert alert-danger">Error loading registrations: {0}</div>'.format(str(e))

    print '</div>'  # End review-container

    # Add Approval Modal CSS, HTML, and JavaScript
    print '''
    <style>
    /* Approval Modal Styles */
    .approval-modal-overlay {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.6);
        z-index: 10001;
        opacity: 0;
        transition: opacity 0.3s ease;
    }
    .approval-modal-overlay.active {
        display: flex;
        align-items: center;
        justify-content: center;
        opacity: 1;
    }
    .approval-modal {
        background: white;
        border-radius: 12px;
        width: 95%;
        max-width: 600px;
        max-height: 90vh;
        overflow-y: auto;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        transform: scale(0.9);
        transition: transform 0.3s ease;
    }
    .approval-modal-overlay.active .approval-modal {
        transform: scale(1);
    }
    .approval-modal-header {
        background: linear-gradient(135deg, #003366 0%, #4A90E2 100%);
        color: white;
        padding: 20px;
        border-radius: 12px 12px 0 0;
        position: relative;
    }
    .approval-modal-header h3 {
        margin: 0;
        font-size: 1.5rem;
    }
    .approval-modal-header .modal-subtitle {
        opacity: 0.9;
        font-size: 1rem;
        margin-top: 5px;
    }
    .approval-modal-close {
        position: absolute;
        top: 15px;
        right: 15px;
        font-size: 24px;
        cursor: pointer;
        opacity: 0.8;
        background: none;
        border: none;
        color: white;
    }
    .approval-modal-close:hover {
        opacity: 1;
    }
    .approval-modal-body {
        padding: 20px;
        font-size: 12px;
    }
    .approval-section {
        margin-bottom: 20px;
    }
    .approval-section h4 {
        font-size: 12px;
        color: #495057;
        margin-bottom: 10px;
        text-transform: uppercase;
        font-weight: 600;
        border-bottom: 1px solid #e9ecef;
        padding-bottom: 5px;
    }
    .person-info-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 12px;
    }
    .person-info-item {
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        text-align: left;
    }
    .person-info-item .label {
        font-size: 10px;
        color: #6c757d;
        text-transform: uppercase;
        text-align: left;
        margin-bottom: 2px;
    }
    .person-info-item .value {
        font-size: 12px;
        text-align: left;
        color: #212529;
    }
    .registration-qa-list {
        max-height: 350px;
        overflow-y: auto;
        background: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
    }
    .registration-qa-item {
        margin-bottom: 15px;
        padding-bottom: 15px;
        border-bottom: 1px solid #e9ecef;
    }
    .registration-qa-item:last-child {
        margin-bottom: 0;
        padding-bottom: 0;
        border-bottom: none;
    }
    .registration-qa-item .question {
        font-weight: 600;
        color: #495057;
        font-size: 12px;
        margin-bottom: 5px;
    }
    .registration-qa-item .answer {
        color: #212529;
        font-size: 12px;
        padding-left: 10px;
        border-left: 3px solid #667eea;
        line-height: 1.5;
    }
    .denial-reason-input {
        width: 100%;
        min-height: 80px;
        padding: 10px;
        border: 1px solid #ced4da;
        border-radius: 6px;
        font-size: 12px;
        resize: vertical;
    }
    .denial-reason-input:focus {
        outline: none;
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102,126,234,0.15);
    }
    .approval-modal-footer {
        display: flex;
        gap: 10px;
        padding: 20px;
        background: #f8f9fa;
        border-radius: 0 0 12px 12px;
        justify-content: flex-end;
    }
    .btn-approve, .btn-deny, .btn-cancel {
        padding: 10px 20px;
        border-radius: 6px;
        font-size: 0.95rem;
        font-weight: 500;
        cursor: pointer;
        border: none;
        transition: all 0.2s ease;
    }
    .btn-approve {
        background: #28a745;
        color: white;
    }
    .btn-approve:hover {
        background: #218838;
    }
    .btn-deny {
        background: #dc3545;
        color: white;
    }
    .btn-deny:hover {
        background: #c82333;
    }
    .btn-cancel {
        background: #6c757d;
        color: white;
    }
    .btn-cancel:hover {
        background: #5a6268;
    }
    .approval-loading {
        text-align: center;
        padding: 40px;
        color: #6c757d;
    }
    .current-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 5px 12px;
        border-radius: 15px;
        font-size: 0.85rem;
        font-weight: 500;
        margin-top: 5px;
    }
    .denial-history {
        background: #fff3cd;
        border: 1px solid #ffc107;
        border-radius: 6px;
        padding: 10px;
        margin-top: 10px;
    }
    .denial-history .denial-label {
        font-weight: 600;
        color: #856404;
        font-size: 0.85rem;
    }
    .denial-history .denial-reason-text {
        color: #856404;
        font-size: 0.9rem;
        margin-top: 5px;
    }
    .denial-history .denial-meta {
        color: #856404;
        font-size: 0.75rem;
        margin-top: 5px;
        opacity: 0.8;
    }
    </style>

    <!-- Approval Modal HTML -->
    <div id="approval-modal-overlay" class="approval-modal-overlay">
        <div class="approval-modal">
            <div class="approval-modal-header">
                <button class="approval-modal-close" onclick="ApprovalWorkflow.closeModal()">&times;</button>
                <h3 id="approval-modal-title">Review Registration</h3>
                <div class="modal-subtitle" id="approval-modal-subtitle"></div>
            </div>
            <div class="approval-modal-body">
                <div id="approval-modal-content">
                    <div class="approval-loading">
                        Loading registration details...
                    </div>
                </div>
            </div>
            <div class="approval-modal-footer" id="approval-modal-footer">
                <button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>
                <button class="btn-deny" id="btn-deny-registration" onclick="ApprovalWorkflow.showDenyForm()">Deny</button>
                <button class="btn-approve" id="btn-approve-registration" onclick="ApprovalWorkflow.approve()">Approve</button>
            </div>
        </div>
    </div>

    <script>
    // Approval Workflow Controller for Review Page
    var ApprovalWorkflow = {
        currentPeopleId: null,
        currentOrgId: null,
        currentPersonName: null,
        currentStatus: null,
        denialReason: '',
        showApprovalButtons: true,

        // Show the approval modal for a person
        // showApprovalBtns parameter controls whether Approve/Deny buttons are shown
        showModal: function(peopleId, personName, orgId, showApprovalBtns) {
            this.currentPeopleId = peopleId;
            this.currentOrgId = orgId;
            this.currentPersonName = personName;
            this.denialReason = '';
            this.showApprovalButtons = (showApprovalBtns !== false); // Default to true

            // Update modal title
            document.getElementById('approval-modal-title').textContent = 'Review Registration';
            document.getElementById('approval-modal-subtitle').textContent = personName;

            // Show loading state
            document.getElementById('approval-modal-content').innerHTML = '<div class="approval-loading">Loading registration details...</div>';

            // Reset footer buttons - conditionally show Approve/Deny based on showApprovalButtons
            if (this.showApprovalButtons) {
                document.getElementById('approval-modal-footer').innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>' +
                    '<button class="btn-deny" id="btn-deny-registration" onclick="ApprovalWorkflow.showDenyForm()">Deny</button>' +
                    '<button class="btn-approve" id="btn-approve-registration" onclick="ApprovalWorkflow.approve()">Approve</button>';
            } else {
                document.getElementById('approval-modal-footer').innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Close</button>';
            }

            // Show modal
            var overlay = document.getElementById('approval-modal-overlay');
            overlay.style.display = 'flex';
            setTimeout(function() {
                overlay.classList.add('active');
            }, 10);

            // Load registration data
            this.loadRegistrationData(peopleId, orgId);
        },

        // Load registration data via AJAX
        loadRegistrationData: function(peopleId, orgId) {
            var self = this;
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'get_registration_answers',
                    people_id: peopleId,
                    org_id: orgId
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.currentStatus = result.status;
                            self.renderRegistrationData(result);
                        } else {
                            document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error loading registration data: ' + (result.error || 'Unknown error') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error parsing response</div>';
                        console.error('Parse error:', e);
                    }
                },
                error: function(xhr, status, error) {
                    document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error loading registration data: ' + error + '</div>';
                }
            });
        },

        // Render the registration data in the modal
        renderRegistrationData: function(data) {
            var html = '';
            var person = data.person || {};

            // Person info section
            html += '<div class="approval-section">';
            html += '<h4>Person Information</h4>';
            html += '<div class="person-info-grid">';
            html += '<div class="person-info-item"><span class="label">Name</span><span class="value">' + this.escapeHtml(person.name || 'N/A') + '</span></div>';
            // Email - clickable to open email popup
            if (person.email) {
                var escapedEmail = this.escapeJsString(person.email);
                var escapedName = this.escapeJsString(person.name || '');
                html += '<div class="person-info-item"><span class="label">Email</span><span class="value"><a href="#" onclick="ApprovalWorkflow.closeModal(); MissionsEmail.openIndividual(' + person.people_id + ', \\'' + escapedEmail + '\\', \\'' + escapedName + '\\', ' + this.currentOrgId + '); return false;" style="color: #007bff; text-decoration: none;">' + this.escapeHtml(person.email) + '</a></span></div>';
            } else {
                html += '<div class="person-info-item"><span class="label">Email</span><span class="value">N/A</span></div>';
            }
            // Phone - clickable tel: link
            if (person.phone) {
                var phoneDigits = person.phone.replace(/[^0-9]/g, '');
                html += '<div class="person-info-item"><span class="label">Phone</span><span class="value"><a href="tel:' + phoneDigits + '" style="color: #007bff; text-decoration: none;">' + this.escapeHtml(person.phone) + '</a></span></div>';
            } else {
                html += '<div class="person-info-item"><span class="label">Phone</span><span class="value">N/A</span></div>';
            }
            html += '<div class="person-info-item"><span class="label">Age</span><span class="value">' + (person.age || 'N/A') + '</span></div>';
            html += '</div>';

            // Current status - only show if approval workflow is enabled
            if (this.showApprovalButtons) {
                html += '<div style="margin-top: 10px;">';
                html += '<span class="label">Current Status: </span>';
                if (data.status === 'approved') {
                    html += '<span class="current-status-badge approved" style="background:#d4edda;color:#155724;">&#10003; Approved</span>';
                } else if (data.status === 'denied') {
                    html += '<span class="current-status-badge denied" style="background:#f8d7da;color:#721c24;">&#10007; Denied</span>';
                } else {
                    html += '<span class="current-status-badge pending" style="background:#fff3cd;color:#856404;">&#8987; Pending</span>';
                }
                html += '</div>';

                // Show denial reason if exists
                if (data.status === 'denied' && data.denial_reason) {
                    html += '<div class="denial-history">';
                    html += '<div class="denial-label">Previous Denial Reason:</div>';
                    html += '<div class="denial-reason-text">' + this.escapeHtml(data.denial_reason) + '</div>';
                    if (data.denied_by_name || data.denied_date) {
                        html += '<div class="denial-meta">';
                        if (data.denied_by_name) html += 'By: ' + this.escapeHtml(data.denied_by_name);
                        if (data.denied_date) html += ' on ' + data.denied_date;
                        html += '</div>';
                    }
                    html += '</div>';
                }
            }

            html += '</div>';

            // Registration Q&A section
            html += '<div class="approval-section">';
            html += '<h4>Registration Answers</h4>';
            if (data.answers && data.answers.length > 0) {
                html += '<div class="registration-qa-list">';
                for (var i = 0; i < data.answers.length; i++) {
                    var qa = data.answers[i];
                    html += '<div class="registration-qa-item">';
                    html += '<div class="question">' + this.escapeHtml(qa.question || 'Question ' + (i + 1)) + '</div>';
                    html += '<div class="answer">' + this.escapeHtml(qa.answer || 'No answer provided') + '</div>';
                    html += '</div>';
                }
                html += '</div>';
            } else {
                html += '<div style="color: #6c757d; font-style: italic;">No registration questions found for this trip.</div>';
            }
            html += '</div>';

            document.getElementById('approval-modal-content').innerHTML = html;

            // Update buttons based on status
            this.updateFooterButtons(data.status);
        },

        // Update footer buttons based on current status
        updateFooterButtons: function(status) {
            var footer = document.getElementById('approval-modal-footer');

            // If approval buttons not enabled, just show Close button
            if (!this.showApprovalButtons) {
                footer.innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Close</button>';
                return;
            }

            var html = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>';

            if (status === 'approved') {
                html += '<button class="btn-deny" onclick="ApprovalWorkflow.revoke()">Revoke Approval</button>';
            } else if (status === 'denied') {
                html += '<button class="btn-approve" onclick="ApprovalWorkflow.approve()">Approve</button>';
                html += '<button class="btn-deny" onclick="ApprovalWorkflow.revoke()">Reset to Pending</button>';
            } else {
                html += '<button class="btn-deny" id="btn-deny-registration" onclick="ApprovalWorkflow.showDenyForm()">Deny</button>';
                html += '<button class="btn-approve" id="btn-approve-registration" onclick="ApprovalWorkflow.approve()">Approve</button>';
            }

            footer.innerHTML = html;
        },

        // Show the denial reason form
        showDenyForm: function() {
            var content = document.getElementById('approval-modal-content');
            var existingHtml = content.innerHTML;

            // Add denial reason textarea
            var denyFormHtml = '<div class="approval-section" id="denial-reason-section">' +
                '<h4>Denial Reason</h4>' +
                '<textarea id="denial-reason-textarea" class="denial-reason-input" placeholder="Please provide a reason for denying this registration...">' + this.escapeHtml(this.denialReason) + '</textarea>' +
                '</div>';

            content.innerHTML = existingHtml + denyFormHtml;

            // Update buttons
            var footer = document.getElementById('approval-modal-footer');
            footer.innerHTML = '<button class="btn-cancel" onclick="ApprovalWorkflow.closeModal()">Cancel</button>' +
                '<button class="btn-approve" onclick="ApprovalWorkflow.showModal(' + this.currentPeopleId + ', \\'' + this.escapeJsString(this.currentPersonName) + '\\', ' + this.currentOrgId + ', ' + this.showApprovalButtons + ')">Back</button>' +
                '<button class="btn-deny" onclick="ApprovalWorkflow.submitDenial()">Confirm Denial</button>';

            // Focus textarea
            document.getElementById('denial-reason-textarea').focus();
        },

        // Submit approval
        approve: function() {
            var self = this;
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            // Show loading
            document.getElementById('approval-modal-content').innerHTML = '<div class="approval-loading">Processing approval...</div>';

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'approve_member',
                    people_id: self.currentPeopleId,
                    org_id: self.currentOrgId
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            // Check if we should prompt to send approval notification email
                            if (result.prompt_email && result.person && result.person.email) {
                                ApprovalWorkflow.showApprovalEmailPrompt(result);
                            } else {
                                // Reload the page to show updated status
                                window.location.reload();
                            }
                        } else {
                            document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + (result.error || 'Unknown error') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error processing response</div>';
                    }
                },
                error: function(xhr, status, error) {
                    document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + error + '</div>';
                }
            });
        },

        // Show approval email prompt with editable email form
        showApprovalEmailPrompt: function(approvalData) {
            var self = this;
            var person = approvalData.person;
            var trip = approvalData.trip;
            var depositStr = trip.deposit > 0 ? '$' + trip.deposit.toLocaleString() : '[Deposit Amount]';
            var costStr = trip.cost > 0 ? '$' + trip.cost.toLocaleString() : '[Trip Cost]';

            // Store the approval data for the email
            ApprovalWorkflow.pendingApprovalEmail = {
                person: person,
                trip: trip
            };

            // Load the approval notification template from MissionsEmail templates
            var defaultSubject = trip.name + ' - Registration Approved!';
            var defaultBody = 'Dear ' + person.first_name + ',\\n\\n' +
                'Great news! Your registration for ' + trip.name + ' has been approved!\\n\\n' +
                'TRIP DETAILS:\\n' +
                '- Trip Cost: ' + costStr + '\\n' +
                '- Required Deposit: ' + depositStr + '\\n\\n' +
                'NEXT STEPS:\\n' +
                '1. Pay your deposit as soon as possible to secure your spot\\n' +
                '2. Start fundraising for the remaining balance\\n' +
                '3. Share your personal fundraising page with friends and family\\n\\n' +
                'PAYMENT OPTIONS:\\n' +
                'View your payment status and make payments here:\\n' +
                '{{MyGivingLink}}\\n\\n' +
                'FUNDRAISING:\\n' +
                'Share this link with supporters who want to help fund your trip:\\n' +
                '{{SupportLink}}\\n\\n' +
                'If you have any questions, please do not hesitate to reach out.\\n\\n' +
                'We are excited to have you on this mission trip!\\n\\n' +
                'Blessings';

            // Try to load the approval_notification template
            if (typeof MissionsEmail !== 'undefined' && MissionsEmail.templates && MissionsEmail.templates.length > 0) {
                for (var i = 0; i < MissionsEmail.templates.length; i++) {
                    if (MissionsEmail.templates[i].id === 'approval_notification') {
                        var tpl = MissionsEmail.templates[i];
                        // Replace placeholders
                        defaultSubject = (tpl.subject || defaultSubject)
                            .replace(/\\{\\{TripName\\}\\}/g, trip.name)
                            .replace(/\\{\\{PersonName\\}\\}/g, person.first_name);
                        defaultBody = (tpl.body || defaultBody)
                            .replace(/\\{\\{TripName\\}\\}/g, trip.name)
                            .replace(/\\{\\{PersonName\\}\\}/g, person.first_name)
                            .replace(/\\{\\{TripCost\\}\\}/g, costStr)
                            .replace(/\\{\\{DepositAmount\\}\\}/g, depositStr);
                        break;
                    }
                }
            }

            // Convert \\n to actual newlines for textarea display
            var bodyForTextarea = defaultBody.replace(/\\\\n/g, '\\n');

            // Create the approval email modal with editable form
            var modalHtml = '<div id="approval-email-prompt-modal" class="email-modal-overlay active">' +
                '<div class="email-modal" style="max-width: 650px;">' +
                    '<div class="email-modal-header" style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white;">' +
                        '<h3 style="margin: 0;">&#10003; Member Approved - Send Notification</h3>' +
                        '<button class="email-close-btn" onclick="ApprovalWorkflow.closeApprovalEmailPrompt()" style="color: white;">&times;</button>' +
                    '</div>' +
                    '<div class="email-modal-body" style="padding: 20px;">' +
                        '<div style="background: #d4edda; border: 1px solid #c3e6cb; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px;">' +
                            '<strong style="color: #155724;">' + person.name + '</strong> has been approved for <strong style="color: #155724;">' + trip.name + '</strong>' +
                        '</div>' +
                        '<div style="background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px;">' +
                            '<div style="display: flex; align-items: center; gap: 10px;">' +
                                '<span style="font-size: 20px;">&#128176;</span>' +
                                '<div>' +
                                    '<strong style="color: #856404;">Trip Fee to be Set:</strong> ' +
                                    '<span style="font-size: 18px; font-weight: bold; color: #856404;">' + costStr + '</span>' +
                                    (trip.deposit > 0 ? '<br><small style="color: #856404;">Deposit required: ' + depositStr + '</small>' : '') +
                                '</div>' +
                            '</div>' +
                        '</div>' +
                        '<div style="margin-bottom: 12px;">' +
                            '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">To:</label>' +
                            '<input type="text" id="approval-email-to" value="' + person.email + '" readonly ' +
                                'style="width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; background: #f8f9fa; box-sizing: border-box;">' +
                        '</div>' +
                        '<div style="margin-bottom: 12px;">' +
                            '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">Subject:</label>' +
                            '<input type="text" id="approval-email-subject" value="' + defaultSubject.replace(/"/g, '&quot;') + '" ' +
                                'style="width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; box-sizing: border-box;">' +
                        '</div>' +
                        '<div style="margin-bottom: 12px;">' +
                            '<label style="display: block; font-weight: 600; margin-bottom: 6px; color: #333;">Message:</label>' +
                            '<textarea id="approval-email-body" rows="12" ' +
                                'style="width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-family: inherit; font-size: 14px; resize: vertical; box-sizing: border-box;">' + bodyForTextarea + '</textarea>' +
                            '<p style="margin: 6px 0 0 0; font-size: 12px; color: #888;">Tip: Edit the default template in Settings > Quick Email Templates > "Registration Approved"</p>' +
                        '</div>' +
                    '</div>' +
                    '<div class="email-modal-footer" style="display: flex; justify-content: flex-end; gap: 10px; padding: 15px 20px; background: #f8f9fa;">' +
                        '<button onclick="ApprovalWorkflow.closeApprovalEmailPrompt()" style="padding: 10px 20px; border: 1px solid #ddd; border-radius: 6px; background: white; cursor: pointer;">Skip (No Fee Set)</button>' +
                        '<button id="approval-send-btn" onclick="ApprovalWorkflow.sendApprovalEmail()" style="padding: 10px 20px; border: none; border-radius: 6px; background: #28a745; color: white; cursor: pointer;">&#128176; Set Fee &amp; Send Email</button>' +
                    '</div>' +
                '</div>' +
            '</div>';

            document.body.insertAdjacentHTML('beforeend', modalHtml);
            document.getElementById('approval-email-prompt-modal').addEventListener('click', function(e) {
                if (e.target === this) ApprovalWorkflow.closeApprovalEmailPrompt();
            });
        },

        closeApprovalEmailPrompt: function() {
            var modal = document.getElementById('approval-email-prompt-modal');
            if (modal) modal.remove();
            ApprovalWorkflow.pendingApprovalEmail = null;
            // Reload to show updated approval status
            window.location.reload();
        },

        sendApprovalEmail: function() {
            var data = ApprovalWorkflow.pendingApprovalEmail;
            if (!data) {
                window.location.reload();
                return;
            }

            var toEmail = document.getElementById('approval-email-to').value;
            var subject = document.getElementById('approval-email-subject').value;
            var body = document.getElementById('approval-email-body').value;

            if (!subject.trim()) {
                alert('Please enter a subject.');
                document.getElementById('approval-email-subject').focus();
                return;
            }

            if (!body.trim()) {
                alert('Please enter a message.');
                document.getElementById('approval-email-body').focus();
                return;
            }

            // Disable the send button and show loading
            var sendBtn = document.getElementById('approval-send-btn');
            if (sendBtn) {
                sendBtn.disabled = true;
                sendBtn.innerHTML = 'Setting fee...';
            }

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            // Encode body to avoid ASP.NET request validation blocking HTML
            var encodedBody = body.replace(/</g, '&lt;').replace(/>/g, '&gt;')
                                  .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

            // Step 1: First adjust the fee (trip cost)
            var tripCost = data.trip.cost || 0;

            function sendEmailAfterFee() {
                // Step 2: Now send the email
                if (sendBtn) sendBtn.innerHTML = 'Sending email...';

                $.ajax({
                    url: ajaxUrl,
                    type: 'POST',
                    data: {
                        action: 'send_email',
                        to_emails: toEmail,
                        subject: subject,
                        body: encodedBody,
                        people_id: data.person.people_id,
                        org_id: data.trip.org_id
                    },
                    success: function(response) {
                        try {
                            var result = typeof response === 'string' ? JSON.parse(response) : response;
                            if (result.success) {
                                // Close modal and reload
                                var modal = document.getElementById('approval-email-prompt-modal');
                                if (modal) modal.remove();
                                ApprovalWorkflow.pendingApprovalEmail = null;
                                window.location.reload();
                            } else {
                                alert('Fee was set but email failed: ' + (result.message || 'Failed to send email.'));
                                // Still reload since fee was set
                                window.location.reload();
                            }
                        } catch (e) {
                            alert('Fee was set. Email status unclear - please check sent emails.');
                            window.location.reload();
                        }
                    },
                    error: function(xhr, status, error) {
                        alert('Fee was set but email failed: ' + error);
                        // Still reload since fee was set
                        window.location.reload();
                    }
                });
            }

            // Only adjust fee if trip cost is greater than 0
            if (tripCost > 0) {
                $.ajax({
                    url: ajaxUrl,
                    type: 'POST',
                    data: {
                        ajax: 'true',
                        action: 'adjust_fee',
                        people_id: data.person.people_id,
                        org_id: data.trip.org_id,
                        amount: -tripCost,  // Negative to increase what they owe
                        description: 'Trip fee set upon approval'
                    },
                    success: function(response) {
                        try {
                            var result = typeof response === 'string' ? JSON.parse(response) : response;
                            if (result.success) {
                                // Fee set successfully, now send email
                                sendEmailAfterFee();
                            } else {
                                alert('Error setting fee: ' + (result.message || 'Unknown error'));
                                if (sendBtn) {
                                    sendBtn.disabled = false;
                                    sendBtn.innerHTML = '&#128176; Set Fee &amp; Send Email';
                                }
                            }
                        } catch (e) {
                            alert('Error processing fee response');
                            if (sendBtn) {
                                sendBtn.disabled = false;
                                sendBtn.innerHTML = '&#128176; Set Fee &amp; Send Email';
                            }
                        }
                    },
                    error: function(xhr, status, error) {
                        alert('Error setting fee: ' + error);
                        if (sendBtn) {
                            sendBtn.disabled = false;
                            sendBtn.innerHTML = '&#128176; Set Fee &amp; Send Email';
                        }
                    }
                });
            } else {
                // No trip cost, just send the email
                sendEmailAfterFee();
            }
        },

        pendingApprovalEmail: null,

        // Submit denial
        submitDenial: function() {
            var self = this;
            var reason = document.getElementById('denial-reason-textarea').value.trim();

            if (!reason) {
                alert('Please provide a reason for denial.');
                document.getElementById('denial-reason-textarea').focus();
                return;
            }

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            // Show loading
            document.getElementById('approval-modal-content').innerHTML = '<div class="approval-loading">Processing denial...</div>';

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'deny_member',
                    people_id: self.currentPeopleId,
                    org_id: self.currentOrgId,
                    reason: reason
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            window.location.reload();
                        } else {
                            document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + (result.error || 'Unknown error') + '</div>';
                        }
                    } catch (e) {
                        document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error processing response</div>';
                    }
                },
                error: function(xhr, status, error) {
                    document.getElementById('approval-modal-content').innerHTML = '<div class="alert alert-danger">Error: ' + error + '</div>';
                }
            });
        },

        // Revoke approval (return to pending)
        revoke: function(peopleId, orgId) {
            var self = this;
            var targetPeopleId = peopleId || this.currentPeopleId;
            var targetOrgId = orgId || this.currentOrgId;

            if (!confirm('Are you sure you want to revoke this status and return to pending?')) {
                return;
            }

            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

            $.ajax({
                url: ajaxUrl,
                type: 'POST',
                data: {
                    ajax: 'true',
                    action: 'revoke_approval',
                    people_id: targetPeopleId,
                    org_id: targetOrgId
                },
                success: function(response) {
                    try {
                        var result = typeof response === 'string' ? JSON.parse(response) : response;
                        if (result.success) {
                            self.closeModal();
                            window.location.reload();
                        } else {
                            alert('Error: ' + (result.error || 'Unknown error'));
                        }
                    } catch (e) {
                        alert('Error processing response');
                    }
                },
                error: function(xhr, status, error) {
                    alert('Error: ' + error);
                }
            });
        },

        // Close the modal
        closeModal: function() {
            var overlay = document.getElementById('approval-modal-overlay');
            overlay.classList.remove('active');
            setTimeout(function() {
                overlay.style.display = 'none';
            }, 300);
        },

        // Helper: Escape HTML
        escapeHtml: function(text) {
            if (!text) return '';
            var div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        },

        // Helper: Escape JS string (uses fromCharCode to avoid Python/JS escaping issues)
        escapeJsString: function(str) {
            if (!str) return '';
            var bs = String.fromCharCode(92);  // backslash
            var sq = String.fromCharCode(39);  // single quote
            var dq = String.fromCharCode(34);  // double quote
            var result = str.split(bs).join(bs + bs);
            result = result.split(sq).join(bs + sq);
            result = result.split(dq).join(bs + dq);
            return result;
        }
    };

    // Initialize event listeners when DOM is ready
    (function() {
        function initApprovalWorkflowEvents() {
            // Close modal when clicking outside
            var overlay = document.getElementById('approval-modal-overlay');
            if (overlay) {
                overlay.addEventListener('click', function(e) {
                    if (e.target === this) {
                        ApprovalWorkflow.closeModal();
                    }
                });
            }

            // Close modal with Escape key
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') {
                    var modalOverlay = document.getElementById('approval-modal-overlay');
                    if (modalOverlay && modalOverlay.classList.contains('active')) {
                        ApprovalWorkflow.closeModal();
                    }
                }
            });
        }

        // Run immediately if DOM is ready, otherwise wait
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initApprovalWorkflowEvents);
        } else {
            initApprovalWorkflowEvents();
        }
    })();

    // Make ApprovalWorkflow globally available
    window.ApprovalWorkflow = ApprovalWorkflow;
    </script>
    '''

    print end_timer(timer, "Review View Load")

# ::END:: Review View

# ::START:: Finance View
def render_finance_view():
    """Finance view with payment tracking"""
    
    timer = start_timer()
    
    # Get totals for both open and all trips, including fully paid counts
    totals_open_sql = '''
    {0}
    SELECT 
        SUM(Due) AS TotalDue,
        COUNT(CASE WHEN Due = 0 THEN 1 END) AS FullyPaidCount,
        COUNT(CASE WHEN Due > 0 THEN 1 END) AS OutstandingCount,
        COUNT(*) AS TotalPeople
    FROM MissionTripTotals
    WHERE PeopleId IS NOT NULL
    '''.format(get_mission_trip_totals_cte(include_closed=False))
    
    totals_all_sql = '''
    {0}
    SELECT 
        SUM(Due) AS TotalDue,
        COUNT(CASE WHEN Due = 0 THEN 1 END) AS FullyPaidCount,
        COUNT(CASE WHEN Due > 0 THEN 1 END) AS OutstandingCount,
        COUNT(*) AS TotalPeople
    FROM MissionTripTotals
    WHERE PeopleId IS NOT NULL
    '''.format(get_mission_trip_totals_cte(include_closed=True))
    
    totals_open = execute_query_with_debug(totals_open_sql, "Finance Totals Open Query", "top1")
    totals_all = execute_query_with_debug(totals_all_sql, "Finance Totals All Query", "top1")
    
    # Display totals at top
    print '''
    <div class="kpi-container" style="margin-bottom: 20px;">
        <div class="kpi-card">
            <div class="value">{0}</div>
            <div class="label">Total Due (Open Trips)</div>
            <div style="font-size: 0.75em; color: var(--text-muted); margin-top: 5px;">
                {1} outstanding / {2} paid in full
            </div>
        </div>
        <div class="kpi-card">
            <div class="value">{3}</div>
            <div class="label">Total Due (All Trips)</div>
            <div style="font-size: 0.75em; color: var(--text-muted); margin-top: 5px;">
                {4} outstanding / {5} paid in full
            </div>
        </div>
        <div class="kpi-card">
            <div class="value">{6}%</div>
            <div class="label">Completion Rate (All)</div>
            <div style="font-size: 0.75em; color: var(--text-muted); margin-top: 5px;">
                {7} of {8} fully paid
            </div>
        </div>
    </div>
    '''.format(
        format_currency(totals_open.TotalDue or 0), 
        totals_open.OutstandingCount or 0,
        totals_open.FullyPaidCount or 0,
        format_currency(totals_all.TotalDue or 0),
        totals_all.OutstandingCount or 0,
        totals_all.FullyPaidCount or 0,
        int((totals_all.FullyPaidCount or 0) * 100.0 / (totals_all.TotalPeople or 1)) if totals_all.TotalPeople > 0 else 0,
        totals_all.FullyPaidCount or 0,
        totals_all.TotalPeople or 0
    )
    
    print '<h3>All Participants - Payment Status (All Trips)</h3>'
    
    # Get outstanding payments - for ALL trips (including closed)
    outstanding_sql = '''
        {0}
        SELECT
            mtt.PeopleId,
            mtt.InvolvementId as OrganizationId,
            mtt.Trip as OrganizationName,
            mtt.Name as Name2,
            p.EmailAddress,
            mtt.TripCost,
            mtt.Raised AS Paid,
            mtt.Due AS Outstanding,
            CASE
                WHEN EXISTS (
                    SELECT 1 FROM OrganizationExtra oe
                    WHERE oe.OrganizationId = mtt.InvolvementId
                    AND oe.Field = 'Close'
                    AND oe.DateValue <= GETDATE()
                ) THEN 'Closed'
                ELSE 'Open'
            END AS TripStatus
        FROM MissionTripTotals mtt
        LEFT JOIN People p WITH (NOLOCK) ON p.PeopleId = mtt.PeopleId
        WHERE mtt.PeopleId IS NOT NULL
        ORDER BY
            CASE
                WHEN EXISTS (
                    SELECT 1 FROM OrganizationExtra oe
                    WHERE oe.OrganizationId = mtt.InvolvementId
                    AND oe.Field = 'Close'
                    AND oe.DateValue <= GETDATE()
                ) THEN 1
                ELSE 0
            END,
            mtt.Trip, mtt.Name
    '''.format(get_mission_trip_totals_cte(include_closed=True))
    
    outstanding = execute_query_with_debug(outstanding_sql, "Outstanding Payments Query", "sql")
    
    print '<table class="mission-table">'
    print '''
    <thead>
        <tr>
            <th>Name</th>
            <th>Mission</th>
            <th>Trip Status</th>
            <th>Payment Status</th>
            <th>Paid</th>
            <th>Outstanding</th>
        </tr>
    </thead>
    <tbody>
    '''
    
    last_org = None
    last_status = None
    for payment in outstanding:
        # Status separator
        if payment.TripStatus != last_status:
            print '''
            <tr class="status-separator">
                <td colspan="6" style="text-align: center; background: {0}; color: {1}; font-weight: bold;">
                    {2} Trips
                </td>
            </tr>
            '''.format(
                '#ffebee' if payment.TripStatus == 'Closed' else '#e8f5e9',
                '#c62828' if payment.TripStatus == 'Closed' else '#2e7d32',
                payment.TripStatus
            )
            last_status = payment.TripStatus
            last_org = None  # Reset org separator when status changes

        # Organization separator - link to Budget & Fundraising tab
        if payment.OrganizationName != last_org:
            print '''
            <tr class="org-separator">
                <td colspan="6" style="background: var(--light-bg); font-weight: bold;">
                    <a href="?trip={0}&section=budget" style="color: inherit; text-decoration: none;">
                        {1} &rarr;
                    </a>
                </td>
            </tr>
            '''.format(payment.OrganizationId, _escape_html(payment.OrganizationName))
            last_org = payment.OrganizationName

        # Determine payment status
        if payment.Outstanding == 0:
            payment_status = "Paid in Full"
            payment_status_class = "status-normal"
            outstanding_style = "color: var(--success-color); font-weight: bold;"
        else:
            payment_status = "Outstanding"
            payment_status_class = "status-urgent"
            outstanding_style = "color: var(--danger-color); font-weight: bold;"

        # Build name cell - clickable for payment reminder if email available
        member_email = getattr(payment, 'EmailAddress', '') or ''
        member_name_js = _escape_js_string(payment.Name2 or '')
        email_js = _escape_js_string(member_email)
        trip_name_js = _escape_js_string(payment.OrganizationName or '')
        trip_cost = float(payment.TripCost) if payment.TripCost else 0.0
        outstanding_amount = float(payment.Outstanding) if payment.Outstanding else 0.0

        if member_email and payment.Outstanding and payment.Outstanding > 0:
            # Clickable name with payment reminder popup
            name_cell = '''<a href="#" onclick="MissionsEmail.openIndividualGoalReminder({0}, '{1}', '{2}', '{3}', {4}, {5}, {6}); return false;" style="color: #007bff; cursor: pointer;" title="Send payment reminder">{7}</a>'''.format(
                payment.PeopleId,
                email_js,
                member_name_js,
                trip_name_js,
                trip_cost,
                outstanding_amount,
                payment.OrganizationId,
                _escape_html(payment.Name2)
            )
        else:
            # Plain text name (no email or fully paid)
            name_cell = _escape_html(payment.Name2)

        print '''
        <tr>
            <td data-label="Name">
                {0}
            </td>
            <td data-label="Mission">
                <a href="?trip={1}&section=budget" style="color: #667eea;">{2}</a>
            </td>
            <td data-label="Trip Status">
                <span class="status-badge status-{3}">{4}</span>
            </td>
            <td data-label="Payment Status">
                <span class="status-badge {5}">{6}</span>
            </td>
            <td data-label="Paid">{7}</td>
            <td data-label="Outstanding" style="{8}">
                {9}
            </td>
        </tr>
        '''.format(
            name_cell,
            payment.OrganizationId,
            _escape_html(payment.OrganizationName),
            'closed' if payment.TripStatus == 'Closed' else 'active',
            payment.TripStatus,
            payment_status_class,
            payment_status,
            format_currency(payment.Paid),
            outstanding_style,
            format_currency(payment.Outstanding)
        )
    
    print '</tbody></table>'
    
    # Recent transactions
    print '<h3 style="margin-top: 30px;">Recent Transactions</h3>'
    
    transactions_sql = '''
        SELECT TOP 50
            tl.TransactionDate,
            tl.People as Name,
            o.OrganizationName,
            tl.BegBal,
            tl.TotalPayment,
            tl.TotDue,
            tl.Message
        FROM TransactionList tl WITH (NOLOCK)
        INNER JOIN Organizations o WITH (NOLOCK) ON o.OrganizationId = tl.OrgId
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND tl.BegBal IS NOT NULL
        ORDER BY tl.Id DESC
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)
    
    transactions = execute_query_with_debug(transactions_sql, "Recent Transactions Query", "sql")
    
    print '<table class="mission-table">'
    print '''
    <thead>
        <tr>
            <th>Date</th>
            <th>Name</th>
            <th>Beginning</th>
            <th>Payment</th>
            <th>Due</th>
        </tr>
    </thead>
    <tbody>
    '''
    
    for trans in transactions:
        print '''
        <tr>
            <td data-label="Date">{0}</td>
            <td data-label="Name">{1}</td>
            <td data-label="Beginning">{2}</td>
            <td data-label="Payment" style="color: var(--success-color);">{3}</td>
            <td data-label="Due">{4}</td>
        </tr>
        <tr>
            <td colspan="5" style="padding-left: 20px; font-style: italic; color: var(--text-muted);">
                {5} - {6}
            </td>
        </tr>
        '''.format(
            format_date(str(trans.TransactionDate)),
            trans.Name,
            format_currency(trans.BegBal),
            format_currency(trans.TotalPayment),
            format_currency(trans.TotDue),
            trans.OrganizationName,
            trans.Message or "No message"
        )
    
    print '</tbody></table>'
    
    print end_timer(timer, "Finance View Load")

# ::END:: Finance View

# ::START:: Stats View
def render_stats_view():
    """Enhanced statistics view for missions pastors"""
    
    timer = start_timer()
    
    print '<h3>Mission Statistics & Analytics</h3>'
    
    # Get all enhanced stat queries
    stat_queries = get_enhanced_stats_queries()
    
    # Financial Summary - now returns 3 rows (Open, Closed, All)
    financial_breakdown = execute_query_with_debug(stat_queries['financial_summary'], "Financial Summary Query", "sql")
    
    # Process the breakdown into a dictionary for easy access
    financial_data = {}
    for row in financial_breakdown:
        financial_data[row.TripStatus] = row
    
    # Display the All trips summary first
    all_data = financial_data.get('All', None)
    if all_data:
        print '''
        <div class="stat-card" style="margin-bottom: 20px;">
            <h4>Financial Overview - All Trips</h4>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px;">
                <div>
                    <div style="font-size: 24px; font-weight: bold; color: var(--primary-color);">{0}</div>
                    <div style="color: var(--text-muted);">Total Trips</div>
                </div>
                <div>
                    <div style="font-size: 24px; font-weight: bold; color: var(--success-color);">{1}</div>
                    <div style="color: var(--text-muted);">Total Goal</div>
                </div>
                <div>
                    <div style="font-size: 24px; font-weight: bold; color: var(--secondary-color);">{2}</div>
                    <div style="color: var(--text-muted);">Total Raised</div>
                </div>
                <div>
                    <div style="font-size: 24px; font-weight: bold; color: var(--danger-color);">{3}</div>
                    <div style="color: var(--text-muted);">Outstanding</div>
                </div>
                <div>
                    <div style="font-size: 24px; font-weight: bold;">{4:.0f}%</div>
                    <div style="color: var(--text-muted);">Funded</div>
                </div>
            </div>
            <div style="margin-top: 15px; padding: 10px; background: #e3f2fd; border-radius: 5px; font-size: 0.85em; color: #1976d2;">
                <strong>📊 Comprehensive View:</strong> This overview includes ALL participants in mission trips, 
                including those who have already paid in full. The outstanding amount represents the total still owed 
                across all participants.
            </div>
        </div>
        '''.format(
            all_data.TotalTrips,
            format_currency(all_data.TotalGoal),
            format_currency(all_data.TotalRaised),
            format_currency(all_data.TotalOutstanding),
            all_data.PercentRaised
        )
    
    # Display Open vs Closed breakdown
    open_data = financial_data.get('Open', None)
    closed_data = financial_data.get('Closed', None)
    
    if open_data or closed_data:
        print '''
        <div class="stat-card" style="margin-bottom: 20px;">
            <h4>Outstanding Balance Breakdown</h4>
            <table class="stat-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>Trip Status</th>
                        <th style="text-align: center;">Trips</th>
                        <th style="text-align: right;">Goal</th>
                        <th style="text-align: right;">Raised</th>
                        <th style="text-align: right;">Outstanding</th>
                        <th style="text-align: center;">Funded</th>
                    </tr>
                </thead>
                <tbody>
        '''
        
        if open_data:
            print '''
            <tr>
                <td><strong style="color: #2e7d32;">Open Trips</strong></td>
                <td style="text-align: center;">{0}</td>
                <td style="text-align: right;">{1}</td>
                <td style="text-align: right; color: var(--success-color);">{2}</td>
                <td style="text-align: right; color: var(--danger-color);"><strong>{3}</strong></td>
                <td style="text-align: center;"><strong>{4:.0f}%</strong></td>
            </tr>
            '''.format(
                open_data.TotalTrips,
                format_currency(open_data.TotalGoal),
                format_currency(open_data.TotalRaised),
                format_currency(open_data.TotalOutstanding),
                open_data.PercentRaised
            )
        
        if closed_data:
            print '''
            <tr>
                <td><strong style="color: #c62828;">Closed Trips</strong></td>
                <td style="text-align: center;">{0}</td>
                <td style="text-align: right;">{1}</td>
                <td style="text-align: right; color: var(--success-color);">{2}</td>
                <td style="text-align: right; color: var(--danger-color);"><strong>{3}</strong></td>
                <td style="text-align: center;"><strong>{4:.0f}%</strong></td>
            </tr>
            '''.format(
                closed_data.TotalTrips,
                format_currency(closed_data.TotalGoal),
                format_currency(closed_data.TotalRaised),
                format_currency(closed_data.TotalOutstanding),
                closed_data.PercentRaised
            )
        
        print '''
                </tbody>
            </table>
        </div>
        '''
    
    # Calendar Year Financial Breakdown
    financial_by_year = execute_query_with_debug(stat_queries['financial_by_year'], "Financial by Year Query", "sql")
    
    print '''
    <div class="stat-card" style="margin-bottom: 20px;">
        <h4>Financial Overview by Calendar Year</h4>
        <table class="stat-table" style="width: 100%;">
            <thead>
                <tr>
                    <th>Year</th>
                    <th style="text-align: center;">Trips</th>
                    <th style="text-align: right;">Goal</th>
                    <th style="text-align: right;">Raised</th>
                    <th style="text-align: right;">Outstanding</th>
                    <th style="text-align: center;">Funded</th>
                </tr>
            </thead>
            <tbody>
    '''
    
    for year_data in financial_by_year:
        print '''
        <tr>
            <td><strong>{0}</strong></td>
            <td style="text-align: center;">{1}</td>
            <td style="text-align: right;">{2}</td>
            <td style="text-align: right; color: var(--success-color);">{3}</td>
            <td style="text-align: right; color: var(--danger-color);">{4}</td>
            <td style="text-align: center;"><strong>{5:.0f}%</strong></td>
        </tr>
        '''.format(
            year_data.Year,
            year_data.TripCount,
            format_currency(year_data.TotalGoal),
            format_currency(year_data.TotalRaised),
            format_currency(year_data.TotalOutstanding),
            year_data.PercentRaised
        )
    
    print '''
            </tbody>
        </table>
    </div>
    '''
    
    # Grid layout for other stats
    print '<div class="stats-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px;">'
    
    # Trip Trends
    trip_trends = execute_query_with_debug(stat_queries['trip_trends'], "Trip Trends Query", "sql")
    
    print '<div class="stat-card">'
    print '<h4>Mission Trip Trends (5 Years)</h4>'
    print '<table class="stat-table">'
    print '<tr><th>Year</th><th>Trips</th><th>Participants</th></tr>'
    for trend in trip_trends:
        print '<tr><td>{0}</td><td style="text-align: center;"><strong>{1}</strong></td><td style="text-align: center;">{2}</td></tr>'.format(
            trend.TripYear, trend.TripCount, trend.TotalParticipants
        )
    print '</table>'
    print '</div>'
    
    # Participation by Age
    age_participation = execute_query_with_debug(stat_queries['participation_by_age'], "Age Participation Query", "sql")
    
    print '<div class="stat-card">'
    print '<h4>Participation by Age Group</h4>'
    print '<table class="stat-table">'
    print '<tr><th>Age Group</th><th>Count</th><th>Avg Age</th></tr>'
    
    # Sort by the SortOrder field that was calculated in the query
    age_participation_sorted = sorted(age_participation, key=lambda x: x.SortOrder if hasattr(x, 'SortOrder') else 99)
    
    for age_group in age_participation_sorted:
        print '<tr><td>{0}</td><td style="text-align: center;"><strong>{1}</strong></td><td style="text-align: center;">{2}</td></tr>'.format(
            age_group.AgeGroup, age_group.ParticipantCount, age_group.AvgAge
        )
    print '</table>'
    print '</div>'
    
    # Repeat Participants
    repeat_participants = execute_query_with_debug(stat_queries['repeat_participants'], "Repeat Participants Query", "sql")
    
    print '<div class="stat-card">'
    print '<h4>Repeat Participation</h4>'
    print '<table class="stat-table">'
    print '<tr><th>Number of Trips</th><th>People</th></tr>'
    for repeat in repeat_participants:
        trip_label = "{0} trip{1}".format(repeat.TripCount, "s" if repeat.TripCount > 1 else "")
        print '<tr><td>{0}</td><td style="text-align: center;"><strong>{1}</strong></td></tr>'.format(
            trip_label, repeat.PeopleCount
        )
    print '</table>'
    print '</div>'
    
    # Original basic stats
    membership_stats_sql = '''
        SELECT ms.Description AS MemberStatus, COUNT(DISTINCT p.PeopleId) AS Count
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN OrganizationMembers om WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
        LEFT JOIN lookup.MemberStatus ms WITH (NOLOCK) ON ms.Id = p.MemberStatusId
        WHERE o.IsMissionTrip = {0}
          AND om.InactiveDate IS NULL
          AND o.OrganizationStatusId = {1}
        GROUP BY ms.Description
        HAVING COUNT(DISTINCT p.PeopleId) > 0  
        ORDER BY COUNT(DISTINCT p.PeopleId) DESC
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)
    
    membership_stats = execute_query_with_debug(membership_stats_sql, "Membership Stats Query", "sql")
    
    print '<div class="stat-card">'
    print '<h4>Church Membership Status</h4>'
    print '<table class="stat-table">'
    for stat in membership_stats:
        print '<tr><td>{0}</td><td style="text-align: right;"><strong>{1}</strong></td></tr>'.format(
            stat.MemberStatus or "Unknown", stat.Count
        )
    print '</table>'
    print '</div>'
    
    # Gender stats
    gender_stats_sql = '''
        SELECT g.Description as Gender, COUNT(DISTINCT p.PeopleId) AS Count 
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN OrganizationMembers om WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
        LEFT JOIN lookup.Gender g WITH (NOLOCK) ON g.Id = p.GenderId
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND g.Description IS NOT NULL
        GROUP BY g.Description
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)
    
    gender_stats = execute_query_with_debug(gender_stats_sql, "Gender Stats Query", "sql")
    
    print '<div class="stat-card">'
    print '<h4>Gender Distribution</h4>'
    print '<table class="stat-table">'
    for stat in gender_stats:
        print '<tr><td>{0}</td><td style="text-align: right;"><strong>{1}</strong></td></tr>'.format(
            stat.Gender, stat.Count
        )
    print '</table>'
    print '</div>'
    
    # Baptism Status
    baptism_stats_sql = '''
        SELECT bs.Description AS BaptismStatus, COUNT(DISTINCT p.PeopleId) AS Count
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN OrganizationMembers om WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
        LEFT JOIN lookup.BaptismStatus bs WITH (NOLOCK) ON bs.Id = p.BaptismStatusId
        WHERE o.IsMissionTrip = {0}
          AND om.InactiveDate IS NULL
          AND o.OrganizationStatusId = {1}
        GROUP BY bs.Description
        HAVING COUNT(DISTINCT p.PeopleId) > 0  
        ORDER BY COUNT(DISTINCT p.PeopleId) DESC
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID)
    
    baptism_stats = execute_query_with_debug(baptism_stats_sql, "Baptism Stats Query", "sql")
    
    print '<div class="stat-card">'
    print '<h4>Baptism Status</h4>'
    print '<table class="stat-table">'
    for stat in baptism_stats:
        print '<tr><td>{0}</td><td style="text-align: right;"><strong>{1}</strong></td></tr>'.format(
            stat.BaptismStatus or "Unknown", stat.Count
        )
    print '</table>'
    print '</div>'
    
    print '</div>'  # End stats-grid
    
    # Upcoming Deadlines
    deadlines = execute_query_with_debug(stat_queries['upcoming_deadlines'], "Stats Deadlines Query", "sql")
    
    print '<div class="stat-card" style="margin-top: 20px;">'
    print '<h4>Upcoming Deadlines</h4>'
    print '<table class="mission-table">'
    print '<thead><tr><th>Mission</th><th>Deadline</th><th>Type</th><th>Days Until</th><th>Amount Remaining</th></tr></thead>'
    print '<tbody>'
    for deadline in deadlines:
        urgency_class = 'status-urgent' if deadline.DaysUntil < 7 else 'status-medium' if deadline.DaysUntil < 30 else 'status-normal'
        print '''
        <tr>
            <td><a href="?OrgView={0}">{1}</a></td>
            <td>{2}</td>
            <td>{3}</td>
            <td><span class="status-badge {4}">{5} days</span></td>
            <td style="color: {6}; font-weight: bold;">{7}</td>
        </tr>
        '''.format(
            deadline.OrganizationId,
            deadline.OrganizationName,
            format_date(str(deadline.DeadlineDate)),
            deadline.DeadlineType,
            urgency_class,
            deadline.DaysUntil,
            '#d32f2f' if deadline.AmountRemaining > 0 else '#2e7d32',
            format_currency(deadline.AmountRemaining) if deadline.AmountRemaining > 0 else "Fully paid"
        )
    print '</tbody></table>'
    print '</div>'
    
    print end_timer(timer, "Stats View Load")

# ::END:: Stats View

# ::START:: Calendar View
def render_calendar_view():
    """Monthly calendar view showing meetings and trip dates for active mission trips"""

    timer = start_timer()

    # Get current month/year from URL parameters or use current date
    today = datetime.date.today()
    try:
        cal_year = int(model.Data.year) if hasattr(model.Data, 'year') and model.Data.year else today.year
        cal_month = int(model.Data.month) if hasattr(model.Data, 'month') and model.Data.month else today.month
    except:
        cal_year = today.year
        cal_month = today.month

    # Calculate first and last day of month
    first_day = datetime.date(cal_year, cal_month, 1)
    if cal_month == 12:
        last_day = datetime.date(cal_year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        last_day = datetime.date(cal_year, cal_month + 1, 1) - datetime.timedelta(days=1)

    # Navigation links
    prev_month = cal_month - 1
    prev_year = cal_year
    if prev_month < 1:
        prev_month = 12
        prev_year -= 1

    next_month = cal_month + 1
    next_year = cal_year
    if next_month > 12:
        next_month = 1
        next_year += 1

    month_names = ['', 'January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December']

    print '''
    <div class="calendar-header" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
        <a href="?view=calendar&year={0}&month={1}" class="btn btn-outline-primary">&larr; {2}</a>
        <h3 style="margin: 0;">{3} {4}</h3>
        <a href="?view=calendar&year={5}&month={6}" class="btn btn-outline-primary">{7} &rarr;</a>
    </div>
    '''.format(
        prev_year, prev_month, month_names[prev_month],
        month_names[cal_month], cal_year,
        next_year, next_month, month_names[next_month]
    )

    # Quick navigation to today
    if cal_year != today.year or cal_month != today.month:
        print '<p class="text-center mb-3"><a href="?view=calendar" class="btn btn-sm btn-secondary">Go to Today</a></p>'

    # Query all meetings for active mission trips in this month
    meetings_sql = '''
    SELECT
        m.MeetingId,
        m.MeetingDate,
        COALESCE(me_desc.Data, m.Description) as Description,
        COALESCE(me_loc.Data, m.Location) as Location,
        o.OrganizationId,
        o.OrganizationName
    FROM Meetings m WITH (NOLOCK)
    INNER JOIN Organizations o WITH (NOLOCK) ON m.OrganizationId = o.OrganizationId
    LEFT JOIN MeetingExtra me_desc WITH (NOLOCK) ON m.MeetingId = me_desc.MeetingId AND me_desc.Field = 'Description'
    LEFT JOIN MeetingExtra me_loc WITH (NOLOCK) ON m.MeetingId = me_loc.MeetingId AND me_loc.Field = 'Location'
    WHERE o.IsMissionTrip = {0}
      AND o.OrganizationStatusId = {1}
      AND m.MeetingDate >= '{2}'
      AND m.MeetingDate < '{3}'
      AND NOT EXISTS (
          SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
          WHERE oe.OrganizationId = o.OrganizationId
            AND oe.Field = 'Close'
            AND oe.DateValue IS NOT NULL
            AND oe.DateValue <= GETDATE()
      )
    ORDER BY m.MeetingDate
    '''.format(
        config.MISSION_TRIP_FLAG,
        config.ACTIVE_ORG_STATUS_ID,
        first_day.strftime('%Y-%m-%d'),
        (last_day + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )

    # Query trip dates (start and end) for trips in this month range
    trips_sql = '''
    SELECT
        o.OrganizationId,
        o.OrganizationName,
        oe_start.DateValue as TripStart,
        oe_end.DateValue as TripEnd
    FROM Organizations o WITH (NOLOCK)
    LEFT JOIN OrganizationExtra oe_start ON o.OrganizationId = oe_start.OrganizationId
        AND oe_start.Field = 'Main Event Start'
    LEFT JOIN OrganizationExtra oe_end ON o.OrganizationId = oe_end.OrganizationId
        AND oe_end.Field = 'Main Event End'
    WHERE o.IsMissionTrip = {0}
      AND o.OrganizationStatusId = {1}
      AND (
          (oe_start.DateValue >= '{2}' AND oe_start.DateValue < '{3}')
          OR (oe_end.DateValue >= '{2}' AND oe_end.DateValue < '{3}')
          OR (oe_start.DateValue < '{2}' AND oe_end.DateValue >= '{3}')
      )
      AND NOT EXISTS (
          SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
          WHERE oe.OrganizationId = o.OrganizationId
            AND oe.Field = 'Close'
            AND oe.DateValue IS NOT NULL
            AND oe.DateValue <= GETDATE()
      )
    ORDER BY oe_start.DateValue
    '''.format(
        config.MISSION_TRIP_FLAG,
        config.ACTIVE_ORG_STATUS_ID,
        first_day.strftime('%Y-%m-%d'),
        (last_day + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )

    try:
        meetings = list(q.QuerySql(meetings_sql))
        trips = list(q.QuerySql(trips_sql))
    except Exception as e:
        print '<div class="alert alert-danger">Error loading calendar data: {0}</div>'.format(str(e))
        return

    # Build events dictionary by date
    events_by_date = {}

    # Add meetings
    for meeting in meetings:
        meeting_date = _convert_to_python_date(meeting.MeetingDate)
        if meeting_date:
            date_key = meeting_date.strftime('%Y-%m-%d')
            if date_key not in events_by_date:
                events_by_date[date_key] = {'meetings': [], 'trip_starts': [], 'trip_ends': [], 'trip_ongoing': []}

            meeting_time = meeting_date.strftime('%I:%M %p') if hasattr(meeting_date, 'strftime') else ''
            events_by_date[date_key]['meetings'].append({
                'id': meeting.MeetingId,
                'time': meeting_time,
                'description': meeting.Description if meeting.Description else 'Team Meeting',
                'location': meeting.Location if meeting.Location else '',
                'org_id': meeting.OrganizationId,
                'org_name': meeting.OrganizationName
            })

    # Add trip start/end dates
    for trip in trips:
        trip_start = _convert_to_python_date(trip.TripStart)
        trip_end = _convert_to_python_date(trip.TripEnd)

        if trip_start:
            start_key = trip_start.strftime('%Y-%m-%d')
            if start_key not in events_by_date:
                events_by_date[start_key] = {'meetings': [], 'trip_starts': [], 'trip_ends': [], 'trip_ongoing': []}
            events_by_date[start_key]['trip_starts'].append({
                'org_id': trip.OrganizationId,
                'org_name': trip.OrganizationName
            })

        if trip_end:
            end_key = trip_end.strftime('%Y-%m-%d')
            if end_key not in events_by_date:
                events_by_date[end_key] = {'meetings': [], 'trip_starts': [], 'trip_ends': [], 'trip_ongoing': []}
            events_by_date[end_key]['trip_ends'].append({
                'org_id': trip.OrganizationId,
                'org_name': trip.OrganizationName
            })

        # Mark ongoing trip days
        if trip_start and trip_end:
            current = trip_start + datetime.timedelta(days=1)
            while current < trip_end:
                if current.month == cal_month and current.year == cal_year:
                    ongoing_key = current.strftime('%Y-%m-%d')
                    if ongoing_key not in events_by_date:
                        events_by_date[ongoing_key] = {'meetings': [], 'trip_starts': [], 'trip_ends': [], 'trip_ongoing': []}
                    events_by_date[ongoing_key]['trip_ongoing'].append({
                        'org_id': trip.OrganizationId,
                        'org_name': trip.OrganizationName
                    })
                current += datetime.timedelta(days=1)

    # Calendar CSS
    print '''
    <style>
        .calendar-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 1px;
            background: #dee2e6;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            overflow: hidden;
        }
        .calendar-header-cell {
            background: #495057;
            color: white;
            padding: 10px;
            text-align: center;
            font-weight: bold;
        }
        .calendar-day {
            background: white;
            min-height: 100px;
            padding: 5px;
            vertical-align: top;
        }
        .calendar-day.other-month {
            background: #f8f9fa;
            color: #adb5bd;
        }
        .calendar-day.today {
            background: #e3f2fd;
        }
        .calendar-day.has-trip {
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
        }
        .day-number {
            font-weight: bold;
            font-size: 1.1rem;
            margin-bottom: 5px;
        }
        .calendar-event {
            font-size: 0.75rem;
            padding: 2px 4px;
            margin: 2px 0;
            border-radius: 3px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: pointer;
        }
        .event-meeting {
            background: #bbdefb;
            color: #1565c0;
            border-left: 3px solid #1976d2;
        }
        .event-trip-start {
            background: #c8e6c9;
            color: #2e7d32;
            border-left: 3px solid #4caf50;
        }
        .event-trip-end {
            background: #ffcdd2;
            color: #c62828;
            border-left: 3px solid #f44336;
        }
        .event-trip-ongoing {
            background: #fff9c4;
            color: #f57f17;
            font-size: 0.7rem;
            border-left: 3px solid #ffc107;
            cursor: pointer;
        }
        /* Different colors for multiple ongoing trips */
        .event-trip-ongoing-0 { background: #fff9c4; color: #f57f17; border-left-color: #ffc107; }
        .event-trip-ongoing-1 { background: #e1f5fe; color: #0277bd; border-left-color: #03a9f4; }
        .event-trip-ongoing-2 { background: #f3e5f5; color: #7b1fa2; border-left-color: #9c27b0; }
        .event-trip-ongoing-3 { background: #e8f5e9; color: #388e3c; border-left-color: #4caf50; }
        .event-trip-ongoing-4 { background: #fce4ec; color: #c2185b; border-left-color: #e91e63; }
        .calendar-event:hover {
            opacity: 0.8;
        }
        @media (max-width: 767px) {
            .calendar-day { min-height: 60px; }
            .calendar-event { font-size: 0.65rem; }
            .day-number { font-size: 0.9rem; }
        }
    </style>
    '''

    # Build calendar grid
    print '<div class="calendar-grid">'

    # Header row
    day_names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
    for day_name in day_names:
        print '<div class="calendar-header-cell">{0}</div>'.format(day_name)

    # Calculate start of calendar (may include days from previous month)
    # First day of month's weekday (0 = Monday in Python, but we want Sunday = 0)
    first_weekday = (first_day.weekday() + 1) % 7  # Convert to Sunday = 0

    # Start from previous month if needed
    calendar_start = first_day - datetime.timedelta(days=first_weekday)

    # Generate 6 weeks of days (42 days max)
    current_date = calendar_start
    for i in range(42):
        date_key = current_date.strftime('%Y-%m-%d')
        is_current_month = current_date.month == cal_month
        is_today = current_date == today

        events = events_by_date.get(date_key, {'meetings': [], 'trip_starts': [], 'trip_ends': [], 'trip_ongoing': []})
        has_trip = len(events['trip_starts']) > 0 or len(events['trip_ends']) > 0 or len(events['trip_ongoing']) > 0

        classes = ['calendar-day']
        if not is_current_month:
            classes.append('other-month')
        if is_today:
            classes.append('today')
        if has_trip and is_current_month:
            classes.append('has-trip')

        print '<div class="{0}">'.format(' '.join(classes))
        print '<div class="day-number">{0}</div>'.format(current_date.day)

        if is_current_month:
            # Trip starts
            for trip in events['trip_starts']:
                short_name = trip['org_name'].replace('STM: ', '')[:20]
                # Add consistent color accent based on org_id
                color_idx = trip['org_id'] % 5
                color_accents = ['#ffc107', '#03a9f4', '#9c27b0', '#4caf50', '#e91e63']
                accent = color_accents[color_idx]
                print '<div class="calendar-event event-trip-start" style="border-right: 4px solid {0};" title="{1} - Trip Starts" onclick="window.location=\'?trip={2}\'">&#9992; {3}</div>'.format(
                    accent,
                    _escape_html(trip['org_name']),
                    trip['org_id'],
                    _escape_html(short_name)
                )

            # Trip ends
            for trip in events['trip_ends']:
                short_name = trip['org_name'].replace('STM: ', '')[:20]
                # Add consistent color accent based on org_id
                color_idx = trip['org_id'] % 5
                color_accents = ['#ffc107', '#03a9f4', '#9c27b0', '#4caf50', '#e91e63']
                accent = color_accents[color_idx]
                print '<div class="calendar-event event-trip-end" style="border-right: 4px solid {0};" title="{1} - Trip Ends" onclick="window.location=\'?trip={2}\'">&#127937; {3}</div>'.format(
                    accent,
                    _escape_html(trip['org_name']),
                    trip['org_id'],
                    _escape_html(short_name)
                )

            # Meetings (limit to 3 shown)
            for j, meeting in enumerate(events['meetings'][:3]):
                short_name = meeting['org_name'].replace('STM: ', '')[:15]
                print '<div class="calendar-event event-meeting" title="{0} - {1}" onclick="window.location=\'?trip={2}&section=meetings\'">{3} {4}</div>'.format(
                    _escape_html(meeting['org_name']),
                    _escape_html(meeting['description']),
                    meeting['org_id'],
                    meeting['time'],
                    _escape_html(short_name)
                )

            if len(events['meetings']) > 3:
                print '<div class="calendar-event" style="background: #e0e0e0; color: #616161;">+{0} more</div>'.format(len(events['meetings']) - 3)

            # Show ongoing trips with individual names and colors
            if len(events['trip_ongoing']) > 0 and len(events['trip_starts']) == 0 and len(events['trip_ends']) == 0:
                # Show each ongoing trip individually (limit to 3 to avoid overflow)
                for ongoing in events['trip_ongoing'][:3]:
                    # Use org_id to consistently assign same color to same trip
                    color_idx = ongoing['org_id'] % 5
                    color_class = 'event-trip-ongoing-{0}'.format(color_idx)
                    short_name = ongoing['org_name'][:15] + '..' if len(ongoing['org_name']) > 15 else ongoing['org_name']
                    print '<div class="calendar-event event-trip-ongoing {0}" title="{1} - In Progress" onclick="window.location=\'?trip={2}\'">&#128205; {3}</div>'.format(
                        color_class,
                        _escape_html(ongoing['org_name']),
                        ongoing['org_id'],
                        _escape_html(short_name)
                    )
                if len(events['trip_ongoing']) > 3:
                    print '<div class="calendar-event event-trip-ongoing" style="background: #e0e0e0; color: #616161;">+{0} more trips</div>'.format(len(events['trip_ongoing']) - 3)

        print '</div>'

        current_date += datetime.timedelta(days=1)

        # Stop after reaching next month and completing the week
        if current_date.month != cal_month and current_date.weekday() == 6:
            break

    print '</div>'  # calendar-grid

    # Legend
    print '''
    <div class="mt-4 p-3" style="background: #f8f9fa; border-radius: 8px;">
        <strong>Legend:</strong>
        <span style="margin-left: 15px;"><span class="calendar-event event-trip-start" style="display: inline-block;">&#9992; Trip Start</span></span>
        <span style="margin-left: 15px;"><span class="calendar-event event-trip-end" style="display: inline-block;">&#127937; Trip End</span></span>
        <span style="margin-left: 15px;"><span class="calendar-event event-meeting" style="display: inline-block;">Meeting</span></span>
        <span style="margin-left: 15px;"><span class="calendar-event event-trip-ongoing" style="display: inline-block;">&#128205; Trip Ongoing</span></span>
    </div>
    '''

    # Upcoming events list (next 30 days)
    upcoming_end = today + datetime.timedelta(days=30)
    upcoming_meetings_sql = '''
    SELECT TOP 10
        m.MeetingDate,
        COALESCE(me_desc.Data, m.Description) as Description,
        COALESCE(me_loc.Data, m.Location) as Location,
        o.OrganizationId,
        o.OrganizationName
    FROM Meetings m WITH (NOLOCK)
    INNER JOIN Organizations o WITH (NOLOCK) ON m.OrganizationId = o.OrganizationId
    LEFT JOIN MeetingExtra me_desc WITH (NOLOCK) ON m.MeetingId = me_desc.MeetingId AND me_desc.Field = 'Description'
    LEFT JOIN MeetingExtra me_loc WITH (NOLOCK) ON m.MeetingId = me_loc.MeetingId AND me_loc.Field = 'Location'
    WHERE o.IsMissionTrip = {0}
      AND o.OrganizationStatusId = {1}
      AND m.MeetingDate >= GETDATE()
      AND m.MeetingDate < '{2}'
      AND NOT EXISTS (
          SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
          WHERE oe.OrganizationId = o.OrganizationId
            AND oe.Field = 'Close'
            AND oe.DateValue IS NOT NULL
            AND oe.DateValue <= GETDATE()
      )
    ORDER BY m.MeetingDate
    '''.format(
        config.MISSION_TRIP_FLAG,
        config.ACTIVE_ORG_STATUS_ID,
        upcoming_end.strftime('%Y-%m-%d')
    )

    try:
        upcoming_meetings = list(q.QuerySql(upcoming_meetings_sql))

        if upcoming_meetings:
            print '''
            <div class="card mt-4">
                <div class="card-header"><h5 style="margin:0;">Upcoming Meetings (Next 30 Days)</h5></div>
                <div class="card-body">
                    <div class="list-group list-group-flush">
            '''

            for meeting in upcoming_meetings:
                meeting_date = _convert_to_python_date(meeting.MeetingDate)
                if meeting_date:
                    date_str = meeting_date.strftime('%a, %b %d at %I:%M %p')
                else:
                    date_str = ''

                description = meeting.Description if meeting.Description else 'Team Meeting'
                org_name = meeting.OrganizationName.replace('STM: ', '')

                print '''
                    <a href="?trip={0}&section=meetings" class="list-group-item list-group-item-action">
                        <div class="d-flex justify-content-between align-items-center">
                            <div>
                                <strong>{1}</strong>
                                <br><small class="text-muted">{2}</small>
                            </div>
                            <span class="badge bg-primary">{3}</span>
                        </div>
                    </a>
                '''.format(
                    meeting.OrganizationId,
                    _escape_html(description),
                    _escape_html(org_name),
                    date_str
                )

            print '''
                    </div>
                </div>
            </div>
            '''
    except:
        pass

    print end_timer(timer, "Calendar View Load")

# ::END:: Calendar View

# ::START:: Settings View
def render_settings_view(user_role):
    """Render the Settings page with Email Template Editor."""

    # Only admins can access settings
    if not user_role.get('is_admin', False):
        print '''
        <div class="alert alert-danger">
            <h4>Access Denied</h4>
            <p>You must have admin privileges to access settings.</p>
        </div>
        '''
        return

    print '''
    <div class="settings-view">
        <h2 style="margin-bottom: 20px;">&#9881; Dashboard Settings</h2>

        <!-- Settings Tabs -->
        <div style="display:flex; border-bottom:2px solid #dee2e6; margin-bottom:20px; gap:4px;">
            <div class="settings-tab active" onclick="switchSettingsTab('approvals')" id="stab-approvals" style="padding:10px 20px; font-size:14px; font-weight:600; color:#888; cursor:pointer; border-bottom:2px solid transparent; margin-bottom:-2px;">Approval Workflow</div>
            <div class="settings-tab" onclick="switchSettingsTab('templates')" id="stab-templates" style="padding:10px 20px; font-size:14px; font-weight:600; color:#888; cursor:pointer; border-bottom:2px solid transparent; margin-bottom:-2px;">Quick Email Templates</div>
            <div class="settings-tab" onclick="switchSettingsTab('config')" id="stab-config" style="padding:10px 20px; font-size:14px; font-weight:600; color:#888; cursor:pointer; border-bottom:2px solid transparent; margin-bottom:-2px;">Dashboard Configuration</div>
        </div>
        <style>
            .settings-tab.active { color: #0d6efd !important; border-bottom-color: #0d6efd !important; }
            .settings-tab:hover { color: #0d6efd !important; }
            .settings-tab-content { display: none; }
            .settings-tab-content.active { display: block; }
        </style>
        <script>
        function switchSettingsTab(name) {
            document.querySelectorAll('.settings-tab-content').forEach(function(el) { el.classList.remove('active'); });
            document.querySelectorAll('.settings-tab').forEach(function(el) { el.classList.remove('active'); });
            document.getElementById('stab-' + name).classList.add('active');
            document.getElementById('stab-content-' + name).classList.add('active');
        }
        </script>

        <!-- TAB: Approval Workflow -->
        <div id="stab-content-approvals" class="settings-tab-content active">

        <!-- Approval Workflow Settings Section -->
        <div class="card" style="margin-bottom: 20px;">
            <div class="card-header">
                <h4>&#10003; Approval Workflow Settings</h4>
            </div>
            <div class="card-body">
                <p class="text-muted" style="margin-bottom: 20px;">
                    Control whether trip registrations require approval before participants are confirmed.
                    You can enable/disable approvals globally, and override the setting for individual trips.
                </p>

                <!-- Global Setting -->
                <div style="display: flex; align-items: center; gap: 15px; padding: 15px; background: #f8f9fa; border-radius: 8px; margin-bottom: 15px;">
                    <div style="flex: 1;">
                        <strong style="font-size: 15px;">Global Approval Workflow</strong>
                        <p class="text-muted" style="margin: 5px 0 0 0; font-size: 13px;">
                            When enabled, all trip registrations will require admin approval by default.
                            Individual trips can override this setting.
                        </p>
                    </div>
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <span id="globalApprovalStatus" style="font-size: 13px; color: #666;">Loading...</span>
                        <label class="toggle-switch" style="position: relative; display: inline-block; width: 50px; height: 26px;">
                            <input type="checkbox" id="globalApprovalsToggle" onchange="ApprovalSettings.saveGlobalSetting(this.checked)" style="opacity: 0; width: 0; height: 0;">
                            <span class="toggle-slider" style="position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .3s; border-radius: 26px;"></span>
                        </label>
                    </div>
                </div>

                <!-- Trip Override Info -->
                <div style="padding: 15px; background: #e3f2fd; border-radius: 8px; border-left: 4px solid #2196f3;">
                    <strong style="color: #1565c0;">&#128712; Per-Trip Override</strong>
                    <p style="margin: 5px 0 0 0; font-size: 13px; color: #1565c0;">
                        To override the global setting for a specific trip, go to that trip's <strong>Team</strong> tab and click the <strong>&#9881; Settings</strong> button next to "Status" filter.
                    </p>
                </div>
            </div>
        </div>

        </div> <!-- end TAB: Approval Workflow -->

        <!-- TAB: Dashboard Configuration -->
        <div id="stab-content-config" class="settings-tab-content">

        <!-- Dashboard Configuration Section -->
        <div class="card" style="margin-bottom: 20px;">
            <div class="card-header" style="display: flex; justify-content: space-between; align-items: center;">
                <h4>&#128295; Dashboard Configuration</h4>
                <button type="button" class="btn btn-primary btn-sm" id="saveConfigBtn" onclick="DashConfig.save()" style="display:none;">Save Changes</button>
            </div>
            <div class="card-body" id="configPanel">
                <p class="text-muted" style="margin-bottom:16px;">These settings are stored separately from the script code and survive updates.</p>
                <div id="configLoading" style="color:#888;">Loading configuration...</div>
                <div id="configForm" style="display:none;"></div>
                <div id="configStatus" style="margin-top:12px; font-size:13px;"></div>
            </div>
        </div>
        <script>
        var DashConfig = {
            original: null,
            load: function() {
                $.ajax({
                    url: (function(){ var p = window.location.pathname; if (p.indexOf('/PyScript/') >= 0) p = p.replace('/PyScript/', '/PyScriptForm/'); return p; })(),
                    type: 'POST',
                    data: { action: 'get_config' },
                    success: function(resp) {
                        try {
                            var data = JSON.parse(resp);
                            if (data.success) {
                                DashConfig.original = data.config;
                                DashConfig.render(data.config);
                            } else {
                                document.getElementById('configLoading').innerHTML = '<span style="color:red;">' + (data.message || 'Error') + '</span>';
                            }
                        } catch(e) {
                            document.getElementById('configLoading').innerHTML = '<span style="color:red;">Error loading config</span>';
                        }
                    }
                });
            },
            render: function(cfg) {
                document.getElementById('configLoading').style.display = 'none';
                document.getElementById('configForm').style.display = 'block';
                document.getElementById('saveConfigBtn').style.display = 'inline-block';
                var h = '<div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">';
                h += DashConfig.field('MY_MISSIONS_LINK', 'Share Link (MyMissions URL)', cfg.MY_MISSIONS_LINK, 'text', 'Leave empty to auto-generate');
                h += DashConfig.field('CURRENCY_SYMBOL', 'Currency Symbol', cfg.CURRENCY_SYMBOL, 'text');
                h += DashConfig.field('ITEMS_PER_PAGE', 'Items Per Page', cfg.ITEMS_PER_PAGE, 'number');
                h += DashConfig.field('MEMBER_TYPE_LEADER', 'Leader Member Type ID', cfg.MEMBER_TYPE_LEADER, 'number');
                h += DashConfig.field('ATTENDANCE_TYPE_LEADER', 'Leader Attendance Type ID', cfg.ATTENDANCE_TYPE_LEADER, 'number');
                h += DashConfig.field('SIDEBAR_BG_COLOR', 'Sidebar Background Color', cfg.SIDEBAR_BG_COLOR, 'color');
                h += DashConfig.field('SIDEBAR_ACTIVE_COLOR', 'Sidebar Active Color', cfg.SIDEBAR_ACTIVE_COLOR, 'color');
                h += DashConfig.check('SHOW_CLOSED_BY_DEFAULT', 'Show Closed Trips by Default', cfg.SHOW_CLOSED_BY_DEFAULT);
                h += DashConfig.check('ENABLE_SQL_DEBUG', 'Enable SQL Debug Output', cfg.ENABLE_SQL_DEBUG);
                h += DashConfig.check('ENABLE_STATS_TAB', 'Enable Stats Tab', cfg.ENABLE_STATS_TAB);
                h += DashConfig.check('ENABLE_FINANCE_TAB', 'Enable Finance Tab', cfg.ENABLE_FINANCE_TAB);
                h += DashConfig.check('ENABLE_MESSAGES_TAB', 'Enable Messages Tab', cfg.ENABLE_MESSAGES_TAB);
                h += '</div>';
                h += '<div style="margin-top:16px; border-top:1px solid #e0e0e0; padding-top:16px;">';
                h += '<h5 style="margin-bottom:12px;">Roles & Permissions</h5>';
                h += DashConfig.field('ADMIN_ROLES', 'Admin Roles (comma-separated)', cfg.ADMIN_ROLES, 'text');
                h += DashConfig.field('FINANCE_ROLES', 'Finance Roles (comma-separated)', cfg.FINANCE_ROLES, 'text');
                h += DashConfig.field('LEADER_MEMBER_TYPES', 'Leader Member Type IDs (comma-separated)', cfg.LEADER_MEMBER_TYPES, 'text');
                h += DashConfig.field('APPLICATION_ORG_IDS', 'Application Org IDs to Exclude (comma-separated)', cfg.APPLICATION_ORG_IDS, 'text');
                h += '</div>';
                document.getElementById('configForm').innerHTML = h;
            },
            field: function(name, label, value, type, hint) {
                var inp = type === 'color'
                    ? '<input type="color" name="' + name + '" value="' + (value || '#000000') + '" style="height:36px;border:1px solid #ccc;border-radius:4px;">'
                    : '<input type="' + (type || 'text') + '" name="' + name + '" value="' + (value || '').toString().replace(/"/g, '&quot;') + '" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;" placeholder="' + (hint || '') + '">';
                return '<div style="margin-bottom:8px;"><label style="font-size:12px;font-weight:600;color:#555;display:block;margin-bottom:3px;">' + label + '</label>' + inp + '</div>';
            },
            check: function(name, label, value) {
                return '<div style="margin-bottom:8px;"><label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;"><input type="checkbox" name="' + name + '"' + (value ? ' checked' : '') + '> ' + label + '</label></div>';
            },
            save: function() {
                var form = document.getElementById('configForm');
                var data = { action: 'save_config' };
                var inputs = form.querySelectorAll('input');
                for (var i = 0; i < inputs.length; i++) {
                    var inp = inputs[i];
                    if (inp.type === 'checkbox') {
                        data[inp.name] = inp.checked ? 'true' : 'false';
                    } else {
                        data[inp.name] = inp.value;
                    }
                }
                document.getElementById('configStatus').innerHTML = '<span style="color:#0078d4;">Saving...</span>';
                $.ajax({
                    url: (function(){ var p = window.location.pathname; if (p.indexOf('/PyScript/') >= 0) p = p.replace('/PyScript/', '/PyScriptForm/'); return p; })(),
                    type: 'POST',
                    data: data,
                    success: function(resp) {
                        try {
                            var r = JSON.parse(resp);
                            document.getElementById('configStatus').innerHTML = r.success
                                ? '<span style="color:#107c10;">Configuration saved! Changes take effect on next page load.</span>'
                                : '<span style="color:red;">' + (r.message || 'Error') + '</span>';
                        } catch(e) {
                            document.getElementById('configStatus').innerHTML = '<span style="color:red;">Error saving</span>';
                        }
                    }
                });
            }
        };
        // Load config when jQuery is ready
        (function waitForjQuery() {
            if (typeof $ !== 'undefined' && typeof $.ajax === 'function') {
                DashConfig.load();
            } else {
                setTimeout(waitForjQuery, 100);
            }
        })();
        </script>

        </div> <!-- end TAB: Dashboard Configuration -->

        <div id="stab-content-templates" class="settings-tab-content">

        <!-- Advanced Options Section -->
        <div class="card" style="margin-bottom:20px;">
            <div class="card-header">
                <h4>&#9881; Advanced Options</h4>
            </div>
            <div class="card-body">
                <p class="text-muted">
                    For more advanced email customization, you can use TouchPoint's built-in email tools:
                </p>
                <ul style="margin-bottom: 20px;">
                    <li><strong>Involvement Email</strong> - Use TouchPoint's email editor with merge fields for each trip</li>
                    <li><strong>Special Content</strong> - Store reusable email templates in Admin > Special Content</li>
                </ul>
                <p class="text-muted" style="font-size: 13px;">
                    Templates edited here are stored in Special Content with the prefix "MissionsDashboard_Template_".
                </p>
            </div>
        </div>

        <!-- Email Dropdown Templates Section -->
        <div class="card" style="margin-top: 20px;">
            <div class="card-header" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div>
                    <h4>&#128232; Quick Email Templates</h4>
                    <p class="text-muted" style="margin: 0; font-size: 14px;">
                        Manage the templates that appear in the email compose dropdown. Set permissions to control who can use each template.
                    </p>
                </div>
                <button type="button" class="btn btn-primary" onclick="DropdownTemplates.showAddForm()" style="padding: 8px 16px;">
                    + Add Template
                </button>
            </div>
            <div class="card-body">
                <!-- Template List -->
                <div id="dropdownTemplateList" style="margin-bottom: 20px;">
                    <div style="text-align: center; padding: 40px; color: #666;">
                        <div style="font-size: 24px;">&#8987;</div>
                        <p>Loading templates...</p>
                    </div>
                </div>
            </div>
        </div>

        <!-- Template Edit Modal -->
        <div id="templateEditModal" class="email-modal-overlay">
            <div class="email-modal" style="max-width: 700px;">
                <div class="email-modal-header">
                    <h3 id="templateModalTitle">Add New Template</h3>
                    <button class="email-close-btn" onclick="DropdownTemplates.closeModal()">&times;</button>
                </div>
                <div class="email-modal-body" style="padding: 20px;">
                    <input type="hidden" id="editTemplateId" value="">

                    <!-- Template Name -->
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="font-weight: bold; display: block; margin-bottom: 5px;">Template Name <span style="color: red;">*</span></label>
                        <input type="text" id="editTemplateName" class="form-control" placeholder="e.g., Fundraising Reminder" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px;">
                    </div>

                    <!-- Row: Context and Type -->
                    <div style="display: flex; gap: 15px; margin-bottom: 15px;">
                        <div class="form-group" style="flex: 1;">
                            <label style="font-weight: bold; display: block; margin-bottom: 5px;">Show In Context</label>
                            <select id="editTemplateContext" class="form-control" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px;">
                                <option value="all">All Contexts</option>
                                <option value="team">Team Tab</option>
                                <option value="budget">Budget Tab</option>
                                <option value="meetings">Meetings Tab</option>
                                <option value="overview">Overview Tab</option>
                                <option value="documents">Documents Tab</option>
                                <option value="tasks">Tasks Tab</option>
                            </select>
                        </div>
                        <div class="form-group" style="flex: 1;">
                            <label style="font-weight: bold; display: block; margin-bottom: 5px;">Email Type</label>
                            <select id="editTemplateType" class="form-control" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px;">
                                <option value="both">Both (Team & Individual)</option>
                                <option value="team">Team Emails Only</option>
                                <option value="individual">Individual Emails Only</option>
                            </select>
                        </div>
                    </div>

                    <!-- Permission -->
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="font-weight: bold; display: block; margin-bottom: 5px;">Who Can Use This Template</label>
                        <select id="editTemplateRole" class="form-control" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px;">
                            <option value="all">Everyone (Admins & Leaders)</option>
                            <option value="admin">Admins Only</option>
                            <option value="leader">Leaders Only</option>
                        </select>
                        <p class="text-muted" style="font-size: 12px; margin-top: 5px;">
                            Control who sees this template in their email dropdown.
                        </p>
                    </div>

                    <!-- Subject Line -->
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="font-weight: bold; display: block; margin-bottom: 5px;">Subject Line <span style="color: red;">*</span></label>
                        <input type="text" id="editTemplateSubject" class="form-control" placeholder="e.g., {{TripName}} - Fundraising Reminder" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px;">
                    </div>

                    <!-- Available Placeholders -->
                    <div style="background: #e3f2fd; padding: 10px; border-radius: 6px; margin-bottom: 15px;">
                        <strong style="font-size: 13px;">Available Placeholders (click to insert):</strong>
                        <div style="margin-top: 8px; display: flex; flex-wrap: wrap; gap: 6px;">
                            <code style="background: #fff; padding: 2px 6px; border-radius: 4px; cursor: pointer; font-size: 12px;" onclick="DropdownTemplates.insertPlaceholder('{{TripName}}')">{{TripName}}</code>
                            <code style="background: #fff; padding: 2px 6px; border-radius: 4px; cursor: pointer; font-size: 12px;" onclick="DropdownTemplates.insertPlaceholder('{{PersonName}}')">{{PersonName}}</code>
                            <code style="background: #fff; padding: 2px 6px; border-radius: 4px; cursor: pointer; font-size: 12px;" onclick="DropdownTemplates.insertPlaceholder('{{TripCost}}')">{{TripCost}}</code>
                            <code style="background: #fff; padding: 2px 6px; border-radius: 4px; cursor: pointer; font-size: 12px;" onclick="DropdownTemplates.insertPlaceholder('{{Outstanding}}')">{{Outstanding}}</code>
                            <code style="background: #fff; padding: 2px 6px; border-radius: 4px; cursor: pointer; font-size: 12px;" onclick="DropdownTemplates.insertPlaceholder('{{MyGivingLink}}')">{{MyGivingLink}}</code>
                            <code style="background: #fff; padding: 2px 6px; border-radius: 4px; cursor: pointer; font-size: 12px;" onclick="DropdownTemplates.insertPlaceholder('{{SupportLink}}')">{{SupportLink}}</code>
                        </div>
                    </div>

                    <!-- Body with Formatting Toolbar -->
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="font-weight: bold; display: block; margin-bottom: 5px;">Email Body <span style="color: red;">*</span></label>
                        <div class="email-formatting-toolbar" style="margin-bottom: 0; border-bottom: none; border-radius: 6px 6px 0 0;">
                            <button type="button" onclick="DropdownTemplates.formatText('bold')" title="Bold"><b>B</b></button>
                            <button type="button" onclick="DropdownTemplates.formatText('italic')" title="Italic"><i>I</i></button>
                            <button type="button" onclick="DropdownTemplates.formatText('underline')" title="Underline"><u>U</u></button>
                            <button type="button" onclick="DropdownTemplates.formatText('insertUnorderedList')" title="Bullet List">&#8226; List</button>
                            <button type="button" onclick="DropdownTemplates.formatText('insertOrderedList')" title="Numbered List">1. List</button>
                            <button type="button" onclick="DropdownTemplates.insertLink()" title="Insert Link">&#128279; Link</button>
                            <button type="button" onclick="DropdownTemplates.formatText('removeFormat')" title="Clear Formatting">Clear</button>
                        </div>
                        <div id="editTemplateBody" class="email-body-editor" contenteditable="true" style="min-height: 200px; border-radius: 0 0 6px 6px; border-top: 1px solid #ddd;" placeholder="Enter email body text..."></div>
                    </div>
                </div>
                <div class="email-modal-footer" style="display: flex; justify-content: space-between; align-items: center;">
                    <button type="button" id="deleteTemplateBtn" class="btn" onclick="DropdownTemplates.deleteTemplate()" style="background: #dc3545; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; display: none;">
                        Delete Template
                    </button>
                    <div style="display: flex; gap: 10px;">
                        <button type="button" class="btn" onclick="DropdownTemplates.closeModal()" style="background: #6c757d; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer;">
                            Cancel
                        </button>
                        <button type="button" class="btn btn-primary" onclick="DropdownTemplates.saveTemplate()" style="background: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer;">
                            Save Template
                        </button>
                    </div>
                </div>
            </div>
        </div>

        </div> <!-- end TAB: Quick Email Templates -->

    </div> <!-- end settings-view -->

    <style>
        .settings-view .card {
            border: 1px solid #ddd;
            border-radius: 8px;
            background: #fff;
        }
        .settings-view .card-header {
            background: #f8f9fa;
            padding: 20px;
            border-bottom: 1px solid #ddd;
            border-radius: 8px 8px 0 0;
        }
        .settings-view .card-header h4 {
            margin: 0 0 5px 0;
        }
        .settings-view .card-body {
            padding: 20px;
        }
        .settings-view .form-control {
            border: 1px solid #ddd;
            border-radius: 6px;
            padding: 10px 12px;
            font-size: 14px;
        }
        .settings-view select.form-control {
            height: 44px;
            line-height: 1.5;
            padding: 8px 12px;
            -webkit-appearance: none;
            -moz-appearance: none;
            appearance: none;
            background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
            background-repeat: no-repeat;
            background-position: right 12px center;
            background-size: 16px;
            padding-right: 40px;
        }
        .settings-view .form-control:focus {
            border-color: #2196F3;
            outline: none;
            box-shadow: 0 0 0 3px rgba(33, 150, 243, 0.1);
        }
        .settings-view .btn {
            border-radius: 6px;
            font-weight: 500;
            cursor: pointer;
        }
        .settings-view .btn-primary {
            background: #2196F3;
            border: none;
            color: white;
        }
        .settings-view .btn-primary:hover {
            background: #1976D2;
        }
        .settings-view .btn-outline-secondary {
            background: transparent;
            border: 1px solid #ddd;
            color: #666;
        }
        .settings-view .btn-outline-secondary:hover {
            background: #f5f5f5;
        }
        .settings-view .btn-outline-info {
            background: transparent;
            border: 1px solid #17a2b8;
            color: #17a2b8;
        }
        .settings-view .btn-outline-info:hover {
            background: #17a2b8;
            color: white;
        }
        .settings-view .badge {
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 12px;
        }
        .settings-view .badge-custom {
            background: #4caf50;
            color: white;
        }
        .settings-view .badge-default {
            background: #9e9e9e;
            color: white;
        }
        /* Toggle switch styles */
        .toggle-switch input:checked + .toggle-slider {
            background-color: #4caf50;
        }
        .toggle-slider:before {
            position: absolute;
            content: "";
            height: 20px;
            width: 20px;
            left: 3px;
            bottom: 3px;
            background-color: white;
            transition: .3s;
            border-radius: 50%;
        }
        .toggle-switch input:checked + .toggle-slider:before {
            transform: translateX(24px);
        }
    </style>

    <script>
    // =========================================================================
    // APPROVAL SETTINGS MANAGER - Manages approval workflow settings
    // =========================================================================
    var ApprovalSettings = {
        globalEnabled: true,

        // Initialize and load settings
        init: function() {
            this.loadGlobalSetting();
        },

        // Load global approval setting from server
        loadGlobalSetting: function() {
            var self = this;
            var xhr = new XMLHttpRequest();
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            xhr.open('POST', ajaxUrl, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        var responseText = xhr.responseText.trim();
                        var jsonStart = responseText.indexOf('{"success"');
                        if (jsonStart > 0) {
                            responseText = responseText.substring(jsonStart);
                        }
                        var response = JSON.parse(responseText);
                        if (response.success && response.settings) {
                            self.globalEnabled = response.settings.approvals_enabled !== false;
                            self.updateUI();
                        }
                    } catch (e) {
                        console.error('Error loading approval settings:', e);
                    }
                }
            };
            xhr.send('ajax=1&action=get_global_settings');
        },

        // Update UI to reflect current settings
        updateUI: function() {
            var toggle = document.getElementById('globalApprovalsToggle');
            var status = document.getElementById('globalApprovalStatus');

            if (toggle) {
                toggle.checked = this.globalEnabled;
            }
            if (status) {
                status.textContent = this.globalEnabled ? 'Enabled' : 'Disabled';
                status.style.color = this.globalEnabled ? '#4caf50' : '#f44336';
            }
        },

        // Save global approval setting
        saveGlobalSetting: function(enabled) {
            var self = this;
            var xhr = new XMLHttpRequest();
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            xhr.open('POST', ajaxUrl, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        var responseText = xhr.responseText.trim();
                        var jsonStart = responseText.indexOf('{"success"');
                        if (jsonStart > 0) {
                            responseText = responseText.substring(jsonStart);
                        }
                        var response = JSON.parse(responseText);
                        if (response.success) {
                            self.globalEnabled = enabled;
                            self.updateUI();
                        } else {
                            alert('Error saving setting: ' + (response.message || 'Unknown error'));
                            // Revert toggle
                            var toggle = document.getElementById('globalApprovalsToggle');
                            if (toggle) toggle.checked = !enabled;
                        }
                    } catch (e) {
                        console.error('Error saving approval setting:', e);
                        alert('Error saving setting. Check console for details.');
                    }
                }
            };
            xhr.send('ajax=1&action=save_global_settings&approvals_enabled=' + (enabled ? 'true' : 'false'));
        }
    };

    // Initialize on page load
    document.addEventListener('DOMContentLoaded', function() {
        ApprovalSettings.init();
    });

    // =========================================================================
    // DROPDOWN TEMPLATES MANAGER - Manages email templates in the compose dropdown
    // =========================================================================
    var DropdownTemplates = {
        templates: [],
        isEditing: false,

        // Initialize and load templates
        init: function() {
            this.loadTemplates();
        },

        // Load all dropdown templates from server
        loadTemplates: function() {
            var self = this;
            var xhr = new XMLHttpRequest();
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            xhr.open('POST', ajaxUrl, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        var responseText = xhr.responseText.trim();
                        var jsonStart = responseText.indexOf('{"success"');
                        if (jsonStart > 0) {
                            responseText = responseText.substring(jsonStart);
                        }
                        var response = JSON.parse(responseText);
                        if (response.success) {
                            self.templates = response.templates || [];
                            self.renderTemplateList();
                        } else {
                            self.showError('Error loading templates: ' + (response.message || 'Unknown error'));
                        }
                    } catch (e) {
                        console.error('Error parsing dropdown templates:', e);
                        self.showError('Error loading templates. Check console for details.');
                    }
                }
            };
            xhr.send('ajax=1&action=get_dropdown_templates');
        },

        // Render the template list
        renderTemplateList: function() {
            var container = document.getElementById('dropdownTemplateList');
            if (!container) return;

            if (this.templates.length === 0) {
                container.innerHTML = '<div style="text-align: center; padding: 40px; color: #666;"><p>No templates configured. Click "Add Template" to create one.</p></div>';
                return;
            }

            // Sort templates alphabetically by name
            var sortedTemplates = this.templates.slice().sort(function(a, b) {
                return (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase());
            });

            // Column explanations
            var html = '<div style="background: #f0f7ff; border: 1px solid #b8daff; border-radius: 6px; padding: 12px; margin-bottom: 15px; font-size: 13px;">';
            html += '<strong style="display: block; margin-bottom: 8px;">Column Definitions:</strong>';
            html += '<ul style="margin: 0; padding-left: 20px; line-height: 1.6;">';
            html += '<li><strong>Context</strong> - Which dashboard tab this template appears in. "All" shows everywhere.</li>';
            html += '<li><strong>Type</strong> - Whether template appears for team emails, individual emails, or both.</li>';
            html += '<li><strong>Permission</strong> - Who can use this template: Everyone, Admins Only, or Leaders Only.</li>';
            html += '</ul></div>';

            html += '<table style="width: 100%; border-collapse: collapse;">';
            html += '<thead><tr style="background: #f8f9fa; text-align: left;">';
            html += '<th style="padding: 12px; border-bottom: 2px solid #ddd;">Name</th>';
            html += '<th style="padding: 12px; border-bottom: 2px solid #ddd;">Context</th>';
            html += '<th style="padding: 12px; border-bottom: 2px solid #ddd;">Type</th>';
            html += '<th style="padding: 12px; border-bottom: 2px solid #ddd;">Permission</th>';
            html += '<th style="padding: 12px; border-bottom: 2px solid #ddd; width: 100px;">Actions</th>';
            html += '</tr></thead><tbody>';

            var roleLabels = { all: 'Everyone', admin: 'Admins Only', leader: 'Leaders Only' };
            var contextLabels = { all: 'All', team: 'Team', budget: 'Budget', meetings: 'Meetings', overview: 'Overview', documents: 'Documents', tasks: 'Tasks' };
            var typeLabels = { both: 'Both', team: 'Team', individual: 'Individual' };

            for (var i = 0; i < sortedTemplates.length; i++) {
                var t = sortedTemplates[i];
                // Default missing fields for backwards compatibility
                var tContext = t.context || 'all';
                var tType = t.type || 'both';
                var tRole = t.role || 'all';
                var roleClass = tRole === 'admin' ? 'color: #dc3545;' : (tRole === 'leader' ? 'color: #17a2b8;' : 'color: #28a745;');
                html += '<tr style="border-bottom: 1px solid #eee;">';
                html += '<td style="padding: 12px;"><strong>' + this.escapeHtml(t.name) + '</strong></td>';
                html += '<td style="padding: 12px;">' + (contextLabels[tContext] || tContext) + '</td>';
                html += '<td style="padding: 12px;">' + (typeLabels[tType] || tType) + '</td>';
                html += '<td style="padding: 12px; ' + roleClass + '">' + (roleLabels[tRole] || tRole) + '</td>';
                html += '<td style="padding: 12px;">';
                html += '<button onclick="DropdownTemplates.editTemplate(\\'' + t.id + '\\')" style="background: #6c757d; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; margin-right: 5px;" title="Edit">Edit</button>';
                html += '</td>';
                html += '</tr>';
            }

            html += '</tbody></table>';
            container.innerHTML = html;
        },

        // Show add form
        showAddForm: function() {
            this.isEditing = false;
            document.getElementById('templateModalTitle').textContent = 'Add New Template';
            document.getElementById('editTemplateId').value = '';
            document.getElementById('editTemplateName').value = '';
            document.getElementById('editTemplateContext').value = 'all';
            document.getElementById('editTemplateType').value = 'both';
            document.getElementById('editTemplateRole').value = 'all';
            document.getElementById('editTemplateSubject').value = '';
            document.getElementById('editTemplateBody').innerHTML = '';
            document.getElementById('deleteTemplateBtn').style.display = 'none';
            document.getElementById('templateEditModal').classList.add('active');
        },

        // Edit existing template
        editTemplate: function(templateId) {
            var template = null;
            for (var i = 0; i < this.templates.length; i++) {
                if (this.templates[i].id === templateId) {
                    template = this.templates[i];
                    break;
                }
            }
            if (!template) return;

            this.isEditing = true;
            document.getElementById('templateModalTitle').textContent = 'Edit Template';
            document.getElementById('editTemplateId').value = template.id;
            document.getElementById('editTemplateName').value = template.name;
            document.getElementById('editTemplateContext').value = template.context || 'all';
            document.getElementById('editTemplateType').value = template.type || 'both';
            document.getElementById('editTemplateRole').value = template.role || 'all';
            document.getElementById('editTemplateSubject').value = template.subject || '';
            // Convert \\n to <br> for contenteditable div
            var body = (template.body || '').replace(/\\\\n/g, '<br>').replace(/\\n/g, '<br>');
            document.getElementById('editTemplateBody').innerHTML = body;
            document.getElementById('deleteTemplateBtn').style.display = 'block';
            document.getElementById('templateEditModal').classList.add('active');
        },

        // Close modal
        closeModal: function() {
            document.getElementById('templateEditModal').classList.remove('active');
        },

        // Format text in the body editor
        formatText: function(command) {
            document.execCommand(command, false, null);
            document.getElementById('editTemplateBody').focus();
        },

        // Insert a link
        insertLink: function() {
            var url = prompt('Enter URL:', 'https://');
            if (url) {
                if (!url.match(/^https?:\\/\\//)) {
                    url = 'https://' + url;
                }
                document.execCommand('createLink', false, url);
                document.getElementById('editTemplateBody').focus();
            }
        },

        // Insert placeholder at cursor
        insertPlaceholder: function(placeholder) {
            var editor = document.getElementById('editTemplateBody');
            editor.focus();
            document.execCommand('insertText', false, placeholder);
        },

        // Save template
        saveTemplate: function() {
            var name = document.getElementById('editTemplateName').value.trim();
            var subject = document.getElementById('editTemplateSubject').value.trim();
            var bodyHtml = document.getElementById('editTemplateBody').innerHTML.trim();

            // Convert HTML to plain text with line breaks for storage
            var body = bodyHtml
                .replace(/<br\\s*\\/?>/gi, '\\n')
                .replace(/<\\/p><p>/gi, '\\n\\n')
                .replace(/<\\/div><div>/gi, '\\n')
                .replace(/<[^>]+>/g, '')
                .trim();

            if (!name || !subject || !body) {
                alert('Please fill in all required fields (Name, Subject, and Body).');
                return;
            }

            var templateData = {
                id: document.getElementById('editTemplateId').value || this.generateId(name),
                name: name,
                context: document.getElementById('editTemplateContext').value,
                type: document.getElementById('editTemplateType').value,
                role: document.getElementById('editTemplateRole').value,
                subject: subject,
                body: body.replace(/\\n/g, '\\\\n')  // Convert newlines for storage
            };

            var self = this;
            var xhr = new XMLHttpRequest();
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            xhr.open('POST', ajaxUrl, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            var responseText = xhr.responseText.trim();
                            // Try to find JSON in the response
                            var jsonStart = responseText.indexOf('{"success"');
                            if (jsonStart === -1) {
                                jsonStart = responseText.indexOf('{"error"');
                            }
                            if (jsonStart > 0) {
                                responseText = responseText.substring(jsonStart);
                            }

                            var response = JSON.parse(responseText);

                            if (response.success === true) {
                                self.closeModal();
                                self.loadTemplates();
                                // Reset MissionsEmail cache so it reloads templates next time
                                if (typeof MissionsEmail !== 'undefined') {
                                    MissionsEmail.templatesLoaded = false;
                                }
                            } else {
                                alert('Error saving template: ' + (response.message || 'Unknown error'));
                            }
                        } catch (e) {
                            console.error('Error parsing save response:', e, xhr.responseText);
                            alert('Error saving template. Check browser console (F12) for details.');
                        }
                    } else {
                        console.error('HTTP error:', xhr.status, xhr.statusText);
                        alert('Error saving template. HTTP status: ' + xhr.status);
                    }
                }
            };

            // Use tpl_ prefix for all params to avoid ASP.NET reserved words
            // Base64 encode subject and body to bypass HTML validation
            function b64Encode(str) {
                return btoa(unescape(encodeURIComponent(str || '')));
            }

            var encodedData = 'ajax=1&action=save_dropdown_template' +
                '&tpl_id=' + encodeURIComponent(templateData.id) +
                '&tpl_name=' + encodeURIComponent(templateData.name) +
                '&tpl_context=' + encodeURIComponent(templateData.context) +
                '&tpl_type=' + encodeURIComponent(templateData.type) +
                '&tpl_role=' + encodeURIComponent(templateData.role) +
                '&tpl_subject=' + encodeURIComponent(b64Encode(templateData.subject)) +
                '&tpl_body=' + encodeURIComponent(b64Encode(templateData.body));
            xhr.send(encodedData);
        },

        // Delete template
        deleteTemplate: function() {
            var templateId = document.getElementById('editTemplateId').value;
            if (!templateId) return;

            if (!confirm('Are you sure you want to delete this template? This cannot be undone.')) {
                return;
            }

            var self = this;
            var xhr = new XMLHttpRequest();
            var ajaxUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            xhr.open('POST', ajaxUrl, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        var responseText = xhr.responseText.trim();
                        var jsonStart = responseText.indexOf('{"success"');
                        if (jsonStart > 0) {
                            responseText = responseText.substring(jsonStart);
                        }
                        var response = JSON.parse(responseText);
                        if (response.success) {
                            self.closeModal();
                            self.loadTemplates();
                            // Reset MissionsEmail cache so it reloads templates next time
                            if (typeof MissionsEmail !== 'undefined') {
                                MissionsEmail.templatesLoaded = false;
                            }
                            alert('Template deleted successfully!');
                        } else {
                            alert('Error deleting template: ' + (response.message || 'Unknown error'));
                        }
                    } catch (e) {
                        console.error('Error parsing delete response:', e);
                        alert('Error deleting template. Check console for details.');
                    }
                }
            };
            xhr.send('ajax=1&action=delete_dropdown_template&template_id=' + encodeURIComponent(templateId));
        },

        // Generate ID from name
        generateId: function(name) {
            return name.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '') + '_' + Date.now();
        },

        // Escape HTML
        escapeHtml: function(text) {
            var div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        },

        // Show error message
        showError: function(message) {
            var container = document.getElementById('dropdownTemplateList');
            if (container) {
                container.innerHTML = '<div style="text-align: center; padding: 40px; color: #dc3545;"><p>' + message + '</p></div>';
            }
        }
    };

    // Initialize on page load
    document.addEventListener('DOMContentLoaded', function() {
        DropdownTemplates.init();
    });
    </script>
    '''

# ::END:: Settings View

# ::START:: Messages View
def render_messages_view():
    """Messages view with email history"""
    
    timer = start_timer()
    
    # Check for show single parameter
    show_single = model.Data.ShowSingle == '1' if hasattr(model.Data, 'ShowSingle') else False
    having_clause = '' if show_single else 'HAVING COUNT(eqt.PeopleId) > 1'
    
    # Toggle link
    toggle_text = "Hide Single Messages" if show_single else "Show Single Messages"
    toggle_value = "0" if show_single else "1"
    print '<p><a href="?simplenav=messages&ShowSingle={0}" class="btn btn-sm btn-secondary">{1}</a></p>'.format(toggle_value, toggle_text)
    
    print '<h3>Email Communications</h3>'
    
    emails_sql = '''
        SELECT TOP 100
            eqt.Id, 
            CONVERT(VARCHAR(10), eqt.Sent, 101) AS SentDate,
            eq.Subject, 
            eq.Body, 
            COUNT(eqt.PeopleId) AS PeopleCount,
            o.OrganizationName,
            CASE 
                WHEN o.RegEnd < GETDATE() THEN 2
                WHEN o.RegStart > GETDATE() THEN 0
                ELSE 1
            END AS SortOrder
        FROM EmailQueueTo eqt WITH (NOLOCK)
        INNER JOIN EmailQueue eq WITH (NOLOCK) ON eq.Id = eqt.Id
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = eqt.PeopleId
        INNER JOIN Organizations o WITH (NOLOCK) ON o.OrganizationId = eqt.OrgId
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND eqt.Sent IS NOT NULL
        GROUP BY eqt.Id, eq.Subject, CONVERT(VARCHAR(10), eqt.Sent, 101), eq.Body, 
                 o.OrganizationName, o.RegStart, o.RegEnd
        {2}
        ORDER BY CONVERT(VARCHAR(10), eqt.Sent, 101) DESC
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, having_clause)
    
    emails = execute_query_with_debug(emails_sql, "Email History Query", "sql")
    
    print '<table class="mission-table">'
    print '''
    <thead>
        <tr>
            <th>Date</th>
            <th>Subject</th>
            <th>Mission</th>
            <th>Recipients</th>
        </tr>
    </thead>
    <tbody>
    '''
    
    for email in emails:
        print '''
        <tr>
            <td data-label="Date">{0}</td>
            <td data-label="Subject">
                <a href="#" onclick="showEmailModal({3}); return false;">{1}</a>
            </td>
            <td data-label="Mission">{2}</td>
            <td data-label="Recipients" class="text-center">{4}</td>
        </tr>
        '''.format(
            format_date(str(email.SentDate)),
            email.Subject,
            email.OrganizationName,
            email.Id,
            email.PeopleCount
        )
    
    print '</tbody></table>'
    
    # Email modal container
    print '''
    <div id="emailModal" class="modal" style="display: none;">
        <div class="modal-content">
            <span class="close" onclick="closeEmailModal()">&times;</span>
            <div id="emailContent"></div>
        </div>
    </div>
    
    <script>
    function showEmailModal(emailId) {
        document.getElementById('emailModal').style.display = 'block';
        // Load email content via AJAX or inline data
    }
    
    function closeEmailModal() {
        document.getElementById('emailModal').style.display = 'none';
    }
    </script>
    '''
    
    print end_timer(timer, "Messages View Load")

# ::END:: Messages View

# ::START:: Organization Detail View
def render_organization_view(org_id):
    """Detailed view for a specific organization"""
    
    timer = start_timer()
    
    # Get organization info with optimized CTE for outstanding amount
    org_info_sql = '''
        {1}
        SELECT 
            o.OrganizationId,
            o.OrganizationName,
            o.Description,
            o.ImageUrl,
            o.RegStart,
            o.RegEnd,
            CASE 
                WHEN o.RegStart > GETDATE() AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Pending'
                WHEN o.RegStart IS NULL AND (o.RegEnd IS NULL OR o.RegEnd > GETDATE()) THEN 'Active'
                WHEN o.RegEnd < GETDATE() THEN 'Closed'
                ELSE 'Active'
            END AS RegistrationStatus,
            ISNULL((SELECT SUM(Due) FROM MissionTripTotals WHERE InvolvementId = o.OrganizationId), 0) AS Outstanding
        FROM Organizations o WITH (NOLOCK)
        WHERE o.OrganizationId = {0}
    '''.format(org_id, get_mission_trip_totals_cte(include_closed=True))
    
    org_info = execute_query_with_debug(org_info_sql, "Organization Info Query", "top1")
    
    if not org_info:
        print '<div class="alert alert-danger">Organization not found.</div>'
        return
    
    # Header with image
    if org_info.ImageUrl:
        print '<img src="{0}" alt="{1}" style="width: 100%; max-height: 300px; object-fit: cover; border-radius: var(--radius); margin-bottom: 20px;">'.format(
            org_info.ImageUrl, org_info.OrganizationName
        )
    
    print '<h2>{0} <a href="/Org/{1}" target="_blank" class="btn btn-sm btn-primary">View in TouchPoint</a></h2>'.format(
        org_info.OrganizationName, org_info.OrganizationId
    )
    
    # Organization details
    print '''
    <div class="org-details" style="background: white; padding: 20px; border-radius: var(--radius); margin-bottom: 20px;">
        <p><strong>Status:</strong> <span class="status-badge status-{0}">{0}</span></p>
        <p><strong>Total Outstanding:</strong> {1}</p>
        <p><strong>Registration:</strong> {2} to {3}</p>
    </div>
    '''.format(
        org_info.RegistrationStatus.lower(),
        format_currency(org_info.Outstanding),
        format_date(str(org_info.RegStart)) if org_info.RegStart else "Not set",
        format_date(str(org_info.RegEnd)) if org_info.RegEnd else "Not set"
    )
    
    # Meetings section
    print '<h3>Meetings</h3>'
    meetings_sql = '''
        SELECT 
            Description,
            Location,
            MeetingDate,
            CASE WHEN MeetingDate < GETDATE() THEN 'Past' ELSE 'Upcoming' END AS Status
        FROM Meetings WITH (NOLOCK)
        WHERE OrganizationId = {0}
        ORDER BY MeetingDate DESC
    '''.format(org_id)
    
    meetings = execute_query_with_debug(meetings_sql, "Organization Meetings Query", "sql")
    
    if meetings:
        print '<table class="mission-table">'
        print '<tbody>'
        for meeting in meetings:
            status_class = 'text-muted' if meeting.Status == 'Past' else 'text-primary'
            print '''
            <tr>
                <td class="{0}">
                    <strong>{1}</strong><br>
                    📅 {2}<br>
                    📍 {3}
                </td>
            </tr>
            '''.format(
                status_class,
                meeting.Description or "No description",
                format_date(str(meeting.MeetingDate)),
                meeting.Location or "TBD"
            )
        print '</tbody></table>'
    else:
        print '<p>No meetings scheduled.</p>'
    
    # Team members section
    print '<h3 style="margin-top: 30px;">Team Members</h3>'
    
    members_sql = '''
        {2}
        SELECT 
            p.PeopleId,
            p.Name2,
            p.Age,
            p.CellPhone,
            p.EmailAddress,
            mt.Description AS Role,
            pic.SmallUrl AS Picture,
            rr.passportnumber,
            rr.passportexpires,
            v.ProcessedDate AS BackgroundCheckDate,
            vs.Description AS BackgroundCheckStatus,
            mtt.Raised AS Paid,
            mtt.Due AS Outstanding,
            mtt.TripCost AS TotalFee
        FROM OrganizationMembers om WITH (NOLOCK)
        INNER JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
        LEFT JOIN lookup.MemberType mt WITH (NOLOCK) ON mt.Id = om.MemberTypeId
        LEFT JOIN Picture pic WITH (NOLOCK) ON pic.PictureId = p.PictureId
        LEFT JOIN RecReg rr WITH (NOLOCK) ON rr.PeopleId = p.PeopleId
        LEFT JOIN Volunteer v WITH (NOLOCK) ON v.PeopleId = p.PeopleId
        LEFT JOIN lookup.VolApplicationStatus vs WITH (NOLOCK) ON vs.Id = v.StatusId
        LEFT JOIN MissionTripTotals mtt ON mtt.PeopleId = p.PeopleId AND mtt.InvolvementId = om.OrganizationId
        WHERE om.OrganizationId = {0}
          AND om.MemberTypeId <> {1}
        ORDER BY mt.Description, p.Name2
    '''.format(org_id, config.MEMBER_TYPE_LEADER, get_mission_trip_totals_cte(include_closed=True))
    
    members = execute_query_with_debug(members_sql, "Organization Members Query", "sql")
    
    # Wrap all members in a container div
    print '<div style="max-width: 100%; overflow: hidden;">'
    
    last_role = None
    current_role_container_open = False
    
    for member in members:
        # Role separator
        if member.Role != last_role:
            # Close previous role container if open
            if current_role_container_open:
                print '</div><!-- End role container -->'
            
            # Role header
            print '<h4 style="margin-top: 20px; padding: 10px; background: var(--primary-color); color: white; border-radius: var(--radius) var(--radius) 0 0;">{0}</h4>'.format(member.Role)
            # Start new role container
            print '<div style="background: #f9f9f9; padding: 15px; border-radius: 0 0 var(--radius) var(--radius); margin-bottom: 20px;">'
            current_role_container_open = True
            last_role = member.Role
        
        # Member card
        print '''
        <div id="person_{0}" class="member-card" style="background: white; padding: 15px; margin-bottom: 15px; border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden;">
            <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                <div style="flex-shrink: 0;">
                    <img src="{1}" alt="{2}" style="width: 60px; height: 60px; border-radius: 50%; object-fit: cover;">
                </div>
                <div style="flex: 1; min-width: 0;">
                    <h4 style="margin: 0; overflow: hidden; text-overflow: ellipsis;"><a href="/Person2/{0}" target="_blank">{2}</a> ({3})</h4>
                    <p style="margin: 5px 0; overflow: hidden; text-overflow: ellipsis;">📱 {4} | ✉️ {5}</p>
                    <div style="display: flex; gap: 10px; margin-top: 10px; flex-wrap: wrap;">
                        <span class="badge {6}">🛂 Passport</span>
                        <span class="badge {7}">🛡️ Background</span>
                    </div>'''.format(
            member.PeopleId,
            member.Picture or '/Content/default-avatar.png',
            member.Name2,
            member.Age,
            model.FmtPhone(member.CellPhone) if member.CellPhone else 'No phone',
            member.EmailAddress or 'No email',
            'status-normal' if member.passportnumber and member.passportexpires else 'status-urgent',
            'status-normal' if member.BackgroundCheckStatus == 'Complete' and member.BackgroundCheckDate and member.BackgroundCheckDate >= model.DateAddDays(model.DateTime, -1095) else 'status-urgent'
        )
        
        # Payment progress
        if member.TotalFee and member.TotalFee > 0:
            percentage = (float(member.Paid or 0) / float(member.TotalFee)) * 100
            progress_label = "${:,.0f} paid of ${:,.0f}".format(member.Paid or 0, member.TotalFee)
            print '''
                    <div style="margin-top: 10px;">
                        <div class="progress-container" style="max-width: 100%;">
                            <div class="progress-bar" style="width: {0:.1f}%;"></div>
                        </div>
                        <div class="progress-text">{1}</div>
                    </div>'''.format(percentage, progress_label)
        
        print '''
                </div>
            </div>
        </div><!-- End member card -->'''
    
    # Close the last role container if it was opened
    if current_role_container_open:
        print '</div><!-- End last role container -->'
    
    # Close the main container
    print '</div><!-- End members container -->'
    
    print end_timer(timer, "Organization View Load")

# ::END:: Organization Detail View

#####################################################################
# MAIN CONTROLLER
#####################################################################

# ::START:: Main Controller
def main():
    """Main entry point for the dashboard with sidebar navigation"""
    try:
        # Check if this is an AJAX request first
        if handle_ajax_request():
            return

        # Get user role for sidebar and access control
        user_role = get_user_role_and_trips()

        # Parse URL parameters - support both new and legacy formats
        # New format: ?trip=123&section=overview or ?view=finance
        # Legacy format: ?simplenav=due or ?OrgView=123

        # New parameters
        trip_id = str(model.Data.trip) if hasattr(model.Data, 'trip') and model.Data.trip else None
        section = str(model.Data.section) if hasattr(model.Data, 'section') and model.Data.section else 'overview'
        view = str(model.Data.view) if hasattr(model.Data, 'view') and model.Data.view else None

        # Legacy parameter handling for backward compatibility
        legacy_nav = str(model.Data.simplenav) if hasattr(model.Data, 'simplenav') and model.Data.simplenav else None
        legacy_org = str(model.Data.OrgView) if hasattr(model.Data, 'OrgView') and model.Data.OrgView else None

        # Convert legacy parameters to new format
        if legacy_org and not trip_id:
            trip_id = legacy_org
            section = 'overview'

        if legacy_nav and not view:
            # Map legacy simplenav values to new view values
            nav_map = {
                'due': 'finance',
                'messages': 'messages',
                'stats': 'stats',
                'dashboard': 'home'
            }
            view = nav_map.get(legacy_nav, 'home')

        # Default view if nothing specified
        if not trip_id and not view:
            view = 'home'

        # Access control for trip views
        if trip_id:
            if not has_trip_access(user_role, trip_id):
                print get_modern_styles()
                print_access_denied()
                return

        # Non-admins can only see their trips, not global admin views
        if not user_role.get('is_admin', False) and view in ['finance', 'stats', 'messages', 'settings', 'review']:
            # Non-admins trying to access admin-only views - show leader dashboard instead
            view = 'leader_home'

        # Output styles
        print get_modern_styles()

        # Output visualization for developers (only for admins)
        if user_role.get('is_admin', False):
            print create_visualization_diagram()

        # Check if this is a regular member (not admin, not leader) viewing My Missions
        is_member_only = (not user_role.get('is_admin', False) and
                         not user_role.get('is_trip_leader', False) and
                         user_role.get('is_trip_member', False))

        # Members viewing My Missions portal don't need the sidebar
        if is_member_only and (view == 'my_missions' or (not trip_id and not view) or view == 'home'):
            # Simplified layout for members - no sidebar
            render_member_dashboard_view(user_role)
            return

        # Determine current view for sidebar highlighting
        current_view = view if not trip_id else None

        # Output app container with sidebar layout
        print '<div class="app-container">'

        # Render sidebar
        print render_sidebar(user_role, trip_id, section, current_view)

        # Main content panel
        print '<main class="main-panel" id="mainContent">'

        # Route to appropriate view
        if trip_id:
            # Trip-specific section view
            print render_trip_section(trip_id, section, user_role)
        elif view == 'finance':
            render_finance_view()
        elif view == 'messages':
            render_messages_view()
        elif view == 'stats':
            render_stats_view()
        elif view == 'calendar':
            render_calendar_view()
        elif view == 'settings':
            render_settings_view(user_role)
        elif view == 'review':
            render_review_view()
        elif view == 'leader_home':
            # Explicit leader dashboard
            render_leader_dashboard_view(user_role)
        elif view == 'my_missions':
            # Explicit member dashboard (My Missions portal)
            render_member_dashboard_view(user_role)
        else:
            # Default to appropriate dashboard based on role
            if user_role.get('is_admin', False):
                render_dashboard_view()
            elif user_role.get('is_trip_leader', False):
                # Leaders see the leader management dashboard
                render_leader_dashboard_view(user_role)
            elif user_role.get('is_trip_member', False):
                # Regular members see their personal My Missions portal
                render_member_dashboard_view(user_role)
            else:
                # User has no trips - show appropriate message
                print '''
                <div style="max-width: 600px; margin: 50px auto; padding: 40px; text-align: center; background: #f8f9fa; border-radius: 12px;">
                    <div style="font-size: 48px; margin-bottom: 20px;">&#127758;</div>
                    <h2 style="color: #333; margin-bottom: 15px;">Welcome to My Missions</h2>
                    <p style="color: #666; margin-bottom: 20px;">
                        You are not currently registered for any mission trips.
                    </p>
                    <p style="color: #666;">
                        If you believe this is an error, please contact your missions director.
                    </p>
                </div>
                '''

        # Add floating action button for email (only when viewing a trip)
        if trip_id:
            # Get team emails for the FAB
            fab_emails_sql = '''
            SELECT p.EmailAddress
            FROM OrganizationMembers om WITH (NOLOCK)
            JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
            WHERE om.OrganizationId = {0}
              AND om.MemberTypeId NOT IN (230, 311)
              AND om.InactiveDate IS NULL
              AND p.EmailAddress IS NOT NULL
              AND p.EmailAddress != ''
            '''.format(trip_id)
            fab_emails_result = list(q.QuerySql(fab_emails_sql))
            fab_emails = [r.EmailAddress for r in fab_emails_result]
            fab_emails_js = ','.join(fab_emails) if fab_emails else ''

            print '''
            <div class="fab-container" id="emailFab">
                <span class="fab-tooltip">Email Team ({1})</span>
                <button class="fab-button" onclick="MissionsEmail.openTeam('{0}', '{2}')" title="Email Team">
                    &#9993;
                </button>
            </div>
            '''.format(trip_id, len(fab_emails), fab_emails_js.replace("'", "\\'"))

        # Add email modal for individual emails
        print '''
        <div class="email-modal-overlay" id="emailModal">
            <div class="email-modal">
                <div class="email-modal-header">
                    <h4 id="emailModalTitle">Send Email</h4>
                    <button class="email-modal-close" onclick="MissionsEmail.close()">&times;</button>
                </div>
                <div class="email-modal-body">
                    <div class="email-form-group">
                        <label>To:</label>
                        <input type="text" id="emailTo" readonly style="background: #f5f5f5;">
                    </div>
                    <div class="email-form-group">
                        <label>Template:</label>
                        <select id="emailTemplate" onchange="MissionsEmail.applyTemplate()" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                            <option value="">-- Select a template (optional) --</option>
                        </select>
                    </div>
                    <div class="email-form-group">
                        <label>Subject:</label>
                        <input type="text" id="emailSubject" placeholder="Enter subject...">
                    </div>
                    <div class="email-form-group">
                        <label>Message:</label>
                        <div class="email-formatting-toolbar">
                            <button type="button" onclick="MissionsEmail.formatText('bold')" title="Bold"><b>B</b></button>
                            <button type="button" onclick="MissionsEmail.formatText('italic')" title="Italic"><i>I</i></button>
                            <button type="button" onclick="MissionsEmail.formatText('underline')" title="Underline"><u>U</u></button>
                            <button type="button" onclick="MissionsEmail.formatText('insertUnorderedList')" title="Bullet List">&#8226; List</button>
                            <button type="button" onclick="MissionsEmail.formatText('insertOrderedList')" title="Numbered List">1. List</button>
                            <button type="button" onclick="MissionsEmail.insertLink()" title="Insert Link">&#128279; Link</button>
                            <button type="button" onclick="MissionsEmail.formatText('removeFormat')" title="Clear Formatting">Clear</button>
                        </div>
                        <div id="emailBody" class="email-body-editor" contenteditable="true" placeholder="Type your message here..."></div>
                    </div>
                </div>
                <div class="email-modal-footer">
                    <a href="#" id="emailAdvancedLink" onclick="MissionsEmail.openAdvanced(); return false;" class="email-advanced-link" style="display: none;">
                        &#9881; Advanced Options
                    </a>
                    <div class="email-footer-buttons">
                        <button class="email-btn email-btn-cancel" onclick="MissionsEmail.close()">Cancel</button>
                        <button class="email-btn email-btn-send" onclick="MissionsEmail.send()">Send via TouchPoint</button>
                    </div>
                </div>
            </div>
        </div>
        '''

        # Add fee adjustment modal for admins
        if user_role.get('is_admin', False):
            print '''
            <div class="email-modal-overlay" id="feeModal">
                <div class="email-modal" style="max-width: 400px;">
                    <div class="email-modal-header" style="background: #ff9800;">
                        <h4 id="feeModalTitle">Adjust Fee</h4>
                        <button class="email-modal-close" onclick="MissionsFee.close()">&times;</button>
                    </div>
                    <div class="email-modal-body">
                        <div class="email-form-group">
                            <label>Current Trip Cost:</label>
                            <div id="feeCurrentCost" style="font-size: 18px; font-weight: bold; color: #333;">$0.00</div>
                        </div>
                        <div class="email-form-group">
                            <label>Adjustment Amount:</label>
                            <input type="number" id="feeAdjustAmount" step="0.01" placeholder="Enter amount (positive to add, negative to reduce)">
                            <small style="color: #666; display: block; margin-top: 4px;">
                                Positive = reduce cost (credit), Negative = increase cost
                            </small>
                        </div>
                        <div class="email-form-group">
                            <label>Description (optional):</label>
                            <input type="text" id="feeDescription" placeholder="Reason for adjustment...">
                        </div>
                    </div>
                    <div class="email-modal-footer">
                        <button class="email-btn email-btn-cancel" onclick="MissionsFee.close()">Cancel</button>
                        <button class="email-btn fee-btn-submit" onclick="MissionsFee.submit()" style="background: #ff9800;">Adjust Fee</button>
                    </div>
                </div>
            </div>
            '''

            # Add dates editing modal for admins
            print '''
            <div class="email-modal-overlay" id="datesModal">
                <div class="email-modal" style="max-width: 450px;">
                    <div class="email-modal-header" style="background: #2196f3;">
                        <h4 id="datesModalTitle">Edit Trip Dates</h4>
                        <button class="email-modal-close" onclick="MissionsDates.close()">&times;</button>
                    </div>
                    <div class="email-modal-body">
                        <div class="email-form-group">
                            <label>Main Event Start:</label>
                            <input type="date" id="dateEventStart" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                        </div>
                        <div class="email-form-group">
                            <label>Main Event End:</label>
                            <input type="date" id="dateEventEnd" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                        </div>
                        <div class="email-form-group">
                            <label>Registration Close Date:</label>
                            <input type="date" id="dateClose" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                        </div>
                        <small style="color: #666; display: block; margin-top: 8px;">
                            Leave fields empty to keep existing values.
                        </small>
                    </div>
                    <div class="email-modal-footer">
                        <button class="email-btn email-btn-cancel" onclick="MissionsDates.close()">Cancel</button>
                        <button class="email-btn dates-btn-submit" onclick="MissionsDates.submit()" style="background: #2196f3;">Save Dates</button>
                    </div>
                </div>
            </div>
            '''

        # Add meeting creation modal
        print '''
        <div class="email-modal-overlay" id="meetingModal">
            <div class="email-modal" style="max-width: 400px;">
                <div class="email-modal-header" style="background: #4caf50;">
                    <h4 id="meetingModalTitle">Create Meeting</h4>
                    <button class="email-modal-close" onclick="MissionsMeeting.close()">&times;</button>
                </div>
                <div class="email-modal-body">
                    <div class="email-form-group">
                        <label>Description:</label>
                        <input type="text" id="meetingDescription" placeholder="e.g., Team Meeting, Training Session, Commissioning" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                    </div>
                    <div class="email-form-group">
                        <label>Meeting Date:</label>
                        <input type="date" id="meetingDate" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;" required>
                    </div>
                    <div class="email-form-group">
                        <label>Meeting Time:</label>
                        <input type="time" id="meetingTime" value="18:00" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                    </div>
                    <div class="email-form-group">
                        <label>Location:</label>
                        <input type="text" id="meetingLocation" placeholder="e.g., Room 201, Fellowship Hall" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                    </div>
                    <small style="color: #666; display: block; margin-top: 8px;">
                        This will create a meeting on the selected date and time.
                    </small>
                </div>
                <div class="email-modal-footer">
                    <button class="email-btn email-btn-cancel" onclick="MissionsMeeting.close()">Cancel</button>
                    <button class="email-btn meeting-btn-submit" onclick="MissionsMeeting.submit()" style="background: #4caf50;">Create Meeting</button>
                </div>
            </div>
        </div>
        '''

        # Close main panel and app container
        print '</main>'
        print '</div>'

        # Output sidebar JavaScript
        print get_sidebar_javascript()

        # Output email JavaScript (with admin status for template filtering)
        print get_email_javascript(user_role.get('is_admin', False))

        # Output fee adjustment JavaScript (admin only)
        if user_role.get('is_admin', False):
            print get_fee_adjustment_javascript()

        # Output dates editing JavaScript (admin only)
        if user_role.get('is_admin', False):
            print get_dates_javascript()

        # Output meeting creation JavaScript
        print get_meeting_javascript()

        # Output popup script for existing functionality
        print get_popup_script()

        # Output person details popup JavaScript
        print get_person_details_javascript()

        # Output approval workflow JavaScript (admin only)
        if user_role.get('is_admin', False):
            print get_approval_workflow_javascript()

    except Exception as e:
        # Error handling
        import traceback
        print '''
        <div class="alert alert-danger" style="background: #ffebee; padding: 15px; border-radius: 8px; margin: 20px;">
            <h3>Error Occurred</h3>
            <p>{0}</p>
            <details>
                <summary>Technical Details</summary>
                <pre style="background: #f5f5f5; padding: 10px; overflow-x: auto;">{1}</pre>
            </details>
        </div>
        '''.format(str(e), traceback.format_exc())

# Execute main controller
try:
    main()
except Exception as e:
    # Catch any errors at the top level
    import traceback
    print("""
    <div style="background: #ffcccc; padding: 20px; margin: 20px; border: 2px solid red;">
        <h2>Critical Error</h2>
        <p><strong>Error:</strong> {0}</p>
        <pre style="background: white; padding: 10px; overflow: auto;">
{1}
        </pre>
    </div>
    """.format(str(e), traceback.format_exc()))

# ::END:: Main Controller
