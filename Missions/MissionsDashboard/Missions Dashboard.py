"""
Mission Dashboard 3.6.3 - Complete Payment Tracking
===================================================
Purpose: Comprehensive mission trip management dashboard with real-time insights
Author: Ben Swaby
Email: bswaby@fbchtn.org

Features:
- Upcoming meetings on right sidebar
- Prominent upcoming deadlines at top
- SQL debug output in HTML comments (when enabled)
- Performance optimized queries (now ~1 second vs 47 seconds)
- AJAX loading for detailed data

Performance Improvements:
- Original MissionTripTotals view: 2+ minutes
- Rewritten query to fix scalar function issue in Touchpoint.  New version now runs in 7 seconds (17x improvement!)

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name: Mission_Dashboard_Complete
4. Paste all this code
5. Test and optionally add to menu
--Upload Instructions End--
"""


print("<!-- Debug: Script file is executing -->")

# Check if we're in TouchPoint environment
try:
    if 'model' in globals():
        print("<!-- Debug: 'model' found in globals -->")
    else:
        print("<!-- Debug: 'model' NOT found in globals -->")
        print("<p style='color: red;'>ERROR: TouchPoint 'model' object not found. This script must be run within TouchPoint.</p>")
    
    if 'q' in globals():
        print("<!-- Debug: 'q' (query) found in globals -->")
    else:
        print("<!-- Debug: 'q' (query) NOT found in globals -->")
        print("<p style='color: red;'>ERROR: TouchPoint 'q' object not found. This script must be run within TouchPoint.</p>")
except Exception as e:
    print("<!-- Debug: Error checking globals: {0} -->".format(str(e)))

#####################################################################
# CONFIGURATION SECTION - Customize for your environment
#####################################################################

# ::START:: Configuration
class Config:
    """Centralized configuration for easy customization"""
    
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
    
    # Currency Settings
    CURRENCY_SYMBOL = "$"
    USE_THOUSANDS_SEPARATOR = True
    
    # Security Roles
    ADMIN_ROLES = ["Admin", "Finance", "MissionsDirector"]
    
    # Cache Settings (in seconds)
    CACHE_DURATION = 300  # 5 minutes
    USE_CACHING = True  # Enable caching for performance
    
    # Application-specific organization IDs to exclude (adjust for your setup)
    APPLICATION_ORG_IDS = [2736, 2737, 2738]  # Organizations used for applications
    
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

# ::END:: Configuration

#####################################################################
# LIBRARY IMPORTS AND INITIALIZATION
#####################################################################

# ::START:: Initialization
# Early debug output
print("<!-- Script initialized - imports starting -->")

import datetime
import re

# Early debug to check if we can access TouchPoint globals
try:
    print("<!-- Attempting to access TouchPoint model... -->")
    # Load TouchPoint globals
    model.Header = "Missions Dashboard"
    print("<!-- TouchPoint model accessed successfully -->")
except Exception as e:
    print("<!-- ERROR accessing TouchPoint model: {0} -->".format(str(e)))

# Initialize configuration
config = Config()

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
    closed_filter = ''
    if not include_closed:
        closed_filter = '''
            AND EXISTS (
                SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
                WHERE oe.OrganizationId = o.OrganizationId 
                  AND oe.Field = 'Close'
                  AND oe.DateValue > GETDATE()
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
            IndPaid + SupporterPaid AS TotalPaid
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
            ISNULL(SUM(gsa.Amount), 0) AS TotalPaid
        FROM ActiveTrips at
        LEFT JOIN dbo.GoerSenderAmounts gsa ON gsa.OrgId = at.OrganizationId 
            AND gsa.GoerId IS NULL AND ISNULL(gsa.InActive, 0) = 0
        GROUP BY at.OrganizationId, at.Trip
    ),
    MissionTripTotals AS (
        -- Final result - all people (including those who have paid in full)
        SELECT 
            OrganizationId AS InvolvementId,
            Trip,
            PeopleId,
            Name,
            SortOrder,
            TripCost,
            TotalPaid AS Raised,
            ISNULL(TripCost, 0) - ISNULL(TotalPaid, 0) AS Due
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
            padding: 4px 6px;
            display: block;
            text-align: left;
        }
        
        .mission-table td[data-label]:before {
            content: attr(data-label) ": ";
            font-weight: bold;
            display: inline-block;
            width: 80px;
        }
        
        .mission-table tr {
            border: 1px solid var(--border-color);
            margin-bottom: 5px;
            border-radius: 4px;
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
          AND EXISTS (
              SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
              WHERE oe.OrganizationId = o.OrganizationId 
                AND oe.Field = 'Close'
                AND oe.DateValue > GETDATE()
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
          AND EXISTS (
              SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
              WHERE oe.OrganizationId = o.OrganizationId 
                AND oe.Field = 'Close'
                AND oe.DateValue > GETDATE()
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
            m.Description,
            m.Location,
            m.MeetingDate,
            CASE WHEN m.MeetingDate < GETDATE() THEN 'Past' ELSE 'Upcoming' END AS Status
        FROM Meetings m WITH (NOLOCK)
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
                  AND EXISTS (
                      SELECT 1 FROM OrganizationExtra oe WITH (NOLOCK)
                      WHERE oe.OrganizationId = o.OrganizationId 
                        AND oe.Field = 'Close'
                        AND oe.DateValue > GETDATE()
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
    let queryStartTime = new Date();
    
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
    """Handle AJAX requests for popup data"""
    # Check if this is a POST request with action parameter
    if not hasattr(model.Data, 'action'):
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
    
    return False  # Not an AJAX request

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
            
            print '''
            <div style="background: white; padding: 10px; border-radius: 5px; border-left: 3px solid {5};">
                <div style="font-weight: bold; font-size: 12px;">
                    {6} <a href="/PyScript/Mission_Dashboard?OrgView={0}" style="color: {5};">{1}</a>
                </div>
                <div style="font-size: 11px; color: #666; margin-top: 3px;">
                    {2}: {3}<br>
                    <strong style="color: {5};">{4} days left</strong><br>
                    <strong style="color: #d32f2f;">{7} remaining</strong>
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
                format_currency(deadline.AmountRemaining) if deadline.AmountRemaining > 0 else "Fully paid"
            )
        
        print '''
            </div>
        </div>
        '''

# ::END:: Helper Functions

#####################################################################
# VIEW CONTROLLERS
#####################################################################

# ::START:: Dashboard View
def render_dashboard_view():
    """Main dashboard view with mission overview"""
    
    timer = start_timer()
    
    # Get total statistics with debug
    stats_query = get_total_stats_query()
    stats = execute_query_with_debug(stats_query, "Total Statistics Query", "top1")
    
    # Render KPI cards with all total due values
    print render_kpi_cards(stats.TotalMembers, stats.TotalApplications, stats.TotalOutstandingOpen, 
                          stats.TotalOutstandingClosed, stats.TotalOutstandingAll)
    
    # Add urgent deadlines section
    render_dashboard_deadlines()
    
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
    
    # Recent signups section (more compact)
    print '<h3 style="margin-top: 30px; margin-bottom: 15px;">Last 20 Sign-ups</h3>'
    recent_signups_sql = '''
        SELECT TOP 20
            o.OrganizationId,
            o.OrganizationName, 
            p.Name,
            om.EnrollmentDate
        FROM Organizations o WITH (NOLOCK)
        INNER JOIN OrganizationMembers om WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
        INNER JOIN People p WITH (NOLOCK) ON p.PeopleId = om.PeopleId
        WHERE o.IsMissionTrip = {0}
          AND o.OrganizationStatusId = {1}
          AND om.MemberTypeId <> {2}
        ORDER BY om.EnrollmentDate DESC
    '''.format(config.MISSION_TRIP_FLAG, config.ACTIVE_ORG_STATUS_ID, config.MEMBER_TYPE_LEADER)
    
    recent_signups = execute_query_with_debug(recent_signups_sql, "Recent Signups Query", "sql")
    
    print '<table class="mission-table">'
    print '''
    <thead>
        <tr>
            <th>Name</th>
            <th>Enrollment Date</th>
            <th>Mission</th>
        </tr>
    </thead>
    <tbody>'''
    for signup in recent_signups:
        print '''
        <tr style="font-size: 13px;">
            <td>{0}</td>
            <td>{1}</td>
            <td>{2}</td>
        </tr>
        '''.format(signup.Name, format_date(str(signup.EnrollmentDate)), signup.OrganizationName)
    print '</tbody></table>'
    
    print '</div>'  # End main-content
    
    # Sidebar with Upcoming Meetings
    print '<div class="sidebar">'
    print '<div style="background: white; padding: 20px; border-radius: var(--radius); box-shadow: var(--shadow);">'
    print '<h3 style="margin-top: 0; color: var(--primary-color);">📅 Upcoming Meetings</h3>'
    
    upcoming_meetings_sql = '''
        SELECT TOP 15
            o.OrganizationName,
            m.Description,
            m.Location,
            m.MeetingDate,
            m.OrganizationId
        FROM Meetings m WITH (NOLOCK)
        INNER JOIN Organizations o WITH (NOLOCK) ON m.OrganizationId = o.OrganizationId
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
                <td colspan="5" style="text-align: center; background: {0}; color: {1}; font-weight: bold;">
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
        
        # Organization separator
        if payment.OrganizationName != last_org:
            print '''
            <tr class="org-separator">
                <td colspan="5" style="background: var(--light-bg); font-weight: bold;">
                    {0}
                </td>
            </tr>
            '''.format(payment.OrganizationName)
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
        
        print '''
        <tr>
            <td data-label="Name">
                <a href="?OrgView={0}#person_{1}">{2}</a>
            </td>
            <td data-label="Mission">{3}</td>
            <td data-label="Trip Status">
                <span class="status-badge status-{4}">{5}</span>
            </td>
            <td data-label="Payment Status">
                <span class="status-badge {6}">{7}</span>
            </td>
            <td data-label="Paid">{8}</td>
            <td data-label="Outstanding" style="{9}">
                {10}
            </td>
        </tr>
        '''.format(
            payment.OrganizationId,
            payment.PeopleId,
            payment.Name2,
            payment.OrganizationName,
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
    """Main entry point for the dashboard"""
    
    try:
        # Check if this is an AJAX request
        if handle_ajax_request():
            return
        
        # Get current view
        current_view = str(model.Data.simplenav) if hasattr(model.Data, 'simplenav') else config.DEFAULT_TAB
        org_view = str(model.Data.OrgView) if hasattr(model.Data, 'OrgView') else ""
        
        # Output styles first - this should be the first visible content
        print get_modern_styles()
        
        # Output visualization for developers
        print create_visualization_diagram()
        
        # Check if viewing specific organization
        if org_view:
            render_organization_view(org_view)
            return
        
        # Output navigation
        print get_navigation_html(current_view)
        
        # Route to appropriate view
        if current_view == 'due':
            render_finance_view()
        elif current_view == 'messages':
            render_messages_view()
        elif current_view == 'stats':
            render_stats_view()
        else:  # Default to dashboard
            render_dashboard_view()
            
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
    # Add initial debug output to verify script is running
    print("<!-- Script starting... -->")
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
