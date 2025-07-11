#------------------------------------------------------
#Widget Quick Links
#------------------------------------------------------
#Quick link is a method to provide links to people to gain
#access to various places quickly.  
#Features Include:
# - Permission Based (both at category and icon level)
# - Break off to give counts on top of icon
# - Categories to organize icons
# - Automatic sub-grouping for large categories
# - Intelligent category display (hides empty categories)
#------------------------------------------------------
#Implementation
#------------------------------------------------------
#-- Steps --
#1. Copy this code to a python script (Admin ~ Advanced ~ Special Content ~ Python Scripts) 
#   and call it something like WidgetQuickLinks and make sure to add the word widget to 
#   the content keywords by the script name
#2. Update the QUICK SETUP section below with your links
#3. Add as a home page widget (Admin ~ Advanced ~ Homepage Widget) by selecting the script, 
#   adding name, and setting permissions that can see it

# written by: Ben Swaby
# email: bswaby@fbchtn.org
#------------------------------------------------------
# ⭐ QUICK SETUP - CONFIGURE YOUR LINKS HERE ⭐
#------------------------------------------------------

# STEP 1: Define your categories
# Format: [category_id, icon_class, display_name, is_expanded_by_default, required_roles]
# - category_id: unique identifier for the category
# - icon_class: Font Awesome icon (https://fontawesome.com/v4/icons/)
# - display_name: What users will see
# - is_expanded_by_default: True = open, False = collapsed
# - required_roles: None = everyone, or ["Role1", "Role2"] = only users with these roles
categories = [
    ["care", "fa-heart", "Pastoral Care", True, ["Edit"]],
    ["general", "fa-church", "General", True, ["Edit"]],
    ["assigned", "fa fa-user-check", "Assigned", True, None],  # Visible to everyone
    ["reports", "fa-chart-bar", "Reports", True, ["Edit"]],
    ["finance", "fa-university", "Finance", False, ["Finance","ManageTransactions"]],
    ["dbtools", "fa-database", "DB Tools", False, ["Admin"]],
    ["vbs", "fa-child", "VBS", False, ["Edit"]],
    ["tools", "fa-tools", "Tools", False, ["Edit"]]
]

# STEP 2: Define your links/icons
# Format: [category, icon, label, link, org_id, query, custom_query, required_roles, priority, subgroup]
# - category: must match a category_id from above
# - icon: Font Awesome icon class
# - label: Text shown under icon
# - link: URL to navigate to
# - org_id: Organization ID for count query (optional)
# - query: SQL query for count badge (optional) 
# - custom_query: Custom SQL for count badge (optional)
# - required_roles: None = everyone, or ["Role1", "Role2"]
# - priority: 1-100, higher = appears first (optional, default 50)
# - subgroup: Subgroup name (optional) - only used if category has >8 visible icons
#
# SORTING TIP: Icons are sorted by priority (high to low), then alphabetically by label
# SUBGROUP TIP: Add a subgroup name to organize large categories (e.g., "Reports", "Processing")
icons = [
    # General category - no subgroups needed (less than 8 icons)
    ["general", "fa-hospital", "Hospital", "/PyScript/HospitalReport", 819, "member_count", None, ["Admin","PastoralCare"], 90, None],
    ["general", "fa-hands-praying", "Pastoral Care", "/PyScript/PastoralCare", 890, "member_count", None, ["Admin","PastoralCare"], 85, None],
    ["general", "fa-globe", "Missions", "/PyScript/Mission_Dashboard", None, None, "mission_count", ["Admin","ManageGroups"], 80, None],
    ["general", "fa-briefcase", "Program Manager", "/PyScript/PM_Start", None, None, None, ["Admin","ProgramManager"], 70, None],
    ["general", "fa-envelope", "Emails", "Person2/0#tab-receivedemails", None, None, None, None, 60, None],
    ["general", "fa fa-check-square", "Tasks", "TaskNoteIndex?v=Action", None, None, "taskCount", None, 50, None],
    ["general", "fa fa-solid fa-person-chalkboard", "Classroom Dashboard", "/PyScript/ClassroomDashboard", None, None, None, None, 40, None],
    
    # Reports category - no subgroups needed (less than 8 icons)
    ["reports", "fa fa-bullseye", "Communication Stats", "/PyScript/TPxi_CommunicationsDashboard", None, None, None, ["Edit","Admin"], 95, None],
    ["reports", "fa-check-circle", "Compliance", "/PyScript/Compliance_Dashboard", None, None, None, ["Admin", "FinanceAdmin", "ManageApplication", "BackgroundCheck", "BackgroundCheckRun", "ViewApplication"], 90, None],
    ["reports", "fa-gauge", "Connect Group", "/pyScript/DashboardAttendance", None, None, None, ["Admin","ManageGroups"], 85, None],
    ["reports", "fa fa-church", "Membership Analysis", "/PyScript/TPxi_MembershipAnalysisReport", None, None, None, ["Edit"], 80, None],
    ["reports", "fa-user-plus", "New Member", "/PyScript/TPxi_NewMemberReport", None, None, None, ["Edit"], 75, None],
    ["reports", "fa-chart-line", "Weekly Attendance", "/PyScript/TPxi_WeeklyAttendance", None, None, None, ["Admin","Beta"], 70, None],
    
    # Finance category - with subgroups (13 icons)
    #sub-reports
    ["finance", "fa-table", "Contribution Age Bins", "/PyScript/Contribution%20Age%20Bins", None, None, None, ["Finance"], 95, "Reports"],
    ["finance", "fa-book", "Contribution Worksheet", "/PyScript/Finance_ContributionTotals", None, None, None, ["Finance"], 90, "Reports"],
    ["finance", "fa fa-chart-line", "Financial Dashboard", "/PyScript/financeDashboard2", None, None, None, ["Finance"], 80, "Reports"],
    #sub-givers
    ["finance", "fa-hand-holding-usd", "Givers By Inv.", "/PyScriptForm/GiversByGroup", None, None, None, ["FinanceAdmin"], 70, "Givers"],
    ["finance", "fa-paper-plane", "Grant/QCD Letter", "/PyScript/Finance_QCDGrantLetters", None, None, None, ["Finance"], 65, "Givers"],
    ["finance", "fa-hourglass-half", "Lapsed Givers", "/PyScript/Finance_StandardDeviation/", None, None, "LapsedGivers", ["Finance"], 55, "Givers"],
    ["finance", "fa-user-slash", "Non-Attending Givers", "/PyScript/Finance_NonAttenders/", None, None, "FinanceNonAttending", ["Finance"], 45, "Givers"],
    #sub-processing
    ["finance", "fa fa-building-columns", "Fortis Fees", "/PyScript/Finance_FortisFees", None, None, None, ["Finance"], 75, "Processing"],
    ["finance", "fa fa-file-invoice-dollar", "Inv. wFees", "/PyScript/TPxi_InvWithFees", None, None, None, ["Finance"], 60, "Processing"],
    ["finance", "fa-hand-holding-usd", "Ministry Deposit Report", "/PyScript/TPxi_MinistryDepositReport", None, None, None, ["Finance","ManageTransactions"], 50, "Processing"],
    ["finance", "fa-wallet", "Payment Manager", "/PyScript/TPxi_PaymentManager", None, None, None, ["Finance","ManageTransactions"], 40, "Processing"],
    #sub-other
    ["finance", "fa fa-envelope", "Envelopes", "/RunScript/Envelope%20Number%20Report2/", None, None, None, ["Finance"], 85, "Other"],
    ["finance", "fa-project-diagram", "Ramp Workflow", "/PyScript/TPxi_FinanceRampApproval", None, None, None, ["FinanceAdmin"], 35, "Other"],
    
    # VBS category - no subgroups needed
    ["vbs", "fa-calendar-check", "VBS Attendance", "/PyScript/VBSDailyAttendance", None, None, None, ["Edit"], 90, None],
    ["vbs", "fa-user-check", "VBS Child Pickup", "/PyScript/VBSChildPickup", None, None, None, ["Edit"], 85, None],
    ["vbs", "fas fa-user-nurse", "VBS Nurse Report", "/PyScript/EmergencyListVBS", None, None, None, ["Edit"], 80, None],
    ["vbs", "fa-hands-praying", "VBS Prayer List", "/PyScript/vbsprayerlist", None, None, None, ["Edit"], 75, None],
    ["vbs", "fa-check-square", "VBS Rollsheet", "/PyScript/VBSChildrenRollsheet", None, None, None, ["Edit"], 70, None],

    # Tools category (13 icons - with subgroups)
    #sub-security
    ["tools", "fa-eye", "Account Activity", "/PyScript/TPxi_UserActivity", None, None, None, ["SuperAdmin"], 95, "Security"],
    ["tools", "fa-user-lock", "Account Security", "/PyScript/TPxi_AccountSecurityMonitor", None, None, None, ["Admin"], 90, "Security"],
    ["tools", "fa fa-lock", "Link Generator", "/PyScript/TPxi_LinkGenerator", None, None, None, ["SuperAdmin"], 55, "Security"],
    #sub-Involvement
    ["tools", "fa-layer-group", "Inv. Activity", "/PyScript/TPxi_InvolvementActivityDashboard", None, None, None, ["Admin"], 70, "Involvement"],
    ["tools", "fa fa-history", "Inv. Owner Search", "/PyScript/TPxi_InvNotificationAuditTool", None, None, None, ["Admin"], 65, "Involvement"],
    ["tools", "fa fa-arrow-right", "Inv. Sync", "/PyScript/TPxi_InvolvementSync", None, None, None, ["Admin"], 60, "Involvement"],
    #sub-data
    ["tools", "fa-download", "Attachment Downloader", "/PyScript/TPxi_AttachmentLinkGenerator", None, None, None, ["Admin","ManageOrgMembers"], 85, "Data"],
    ["tools", "fa-database", "Data Quality", "/PyScript/TPxi_DataQualityDashboard", None, None, None, ["Admin"], 80, "Data"],
    ["tools", "fa-cubes", "Ministry Structure", "/PyScript/TPxi_MinistryStructure", None, None, None, ["Admin","Beta"], 45, "Data"],
    ["tools", "fa-file-alt", "Registration Data Mgr", "/PyScript/TPxi_RegistrationDataManager", None, None, None, ["Edit"], 40, "Data"],
    ["tools", "fa fa-tasks", "TaskNote Activity", "/PyScript/TPxi_TaskNoteActivityDashboard", None, None, None, ["Admin","Beta"], 35, "Data"],
    #sub-events
    ["tools", "fa-sign-in-alt", "Event CheckIn", "/PyScript/FastLaneCheckIn", None, None, None, ["Admin"], 75, "Events"],
    ["tools", "fa-calendar-check", "Meeting Reminder", "/PyScript/TPxi_MeetingReminder", None, None, None, ["Admin","Beta"], 50, "Events"],

]

# STEP 3: Configure subgroup display (optional)
# Define how subgroups appear when categories have >8 icons
# The subgroup names in your icons (step 2) map to these definitions
SUBGROUP_INFO = {
    # Finance subgroups
    "Reports": {"name": "Financial Reports", "icon": "fa-chart-line"},
    "Givers": {"name": "Giver Management", "icon": "fa-hand-holding-usd"},
    "Processing": {"name": "Payment Processing", "icon": "fa-credit-card"},
    "Other": {"name": "Other Tools", "icon": "fa-coins"},
    
    # Tools subgroups
    "Security": {"name": "Security & Access", "icon": "fa-shield-alt"},
    "Involvement": {"name": "Involvement Tools", "icon": "fa-users-cog"},
    "Data": {"name": "Data Management", "icon": "fa-database"},
    "Events": {"name": "Event Tools", "icon": "fa-calendar-check"}
}

# STEP 4: Configure when subgroups appear (optional)
# Change this number to control when categories split into subgroups
AUTO_SUBGROUP_THRESHOLD = 8  # Show subgroups when more than 8 icons are visible

# STEP 5: Configure count queries (optional)
# These are referenced by name in the icons above
count_queries = {
    "member_count": "select count(*) From OrganizationMembers om Where om.OrganizationId = {0}",
    "mission_count": """
        SELECT SUM(o.MemberCount) AS TotalOnMission
        FROM Organizations o  
        INNER JOIN OrganizationExtra oe
            ON oe.OrganizationId = o.OrganizationId 
            AND oe.Field = 'Close'
            AND oe.DateValue > GETDATE()
        WHERE o.IsMissionTrip = 1  
          AND o.OrganizationStatusId = 30
          AND o.OrganizationId NOT IN (2736,2737,2738)
    """,
    "taskCount": """
        Select Count(*) AS [PastDueTasks]
        From TaskNote tn
        LEFT JOIN People paid ON paid.PeopleId = tn.AssigneeId
        LEFT JOIN People apid ON apid.PeopleId = tn.AboutPersonId
        Where AssigneeId Is NOT NULL
        AND tn.CompletedDate IS NULL
        AND tn.StatusId NOT IN (5,1)
        AND tn.CreatedBy <> 1
        AND tn.AssigneeId = {}
    """.format(model.UserPeopleId),
    "LapsedGivers": """
        -- Your lapsed givers query here
        SELECT COUNT(*) FROM People WHERE PeopleId = 0
    """,
    "FinanceNonAttending": """
        -- Your non-attending givers query here
        SELECT COUNT(*) FROM People WHERE PeopleId = 0
    """
}

# Advanced Configuration (loads from settings above)
class Config:
    # These values are loaded from the configuration steps above
    AUTO_SUBGROUP_THRESHOLD = AUTO_SUBGROUP_THRESHOLD
    SUBGROUP_INFO = SUBGROUP_INFO
    
    # Hide categories with only this many icons or fewer (0 = never hide)
    HIDE_SMALL_CATEGORIES = 0

#------------------------------------------------------
# END OF QUICK SETUP - Code below handles the display
#------------------------------------------------------


#------------------------------------------------------
#config .. nothing should need changed below this.
#------------------------------------------------------

# Function to check if user has required roles
def user_has_required_roles(required_roles):
    # If no roles are required, return True (everyone has access)
    if not required_roles:
        return True
    
    # Check if user has any of the required roles
    for role in required_roles:
        if model.UserIsInRole(role):
            return True
    
    # If we get here, user doesn't have any of the required roles
    return False

# Function to get count for an icon
def get_icon_count(query_name, org_id=None):
    """Get count based on query name from count_queries dictionary"""
    if not query_name or query_name not in count_queries:
        return 0
    
    try:
        query = count_queries[query_name]
        if org_id:
            # Format query with org_id if provided
            return q.QuerySqlInt(query.format(org_id))
        else:
            # Use query as-is for custom queries
            return q.QuerySqlInt(query)
    except Exception as e:
        return 0

# Function to generate an app icon with a dynamically-sized badge
def generate_app_icon(icon_class, label, link, org_id=None, query=None, custom_query=None, required_roles=None):
    # Check if user has required roles to see this icon
    if required_roles and not user_has_required_roles(required_roles):
        return ""
    
    count = 0
    
    # Handle new query name format
    if custom_query and isinstance(custom_query, str):
        # It's a query name from count_queries
        count = get_icon_count(custom_query, org_id)
    elif query and isinstance(query, str):
        # It's a query name from count_queries
        count = get_icon_count(query, org_id)
    elif custom_query:
        # Legacy format - direct SQL
        try:
            count = q.QuerySqlInt(custom_query)
        except:
            count = 0
    elif query and org_id:
        # Legacy format - direct SQL with org_id
        try:
            count = q.QuerySqlInt(query.format(org_id))
        except:
            count = 0
    
    # Determine badge class based on digit count
    badge_class = "notification-badge"
    if count >= 100:
        badge_class += " notification-badge-large"
    elif count >= 10:
        badge_class += " notification-badge-medium"
    
    # Start building the HTML for this icon
    html = '''
    <div class="app-item">
        <a href="{link}">
    '''.format(link=link)
    
    # Add badge if count > 0
    if count > 0:
        html += '''
            <div class="{badge_class}">{count}</div>
        '''.format(badge_class=badge_class, count=count)
    
    # Complete the icon HTML
    html += '''
            <div class="app-icon-container">
                <i class="fa {icon_class} app-icon"></i>
            </div>
            <div class="app-label-container">
                <span class="app-label">{label}</span>
            </div>
        </a>
    </div>
    '''.format(icon_class=icon_class, label=label)
    
    return html

# Function to sort icons by priority and label
def sort_icons(icon_list):
    """Sort icons by priority (descending) then by label (ascending)"""
    def get_priority(icon):
        # Get priority from icon array (index 8), default to 50 if not specified
        return icon[8] if len(icon) > 8 else 50
    
    def get_label(icon):
        # Get label from icon array (index 2)
        return icon[2].lower()
    
    # Sort by priority (descending) then label (ascending)
    return sorted(icon_list, key=lambda x: (-get_priority(x), get_label(x)))

# CSS for the dashboard with responsive badge sizes and category styles
dashboard_css = '''
<style>
#divCustom {
    text-align: center;
}

.app-menu {
    display: flex;
    flex-direction: column;
    gap: 2px; /* Further reduced from 3px */
    width: 100%;
    max-width: 480px;
    margin: 0 auto;
}

.category-container {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    overflow: hidden;
    margin-bottom: 2px; /* Reduced from 3px */
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.category-header {
    padding: 5px 10px; /* Reduced top/bottom padding */
    background-color: #f5f5f5;
    display: flex;
    align-items: center;
    cursor: pointer;
    border-bottom: 1px solid #e0e0e0;
}

.category-header:hover {
    background-color: #efefef;
}

.category-icon {
    margin-right: 10px;
    color: #005587;
}

.category-title {
    font-weight: bold;
    flex-grow: 1;
}

.category-toggle {
    transition: transform 0.3s;
    transform: rotate(0deg); /* Start pointing down */
}

.category-toggle.collapsed {
    transform: rotate(-90deg); /* Point to the right when collapsed */
}

.category-content {
    display: flex;
    flex-wrap: wrap;
    gap: 0; /* Removed gap completely */
    padding: 1px 2px 2px 2px; /* Small padding all around */
    transition: max-height 0.3s ease-out;
    overflow: hidden;
    margin: 0; /* Ensure no margins */
    align-items: flex-start; /* Align items to top so they don't stretch */
}

.category-content.collapsed {
    max-height: 0 !important;
    padding: 0;
    border: none; /* Remove border when collapsed */
}

.app-item {
    min-width: 80px;
    min-height: 45px; /* Minimum height for single-line labels */
    height: auto; /* Allow height to grow with content */
    width: calc(25% - 0px); /* No gap adjustment needed */
    margin: 0;
    padding: 2px 0; /* Small vertical padding */
    box-sizing: border-box;
    position: relative;
    display: flex;
    flex-direction: column;
}

.app-icon-container {
    display: flex;
    justify-content: center;
    align-items: center;
    width: 100%;
    margin: 0;
    padding: 0;
    height: 30px; /* Fixed height for icon container */
}

.app-label-container {
    display: flex;
    justify-content: center;
    align-items: flex-start; /* Align to top */
    width: 100%;
    box-sizing: border-box;
    min-height: 15px; /* Minimum height for single line */
    height: auto; /* Grow with content */
    margin: 0;
    padding: 0 4px 2px 4px; /* Add horizontal padding and small bottom padding */
    overflow: visible; /* Allow content to determine height */
}

.app-label {
    font: 11px SegoeUI-Regular-final, Segoe UI, "Segoe UI Web (West European)", Segoe, -apple-system, BlinkMacSystemFont, Roboto, Helvetica Neue, Tahoma, Helvetica, Arial, sans-serif;
    margin: 0;
    padding: 0;
    overflow: hidden;
    line-height: 13px; /* Line height for better readability */
    transition: color 83ms linear;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    text-overflow: ellipsis;
    text-align: center;
    max-height: 26px; /* Maximum 2 lines */
    width: 100%; /* Ensure label uses full width */
    word-break: normal; /* Prevent words from breaking */
    word-wrap: break-word; /* Allow words to wrap */
}

/* Base notification badge */
.notification-badge {
    position: absolute;
    top: -2px; /* Adjusted to sit slightly higher */
    right: 15px;
    background-color: #e53935;
    color: white;
    border-radius: 50%;
    width: 20px;
    height: 20px;
    display: flex;
    justify-content: center;
    align-items: center;
    font-size: 12px;
    font-weight: bold;
    border: 2px solid white;
}

/* Medium-sized badge for 2 digits */
.notification-badge-medium {
    width: 24px;
    height: 24px;
    font-size: 11px;
    right: 13px;
}

/* Large-sized badge for 3+ digits */
.notification-badge-large {
    width: 28px;
    height: 28px;
    font-size: 10px;
    border-radius: 14px; /* Make it more oval for 3 digits */
    right: 11px;
}

@media (max-width: 400px) {
    .app-item {
        width: calc(33.33% - 0px);
    }
}

@media (max-width: 300px) {
    .app-item {
        width: calc(50% - 0px);
    }
}

.app-icon {
    color: #005587;
    font-size: 24px; /* Reduced from 28px */
    margin: 0; /* No margins */
}

#divformat {
  background-color: White;
  border: 1px solid green;
  padding: 2px;
  margin: 0;
  margin-top: 2px;
}

.app-item a {
    text-decoration: none;
    color: #005587;
    display: flex;
    flex-direction: column;
    height: 100%;
    width: 100%;
    margin: 0;
    padding: 0;
}

/* Subgroup styles */
.subgroup-section {
    width: 100%;
    margin-bottom: 6px;  /* Reduced from 10px */
}

.subgroup-header {
    font-size: 13px;
    font-weight: bold;
    color: #005587;
    margin: 3px 0 3px 5px;  /* Reduced top margin from 8px to 3px */
    padding-left: 5px;
    border-left: 3px solid #005587;
}

.subgroup-header i {
    margin-right: 5px;
    font-size: 12px;
}

.subgroup-content {
    display: flex;
    flex-wrap: wrap;
    gap: 0;
    padding: 0;
    align-items: flex-start; /* Align items to top */
}

/* Make subgroup last item not have bottom margin */
.subgroup-section:last-child {
    margin-bottom: 0;
}

/* First subgroup has minimal top spacing */
.subgroup-section:first-child .subgroup-header {
    margin-top: 1px;  /* Even less space for the first subgroup */
}
</style>
'''

# JavaScript for category toggling
category_js = '''
<script>
function toggleCategory(categoryId) {
    const content = document.getElementById('category-content-' + categoryId);
    const toggle = document.getElementById('category-toggle-' + categoryId);
    const container = document.getElementById('category-container-' + categoryId);
    
    if (content.classList.contains('collapsed')) {
        content.classList.remove('collapsed');
        toggle.className = 'fa fa-chevron-down category-toggle'; // Point down when expanded
        container.style.marginBottom = '2px';
        localStorage.setItem('category-' + categoryId, 'expanded');
    } else {
        content.classList.add('collapsed');
        toggle.className = 'fa fa-chevron-right category-toggle'; // Point right when collapsed
        container.style.marginBottom = '1px';
        localStorage.setItem('category-' + categoryId, 'collapsed');
    }
}

document.addEventListener('DOMContentLoaded', function() {
    const categories = document.querySelectorAll('[id^="category-content-"]');
    categories.forEach(function(category) {
        const categoryId = category.id.replace('category-content-', '');
        const toggle = document.getElementById('category-toggle-' + categoryId);
        const container = document.getElementById('category-container-' + categoryId);
        
        // Check localStorage for saved state
        const savedState = localStorage.getItem('category-' + categoryId);
        
        if (savedState === 'collapsed' && !category.classList.contains('collapsed')) {
            // Should be collapsed but isn't
            category.classList.add('collapsed');
            toggle.className = 'fa fa-chevron-right category-toggle';
            container.style.marginBottom = '1px';
        } else if (savedState === 'expanded' && category.classList.contains('collapsed')) {
            // Should be expanded but isn't
            category.classList.remove('collapsed');
            toggle.className = 'fa fa-chevron-down category-toggle';
            container.style.marginBottom = '2px';
        }
    });
});
</script>
'''

# Helper function to print regular category
def print_regular_category(category_id, category_icon, category_name, is_expanded, visible_icons):
    """Print a regular category without subgroups"""
    arrow_icon = "fa-chevron-down" if is_expanded else "fa-chevron-right"
    
    print '''
    <div class="category-container" id="category-container-{0}">
        <div class="category-header" onclick="toggleCategory('{0}')">
            <i class="fa {1} category-icon"></i>
            <span class="category-title">{2}</span>
            <i class="fa {3} category-toggle" id="category-toggle-{0}"></i>
        </div>
        <div class="category-content{4}" id="category-content-{0}" data-default="{5}">
    '''.format(
        category_id,
        category_icon,
        category_name,
        arrow_icon,
        "" if is_expanded else " collapsed",
        "expanded" if is_expanded else "collapsed"
    )
    
    # Generate icons for this category
    for icon in visible_icons:
        # Unpack icon data (handle variable length)
        if len(icon) >= 10:
            _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _ = icon
        elif len(icon) >= 9:
            _, icon_class, label, link, org_id, query, custom_query, required_roles, _ = icon
        else:
            _, icon_class, label, link, org_id, query, custom_query, required_roles = icon
        
        print generate_app_icon(icon_class, label, link, org_id, query, custom_query, required_roles)
    
    print '''
        </div>
    </div>
    '''

# Helper function to print subgrouped category
def print_subgrouped_category(category_id, category_icon, category_name, is_expanded, visible_icons):
    """Print a category with subgroups based on icon subgroup field"""
    # Group icons by their subgroup field
    subgroups = {}
    
    for icon in visible_icons:
        # Get subgroup from icon (index 9 if it exists)
        subgroup_id = icon[9] if len(icon) > 9 else None
        
        if subgroup_id:
            if subgroup_id not in subgroups:
                subgroups[subgroup_id] = []
            subgroups[subgroup_id].append(icon)
        else:
            # Icons without subgroup go to "Other"
            if "Other" not in subgroups:
                subgroups["Other"] = []
            subgroups["Other"].append(icon)
    
    # Print the main category
    arrow_icon = "fa-chevron-down" if is_expanded else "fa-chevron-right"
    
    print '''
    <div class="category-container" id="category-container-{0}">
        <div class="category-header" onclick="toggleCategory('{0}')">
            <i class="fa {1} category-icon"></i>
            <span class="category-title">{2}</span>
            <i class="fa {3} category-toggle" id="category-toggle-{0}"></i>
        </div>
        <div class="category-content{4}" id="category-content-{0}" data-default="{5}">
    '''.format(
        category_id,
        category_icon,
        category_name,
        arrow_icon,
        "" if is_expanded else " collapsed",
        "expanded" if is_expanded else "collapsed"
    )
    
    # Sort subgroups to show in a logical order
    subgroup_order = ["Reports", "Givers", "Processing", "Security", "Involvement", "Data", "Events", "Other"]
    sorted_subgroups = []
    
    # Add subgroups in preferred order if they exist
    for sg in subgroup_order:
        if sg in subgroups:
            sorted_subgroups.append(sg)
    
    # Add any remaining subgroups not in the order list
    for sg in subgroups:
        if sg not in sorted_subgroups:
            sorted_subgroups.append(sg)
    
    # Print each subgroup
    for subgroup_id in sorted_subgroups:
        icons_list = subgroups[subgroup_id]
        
        # Skip empty subgroups
        if not icons_list:
            continue
        
        # Get subgroup info
        subgroup_info = Config.SUBGROUP_INFO.get(subgroup_id, 
                                                 {"name": subgroup_id, "icon": "fa-folder"})
        
        # Print subgroup header
        print '''
        <div class="subgroup-section">
            <h4 class="subgroup-header">
                <i class="fa {0}"></i> {1}
            </h4>
            <div class="subgroup-content">
        '''.format(subgroup_info["icon"], subgroup_info["name"])
        
        # Print icons in this subgroup
        for icon in icons_list:
            # Unpack icon data
            if len(icon) >= 10:
                _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _ = icon
            elif len(icon) >= 9:
                _, icon_class, label, link, org_id, query, custom_query, required_roles, _ = icon
            else:
                _, icon_class, label, link, org_id, query, custom_query, required_roles = icon
            
            print generate_app_icon(icon_class, label, link, org_id, query, custom_query, required_roles)
        
        print '''
            </div>
        </div>
        '''
    
    print '''
        </div>
    </div>
    '''

# Generate the HTML output
print dashboard_css
print '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">'
print '<div id="appsModule" class="app-menu" role="list" aria-label="Apps">'

# Try to generate categories and icons, handling any errors
try:
    # Organize icons by category
    category_icons = {}
    for icon in icons:
        category = icon[0]
        if category not in category_icons:
            category_icons[category] = []
        category_icons[category].append(icon)
    
    # Sort icons within each category
    for category_id in category_icons:
        category_icons[category_id] = sort_icons(category_icons[category_id])
    
    # Generate HTML for each category
    for i, category in enumerate(categories):
        category_id = category[0]
        category_icon = category[1]
        category_name = category[2]
        is_expanded = category[3]
        required_roles = category[4] if len(category) > 4 else None
        
        # Skip this category if user doesn't have required roles
        if not user_has_required_roles(required_roles):
            continue
        
        # Skip categories with no visible icons
        if category_id not in category_icons or not category_icons[category_id]:
            continue
        
        # Check if there are any visible icons for this category (after role filtering)
        visible_icons = []
        for icon in category_icons[category_id]:
            # Get icon roles (handle both old 8-item and new 9-item format)
            icon_roles = None
            if len(icon) >= 8:
                icon_roles = icon[7]  # Required roles is at index 7
            
            if user_has_required_roles(icon_roles):
                visible_icons.append(icon)
        
        # Skip category if there are no visible icons after role filtering
        if not visible_icons:
            continue
        
        # Skip small categories if configured
        if Config.HIDE_SMALL_CATEGORIES > 0 and len(visible_icons) <= Config.HIDE_SMALL_CATEGORIES:
            continue
        
        # Check if we should auto-split this category
        # Split if we have more icons than threshold AND at least one icon has a subgroup
        has_subgroups = any(len(icon) > 9 and icon[9] for icon in visible_icons)
        should_split = len(visible_icons) > Config.AUTO_SUBGROUP_THRESHOLD and has_subgroups
        
        if should_split:
            # Create sub-grouped display
            print_subgrouped_category(category_id, category_icon, category_name, 
                                    is_expanded, visible_icons)
        else:
            # Regular category display
            print_regular_category(category_id, category_icon, category_name, 
                                 is_expanded, visible_icons)
    
    # Print the JavaScript for category toggling
    print category_js
    
except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"

print '</div>'
