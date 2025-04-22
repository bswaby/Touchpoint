#------------------------------------------------------
#Widget Quick Links
#------------------------------------------------------
#Quick link is a method to provide links to people to gain
#access to various places quickly.  
#Features Include:
# - Permission Based (both at category and icon level)
# - Break off to give counts on top of icon
# - Categories to organize icons
#------------------------------------------------------
#Implementation
#------------------------------------------------------
#-- Steps --
#1. Copy this code to a python script (Admin ~ Advanced ~ Special Content ~ Python Scripts) 
#   and call it something like WidgetQuickLinks and make sure to add the word widget to 
#   the content keywords by the script name
#2. Update config parameters below
#3. Add as a home page widget (Admin ~ Advanced ~ Homepage Widget) by selecting the script, 
#   adding name, and setting permissions that can see it


#------------------------------------------------------
#config parameters
#------------------------------------------------------

#### The SQL scripts are only needed if you are including counts on top of icons

# Define the query template for counts
member_count_query = "select count(*) From OrganizationMembers om Where om.OrganizationId = {0}"

# Define the custom mission count query
mission_count_query = """
SELECT SUM(o.MemberCount) AS TotalOnMission
FROM Organizations o  
INNER JOIN OrganizationExtra oe
    ON oe.OrganizationId = o.OrganizationId 
    AND oe.Field = 'Close'
    AND oe.DateValue > GETDATE()
WHERE o.IsMissionTrip = 1  
  AND o.OrganizationStatusId = 30
  AND o.OrganizationId NOT IN (2736,2737,2738)
"""

taskCount = """
Select Count(*) AS [PastDueTasks]
From TaskNote tn
LEFT JOIN People paid ON paid.PeopleId = tn.AssigneeId
LEFT JOIN People apid ON apid.PeopleId = tn.AboutPersonId
Where AssigneeId Is NOT NULL
AND tn.CompletedDate IS NULL
AND tn.StatusId NOT IN (5,1)
AND tn.CreatedBy <> 1
AND tn.AssigneeId = {}
""".format(model.UserPeopleId)


#### Here is where you define categories and their icons
# Format: [category_name, icon_class, display_name, is_expanded_by_default, required_roles]
# Required_roles can be None (accessible to all) or a list of roles
categories = [
    ["care", "fa-heart", "Pastoral Care", True, ["Edit"]],
    ["general", "fa-church", "General", True, ["Edit"]],
    ["assigned", "fa fa-user-check", "Assigned", True, None],  # Visible to everyone
    ["reports", "fa-chart-bar", "Reports", False, ["Edit"]],
    ["tools", "fa-tools", "Tools", False, ["Edit"]]
]

#### Here is where you define the icon parameters
#### icons can be found here https://fontawesome.com/v4/icons/
#
# Format: [category, icon_class, label, link_url, org_id, query, custom_query, required_roles]
# The category must match one of the category names defined above
icons = [
    # General category
    ["general", "fa-hospital", "Hospital", "/PyScript/HospitalReport", 819, member_count_query, None, ["Admin","PastoralCare"]],
    ["general", "fa-hands-praying", "Pastoral Care", "/PyScript/PastoralCare", 890, member_count_query, None, ["Admin","PastoralCare"]],
    ["general", "fa-globe", "Missions", "/PyScript/Mission_Dashboard", None, None, mission_count_query, ["Admin","ManageGroups"]],
    ["general", "fa-briefcase", "Program Manager", "/PyScript/PM_Start", None, None, None, ["Admin","ProgramManager"]],
    ["general", "fa-envelope", "Emails", "Person2/0#tab-receivedemails", None, None, None, None],
    ["general", "fa fa-check-square", "Tasks", "TaskNoteIndex?v=Action", None, None, taskCount, None],
    ["general", "fa fa-solid fa-person-chalkboard", "Classroom Dashboard", "/PyScript/ClassroomDashboard", None, None, None, None],
    
    # Reports category
    ["reports", "fa-check-circle", "Compliance", "/PyScript/Compliance_Dashboard", None, None, None, ["Admin", "FinanceAdmin", "ManageApplication", "BackgroundCheck", "BackgroundCheckRun", "ViewApplication"]],
    ["reports", "fa-gauge", "Connect Group", "/pyScript/DashboardAttendance", None, None, None, ["Admin","ManageGroups"]],
    ["reports", "fa-database", "Data Quality", "/PyScript/TPxi_DataQualityDashboard", None, None, None, ["Admin"]],
    ["reports", "fa-layer-group", "Inv. Activity", "/PyScript/TPxi_InvolvementActivityDashboard", None, None, None, ["Beta"]],
    ["reports", "fa fa-tasks", "TaskNote Activity", "/PyScript/TPxi_TaskNoteActivityDashboard", None, None, None, ["Beta"]],
    ["reports", "fa-chart-line", "Weekly Attendance", "/PyScript/TPxi_WeeklyAttendance", None, None, None, ["Beta"]],
    
    # Tools category
    ["tools", "fa-cubes", "Ministry Structure", "/PyScript/TPxi_MinistryStructure", None, None, None, ["Beta"]],
    ["tools", "fa-download", "Attachment Downloader", "/PyScript/TPxi_AttachmentLinkGenerator", None, None, None, ["Admin","ManageOrgMembers"]],
    ["tools", "fa-calendar-check", "Meeting Reminder", "/PyScript/TPxi_MeetingReminder", None, None, None, ["Beta"]]
]


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

# Function to generate an app icon with a dynamically-sized badge
def generate_app_icon(icon_class, label, link, org_id=None, query=None, custom_query=None, required_roles=None):
    # Check if user has required roles to see this icon
    if required_roles and not user_has_required_roles(required_roles):
        return ""
    
    count = 0
    
    # Get count based on the query type
    try:
        if custom_query:
            # Use custom query directly
            count = q.QuerySqlInt(custom_query)
        elif query and org_id:
            # Use regular query with org_id
            count = q.QuerySqlInt(query.format(org_id))
    except Exception as e:
        # If there's an error with the query, log it but continue
        count = 0
        # You could log the error here, but we'll just suppress it for now
    
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
}

.category-toggle.collapsed {
    transform: rotate(-90deg);
}

.category-content {
    display: flex;
    flex-wrap: wrap;
    gap: 0; /* Removed gap completely */
    padding: 1px 2px 0 2px; /* Adjusted padding - reduced bottom padding */
    transition: max-height 0.3s ease-out;
    overflow: hidden;
    margin: 0; /* Ensure no margins */
}

.category-content.collapsed {
    max-height: 0 !important;
    padding: 0;
    border: none; /* Remove border when collapsed */
}

.app-item {
    min-width: 80px;
    height: 58px; /* Reduced height to eliminate extra space */
    width: calc(25% - 0px); /* No gap adjustment needed */
    margin: 0;
    padding: 0;
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
    height: 28px; /* Height for label container */
    margin: 0;
    padding: 0 4px; /* Add horizontal padding */
    overflow: hidden; /* Keep overflow hidden */
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
    max-height: 26px; /* Slightly reduced from 28px */
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
        toggle.classList.remove('collapsed');
        container.style.marginBottom = '2px'; // Reduced margin when expanded
        // Store expanded state in localStorage
        localStorage.setItem('category-' + categoryId, 'expanded');
    } else {
        content.classList.add('collapsed');
        toggle.classList.add('collapsed');
        container.style.marginBottom = '1px'; // Reduce margin when collapsed
        // Store collapsed state in localStorage
        localStorage.setItem('category-' + categoryId, 'collapsed');
    }
}

// Initialize categories based on stored state or default
document.addEventListener('DOMContentLoaded', function() {
    const categories = document.querySelectorAll('[id^="category-content-"]');
    categories.forEach(function(category) {
        const categoryId = category.id.replace('category-content-', '');
        const toggle = document.getElementById('category-toggle-' + categoryId);
        const container = document.getElementById('category-container-' + categoryId);
        
        // Check localStorage for saved state, use default if not present
        const savedState = localStorage.getItem('category-' + categoryId);
        const defaultState = category.getAttribute('data-default');
        
        if ((savedState === 'collapsed') || (savedState === null && defaultState === 'collapsed')) {
            category.classList.add('collapsed');
            toggle.classList.add('collapsed');
            container.style.marginBottom = '1px'; // Reduce margin when collapsed
        }
    });
});
</script>
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
            if len(icon) >= 8:
                _, _, _, _, _, _, _, icon_roles = icon
                if user_has_required_roles(icon_roles):
                    visible_icons.append(icon)
        
        # Skip category if there are no visible icons after role filtering
        if not visible_icons:
            continue
        
        print '''
        <div class="category-container" id="category-container-{0}">
            <div class="category-header" onclick="toggleCategory('{0}')">
                <i class="fa {1} category-icon"></i>
                <span class="category-title">{2}</span>
                <i class="fa fa-chevron-down category-toggle" id="category-toggle-{0}"></i>
            </div>
            <div class="category-content{3}" id="category-content-{0}" data-default="{4}">
        '''.format(
            category_id,
            category_icon,
            category_name,
            "" if is_expanded else " collapsed",
            "expanded" if is_expanded else "collapsed"
        )
        
        # Generate icons for this category (only for icons that passed role check)
        for icon in visible_icons:
            _, icon_class, label, link, org_id, query, custom_query, required_roles = icon
            print generate_app_icon(icon_class, label, link, org_id, query, custom_query, required_roles)
        
        print '''
            </div>
        </div>
        '''
    
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
