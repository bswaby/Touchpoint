# ---------------------------------------------------------------
# TPxi QuickLinks Widget (Part 1 of 2)
# ---------------------------------------------------------------
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
#
# This is a 2-part system:
#   Part 1: TPxi_QuickLinks (this script) - the homepage widget
#   Part 2: TPxi_QuickLinksAdmin - the visual admin UI
#
# Both share config stored in Special Content as "QuickLinksConfig".
# The admin writes it; this widget reads it.
#
# Features:
#   - Permission-based visibility (category and icon level)
#   - Count badges via SQL queries (tasks, members, etc.)
#   - Categories with expand/collapse
#   - Automatic subgrouping for large categories
#   - Popup subgroups (clickable tiles that open a menu)
#   - Nested subgroups (popups inside inline subgroups)
#   - Expiration dates to auto-hide seasonal icons
#   - Intelligent category display (hides empty categories)
#
# --Upload Instructions--
# 1. Admin ~ Advanced ~ Special Content ~ Python
# 2. New Python Script File named "TPxi_QuickLinks"
#    (add "widget" to content keywords)
# 3. Admin ~ Advanced ~ Homepage Widget - select the script
# 4. Upload TPxi_QuickLinksAdmin for the admin UI (see Part 2)
# ---------------------------------------------------------------
#------------------------------------------------------
# ⭐ QUICK SETUP - CONFIGURE YOUR LINKS HERE ⭐
#------------------------------------------------------
# NOTE: These defaults are overridden if QuickLinksConfig content exists.
# Use /PyScript/TPxi_QuickLinksAdmin to manage config visually.
#------------------------------------------------------

# STEP 1: Define your categories
# Format: [category_id, icon_class, display_name, is_expanded_by_default, required_roles]
# These are starter examples. Use /PyScript/TPxi_QuickLinksAdmin to manage visually.
categories = [
    ["general", "fa-church", "General", True, None],
    ["reports", "fa-chart-bar", "Reports", True, ["Edit"]],
    ["tools", "fa-tools", "Tools", False, ["Admin"]],
]

# STEP 2: Define your links/icons
# Format: [category, icon, label, link, org_id, query, custom_query, required_roles, priority, subgroup, expiration_date]
# - category: must match a category_id from above
# - icon: Font Awesome icon class
# - label: Text shown under icon
# - link: URL to navigate to
# - org_id: Organization ID for simple member count (optional)
# - query: Named query from count_queries for custom count (optional)
# - custom_query: DEPRECATED - use 'query' instead (kept for backward compatibility)
# - required_roles: None = everyone, or ["Role1", "Role2"]
# - priority: 1-100, higher = appears first (optional, default 50)
# - subgroup: Subgroup name (optional) - only used if category has >8 visible icons
# - expiration_date: Date string "YYYY-MM-DD" when icon expires (optional) - icon will not show after this date
#
# COUNT OPTIONS:
# 1. Just org_id: Shows member count for that organization
# 2. Just query: Shows count from named query (no org reference)
# 3. Both org_id + query: Passes org_id to the named query
#
# SORTING TIP: Icons are sorted by priority (high to low), then alphabetically by label
# SUBGROUP TIP: Add a subgroup name to organize large categories (e.g., "Reports", "Processing")
# EXPIRATION TIP: Add expiration date as "YYYY-MM-DD" to auto-hide icons after a certain date
icons = [
    # Starter examples - customize via /PyScriptForm/TPxi_QuickLinksAdmin
    ["general", "fa-envelope", "Emails", "Person2/0#tab-receivedemails", None, None, None, None, 60, None],
    ["general", "fa-check-square", "Tasks", "TaskNoteIndex?v=Action", None, None, None, None, 50, None],
    ["reports", "fa-chart-line", "Weekly Attendance", "/PyScript/TPxi_WeeklyAttendance", None, None, None, ["Admin"], 70, None],
    ["tools", "fa-database", "Data Quality", "/PyScript/TPxi_DataQualityDashboard", None, None, None, ["Admin"], 80, None],
]

# STEP 3: Configure subgroup display (optional)
# Define how subgroups appear when categories have >8 icons
# The subgroup names in your icons (step 2) map to these definitions
# Format: "subgroup_id": {"name": "Display Name", "icon": "fa-icon-class", "popup": True/False}
# - subgroup_id: Must match the subgroup name used in your icons
# - name: Display name shown to users
# - icon: Font Awesome icon class
# - popup: True = clickable popup menu, False = inline display (default)
#
# ORDERING: Subgroups appear in the order listed here
SUBGROUP_INFO = {
    # Starter examples - customize via /PyScriptForm/TPxi_QuickLinksAdmin
    # "Reports": {"name": "Reports", "icon": "fa-chart-line", "popup": False},
    # "Tools": {"name": "Tools", "icon": "fa-wrench", "popup": True},
}

# STEP 4: Configure when subgroups appear (optional)
# Change this number to control when categories split into subgroups
AUTO_SUBGROUP_THRESHOLD = 8  # Show subgroups when more than 8 icons are visible

# STEP 5: Configure count queries (optional)
# These are referenced by name in the icons above
count_queries = {
    # Starter examples - customize via /PyScriptForm/TPxi_QuickLinksAdmin
    # "member_count": "SELECT COUNT(*) FROM OrganizationMembers WHERE OrganizationId = {0}",
    # "taskCount": "SELECT COUNT(*) FROM TaskNote WHERE AssigneeId = {} AND CompletedDate IS NULL".format(model.UserPeopleId),
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
# Override with content storage config if available
# (managed via /PyScriptForm/TPxi_QuickLinksAdmin)
#------------------------------------------------------
import json as _json
try:
    _config_json = model.TextContent("QuickLinksConfig")
    if _config_json and _config_json.strip():
        _config = _json.loads(_config_json)
        if _config.get("version"):
            categories = []
            for _cat in sorted(_config.get("categories", []), key=lambda x: x.get("sort_order", 0)):
                categories.append([
                    _cat["id"], _cat["icon"], _cat["name"],
                    _cat.get("expanded", True),
                    _cat.get("roles") or None
                ])

            icons = []
            for _ic in _config.get("icons", []):
                icons.append([
                    _ic["category_id"], _ic["icon"], _ic["label"], _ic["link"],
                    _ic.get("org_id"), _ic.get("query_name"), None,
                    _ic.get("roles") or None,
                    _ic.get("priority", 50),
                    _ic.get("subgroup"),
                    _ic.get("expiration_date")
                ])

            SUBGROUP_INFO = _config.get("subgroups", {})

            count_queries = {}
            for _qname, _qdata in _config.get("count_queries", {}).items():
                _sql = _qdata["sql"] if isinstance(_qdata, dict) else _qdata
                _uses_user = _qdata.get("uses_current_user", False) if isinstance(_qdata, dict) else False
                if _uses_user and '{}' in _sql:
                    _sql = _sql.format(model.UserPeopleId)
                count_queries[_qname] = _sql

            AUTO_SUBGROUP_THRESHOLD = _config.get("settings", {}).get("auto_subgroup_threshold", 8)

            # Update Config class
            Config.AUTO_SUBGROUP_THRESHOLD = AUTO_SUBGROUP_THRESHOLD
            Config.SUBGROUP_INFO = SUBGROUP_INFO
            Config.HIDE_SMALL_CATEGORIES = _config.get("settings", {}).get("hide_small_categories", 0)
except:
    pass  # Fall back to hardcoded defaults

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

# Function to check if an icon has expired
def is_icon_expired(expiration_date):
    """Check if an icon has passed its expiration date"""
    if not expiration_date:
        return False
    
    try:
        # Parse the expiration date string (format: YYYY-MM-DD)
        exp_parts = expiration_date.split('-')
        if len(exp_parts) != 3:
            return False
        
        exp_year = int(exp_parts[0])
        exp_month = int(exp_parts[1])
        exp_day = int(exp_parts[2])
        
        # Get current date
        current_date = model.DateTime
        
        # Compare dates
        if current_date.Year > exp_year:
            return True
        elif current_date.Year == exp_year:
            if current_date.Month > exp_month:
                return True
            elif current_date.Month == exp_month:
                if current_date.Day > exp_day:
                    return True
        
        return False
    except:
        # If any error occurs parsing the date, don't expire the icon
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
    
    # Handle count queries
    if query:
        # Named query from count_queries
        count = get_icon_count(query, org_id)
    elif custom_query:
        # Legacy support - treat custom_query same as query
        count = get_icon_count(custom_query, org_id)
    elif org_id:
        # Simple member count for organization
        count = get_icon_count("member_count", org_id)
    
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
    background: #fafbfc;
    border-left: 3px solid #e1e4e8;
    padding-left: 8px;
    transition: border-color 0.2s;
}

.subgroup-section:hover {
    border-left-color: #005587;
}

.subgroup-header {
    font-size: 13px;
    font-weight: bold;
    color: #005587;
    margin: 3px 0 3px 0;  /* Removed left margin since section has padding */
    padding-left: 5px;
}

.subgroup-header i {
    margin-right: 5px;
    font-size: 12px;
}

.subgroup-content {
    display: flex;
    flex-wrap: wrap;
    gap: 0;
    padding: 0 8px 4px 0; /* Right and bottom padding to balance the section's left padding */
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

/* Container for popup subgroups - displays them like regular icons */
.popup-subgroups-container {
    display: flex;
    flex-wrap: wrap;
    gap: 0;
    padding: 8px 2px 2px 2px;
    align-items: flex-start;
    margin-top: 8px;
    border-top: 1px solid #e9ecef;
    position: relative;
}

/* Optional: Add a subtle label for popup subgroups */
.popup-subgroups-container::before {
    content: "";
    position: absolute;
    top: -1px;
    left: 0;
    right: 0;
    height: 1px;
    background: linear-gradient(to right, transparent, #e0e0e0 10%, #e0e0e0 90%, transparent);
}

/* Popup subgroup styles - when a subgroup is clickable */
.subgroup-popup-item {
    width: calc(25% - 0px);
    min-width: 80px;
    margin: 0;
    padding: 2px 0;
    box-sizing: border-box;
    position: relative;
    display: flex;
    flex-direction: column;
}

.subgroup-popup-button {
    display: flex;
    flex-direction: column;
    align-items: center;
    text-decoration: none;
    color: #005587;
    width: 100%;
    height: 100%;
    padding: 0;
    background: none;
    border: none;
    cursor: pointer;
    position: relative;
}

.subgroup-popup-button:hover .app-icon-container {
    transform: scale(1.1);
}

.subgroup-dropdown-indicator {
    position: absolute;
    top: -2px;
    right: 12px;
    background: #005587;
    color: white;
    border-radius: 3px;
    width: 16px;
    height: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 10px;
    border: 2px solid white;
}

@media (max-width: 400px) {
    .subgroup-popup-item {
        width: calc(33.33% - 0px);
    }
}

@media (max-width: 300px) {
    .subgroup-popup-item {
        width: calc(50% - 0px);
    }
}

/* Popup overlay and menu styles */
.popup-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: none;
    z-index: 9998;
    opacity: 0;
    transition: opacity 0.3s;
}

.popup-overlay.active {
    display: block;
    opacity: 1;
}

.popup-menu {
    position: fixed;
    background: white;
    border-radius: 12px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
    padding: 20px;
    z-index: 9999;
    max-width: 90%;
    max-height: 85vh;
    width: 400px; /* Default width for desktop */
    display: flex;
    flex-direction: column;
    visibility: hidden;
    opacity: 0;
    transition: opacity 0.3s ease, transform 0.3s ease;
    overflow: hidden; /* Prevent overflow on the container itself */
}

@media (max-width: 480px) {
    .popup-menu {
        width: 90%; /* Full width on mobile */
    }
}

.popup-menu.active {
    visibility: visible;
    opacity: 1;
    /* Transform is handled by JavaScript based on position */
}

.popup-menu-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
    padding-bottom: 10px;
    border-bottom: 2px solid #f0f0f0;
    flex-shrink: 0; /* Prevent header from shrinking */
}

.popup-menu-title {
    font-size: 18px;
    font-weight: bold;
    color: #005587;
    display: flex;
    align-items: center;
}

.popup-menu-title i {
    margin-right: 10px;
}

.popup-close {
    background: none;
    border: none;
    font-size: 24px;
    color: #666;
    cursor: pointer;
    padding: 0;
    width: 32px;
    height: 32px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 50%;
    transition: all 0.2s;
}

.popup-close:hover {
    background: #f0f0f0;
    color: #333;
}

.popup-subgroups {
    display: flex;
    flex-direction: column;
    gap: 15px;
    overflow-y: auto; /* Enable vertical scrolling */
    flex: 1; /* Take remaining space */
    min-height: 0; /* Important for flexbox scrolling */
}

.popup-subgroup {
    background: #f8f8f8;
    border-radius: 8px;
    padding: 12px;
}

.popup-subgroup-header {
    font-size: 14px;
    font-weight: bold;
    color: #005587;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
}

.popup-subgroup-header i {
    margin-right: 8px;
    font-size: 13px;
}

/* Wrapper for scrollable content */
.popup-content-wrapper {
    flex: 1;
    overflow-y: auto;
    min-height: 0; /* Important for flexbox scrolling */
    padding: 10px;
    -webkit-overflow-scrolling: touch; /* Smooth scrolling on iOS */
}

.popup-subgroup-icons {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(80px, 1fr));
    gap: 8px;
}

/* Popup icon styles - slightly different from regular icons */
.popup-icon-item {
    background: white;
    border-radius: 8px;
    padding: 10px 5px;
    text-align: center;
    transition: all 0.2s;
    cursor: pointer;
    text-decoration: none;
    color: #005587;
    display: flex;
    flex-direction: column;
    align-items: center;
    position: relative;
}

.popup-icon-item:hover {
    background: #e8f4f8;
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.popup-icon-item .app-icon {
    font-size: 28px;
    margin-bottom: 5px;
}

.popup-icon-item .app-label {
    font-size: 11px;
    line-height: 13px;
}

/* Mobile adjustments for popup */
@media (max-width: 768px) {
    .popup-menu {
        max-width: 95%;
        max-height: 90vh;
        padding: 15px;
    }
    
    .popup-content-wrapper {
        padding: 8px;
    }
    
    .popup-subgroup-icons {
        grid-template-columns: repeat(auto-fill, minmax(70px, 1fr));
        gap: 6px;
    }
    
    .popup-menu-title {
        font-size: 16px;
    }
    
    .popup-icon-item {
        padding: 8px 3px;
    }
    
    .popup-icon-item .app-icon {
        font-size: 24px;
    }
}

/* Extra small mobile */
@media (max-width: 400px) {
    .popup-subgroup-icons {
        grid-template-columns: repeat(auto-fill, minmax(60px, 1fr));
        gap: 4px;
    }
    
    .popup-icon-item .app-label {
        font-size: 10px;
    }
}

/* Smooth focus animation for popup items */
@keyframes itemFocus {
    from { transform: scale(1); }
    to { transform: scale(1.05); }
}

.popup-icon-item:focus {
    animation: itemFocus 0.2s ease-out forwards;
    outline: 2px solid #005587;
    outline-offset: 2px;
}
</style>
'''

# JavaScript for category toggling and popup functionality
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

// Popup submenu functions
function showSubgroupPopup(subgroupId, subgroupName, subgroupIcon, event) {
    // Get the popup data stored in the data attribute
    const popupData = document.getElementById('subgroup-popup-data-' + subgroupId);
    if (!popupData) return;
    
    const data = JSON.parse(popupData.textContent);
    
    // Get the clicked button position
    const button = event ? event.currentTarget : null;
    const buttonRect = button ? button.getBoundingClientRect() : null;
    
    // Create overlay if it doesn't exist
    let overlay = document.getElementById('popup-overlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'popup-overlay';
        overlay.className = 'popup-overlay';
        overlay.onclick = hidePopupSubmenu;
        document.body.appendChild(overlay);
    }
    
    // Create popup menu
    let popup = document.getElementById('popup-menu');
    if (!popup) {
        popup = document.createElement('div');
        popup.id = 'popup-menu';
        popup.className = 'popup-menu';
        document.body.appendChild(popup);
    }
    
    // Build popup content for single subgroup
    let html = `
        <div class="popup-menu-header">
            <div class="popup-menu-title">
                <i class="fa ${subgroupIcon}"></i>
                ${subgroupName}
            </div>
            <button class="popup-close" onclick="hidePopupSubmenu()">
                <i class="fa fa-times"></i>
            </button>
        </div>
        <div class="popup-content-wrapper">
            <div class="popup-subgroup-icons">
    `;
    
    // Add icons for this subgroup
    for (const icon of data.icons) {
        html += `
            <a href="${icon.link}" class="popup-icon-item">
                ${icon.badge > 0 ? `<div class="${icon.badgeClass}">${icon.badge}</div>` : ''}
                <i class="fa ${icon.iconClass} app-icon"></i>
                <span class="app-label">${icon.label}</span>
            </a>
        `;
    }
    
    html += '</div></div>';
    popup.innerHTML = html;
    
    // Position popup based on screen size and click location
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;
    
    // On larger screens, position near the clicked button
    if (viewportWidth > 768 && buttonRect) {
        // Reset transform for accurate measurements
        popup.style.transform = 'scale(0.95)';
        popup.style.opacity = '0';
        popup.style.visibility = 'visible';
        
        // Get popup dimensions
        const popupRect = popup.getBoundingClientRect();
        
        // Calculate ideal position (below the button)
        let left = buttonRect.left + (buttonRect.width / 2) - (popupRect.width / 2);
        let top = buttonRect.bottom + 10;
        
        // Adjust if popup would go off screen
        if (left + popupRect.width > viewportWidth - 20) {
            left = viewportWidth - popupRect.width - 20;
        }
        if (left < 20) {
            left = 20;
        }
        
        // If popup would go below viewport, position above button
        if (top + popupRect.height > viewportHeight - 20) {
            top = buttonRect.top - popupRect.height - 10;
        }
        
        // If still doesn't fit, center vertically
        if (top < 20) {
            top = (viewportHeight - popupRect.height) / 2;
        }
        
        popup.style.left = left + 'px';
        popup.style.top = top + 'px';
        popup.style.transform = 'scale(0.95)';
    } else {
        // On mobile or if no button reference, center the popup
        popup.style.left = '50%';
        popup.style.top = '50%';
        popup.style.transform = 'translate(-50%, -50%) scale(0.95)';
    }
    
    popup.style.opacity = '0';
    popup.style.visibility = 'visible';
    
    // Force browser to calculate layout
    popup.offsetHeight;
    
    // Show overlay immediately
    overlay.style.display = 'block';
    setTimeout(() => {
        overlay.classList.add('active');
    }, 10);
    
    // Show popup with smooth animation
    setTimeout(() => {
        popup.classList.add('active');
        if (viewportWidth > 768 && buttonRect) {
            popup.style.transform = 'scale(1)';
        } else {
            popup.style.transform = 'translate(-50%, -50%) scale(1)';
        }
        popup.style.opacity = '1';
    }, 50);
}

function hidePopupSubmenu() {
    const overlay = document.getElementById('popup-overlay');
    const popup = document.getElementById('popup-menu');
    
    if (overlay) overlay.classList.remove('active');
    if (popup) {
        popup.classList.remove('active');
        popup.style.transform = 'translate(-50%, -50%) scale(0.95)';
        popup.style.opacity = '0';
    }
    
    // Hide elements after animation
    setTimeout(() => {
        if (overlay) overlay.style.display = 'none';
        if (popup) popup.style.visibility = 'hidden';
    }, 300);
}

// Close popup on escape key
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        hidePopupSubmenu();
    }
});

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
        if len(icon) >= 11:
            _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _, _ = icon
        elif len(icon) >= 10:
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

# Helper function to render popup subgroup tiles
def _render_popup_tiles(category_id, popup_list):
    """Render a list of popup subgroups as clickable tiles.
    popup_list: [(subgroup_id, subgroup_info, icons_list), ...]
    """
    import json

    print '''
    <div class="popup-subgroups-container">
    '''

    for subgroup_id, subgroup_info, icons_list in popup_list:
        unique_subgroup_id = "{0}_{1}".format(category_id, subgroup_id)

        print '''
            <div class="subgroup-popup-item">
                <button class="subgroup-popup-button" onclick="showSubgroupPopup('{0}', '{1}', '{2}', event)">
                    <div class="subgroup-dropdown-indicator">
                        <i class="fa fa-caret-down"></i>
                    </div>
                    <div class="app-icon-container">
                        <i class="fa {2} app-icon"></i>
                    </div>
                    <div class="app-label-container">
                        <span class="app-label">{1}</span>
                    </div>
                </button>
            </div>
        '''.format(unique_subgroup_id, subgroup_info["name"], subgroup_info["icon"])

        popup_data = {"icons": []}

        for icon in icons_list:
            if len(icon) >= 11:
                _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _, _ = icon
            elif len(icon) >= 10:
                _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _ = icon
            elif len(icon) >= 9:
                _, icon_class, label, link, org_id, query, custom_query, required_roles, _ = icon
            else:
                _, icon_class, label, link, org_id, query, custom_query, required_roles = icon

            count = 0
            if custom_query and isinstance(custom_query, str):
                count = get_icon_count(custom_query, org_id)
            elif query and isinstance(query, str):
                count = get_icon_count(query, org_id)
            elif custom_query:
                try:
                    count = q.QuerySqlInt(custom_query)
                except:
                    count = 0
            elif query and org_id:
                try:
                    count = q.QuerySqlInt(query.format(org_id))
                except:
                    count = 0

            badge_class = "notification-badge"
            if count >= 100:
                badge_class += " notification-badge-large"
            elif count >= 10:
                badge_class += " notification-badge-medium"

            popup_data["icons"].append({
                "iconClass": icon_class,
                "label": label,
                "link": link,
                "badge": count,
                "badgeClass": badge_class
            })

        print '''<script type="application/json" id="subgroup-popup-data-{0}">{1}</script>'''.format(
            unique_subgroup_id,
            json.dumps(popup_data)
        )

    print '''
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
    
    # Sort subgroups to show in a logical order based on SUBGROUP_INFO
    # First, get the order from SUBGROUP_INFO keys
    subgroup_order = list(Config.SUBGROUP_INFO.keys())
    # Add "Other" at the end if not already in the list
    if "Other" not in subgroup_order:
        subgroup_order.append("Other")
    
    sorted_subgroups = []
    
    # Add subgroups in preferred order if they exist
    for sg in subgroup_order:
        if sg in subgroups:
            sorted_subgroups.append(sg)
    
    # Add any remaining subgroups not in the order list
    for sg in subgroups:
        if sg not in sorted_subgroups:
            sorted_subgroups.append(sg)
    
    # Build parent_subgroup mapping: parent_id -> [child popup subgroups]
    child_popups = {}
    for sg_id in sorted_subgroups:
        sg_info = Config.SUBGROUP_INFO.get(sg_id, {})
        parent = sg_info.get("parent_subgroup")
        if parent and sg_info.get("popup", False):
            if parent not in child_popups:
                child_popups[parent] = []
            child_popups[parent].append(sg_id)

    # Track if we have any popup subgroups that need to be grouped (top-level only)
    popup_subgroups = []

    # Print each subgroup
    for subgroup_id in sorted_subgroups:
        icons_list = subgroups.get(subgroup_id, [])

        # Get subgroup info
        subgroup_info = Config.SUBGROUP_INFO.get(subgroup_id,
                                                 {"name": subgroup_id, "icon": "fa-folder", "popup": False})

        # Skip child popup subgroups - they render inside their parent
        parent = subgroup_info.get("parent_subgroup")
        if parent and subgroup_info.get("popup", False) and parent in subgroups or (parent in [s for s in sorted_subgroups]):
            continue

        # Check if this subgroup should be a popup
        is_popup_subgroup = subgroup_info.get("popup", False)

        if is_popup_subgroup:
            # Collect top-level popup subgroups to render together
            if icons_list:
                popup_subgroups.append((subgroup_id, subgroup_info, icons_list))
        else:
            # Skip empty inline subgroups that also have no child popups
            if not icons_list and subgroup_id not in child_popups:
                continue

            # Regular inline subgroup display
            print '''
            <div class="subgroup-section">
                <h4 class="subgroup-header">
                    <i class="fa {0}"></i> {1}
                </h4>
                <div class="subgroup-content">
            '''.format(subgroup_info["icon"], subgroup_info["name"])

            # Sort icons in this subgroup by priority and label
            sorted_icons = sort_icons(icons_list)

            # Print icons in this subgroup
            for icon in sorted_icons:
                # Unpack icon data
                if len(icon) >= 11:
                    _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _, _ = icon
                elif len(icon) >= 10:
                    _, icon_class, label, link, org_id, query, custom_query, required_roles, _, _ = icon
                elif len(icon) >= 9:
                    _, icon_class, label, link, org_id, query, custom_query, required_roles, _ = icon
                else:
                    _, icon_class, label, link, org_id, query, custom_query, required_roles = icon

                print generate_app_icon(icon_class, label, link, org_id, query, custom_query, required_roles)

            # Render child popup subgroups inside this inline parent
            if subgroup_id in child_popups:
                child_popup_list = []
                for child_sg_id in child_popups[subgroup_id]:
                    child_icons = subgroups.get(child_sg_id, [])
                    if child_icons:
                        child_info = Config.SUBGROUP_INFO.get(child_sg_id, {"name": child_sg_id, "icon": "fa-folder", "popup": True})
                        child_popup_list.append((child_sg_id, child_info, child_icons))

                if child_popup_list:
                    _render_popup_tiles(category_id, child_popup_list)

            print '''
                </div>
            </div>
            '''

    # Now render top-level popup subgroups together
    if popup_subgroups:
        _render_popup_tiles(category_id, popup_subgroups)
    
    print '''
        </div>
    </div>
    '''

# Hide TouchPoint header/footer on mobile app (works for both widgets and standalone scripts)
print '''<script>
(function(){
    var ua=navigator.userAgent;
    var isTPApp=(/iPhone|iPad|Android/.test(ua)&&/AppleWebKit/.test(ua)&&!/Safari/.test(ua))
        ||(/Android/.test(ua)&&/wv/.test(ua));
    if(!isTPApp)return;
    var rules="#top-navbar,#navbar,.navbar,nav,header,footer,#footer,#contact-footer,.page-header,#page-header,#header,.hidden-print{display:none!important}body{padding-top:0!important}";
    var s=document.createElement("style");s.textContent=rules;document.documentElement.appendChild(s);
    try{if(window.parent&&window.parent.document!==document){var ps=window.parent.document.createElement("style");ps.textContent=rules;window.parent.document.documentElement.appendChild(ps);}}catch(e){}
    try{if(window.top&&window.top.document!==document){var ts=window.top.document.createElement("style");ts.textContent=rules;window.top.document.documentElement.appendChild(ts);}}catch(e){}
})();
</script>'''

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
            
            # Get expiration date if exists (index 10)
            expiration_date = None
            if len(icon) >= 11:
                expiration_date = icon[10]
            
            # Check both role permissions and expiration
            if user_has_required_roles(icon_roles) and not is_icon_expired(expiration_date):
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
            # Use sub-grouped display (which handles both inline and popup subgroups)
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
