# Role=Admin
###############################################################
# TouchPoint Authenticated Link Generator with Live Search
###############################################################
# This script provides a powerful interface for creating authenticated links
# to various TouchPoint features for a specific person. It includes:
# - Live search to quickly find people
# - Generate links to organizations, profile, and custom URLs
# - Copy to clipboard functionality for easy sharing
# - Simple and intuitive interface

# To upload code to Touchpoint:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the script with your desired name and paste this code
# 4. Test and add to menu

# Written by: Ben Swaby
# Email: bswaby@fbchtn.org
# Last updated: 05/12/2025

import sys
import traceback

# Configuration options
MAX_RESULTS = 15  # Maximum number of results to display in search
SEARCH_DELAY = 300  # Milliseconds to wait after typing before searching
MAX_ORGS_SHOWN = 10  # Maximum number of organizations to display

# Class for handling search functionality
class PeopleSearch:
    @staticmethod
    def search_people(search_term, max_results=MAX_RESULTS):
        """Search for people using the provided search term"""
        try:
            # Helper function to normalize phone numbers for searching
            def normalize_phone(phone_str):
                if not phone_str:
                    return ""
                # Remove all non-digit characters
                return ''.join(c for c in phone_str if c.isdigit())
            
            # Check if the search term is a phone number
            def is_phone_search(term):
                # A phone search is when we have at least 3 digits in the search term
                digits = ''.join(c for c in term if c.isdigit())
                return len(digits) >= 3
            
            is_phone_number_search = is_phone_search(search_term)
            normalized_search = normalize_phone(search_term)
            
            # People search results list
            people = []
            
            # Try to use a simple query for people
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
                    """.format(normalized_search, max_results)
                    people = q.QuerySql(people_sql)
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
                    """.format(search_term.replace("'", "''"), max_results)
                    people = q.QuerySql(people_sql)
            except Exception as e:
                # Log the error but continue
                print "<!-- Error in people search: " + str(e) + " -->"
                
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
                    """.format(search_term.replace("'", "''"), max_results)
                    people = q.QuerySql(people_sql)
                except:
                    people = []
            
            return people
        except Exception as e:
            print "<!-- Error in people_search method: " + str(e) + " -->"
            return []

# Class for generating authenticated links
class LinkGenerator:
    @staticmethod
    def get_authenticated_url(people_id, url, include_host=True):
        """Generate an authenticated URL for the given people ID and URL"""
        try:
            return model.GetAuthenticatedUrl(int(people_id), url, include_host)
        except Exception as e:
            print "<!-- Error in get_authenticated_url: " + str(e) + " -->"
            return None
    
    @staticmethod
    def get_person_organizations(people_id, max_orgs=MAX_ORGS_SHOWN):
        """Get organizations a person is a member of, with focus on leadership and financial info"""
        try:
            # Get orgs where the person is a member
            orgs_sql = """
            SELECT TOP {1}
                o.OrganizationId,
                o.OrganizationName,
                mt.Description AS MemberType,
                mt.Id AS MemberTypeId,
                (SELECT SUM(Amount - AmountPaid) FROM OrganizationMembers 
                 WHERE OrganizationId = o.OrganizationId AND PeopleId = {0} AND Amount > AmountPaid) AS BalanceDue,
                p.Name AS ProgramName,
                d.Name AS DivisionName,
                om.AttendPct AS AttendancePercentage,
                o.MemberCount
            FROM OrganizationMembers om
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program p ON d.ProgId = p.Id
            WHERE om.PeopleId = {0}
                AND o.OrganizationStatusId = 30 -- Active orgs only
            ORDER BY 
                CASE WHEN mt.AttendanceTypeId IN (10, 60) THEN 0 ELSE 1 END, -- Leaders first
                CASE WHEN (SELECT SUM(Amount - AmountPaid) FROM OrganizationMembers 
                         WHERE OrganizationId = o.OrganizationId AND PeopleId = {0} AND Amount > AmountPaid) > 0 
                     THEN 0 ELSE 1 END, -- Balance due orgs next
                o.OrganizationName
            """.format(people_id, max_orgs)
            
            return q.QuerySql(orgs_sql)
        except Exception as e:
            print "<!-- Error in get_person_organizations: " + str(e) + " -->"
            return []

    @staticmethod
    def get_person_leader_orgs(people_id, max_orgs=MAX_ORGS_SHOWN):
        """Get organizations where the person is a leader"""
        try:
            # Get orgs where the person is a leader
            leader_orgs_sql = """
            SELECT TOP {1}
                o.OrganizationId,
                o.OrganizationName,
                mt.Description AS MemberType,
                p.Name AS ProgramName,
                d.Name AS DivisionName,
                o.MemberCount
            FROM OrganizationMembers om
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program p ON d.ProgId = p.Id
            WHERE om.PeopleId = {0}
                AND o.OrganizationStatusId = 30 -- Active orgs only
                AND mt.AttendanceTypeId IN (10, 60) -- Leader types
            ORDER BY 
                o.OrganizationName
            """.format(people_id, max_orgs)
            
            return q.QuerySql(leader_orgs_sql)
        except Exception as e:
            print "<!-- Error in get_person_leader_orgs: " + str(e) + " -->"
            return []

# Helper function to get script name from URL
def get_script_name():
    """Get the current script name from the URL or use a default"""
    script_name = "TPxi_LinkGenerator"  # Default name
    try:
        if hasattr(model.Data, "PATH_INFO"):
            path_parts = model.Data.PATH_INFO.split('/')
            if len(path_parts) > 2 and path_parts[1] == "PyScript":
                script_name = path_parts[2]
    except:
        pass  # Use default name if we can't extract from PATH_INFO
    return script_name

# Main script logic starts here
try:
    # Initialize parameters
    search_term = ""
    ajax_mode = False
    selected_person_id = None
    get_auth_url = False
    url_to_auth = ""
    pid_for_auth = ""
    auth_url_result = None
    
    # Get query parameters
    try:
        if hasattr(model.Data, "q"):
            search_term = str(model.Data.q)
        if hasattr(model.Data, "ajax") and model.Data.ajax == "1":
            ajax_mode = True
        if hasattr(model.Data, "person_id") and model.Data.person_id:
            selected_person_id = int(model.Data.person_id)
        if hasattr(model.Data, "url") and model.Data.url:
            url_to_auth = str(model.Data.url)
        if hasattr(model.Data, "pid") and model.Data.pid:
            pid_for_auth = str(model.Data.pid)
            
        # Check if we have both URL and PID parameters
        if url_to_auth and pid_for_auth:
            get_auth_url = True
            
            # Generate authenticated URL
            try:
                auth_url_result = LinkGenerator.get_authenticated_url(int(pid_for_auth), url_to_auth, True)
            except Exception as e:
                print "<div class='error-message'>Error generating authenticated URL: {0}</div>".format(str(e))
    except Exception as e:
        print "<div class='error-message'>Error processing parameters: {0}</div>".format(str(e))
    
    # Set page title
    model.Header = "Authenticated Link Generator"
    
    # Get script name for form submission
    script_name = get_script_name()
    
    # AJAX mode - search results
    if ajax_mode and search_term:
        people = PeopleSearch.search_people(search_term)
        
        # Helper function to format phone numbers
        def format_phone(phone):
            try:
                if phone:
                    return model.FmtPhone(phone)
                return ""
            except:
                return phone or ""
        
        # Output search results as HTML
        print "<div class='search-results-container'>"
        
        if people and len(people) > 0:
            for person in people:
                # Get person attributes safely
                person_id = getattr(person, 'PeopleId', '')
                person_name = getattr(person, 'Name', '')
                person_email = getattr(person, 'EmailAddress', '') or ''
                person_cell = getattr(person, 'CellPhone', '') or ''
                person_home = getattr(person, 'HomePhone', '') or ''
                person_age = getattr(person, 'Age', '') or ''
                
                # Format the phones
                formatted_cell = format_phone(person_cell)
                formatted_home = format_phone(person_home)
                
                print """
                <div class="result-item" data-person-id="{0}" data-person-name="{1}">
                    <div class="result-name">{1}</div>
                    <div class="result-meta">
                """.format(person_id, person_name)
                
                if person_age:
                    print "<span>Age: {0}</span>".format(person_age)
                
                print "</div>"
                
                # Contact information
                if person_email or person_cell or person_home:
                    print "<div class='result-meta'>"
                    
                    if person_email:
                        print "<div><i class='fa fa-envelope'></i> {0}</div>".format(person_email)
                    
                    if person_cell:
                        print "<div><i class='fa fa-mobile'></i> {0}</div>".format(formatted_cell)
                    
                    if person_home and person_home != person_cell:
                        print "<div><i class='fa fa-phone'></i> {0}</div>".format(formatted_home)
                    
                    print "</div>"
                
                print """
                    <div class="result-button">
                        <button class="select-person-btn">Select</button>
                    </div>
                </div>
                """
        else:
            print "<div class='no-results'>No people found matching your search.</div>"
        
        print "</div>"
    
    # If person selected, show link generation options
    elif selected_person_id:
        try:
            # Get person details
            person_sql = """
            SELECT p.PeopleId, p.Name, p.EmailAddress, p.Age, ms.Description AS MemberStatus
            FROM People p
            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
            WHERE p.PeopleId = {0}
            """.format(selected_person_id)
            
            person_details = q.QuerySqlTop1(person_sql)
            
            if person_details:
                # Get organizations where the person is a member
                orgs = LinkGenerator.get_person_organizations(selected_person_id)
                
                # Get leadership organizations specifically
                leader_orgs = LinkGenerator.get_person_leader_orgs(selected_person_id)
                
                # Generate profile link
                profile_url = "/Person2/{0}".format(selected_person_id)
                profile_link = LinkGenerator.get_authenticated_url(selected_person_id, profile_url)
                
                # Generate dashboard link
                dashboard_url = "/Person2/{0}#tab-touchpoints".format(selected_person_id)
                dashboard_link = LinkGenerator.get_authenticated_url(selected_person_id, dashboard_url)
                
                # Generate giving link
                giving_url = "/OnlineReg/GiveNow?PeopleId={0}".format(selected_person_id)
                giving_link = LinkGenerator.get_authenticated_url(selected_person_id, giving_url)
                
                # Output person details and link generation form
                print """
                <div class="person-details">
                    <h2>{0}</h2>
                    <p>{1} • Age: {2}</p>
                    <p>{3}</p>
                </div>
                
                <div class="link-generation">
                    <h3>Select a Link Type</h3>
                    
                    <div class="incognito-tip">
                        <i class="fa fa-info-circle"></i> <strong>Tip:</strong> For testing links as the user would see them, 
                        open them in an <strong>incognito/private window</strong> to avoid your current admin session.
                    </div>
                    
                    <div class="link-section">
                        <h4>Profile Links</h4>
                        <div class="link-option" id="profile-link">
                            <div class="link-label">Profile Page</div>
                            <div class="link-value">{4}</div>
                            <button class="copy-link-btn" data-link="{4}">Copy Link</button>
                            <a href="{4}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                        </div>
                        
                        <div class="link-option" id="dashboard-link">
                            <div class="link-label">Touchpoints Dashboard</div>
                            <div class="link-value">{5}</div>
                            <button class="copy-link-btn" data-link="{5}">Copy Link</button>
                            <a href="{5}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                        </div>
                        
                        <div class="link-option" id="giving-link">
                            <div class="link-label">Online Giving</div>
                            <div class="link-value">{6}</div>
                            <button class="copy-link-btn" data-link="{6}">Copy Link</button>
                            <a href="{6}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                        </div>
                    </div>
                """.format(
                    person_details.Name,
                    person_details.MemberStatus,
                    person_details.Age or "N/A",
                    person_details.EmailAddress or "No email available",
                    profile_link,
                    dashboard_link,
                    giving_link
                )
                
                # If they have organizations with balance due, show those
                finance_orgs_available = False
                print """
                <div class="link-section">
                    <h4>Financial Links</h4>
                """
                
                for org in orgs:
                    balance_due = getattr(org, 'BalanceDue', 0) or 0
                    if balance_due > 0:
                        finance_orgs_available = True
                        org_url = "/OnlineReg/{0}?PeopleId={1}".format(org.OrganizationId, selected_person_id)
                        org_link = LinkGenerator.get_authenticated_url(selected_person_id, org_url)
                        
                        print """
                        <div class="link-option">
                            <div class="link-label">{0} - Balance: ${1}</div>
                            <div class="link-value">{2}</div>
                            <button class="copy-link-btn" data-link="{2}">Copy Link</button>
                            <a href="{2}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                        </div>
                        """.format(
                            org.OrganizationName,
                            balance_due,
                            org_link
                        )
                
                if not finance_orgs_available:
                    print "<p>No organizations with balance due.</p>"
                
                print "</div>"  # End finance section
                
                # If they are a leader in any organizations, show those
                print """
                <div class="link-section">
                    <h4>Leadership Links</h4>
                """
                
                if leader_orgs and len(leader_orgs) > 0:
                    for org in leader_orgs:
                        org_url = "/Org/{0}".format(org.OrganizationId)
                        org_link = LinkGenerator.get_authenticated_url(selected_person_id, org_url)
                        
                        print """
                        <div class="link-option">
                            <div class="link-label">{0} - {1} ({2} members)</div>
                            <div class="link-value">{3}</div>
                            <button class="copy-link-btn" data-link="{3}">Copy Link</button>
                            <a href="{3}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                        </div>
                        """.format(
                            org.OrganizationName,
                            org.MemberType,
                            org.MemberCount,
                            org_link
                        )
                else:
                    print "<p>No leadership roles found.</p>"
                
                print "</div>"  # End leadership section
                
                # All organizations section
                print """
                <div class="link-section">
                    <h4>All Organizations</h4>
                """
                
                if orgs and len(orgs) > 0:
                    for org in orgs:
                        org_url = "/Org/{0}".format(org.OrganizationId)
                        org_link = LinkGenerator.get_authenticated_url(selected_person_id, org_url)
                        
                        print """
                        <div class="link-option">
                            <div class="link-label">{0} - {1}</div>
                            <div class="link-value">{2}</div>
                            <button class="copy-link-btn" data-link="{2}">Copy Link</button>
                            <a href="{2}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                        </div>
                        """.format(
                            org.OrganizationName,
                            org.MemberType,
                            org_link
                        )
                else:
                    print "<p>No organizations found.</p>"
                
                print "</div>"  # End all orgs section
                
                # Custom URL section
                print """
                <div class="link-section">
                    <h4>Custom URL</h4>
                    <div class="custom-url-form">
                        <form method="get" action="/PyScript/{0}">
                            <input type="hidden" name="person_id" value="{1}" />
                            <input type="hidden" name="pid" value="{1}" />
                            <div class="input-group">
                                <input type="text" id="customUrl" name="url" placeholder="Enter URL path (e.g., /Settings)" 
                                       class="custom-url-input" required />
                                <button type="submit" class="custom-url-button">Generate Link</button>
                            </div>
                            <div class="form-help" style="margin-top: 5px; font-size: 12px; color: #666;">
                                Enter a full URL (e.g., https://myfbch.com/Person2/12345) or just a path (e.g., /Person2/12345)
                            </div>
                        </form>
                    </div>
                """.format(script_name, selected_person_id)
                
                # If we have a custom URL result to show
                if get_auth_url and auth_url_result and str(selected_person_id) == pid_for_auth:
                    print """
                    <div id="customLinkResult" class="link-option" style="margin-top: 15px;">
                        <div class="link-label">Custom Link for: {0}</div>
                        <div class="link-value">{1}</div>
                        <button class="copy-link-btn" data-link="{1}">Copy Link</button>
                        <a href="{1}" target="_blank" class="open-link-btn"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                    </div>
                    """.format(url_to_auth, auth_url_result)
                
                print """
                </div>  <!-- End custom URL section -->
                </div>  <!-- End link-generation -->
                
                <div class="back-section">
                    <button id="backToSearch">← Back to Search</button>
                </div>
                """
            else:
                print "<div class='error-message'>Person not found.</div>"
                print "<div class='back-section'><button id='backToSearch'>← Back to Search</button></div>"
        
        except Exception as e:
            print "<div class='error-message'>Error loading person details: {0}</div>".format(str(e))
            print "<pre class='error-details'>{0}</pre>".format(traceback.format_exc())
            print "<div class='back-section'><button id='backToSearch'>← Back to Search</button></div>"
    
    # If just URL and PID but no person selected, show the authenticated URL result
    elif get_auth_url and auth_url_result:
        print """
        <div style="max-width: 800px; margin: 0 auto; padding: 20px;">
            <h2>Authenticated URL Generated</h2>
            
            <div style="margin-top: 20px; padding: 15px; background: #f5f5f5; border-radius: 4px;">
                <div style="font-weight: bold; margin-bottom: 10px;">URL Path: {0}</div>
                <div style="font-family: monospace; padding: 10px; background: white; border: 1px solid #ddd; border-radius: 4px; word-break: break-all; margin-bottom: 15px;">{1}</div>
                
                <div>
                    <button onclick="copyToClipboard('{1}')" style="padding: 8px 15px; background: #0B3D4C; color: white; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px;">Copy Link</button>
                    <a href="{1}" target="_blank" style="padding: 8px 15px; background: #4CAF50; color: white; text-decoration: none; display: inline-block; border-radius: 4px;"><i class="fa fa-user-secret"></i> Open in Incognito</a>
                </div>
            </div>
            
            <div style="margin-top: 20px;">
                <button onclick="window.history.back()" style="padding: 10px 20px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer;">← Back</button>
            </div>
        </div>
        
        <script>
        function copyToClipboard(text) {{
            var tempInput = document.createElement('textarea');
            tempInput.value = text;
            document.body.appendChild(tempInput);
            tempInput.select();
            
            try {{
                var successful = document.execCommand('copy');
                if (successful) {{
                    alert('Link copied to clipboard!');
                }} else {{
                    alert('Failed to copy link');
                }}
            }} catch (err) {{
                alert('Failed to copy link: ' + err);
            }}
            
            document.body.removeChild(tempInput);
        }}
        </script>
        """.format(url_to_auth, auth_url_result)
    
    # Main search interface (default view)
    else:
        # Display search interface
        print """
        <div class="search-container">
            <h2>Authenticated Link Generator<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
            <p>Search for a person to generate authenticated links for TouchPoint features</p>
            
            <div class="security-warning">
                <h3><i class="fa fa-exclamation-triangle"></i> Security Notice</h3>
                <p>Pre-authenticated links provide direct access to TouchPoint data <strong>without requiring login</strong>. 
                These links should be:</p>
                <ul>
                    <li>Shared only with the intended recipient</li>
                    <li>Sent through secure channels (e.g., direct messages, not public forums)</li>
                    <li>Used for limited time purposes (links will eventually expire)</li>
                    <li>Never posted publicly or in group communications</li>
                </ul>
                <p>By using this tool, you accept responsibility for the secure handling of these links.</p>
            </div>
            
            <div class="incognito-tip">
                <i class="fa fa-info-circle"></i> <strong>How to Test Links:</strong> Use an <strong>incognito/private window</strong> to test links. 
                This avoids your active admin session and shows exactly what the recipient will see.
            </div>
            
            <div class="search-box">
                <input type="text" id="searchInput" class="search-input" placeholder="Search by name, phone, or email...">
                <span id="clearSearch" class="search-clear" style="display: none;">×</span>
                <div class="search-loading" id="searchLoading">
                    <div class="spinner"></div>
                </div>
                <button type="button" id="searchButton" class="search-button">
                    <i class="fa fa-search"></i> Search
                </button>
            </div>
            
            <div id="resultsContainer" class="results-container">
                <!-- Results will be loaded here -->
            </div>
        </div>
        """

    # Include CSS and JavaScript for the page
    print """
        <style>
        .search-container {{
            max-width: 800px; 
            margin: 0 auto;
            font-family: Arial, sans-serif;
            padding: 20px;
            width: 100%;
            box-sizing: border-box;
        }}
        .security-warning {{
            background-color: #fff3cd;
            border: 1px solid #ffeeba;
            border-left: 5px solid #ffc107;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }}
        .security-warning h3 {{
            color: #856404;
            margin-top: 0;
        }}
        .security-warning ul {{
            margin-left: 20px;
            padding-left: 0;
        }}
        .security-warning li {{
            margin-bottom: 5px;
        }}
        .search-box {{
            position: relative;
            margin-bottom: 20px;
            width: 100%;
            box-sizing: border-box;
        }}
        .search-input {{
            width: 100%;
            padding: 12px 50px 12px 15px;
            font-size: 16px;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            box-sizing: border-box;
        }}
        .search-button {{
            position: absolute;
            right: 5px;
            top: 5px;
            padding: 8px 15px;
            background: #0B3D4C;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }}
        .search-loading {{
            display: none;
            position: absolute;
            right: 60px;
            top: 12px;
        }}
        .search-clear {{
            position: absolute;
            right: 60px;
            top: 12px;
            color: #999;
            font-size: 24px;
            cursor: pointer;
            display: none;
        }}
        .search-clear:hover {{
            color: #333;
        }}
        .results-container {{
            margin-top: 20px;
            width: 100%;
        }}
        .search-results-container {{
            max-height: 500px;
            overflow-y: auto;
            border: 1px solid #eee;
            border-radius: 4px;
        }}
        .result-item {{
            padding: 15px;
            margin-bottom: 10px;
            border: 1px solid #eee;
            border-radius: 4px;
            transition: all 0.2s;
            display: flex;
            flex-direction: column;
            background-color: white;
            cursor: pointer;
        }}
        .result-item:hover {{
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            background-color: #f8f9fa;
        }}
        .result-name {{
            font-weight: bold;
            margin-bottom: 5px;
            font-size: 16px;
        }}
        .result-meta {{
            font-size: 14px;
            color: #666;
            margin-bottom: 10px;
        }}
        .result-button {{
            margin-top: 10px;
            text-align: right;
        }}
        .select-person-btn {{
            padding: 8px 15px;
            background: #0B3D4C;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }}
        .select-person-btn:hover {{
            background: #0a2933;
        }}
        .no-results {{
            padding: 20px;
            text-align: center;
            color: #777;
            background: #f9f9f9;
            border-radius: 4px;
        }}
        @keyframes spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
        .spinner {{
            border: 3px solid #f3f3f3;
            border-top: 3px solid #0B3D4C;
            border-radius: 50%;
            width: 20px;
            height: 20px;
            animation: spin 1s linear infinite;
        }}
        
        /* Link generation styles */
        .person-details {{
            padding: 20px;
            margin-bottom: 20px;
            background: #f5f5f5;
            border-radius: 4px;
        }}
        .person-details h2 {{
            margin-top: 0;
            color: #0B3D4C;
        }}
        .link-generation {{
            margin-bottom: 30px;
        }}
        .link-section {{
            margin-bottom: 25px;
            padding: 15px;
            background: #f9f9f9;
            border-radius: 4px;
            border-left: 4px solid #0B3D4C;
        }}
        .link-section h4 {{
            margin-top: 0;
            color: #0B3D4C;
            border-bottom: 1px solid #ddd;
            padding-bottom: 10px;
        }}
        .link-option {{
            margin-bottom: 15px;
            padding: 10px;
            background: white;
            border: 1px solid #eee;
            border-radius: 4px;
        }}
        .link-label {{
            font-weight: bold;
            margin-bottom: 8px;
        }}
        .link-value {{
            font-family: monospace;
            padding: 8px;
            background: #f5f5f5;
            border: 1px solid #ddd;
            border-radius: 4px;
            word-break: break-all;
            margin-bottom: 10px;
            font-size: 12px;
            overflow-x: auto;
        }}
        .copy-link-btn, .open-link-btn {{
            padding: 8px 15px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            display: inline-block;
            margin-right: 10px;
            text-decoration: none;
            font-size: 14px;
        }}
        .copy-link-btn {{
            background: #0B3D4C;
            color: white;
        }}
        .copy-link-btn:hover {{
            background: #0a2933;
        }}
        .open-link-btn {{
            background: #4CAF50;
            color: white;
        }}
        .open-link-btn:hover {{
            background: #3e8e41;
        }}
        .custom-url-form {{
            margin-bottom: 15px;
        }}
        .custom-url-form form {{
            width: 100%;
        }}
        .input-group {{
            display: flex;
            width: 100%;
        }}
        .custom-url-input {{
            flex: 1;
            padding: 10px;
            border: 1px solid #ccc;
            border-radius: 4px 0 0 4px;
            font-size: 14px;
        }}
        .custom-url-button {{
            padding: 10px 15px;
            background: #0B3D4C;
            color: white;
            border: none;
            border-radius: 0 4px 4px 0;
            cursor: pointer;
        }}
        .back-section {{
            margin-top: 20px;
        }}
        #backToSearch {{
            padding: 10px 20px;
            background: #f5f5f5;
            border: 1px solid #ddd;
            border-radius: 4px;
            cursor: pointer;
        }}
        #backToSearch:hover {{
            background: #e5e5e5;
        }}
        .error-message {{
            padding: 15px;
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
            border-radius: 4px;
            margin-bottom: 20px;
        }}
        .error-details {{
            background: #f8f9fa;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            overflow: auto;
            font-size: 12px;
            margin-bottom: 20px;
        }}
        /* Toast notification */
        .toast {{
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 10px 20px;
            background-color: #4CAF50;
            color: white;
            border-radius: 4px;
            z-index: 1000;
            display: none;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }}
        .incognito-tip {{
            background-color: #e7f3ff;
            border: 1px solid #b3d7ff;
            border-radius: 4px;
            padding: 10px 15px;
            margin-bottom: 20px;
            color: #004085;
        }}
        </style>
        
        <div id="toastNotification" class="toast"></div>
        
        <script>
        // Helper function to show toast notification
        function showToast(message, duration) {{
            var toast = document.getElementById('toastNotification');
            toast.textContent = message;
            toast.style.display = 'block';
            
            setTimeout(function() {{
                toast.style.display = 'none';
            }}, duration || 3000);
        }}
        
        // Copy to clipboard functionality
        function copyToClipboard(text) {{
            var tempInput = document.createElement('textarea');
            tempInput.value = text;
            document.body.appendChild(tempInput);
            tempInput.select();
            
            try {{
                var successful = document.execCommand('copy');
                if (successful) {{
                    showToast('Link copied to clipboard!', 2000);
                }} else {{
                    showToast('Failed to copy link', 2000);
                }}
            }} catch (err) {{
                showToast('Failed to copy link: ' + err, 3000);
            }}
            
            document.body.removeChild(tempInput);
        }}
        
        // Function to perform search
        function performSearch() {{
            var searchTerm = document.getElementById('searchInput').value.trim();
            var resultsContainer = document.getElementById('resultsContainer');
            var searchLoading = document.getElementById('searchLoading');
            
            // Don't search if term is too short
            if (searchTerm.length < 2) {{
                resultsContainer.innerHTML = '<div class="no-results">Please enter at least 2 characters to search</div>';
                return;
            }}
            
            // Show loading indicator
            searchLoading.style.display = 'block';
            resultsContainer.innerHTML = '<div class="loading-message">Searching...</div>';
            
            // Make AJAX request
            var xhr = new XMLHttpRequest();
            xhr.open('GET', window.location.pathname + '?q=' + encodeURIComponent(searchTerm) + '&ajax=1', true);
            xhr.onreadystatechange = function() {{
                if (xhr.readyState === 4) {{
                    searchLoading.style.display = 'none';
                    
                    if (xhr.status === 200) {{
                        resultsContainer.innerHTML = xhr.responseText;
                        
                        // Attach click handlers to select buttons
                        var selectButtons = document.querySelectorAll('.select-person-btn');
                        for (var i = 0; i < selectButtons.length; i++) {{
                            selectButtons[i].addEventListener('click', function() {{
                                var personItem = this.closest('.result-item');
                                var personId = personItem.getAttribute('data-person-id');
                                var personName = personItem.getAttribute('data-person-name');
                                
                                // Navigate to the link generation page for this person
                                window.location.href = window.location.pathname + '?person_id=' + personId;
                            }});
                        }}
                        
                        // Attach click handlers to result items for easier selection
                        var resultItems = document.querySelectorAll('.result-item');
                        for (var i = 0; i < resultItems.length; i++) {{
                            resultItems[i].addEventListener('click', function(e) {{
                                // Don't trigger if clicking on the button (it has its own handler)
                                if (!e.target.classList.contains('select-person-btn') && 
                                    !e.target.parentElement.classList.contains('select-person-btn')) {{
                                    var personId = this.getAttribute('data-person-id');
                                    window.location.href = window.location.pathname + '?person_id=' + personId;
                                }}
                            }});
                        }}
                    }} else {{
                        resultsContainer.innerHTML = '<div class="error-message">An error occurred while searching. Please try again.</div>';
                    }}
                }}
            }};
            xhr.send();
        }}
        
        // Document ready
        document.addEventListener('DOMContentLoaded', function() {{
            var searchInput = document.getElementById('searchInput');
            var searchButton = document.getElementById('searchButton');
            var clearButton = document.getElementById('clearSearch');
            var searchTimeout = null;
            
            // Search as you type with a delay
            if (searchInput) {{
                searchInput.focus();
                
                // Search as you type with a delay
                searchInput.addEventListener('input', function() {{
                    clearTimeout(searchTimeout);
                    
                    // Show/hide clear button
                    clearButton.style.display = this.value.length > 0 ? 'block' : 'none';
                    
                    if (this.value.length >= 2) {{
                        searchTimeout = setTimeout(performSearch, {0});
                    }} else {{
                        document.getElementById('resultsContainer').innerHTML = '';
                    }}
                }});
                
                // Search when Enter is pressed
                searchInput.addEventListener('keypress', function(e) {{
                    if (e.key === 'Enter') {{
                        e.preventDefault();
                        performSearch();
                    }}
                }});
            }}
            
            // Clear search
            if (clearButton) {{
                clearButton.addEventListener('click', function() {{
                    searchInput.value = '';
                    clearButton.style.display = 'none';
                    document.getElementById('resultsContainer').innerHTML = '';
                    searchInput.focus();
                }});
            }}
            
            // Search button click
            if (searchButton) {{
                searchButton.addEventListener('click', performSearch);
            }}
            
            // Back to search button
            var backButton = document.getElementById('backToSearch');
            if (backButton) {{
                backButton.addEventListener('click', function() {{
                    window.location.href = window.location.pathname;
                }});
            }}
            
            // Copy link buttons
            var copyButtons = document.querySelectorAll('.copy-link-btn');
            for (var i = 0; i < copyButtons.length; i++) {{
                copyButtons[i].addEventListener('click', function() {{
                    var link = this.getAttribute('data-link');
                    copyToClipboard(link);
                }});
            }}
        }});
        </script>
        """.format(SEARCH_DELAY)  # Pass the search delay to the JavaScript

except Exception as e:
    # Print any errors that occurred
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
