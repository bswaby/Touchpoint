"""
Involvement Notification Auditor

This tool identifies involvements where former staff members might still be receiving notifications.
It helps administrators clean up notification settings when staff leave the organization.

Features:
- Identifies former staff members based on status flags and login dates
- Finds all involvements where former staff are still listed in NotifyIds
- Provides direct links to edit involvement notification settings
- Allows export to CSV for bulk processing

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps.
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python "InvolvementNotificationAuditor" and paste all this code
4. Test
--Upload Instructions End--

Written By: Ben Swaby
Email: bswaby@fbchtn.org
"""

import json
import traceback

class NotificationAuditor:
    """Class to handle finding and displaying former staff notifications"""
    
    def __init__(self, days_inactive=30, include_inactive_orgs=False, identification_method="all", org_ids=None, staff_ids=None):
        """Initialize with filtering parameters"""
        self.days_inactive = int(days_inactive)
        self.include_inactive_orgs = bool(include_inactive_orgs)
        self.identification_method = str(identification_method)
        self.org_ids = org_ids if org_ids else []
        self.staff_ids = staff_ids if staff_ids else []
        self.former_staff = []
        self.audit_results = []
    
    def get_former_staff_members(self):
        """
        Identify former staff members based on the selected identification method
        Returns all users who haven't logged in for the specified days
        """
        # This is a more detailed query to get all login accounts for each person
        sql = """
        SELECT 
            p.PeopleId, 
            p.Name,
            p.EmailAddress,
            u.UserId,
            u.Username,
            u.LastLoginDate,
            u.LastActivityDate,
            CASE 
                WHEN u.LastLoginDate IS NULL THEN 999999 -- Treat NULL as a very large number of days
                ELSE DATEDIFF(day, u.LastLoginDate, GETDATE()) 
            END AS DaysSinceLogin
        FROM 
            People p
            JOIN Users u ON p.PeopleId = u.PeopleId
        WHERE 
            CASE 
                WHEN u.LastLoginDate IS NULL THEN 999999 -- Treat NULL as a very large number of days
                ELSE DATEDIFF(day, u.LastLoginDate, GETDATE()) 
            END > {0}
            AND u.Username IS NOT NULL
        ORDER BY 
            p.Name, u.LastLoginDate DESC
        """.format(self.days_inactive)
        
        print '<!-- DEBUG: get_former_staff_members SQL: {0} -->'.format(sql)
        
        self.former_staff = q.QuerySql(sql)
        return self.former_staff
        
    def get_involvement_notifications(self):
        """
        Find all involvements where former staff members are listed in NotifyIds
        """
        # Build a SQL WHERE clause for organization status
        org_status_filter = ""
        if not self.include_inactive_orgs:
            org_status_filter = "AND o.OrganizationStatusId = 30 /* Active organizations only */"
        
        # For people_list method, directly search for those IDs in NotifyIds
        notifyids_filter = ""
        if (self.identification_method == "people_list" or self.identification_method == "organization") and self.staff_ids:
            # Create conditions for each People ID to match in NotifyIds - using simpler LIKE approach
            notifyids_conditions = []
            for staff_id in self.staff_ids:
                # Just use a simple LIKE with % on both sides - most reliable approach
                notifyids_conditions.append("o.NotifyIds LIKE '%{0}%'".format(staff_id))
            
            if notifyids_conditions:
                notifyids_filter = "AND ({0})".format(" OR ".join(notifyids_conditions))
        
        print "<!-- DEBUG: NotifyIds filter: {0} -->".format(notifyids_filter)
        
        # Build the base SQL to get all organizations with NotifyIds
        sql = """
        SELECT 
            o.OrganizationId,
            o.OrganizationName,
            o.NotifyIds,
            o.CreatedDate,
            o.MemberCount,
            o.OrganizationStatusId,
            CASE WHEN o.OrganizationStatusId = 30 THEN 'Active' ELSE 'Inactive' END AS Status,
            p.Program,
            d.Name AS Division,
            DATEDIFF(day, ISNULL(m.MeetingDate, o.CreatedDate), GETDATE()) AS DaysSinceLastMeeting
        FROM 
            Organizations o
            LEFT JOIN OrganizationStructure p ON o.OrganizationId = p.OrgId
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN (
                SELECT OrganizationId, MAX(MeetingDate) AS MeetingDate
                FROM Meetings
                GROUP BY OrganizationId
            ) m ON o.OrganizationId = m.OrganizationId
        WHERE 
            o.NotifyIds IS NOT NULL
            AND o.NotifyIds <> ''
            {0}
            {1}
        ORDER BY
            o.OrganizationName
        """.format(org_status_filter, notifyids_filter)
        
        print  '<!-- DEBUG: get_involvement_notifications SQL: {0} -->'.format(sql)
        
        involvements = q.QuerySql(sql)
        
        # Process each involvement to check its NotifyIds
        results = []
        
        # Process based on identification method
        if self.identification_method == "organization" and self.org_ids:
            # Get all members of the specified organizations
            org_ids_str = ",".join([str(oid) for oid in self.org_ids])
            
            print "<!-- DEBUG: Using org_ids: {0} -->".format(org_ids_str)
            
            # This query gets all active members of the specified organizations
            members_sql = """
            SELECT DISTINCT
                p.PeopleId,
                p.Name,
                p.EmailAddress
            FROM 
                OrganizationMembers om
                JOIN People p ON om.PeopleId = p.PeopleId
            WHERE 
                om.OrganizationId IN ({0})
            ORDER BY
                p.Name
            """.format(org_ids_str)
            
            print  '<!-- DEBUG: members_sql: {0} -->'.format(members_sql)
            
            current_members = q.QuerySql(members_sql)
            
            # Also get their login information
            member_ids = [member.PeopleId for member in current_members]
            if member_ids:
                member_ids_str = ",".join([str(mid) for mid in member_ids])
                
                users_sql = """
                SELECT 
                    u.PeopleId,
                    u.UserId,
                    u.Username,
                    u.LastLoginDate,
                    u.LastActivityDate,
                    CASE 
                        WHEN u.LastLoginDate IS NULL THEN 999999 -- Treat NULL as a very large number of days
                        ELSE DATEDIFF(day, u.LastLoginDate, GETDATE()) 
                    END AS DaysSinceLogin
                FROM 
                    Users u
                WHERE 
                    u.PeopleId IN ({0})
                    AND CASE 
                        WHEN u.LastLoginDate IS NULL THEN 999999
                        ELSE DATEDIFF(day, u.LastLoginDate, GETDATE())
                    END > {1}
                ORDER BY
                    u.PeopleId, u.LastLoginDate DESC
                """.format(member_ids_str, self.days_inactive)
                
                print "<!-- DEBUG: User login query: {0} -->".format(users_sql)
                users_data = q.QuerySql(users_sql)
                
                # Create a set of users that match our filter
                filtered_user_ids = set()
                for user in users_data:
                    filtered_user_ids.add(user.PeopleId)
                
                print "<!-- DEBUG: Found {0} users that match the login filter -->".format(len(filtered_user_ids))
                
                # Create a dictionary of user login data by PeopleId
                user_login_dict = {}
                for user in users_data:
                    if user.PeopleId not in user_login_dict:
                        user_login_dict[user.PeopleId] = []
                    user_login_dict[user.PeopleId].append(user)
                
                # Filter current_members to only those in filtered_user_ids
                filtered_members = []
                for member in current_members:
                    if member.PeopleId in filtered_user_ids:
                        # Attach the login data
                        if member.PeopleId in user_login_dict:
                            member.login_data = user_login_dict[member.PeopleId]
                        else:
                            member.login_data = []
                        filtered_members.append(member)
                
                current_members = filtered_members
            
            # Now check each involvement against these members
            for involvement in involvements:
                if not involvement.NotifyIds:
                    continue
                
                # Parse the NotifyIds into a list of integers
                notify_ids = []
                for id_str in involvement.NotifyIds.split(','):
                    id_str = id_str.strip()
                    if id_str and id_str.isdigit():
                        notify_ids.append(int(id_str))
                
                # Find staff members from our organization in the NotifyIds
                staff_in_notify = []
                for member in current_members:
                    if member.PeopleId in notify_ids:
                        staff_in_notify.append(member)
                
                if staff_in_notify:
                    # Calculate updated NotifyIds without staff
                    updated_notify_ids = []
                    for notify_id in notify_ids:
                        if not any(member.PeopleId == notify_id for member in staff_in_notify):
                            updated_notify_ids.append(str(notify_id))
                    
                    updated_notify_ids_str = ','.join(updated_notify_ids) if updated_notify_ids else ''
                    
                    results.append({
                        'involvement': involvement,
                        'former_staff': staff_in_notify,
                        'current_notify_ids': involvement.NotifyIds,
                        'updated_notify_ids': updated_notify_ids_str
                    })
                    
        elif self.identification_method == "people_list" and self.staff_ids:
            # Direct search for people IDs in NotifyIds
            staff_ids_str = ",".join([str(pid) for pid in self.staff_ids])
            
            # Get information about these people
            people_sql = """
            SELECT 
                p.PeopleId,
                p.Name,
                p.EmailAddress
            FROM 
                People p
            WHERE 
                p.PeopleId IN ({0})
            ORDER BY
                p.Name
            """.format(staff_ids_str)
            
            staff_people = q.QuerySql(people_sql)
            
            # Get login information
            users_sql = """
            SELECT 
                u.PeopleId,
                u.UserId,
                u.Username,
                u.LastLoginDate,
                u.LastActivityDate,
                DATEDIFF(day, ISNULL(u.LastLoginDate, u.CreationDate), GETDATE()) AS DaysSinceLogin
            FROM 
                Users u
            WHERE 
                u.PeopleId IN ({0})
            ORDER BY
                u.PeopleId, u.LastLoginDate DESC
            """.format(staff_ids_str)
            
            users_data = q.QuerySql(users_sql)
            
            # Create a dictionary of user login data by PeopleId
            user_login_dict = {}
            for user in users_data:
                if user.PeopleId not in user_login_dict:
                    user_login_dict[user.PeopleId] = []
                user_login_dict[user.PeopleId].append(user)
            
            # Now attach the login data to each staff person
            for person in staff_people:
                if person.PeopleId in user_login_dict:
                    person.login_data = user_login_dict[person.PeopleId]
                else:
                    person.login_data = []
            
            # Now check each involvement
            for involvement in involvements:
                if not involvement.NotifyIds:
                    continue
                
                # Parse the NotifyIds into a list of integers
                notify_ids = []
                for id_str in involvement.NotifyIds.split(','):
                    id_str = id_str.strip()
                    if id_str and id_str.isdigit():
                        notify_ids.append(int(id_str))
                
                # Find staff members in the NotifyIds
                staff_in_notify = []
                for person in staff_people:
                    if person.PeopleId in notify_ids:
                        staff_in_notify.append(person)
                
                if staff_in_notify:
                    # Calculate updated NotifyIds without staff
                    updated_notify_ids = []
                    for notify_id in notify_ids:
                        if not any(person.PeopleId == notify_id for person in staff_in_notify):
                            updated_notify_ids.append(str(notify_id))
                    
                    updated_notify_ids_str = ','.join(updated_notify_ids) if updated_notify_ids else ''
                    
                    results.append({
                        'involvement': involvement,
                        'former_staff': staff_in_notify,
                        'current_notify_ids': involvement.NotifyIds,
                        'updated_notify_ids': updated_notify_ids_str
                    })
        
        else:  # Default "all" method - find users who haven't logged in for a while
            # Group staff by PeopleId (to handle multiple login accounts)
            staff_by_id = {}
            for person in self.former_staff:
                if person.PeopleId not in staff_by_id:
                    staff_by_id[person.PeopleId] = {
                        'PeopleId': person.PeopleId,
                        'Name': person.Name,
                        'EmailAddress': person.EmailAddress,
                        'logins': []
                    }
                
                staff_by_id[person.PeopleId]['logins'].append({
                    'UserId': person.UserId,
                    'Username': person.Username,
                    'LastLoginDate': person.LastLoginDate,
                    'LastActivityDate': person.LastActivityDate,
                    'DaysSinceLogin': person.DaysSinceLogin
                })
            
            # Check for each involvement
            for involvement in involvements:
                if not involvement.NotifyIds:
                    continue
                
                # Parse the NotifyIds
                notify_ids = []
                for id_str in involvement.NotifyIds.split(','):
                    id_str = id_str.strip()
                    if id_str and id_str.isdigit():
                        notify_ids.append(int(id_str))
                
                # Find former staff in the NotifyIds
                former_staff_in_notify = []
                for staff_id, staff_data in staff_by_id.items():
                    if staff_id in notify_ids:
                        former_staff_in_notify.append(staff_data)
                
                if former_staff_in_notify:
                    # Calculate updated NotifyIds without former staff
                    updated_notify_ids = []
                    for notify_id in notify_ids:
                        if notify_id not in staff_by_id:
                            updated_notify_ids.append(str(notify_id))
                    
                    updated_notify_ids_str = ','.join(updated_notify_ids) if updated_notify_ids else ''
                    
                    results.append({
                        'involvement': involvement,
                        'former_staff': former_staff_in_notify,
                        'current_notify_ids': involvement.NotifyIds,
                        'updated_notify_ids': updated_notify_ids_str
                    })
                
        self.audit_results = results
        return results

class NotificationAuditorUI:
    """Class to handle the UI for the audit tool"""
    
    @staticmethod
    def print_header():
        """Print the page header and form"""
        include_inactive = "checked" if hasattr(model.Data, 'include_inactive') and model.Data.include_inactive else ""
        identification_method = getattr(model.Data, 'identification_method', 'all')
        
        print """
        <div class="container-fluid">
            <div class="row">
                <div class="col-md-12">
                    <h1><i class="fa fa-bell"></i> Involvement Notification Audit Tool<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                  </svg></h1>
                    <p class="lead">Identify involvements where former staff or inactive users may still be receiving notifications</p>
                    <hr>
                </div>
            </div>
            
            <form id="auditForm" class="form">
                <div class="row">
                    <div class="col-md-6">
                        <div class="form-group">
                            <label for="identification_method">Identification Method:</label>
                            <select class="form-control" name="identification_method" id="identification_method">
                                <option value="all" {0}>All Users (based on login date only)</option>
                                <option value="organization" {1}>Organization Members (filter by organization)</option>
                                <option value="people_list" {2}>People IDs (filter by specific People IDs)</option>
                            </select>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label for="days_inactive">Min Days Since Last Login (Only for All Users Search):</label>
                            <input type="number" class="form-control" name="days_inactive" id="days_inactive" 
                                placeholder="Enter minimum days inactive (e.g. 30)" 
                                value="{3}">
                            <small class="form-text text-muted">Users who haven't logged in for at least this many days will be considered inactive</small>
                        </div>
                    </div>
                </div>
                
                <div class="row" id="org_ids_section" style="display: {4};">
                    <div class="col-md-12">
                        <div class="form-group">
                            <label for="org_ids">Current Staff Organization IDs:</label>
                            <textarea class="form-control" name="org_ids" id="org_ids" 
                                placeholder="Enter Organization IDs for staff organizations, comma separated (e.g. 101,102,103)" 
                                rows="3">{5}</textarea>
                            <small class="form-text text-muted">Enter the Organization IDs that contain current staff members (comma separated)</small>
                        </div>
                    </div>
                </div>
                
                <div class="row" id="people_ids_section" style="display: {6};">
                    <div class="col-md-12">
                        <div class="form-group">
                            <label for="staff_ids">Current Staff People IDs:</label>
                            <textarea class="form-control" name="staff_ids" id="staff_ids" 
                                placeholder="Enter People IDs for current staff, comma separated (e.g. 1001,1002,1003)" 
                                rows="3">{7}</textarea>
                            <small class="form-text text-muted">Enter the PeopleIDs of current staff members (comma separated)</small>
                        </div>
                    </div>
                </div>
                
                <div class="row">
                    <div class="col-md-12">
                        <div class="form-group">
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input" name="include_inactive" id="include_inactive" {8}>
                                <label class="form-check-label" for="include_inactive">Include inactive involvements</label>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="row">
                    <div class="col-md-12">
                        <div class="form-group">
                            <button type="submit" name="submit" value="Run Report" class="btn btn-primary">
                                <i class="fa fa-search"></i> Run Report
                            </button>
                        </div>
                    </div>
                </div>
            </form>
            <div id="loading" style="display:none;" class="text-center">
                <p><i class="fa fa-spinner fa-spin fa-3x"></i></p>
                <p>Processing... please wait</p>
            </div>
        """.format(
            'selected' if identification_method == 'all' else '',
            'selected' if identification_method == 'organization' else '',
            'selected' if identification_method == 'people_list' else '',
            getattr(model.Data, 'days_inactive', '30'),
            'block' if identification_method == 'organization' else 'none',
            getattr(model.Data, 'org_ids', ''),
            'block' if identification_method == 'people_list' else 'none',
            getattr(model.Data, 'staff_ids', ''),
            include_inactive
        )
        
        # Add JavaScript to toggle form sections based on identification method
        print """
        <script type="text/javascript">
            (function() {
                // Toggle form sections based on identification method
                var identificationMethod = document.getElementById('identification_method');
                var orgIdsSection = document.getElementById('org_ids_section');
                var peopleIdsSection = document.getElementById('people_ids_section');
                var orgIds = document.getElementById('org_ids');
                var staffIds = document.getElementById('staff_ids');
                
                function updateSections() {
                    var method = identificationMethod.value;
                    
                    // Hide all sections first
                    orgIdsSection.style.display = 'none';
                    peopleIdsSection.style.display = 'none';
                    
                    // Show the appropriate section
                    if (method === 'organization') {
                        orgIdsSection.style.display = 'block';
                        if (orgIds) orgIds.focus();
                    } else if (method === 'people_list') {
                        peopleIdsSection.style.display = 'block';
                        if (staffIds) staffIds.focus();
                    }
                }
                
                // Add change event listener
                if (identificationMethod) {
                    identificationMethod.addEventListener('change', updateSections);
                }
                
                // Initialize on page load
                updateSections();
                
                // Handle the form submission 
                var auditForm = document.getElementById('auditForm');
                var loadingIndicator = document.getElementById('loading');
                
                if (auditForm) {
                    auditForm.addEventListener('submit', function(e) {
                        if (loadingIndicator) loadingIndicator.style.display = 'block';
                    });
                }
            })();
        </script>
        """
    
    @staticmethod
    def print_results(auditor):
        """Display the results of the audit"""
        if not auditor.audit_results:
            print """
                <div class="alert alert-info">
                    <i class="fa fa-info-circle"></i> 
                    No involvements found with notifications going to former staff members.
                </div>
            """
            return
            
        # Generate CSV data for export
        csv_data = [
            "Involvement ID,Involvement Name,Program,Division,Status,Member Count,Days Since Last Meeting,Former Staff IDs,Current NotifyIds,Updated NotifyIds"
        ]
        
        for result in auditor.audit_results:
            involvement = result['involvement']
            former_staff = result['former_staff']
            
            # For CSV generation, handle different structures based on identification method
            if auditor.identification_method == "all":
                # Get comma-separated list of former staff IDs for CSV
                former_staff_ids = ",".join([str(person['PeopleId']) for person in former_staff])
            else:
                # For other methods
                former_staff_ids = ",".join([str(person.PeopleId) for person in former_staff])
            
            # Add a row to CSV data
            csv_data.append('{0},"{1}","{2}","{3}",{4},{5},{6},"{7}","{8}","{9}"'.format(
                involvement.OrganizationId,
                involvement.OrganizationName.replace('"', '""'),  # Escape quotes
                involvement.Program.replace('"', '""') if involvement.Program else "",
                involvement.Division.replace('"', '""') if involvement.Division else "",
                involvement.Status,
                involvement.MemberCount,
                involvement.DaysSinceLastMeeting,
                former_staff_ids,
                result['current_notify_ids'],
                result['updated_notify_ids']
            ))
        
        # Convert CSV data to a single string
        csv_content = "\n".join(csv_data)
        
        print """
            <div class="row mt-4">
                <div class="col-md-12">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h2>Audit Results</h2>
                        <div>
                            <button class="btn btn-success" id="exportCsv">
                                <i class="fa fa-download"></i> Export to CSV
                            </button>
                        </div>
                    </div>
                    
                    <p>Found {0} involvements where former staff may be receiving notifications</p>
                    
                    <div class="table-responsive">
                        <table class="table table-striped table-hover" id="resultsTable">
                            <thead>
                                <tr>
                                    <th>Involvement</th>
                                    <th>Program / Division</th>
                                    <th>Status</th>
                                    <th>Members</th>
                                    <th>Last Activity<br><small class="text-muted">(days since last meeting)</small></th>
                                    <th>Former Staff Receiving Notifications</th>
                                    <th>All Notify IDs<br><small class="text-muted">(Current & Suggested Update)</small></th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
        """.format(len(auditor.audit_results))
        
        # Loop through each involvement and display details
        for result in auditor.audit_results:
            involvement = result['involvement']
            former_staff = result['former_staff']
            
            # Format the list of former staff members - handling different structures
            staff_list = ""
            
            if auditor.identification_method == "all":
                # The 'all' method groups by person with multiple logins
                for person in former_staff:
                    staff_list += """
                        <div class="mb-2 border-bottom pb-1">
                            <a href="{0}/Person2/{1}" target="_blank">{2}</a> 
                            <small class="text-muted">({3})</small>
                    """.format(
                        model.CmsHost,
                        person['PeopleId'],
                        person['Name'],
                        person['EmailAddress'] if person['EmailAddress'] else 'No email'
                    )
                    
                    # Add login information for each account
                    for login in person['logins']:
                        login_date_str = "Never"
                        if login['LastLoginDate']:
                            try:
                                login_date_str = login['LastLoginDate'].strftime('%Y-%m-%d %H:%M:%S')
                            except:
                                login_date_str = str(login['LastLoginDate'])
                        
                        staff_list += """
                            <div class="ml-3 small">
                                Login: <span class="text-monospace">{0}</span><br>
                                Last login: {1} ({2} days ago)
                            </div>
                        """.format(
                            login['Username'] if login['Username'] else 'No username',
                            login_date_str,
                            login['DaysSinceLogin']
                        )
                    
                    staff_list += "</div>"
            
            else:
                # The 'organization' and 'people_list' methods
                for person in former_staff:
                    staff_list += """
                        <div class="mb-2 border-bottom pb-1">
                            <a href="{0}/Person2/{1}" target="_blank">{2}</a> 
                            <small class="text-muted">({3})</small>
                    """.format(
                        model.CmsHost,
                        person.PeopleId,
                        person.Name,
                        person.EmailAddress if hasattr(person, 'EmailAddress') and person.EmailAddress else 'No email'
                    )
                    
                    # Check if this person has login data
                    if hasattr(person, 'login_data') and person.login_data:
                        for login in person.login_data:
                            login_date_str = "Never"
                            if hasattr(login, 'LastLoginDate') and login.LastLoginDate:
                                try:
                                    login_date_str = login.LastLoginDate.strftime('%Y-%m-%d %H:%M:%S')
                                except:
                                    login_date_str = str(login.LastLoginDate)
                            
                            staff_list += """
                                <div class="ml-3 small">
                                    Login: <span class="text-monospace">{0}</span><br>
                                    Last login: {1} ({2} days ago)
                                </div>
                            """.format(
                                login.Username if hasattr(login, 'Username') and login.Username else 'No username',
                                login_date_str,
                                login.DaysSinceLogin if hasattr(login, 'DaysSinceLogin') else 'unknown'
                            )
                    else:
                        staff_list += """
                            <div class="ml-3 small">
                                <em>No login information available</em>
                            </div>
                        """
                    
                    staff_list += "</div>"
            
            # Format Program/Division info
            prog_div = ""
            if involvement.Program:
                prog_div += """<div>{0}</div>""".format(involvement.Program)
            if involvement.Division:
                prog_div += """<div><small>{0}</small></div>""".format(involvement.Division)
            
            # Print the row for this involvement
            print """
                <tr>
                    <td>
                        <a href="{0}/Org/{1}" target="_blank">{2}</a>
                    </td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6}</td>
                    <td>{7}</td>
                    <td>
                        <div><strong>Current:</strong> <code>{8}</code></div>
                        <div class="mt-1"><strong>Updated:</strong> <code>{9}</code></div>
                    </td>
                    <td>
                        <a href="{0}/Org/{1}#options" target="_blank" class="btn btn-sm btn-primary">
                            <i class="fa fa-edit"></i> Edit
                        </a>
                    </td>
                </tr>
            """.format(
                model.CmsHost,
                involvement.OrganizationId,
                involvement.OrganizationName,
                prog_div,
                involvement.Status,
                involvement.MemberCount,
                involvement.DaysSinceLastMeeting,
                staff_list,
                result['current_notify_ids'],
                result['updated_notify_ids']
            )
        
        print """
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        """
        
        # Add CSV export script with proper JSON handling
        csv_json = json.dumps(csv_content)
        print """
            <script>
                // Function to download CSV data
                document.getElementById('exportCsv').addEventListener('click', function() {
                    var csvContent = %s;
                    var encodedUri = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvContent);
                    var link = document.createElement('a');
                    link.setAttribute('href', encodedUri);
                    link.setAttribute('download', 'involvement_notification_audit.csv');
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                });
            </script>
        """ % csv_json
    
    @staticmethod
    def print_former_staff_summary(former_staff):
        """Display a summary of former staff members identified"""
        if not former_staff or len(former_staff) == 0:
            print """
                <div class="alert alert-info">
                    <i class="fa fa-info-circle"></i> 
                    No former staff or inactive users found based on your criteria.
                </div>
            """
            return
            
        print """
            <div class="row mt-4">
                <div class="col-md-12">
                    <h3>Inactive Users Summary</h3>
                    <p>Found {0} inactive users based on your criteria</p>
                    
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>Name</th>
                                    <th>Email</th>
                                    <th>Username</th>
                                    <th>Last Login</th>
                                    <th>Days Since Login</th>
                                </tr>
                            </thead>
                            <tbody>
        """.format(len(former_staff))
        
        for person in former_staff:
            login_date = "Never"
            days_since = "N/A"
            username = "-"
            
            if hasattr(person, 'LastLoginDate') and person.LastLoginDate:
                try:
                    login_date = person.LastLoginDate.strftime('%Y-%m-%d')
                except:
                    login_date = str(person.LastLoginDate)
            
            if hasattr(person, 'DaysSinceLogin') and person.DaysSinceLogin:
                days_since = str(person.DaysSinceLogin)
                
            if hasattr(person, 'Username') and person.Username:
                username = person.Username
                
            print """
                <tr>
                    <td><a href="{0}/Person2/{1}" target="_blank">{2}</a></td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6}</td>
                </tr>
            """.format(
                model.CmsHost,
                person.PeopleId,
                person.Name if hasattr(person, 'Name') else "Unknown",
                person.EmailAddress if hasattr(person, 'EmailAddress') and person.EmailAddress else '-',
                username,
                login_date,
                days_since
            )
            
        print """
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        """
    
    @staticmethod
    def print_footer():
        """Print the page footer"""
        print """
            </div>
            <script>
                $(function() {
                    // Function to convert path from /PyScript/ to /PyScriptForm/
                    function getPyScriptAddress() {
                        let path = window.location.pathname;
                        return path.replace("/PyScript/", "/PyScriptForm/");
                    }
                    
                    // Handle the form submission with AJAX
                    $('#auditForm').on('submit', function(e) {
                        e.preventDefault();
                        $('#loading').show();
                        
                        $.ajax({
                            url: getPyScriptAddress(),
                            type: 'POST',
                            data: $(this).serialize() + '&submit=Run+Report',
                            success: function(response) {
                                // Replace the current page content with the response
                                document.open();
                                document.write(response);
                                document.close();
                            },
                            error: function() {
                                $('#loading').hide();
                                alert('An error occurred while processing your request.');
                            }
                        });
                    });
                });
            </script>
        """

# Main execution logic
try:
    # Check if running as a page or from a form submission
    is_form_submission = hasattr(model.Data, 'submit') and model.Data.submit == "Run Report"
    
    # Print the page header and form
    NotificationAuditorUI.print_header()
    
    if is_form_submission:
        # Handle the days_inactive parameter safely
        try:
            days_inactive_str = str(getattr(model.Data, 'days_inactive', '30'))
            days_inactive = int(days_inactive_str) if days_inactive_str.strip() else 30
        except ValueError:
            days_inactive = 30  # Default to 30 if conversion fails
        
        # Handle the include_inactive checkbox properly
        include_inactive = False
        if hasattr(model.Data, 'include_inactive'):
            if model.Data.include_inactive == "on" or model.Data.include_inactive == "true" or model.Data.include_inactive == True:
                include_inactive = True
        
        # Get the identification method
        identification_method = getattr(model.Data, 'identification_method', 'all')
        
        # Initialize parameters for the auditor
        org_ids = []
        staff_ids = []
        
        # Parse organization IDs if that method is selected
        if identification_method == 'organization':
            org_ids_str = str(getattr(model.Data, 'org_ids', ''))
            if org_ids_str.strip():
                for id_str in org_ids_str.split(','):
                    id_str = id_str.strip()
                    if id_str and id_str.isdigit():
                        org_ids.append(int(id_str))
                        
            # Display info about the organizations searched
            if org_ids:
                org_names_sql = """
                SELECT o.OrganizationId, o.OrganizationName 
                FROM Organizations o 
                WHERE o.OrganizationId IN ({0})
                """.format(','.join([str(oid) for oid in org_ids]))
                
                try:
                    org_names = q.QuerySql(org_names_sql)
                    org_names_list = ["{0} (ID: {1})".format(org.OrganizationName, org.OrganizationId) for org in org_names]
                    
                    print """
                        <div class="alert alert-info">
                            <i class="fa fa-info-circle"></i> 
                            Searching for members of: {0}
                        </div>
                    """.format(', '.join(org_names_list))
                except:
                    print """
                        <div class="alert alert-info">
                            <i class="fa fa-info-circle"></i> 
                            Searching for members of Organization IDs: {0}
                        </div>
                    """.format(', '.join(map(str, org_ids)))
        
        # Parse staff IDs if that method is selected
        if identification_method == 'people_list':
            staff_ids_str = str(getattr(model.Data, 'staff_ids', ''))
            if staff_ids_str.strip():
                for id_str in staff_ids_str.split(','):
                    id_str = id_str.strip()
                    if id_str and id_str.isdigit():
                        staff_ids.append(int(id_str))
                        
            # Display info about the people IDs searched
            if staff_ids:
                # Try to get names for the IDs
                people_names_sql = """
                SELECT p.PeopleId, p.Name 
                FROM People p 
                WHERE p.PeopleId IN ({0})
                """.format(','.join([str(pid) for pid in staff_ids]))
                
                try:
                    people_names = q.QuerySql(people_names_sql)
                    people_names_list = ["{0} (ID: {1})".format(p.Name, p.PeopleId) for p in people_names]
                    
                    print """
                        <div class="alert alert-info">
                            <i class="fa fa-info-circle"></i> 
                            Searching for: {0} in NotifyIds
                        </div>
                    """.format(', '.join(people_names_list))
                except:
                    print """
                        <div class="alert alert-info">
                            <i class="fa fa-info-circle"></i> 
                            Searching for People IDs: {0} in NotifyIds
                        </div>
                    """.format(', '.join(map(str, staff_ids)))
        
        # Create the auditor with all parameters
        auditor = NotificationAuditor(
            days_inactive=days_inactive,
            include_inactive_orgs=include_inactive,
            identification_method=identification_method,
            org_ids=org_ids,
            staff_ids=staff_ids
        )
        
        # First get all users who haven't logged in recently - only needed for 'all' method
        # For other methods, we directly search based on IDs
        if identification_method == 'all':
            former_staff = auditor.get_former_staff_members()
        elif identification_method == 'organization':
            # For organization method, get all members of the specified orgs
            if org_ids:
                # Build NotifyIds filter for all members of the organization
                org_member_ids_sql = """
                SELECT DISTINCT om.PeopleId
                FROM OrganizationMembers om
                WHERE om.OrganizationId IN ({0})
                """.format(','.join([str(oid) for oid in org_ids]))
                
                members = q.QuerySql(org_member_ids_sql)
                member_ids = [member.PeopleId for member in members]
                
                # Convert member IDs to staff_ids for use in the notification search
                auditor.staff_ids = member_ids
                auditor.identification_method = "people_list"  # Use the people_list method logic
                
                # Get people info
                if member_ids:
                    member_ids_str = ",".join([str(pid) for pid in member_ids])
                    sql = """
                    SELECT 
                        p.PeopleId, 
                        p.Name,
                        p.EmailAddress,
                        u.UserId,
                        u.Username,
                        u.LastLoginDate,
                        u.LastActivityDate,
                        DATEDIFF(day, ISNULL(u.LastLoginDate, u.CreationDate), GETDATE()) AS DaysSinceLogin
                    FROM 
                        People p
                        LEFT JOIN Users u ON p.PeopleId = u.PeopleId
                    WHERE 
                        p.PeopleId IN ({0})
                    """.format(member_ids_str)
                    auditor.former_staff = q.QuerySql(sql)
                else:
                    auditor.former_staff = []
            else:
                auditor.former_staff = []
        else:  # 'people_list'
            # For 'people_list', just run a query to get basic info about the people IDs
            if staff_ids:
                staff_ids_str = ",".join([str(pid) for pid in staff_ids])
                sql = """
                SELECT 
                    p.PeopleId, 
                    p.Name,
                    p.EmailAddress,
                    u.UserId,
                    u.Username,
                    u.LastLoginDate,
                    u.LastActivityDate,
                    CASE 
                        WHEN u.LastLoginDate IS NULL THEN 999999 -- Treat NULL as a very large number of days
                        ELSE DATEDIFF(day, u.LastLoginDate, GETDATE()) 
                    END AS DaysSinceLogin
                FROM 
                    People p
                    LEFT JOIN Users u ON p.PeopleId = u.PeopleId
                WHERE 
                    p.PeopleId IN ({0})
                    AND (u.PeopleId IS NULL OR 
                        CASE 
                            WHEN u.LastLoginDate IS NULL THEN 999999 -- Treat NULL as a very large number of days
                            ELSE DATEDIFF(day, u.LastLoginDate, GETDATE()) 
                        END > {1}
                    )
                ORDER BY
                    p.Name, u.LastLoginDate DESC
                """.format(staff_ids_str, days_inactive)
                
                print '<!-- DEBUG: staff users SQL: {0} -->'.format(sql)
                auditor.former_staff = q.QuerySql(sql)
            else:
                auditor.former_staff = []
        
        # Run the audit and display results
        audit_results = auditor.get_involvement_notifications()
        NotificationAuditorUI.print_results(auditor)
        
        # Only display former staff summary if specifically requested or no results found
        if (not audit_results or len(audit_results) == 0) and identification_method == 'all' and auditor.former_staff and len(auditor.former_staff) > 0:
            NotificationAuditorUI.print_former_staff_summary(auditor.former_staff)
    
    # Print the page footer
    NotificationAuditorUI.print_footer()

except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
