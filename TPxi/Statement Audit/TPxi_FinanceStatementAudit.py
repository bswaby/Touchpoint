# TPxi Statement Audit Dashboard
# Analyzes electronic vs printed statement preferences with email/address validation
# Allows drill-down into people with missing contact info

# Created by: Ben Swaby
# Email: bswaby@fbchtn.org
# Created on: 2025-12-11
# Updated on: 2025-12-11

# --Upload Instructions Start--
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python "TPxi_StatementAudit" and paste all this code
# 4. Test and optionally add to menu
# --Upload Instructions End--

import json

# Get current year and build year options
current_year = model.DateTime.Now.Year
year_options = []
for y in range(current_year, current_year - 5, -1):
    year_options.append(y)

# Handle AJAX requests
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    if action == "get_data":
        selected_year = int(Data.year) if hasattr(Data, 'year') and Data.year else current_year
        start_date = "{0}-01-01".format(selected_year)
        end_date = "{0}-12-31".format(selected_year)

        # Get givers summary
        sql_givers = '''
        WITH Givers AS (
            SELECT DISTINCT PeopleId
            FROM Contribution
            WHERE ContributionDate >= '{0}' AND ContributionDate <= '{1}'
              AND ContributionTypeId NOT IN (6, 7, 8, 99)
              AND ContributionStatusId = 0
        )
        SELECT
            CASE WHEN p.ElectronicStatement = 1 THEN 'Electronic' ELSE 'Printed' END as StatementPref,
            CASE
                WHEN p.EmailAddress IS NOT NULL AND p.EmailAddress != '' THEN 'HasEmail'
                ELSE 'NoEmail'
            END as EmailStatus,
            COUNT(*) as Cnt
        FROM People p
        INNER JOIN Givers g ON p.PeopleId = g.PeopleId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
        GROUP BY
            CASE WHEN p.ElectronicStatement = 1 THEN 'Electronic' ELSE 'Printed' END,
            CASE WHEN p.EmailAddress IS NOT NULL AND p.EmailAddress != '' THEN 'HasEmail' ELSE 'NoEmail' END
        '''.format(start_date, end_date)

        givers_data = q.QuerySql(sql_givers)
        givers_result = {}
        for row in givers_data:
            key = "{0}_{1}".format(row.StatementPref, row.EmailStatus)
            givers_result[key] = row.Cnt

        # Get non-givers summary
        sql_nongivers = '''
        WITH Givers AS (
            SELECT DISTINCT PeopleId
            FROM Contribution
            WHERE ContributionDate >= '{0}' AND ContributionDate <= '{1}'
              AND ContributionTypeId NOT IN (6, 7, 8, 99)
              AND ContributionStatusId = 0
        )
        SELECT
            CASE WHEN p.ElectronicStatement = 1 THEN 'Electronic' ELSE 'Printed' END as StatementPref,
            CASE
                WHEN p.EmailAddress IS NOT NULL AND p.EmailAddress != '' THEN 'HasEmail'
                ELSE 'NoEmail'
            END as EmailStatus,
            COUNT(*) as Cnt
        FROM People p
        LEFT JOIN Givers g ON p.PeopleId = g.PeopleId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0 AND g.PeopleId IS NULL
        GROUP BY
            CASE WHEN p.ElectronicStatement = 1 THEN 'Electronic' ELSE 'Printed' END,
            CASE WHEN p.EmailAddress IS NOT NULL AND p.EmailAddress != '' THEN 'HasEmail' ELSE 'NoEmail' END
        '''.format(start_date, end_date)

        nongivers_data = q.QuerySql(sql_nongivers)
        nongivers_result = {}
        for row in nongivers_data:
            key = "{0}_{1}".format(row.StatementPref, row.EmailStatus)
            nongivers_result[key] = row.Cnt

        # Get address issues for printed statement givers
        sql_address = '''
        WITH Givers AS (
            SELECT DISTINCT PeopleId
            FROM Contribution
            WHERE ContributionDate >= '{0}' AND ContributionDate <= '{1}'
              AND ContributionTypeId NOT IN (6, 7, 8, 99)
              AND ContributionStatusId = 0
        )
        SELECT
            CASE
                WHEN (f.AddressLineOne IS NULL OR f.AddressLineOne = '') THEN 'MissingStreet'
                WHEN (f.CityName IS NULL OR f.CityName = '') THEN 'MissingCity'
                WHEN (f.StateCode IS NULL OR f.StateCode = '') THEN 'MissingState'
                WHEN (f.ZipCode IS NULL OR f.ZipCode = '') THEN 'MissingZip'
                ELSE 'Complete'
            END as AddressStatus,
            COUNT(*) as Cnt
        FROM People p
        INNER JOIN Givers g ON p.PeopleId = g.PeopleId
        LEFT JOIN Families f ON p.FamilyId = f.FamilyId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
          AND (p.ElectronicStatement = 0 OR p.ElectronicStatement IS NULL)
        GROUP BY
            CASE
                WHEN (f.AddressLineOne IS NULL OR f.AddressLineOne = '') THEN 'MissingStreet'
                WHEN (f.CityName IS NULL OR f.CityName = '') THEN 'MissingCity'
                WHEN (f.StateCode IS NULL OR f.StateCode = '') THEN 'MissingState'
                WHEN (f.ZipCode IS NULL OR f.ZipCode = '') THEN 'MissingZip'
                ELSE 'Complete'
            END
        '''.format(start_date, end_date)

        address_data = q.QuerySql(sql_address)
        address_result = {}
        for row in address_data:
            address_result[row.AddressStatus] = row.Cnt

        response = {
            'success': True,
            'year': selected_year,
            'givers': givers_result,
            'nongivers': nongivers_result,
            'address': address_result
        }
        print json.dumps(response)

    elif action == "get_electronic_no_email":
        selected_year = int(Data.year) if hasattr(Data, 'year') and Data.year else current_year
        start_date = "{0}-01-01".format(selected_year)
        end_date = "{0}-12-31".format(selected_year)

        sql = '''
        WITH Givers AS (
            SELECT DISTINCT PeopleId
            FROM Contribution
            WHERE ContributionDate >= '{0}' AND ContributionDate <= '{1}'
              AND ContributionTypeId NOT IN (6, 7, 8, 99)
              AND ContributionStatusId = 0
        )
        SELECT
            p.PeopleId,
            p.Name2 as Name,
            p.CellPhone,
            p.HomePhone,
            f.AddressLineOne as Address,
            f.CityName as City,
            f.StateCode as State,
            f.ZipCode as Zip
        FROM People p
        INNER JOIN Givers g ON p.PeopleId = g.PeopleId
        LEFT JOIN Families f ON p.FamilyId = f.FamilyId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
          AND p.ElectronicStatement = 1
          AND (p.EmailAddress IS NULL OR p.EmailAddress = '')
        ORDER BY p.Name2
        '''.format(start_date, end_date)

        results = q.QuerySql(sql)
        people = []
        for row in results:
            people.append({
                'id': row.PeopleId,
                'name': row.Name or '',
                'cell': row.CellPhone or '',
                'home': row.HomePhone or '',
                'address': row.Address or '',
                'city': row.City or '',
                'state': row.State or '',
                'zip': row.Zip or ''
            })

        print json.dumps({'success': True, 'people': people})

    elif action == "get_address_issues":
        selected_year = int(Data.year) if hasattr(Data, 'year') and Data.year else current_year
        issue_type = str(Data.issue_type) if hasattr(Data, 'issue_type') and Data.issue_type else 'all'
        start_date = "{0}-01-01".format(selected_year)
        end_date = "{0}-12-31".format(selected_year)

        # Build WHERE clause for issue type
        if issue_type == 'MissingStreet':
            issue_filter = "AND (f.AddressLineOne IS NULL OR f.AddressLineOne = '')"
        elif issue_type == 'MissingCity':
            issue_filter = "AND f.AddressLineOne IS NOT NULL AND f.AddressLineOne != '' AND (f.CityName IS NULL OR f.CityName = '')"
        elif issue_type == 'MissingState':
            issue_filter = "AND f.AddressLineOne IS NOT NULL AND f.AddressLineOne != '' AND f.CityName IS NOT NULL AND f.CityName != '' AND (f.StateCode IS NULL OR f.StateCode = '')"
        elif issue_type == 'MissingZip':
            issue_filter = "AND f.AddressLineOne IS NOT NULL AND f.AddressLineOne != '' AND f.CityName IS NOT NULL AND f.CityName != '' AND f.StateCode IS NOT NULL AND f.StateCode != '' AND (f.ZipCode IS NULL OR f.ZipCode = '')"
        else:
            # All issues
            issue_filter = '''AND (
                (f.AddressLineOne IS NULL OR f.AddressLineOne = '')
                OR (f.CityName IS NULL OR f.CityName = '')
                OR (f.StateCode IS NULL OR f.StateCode = '')
                OR (f.ZipCode IS NULL OR f.ZipCode = '')
            )'''

        sql = '''
        WITH Givers AS (
            SELECT DISTINCT PeopleId
            FROM Contribution
            WHERE ContributionDate >= '{0}' AND ContributionDate <= '{1}'
              AND ContributionTypeId NOT IN (6, 7, 8, 99)
              AND ContributionStatusId = 0
        )
        SELECT
            p.PeopleId,
            p.Name2 as Name,
            p.EmailAddress as Email,
            p.CellPhone,
            f.AddressLineOne as Address,
            f.CityName as City,
            f.StateCode as State,
            f.ZipCode as Zip,
            CASE
                WHEN (f.AddressLineOne IS NULL OR f.AddressLineOne = '') THEN 'Missing Street'
                WHEN (f.CityName IS NULL OR f.CityName = '') THEN 'Missing City'
                WHEN (f.StateCode IS NULL OR f.StateCode = '') THEN 'Missing State'
                WHEN (f.ZipCode IS NULL OR f.ZipCode = '') THEN 'Missing Zip'
                ELSE 'Unknown'
            END as Issue
        FROM People p
        INNER JOIN Givers g ON p.PeopleId = g.PeopleId
        LEFT JOIN Families f ON p.FamilyId = f.FamilyId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
          AND (p.ElectronicStatement = 0 OR p.ElectronicStatement IS NULL)
          {2}
        ORDER BY p.Name2
        '''.format(start_date, end_date, issue_filter)

        results = q.QuerySql(sql)
        people = []
        for row in results:
            people.append({
                'id': row.PeopleId,
                'name': row.Name or '',
                'email': row.Email or '',
                'cell': row.CellPhone or '',
                'address': row.Address or '',
                'city': row.City or '',
                'state': row.State or '',
                'zip': row.Zip or '',
                'issue': row.Issue or ''
            })

        print json.dumps({'success': True, 'people': people})

    else:
        print json.dumps({'success': False, 'error': 'Unknown action'})

else:
    # Build year options HTML
    year_options_html = ''
    for y in year_options:
        selected = ' selected' if y == current_year else ''
        year_options_html += '<option value="{0}"{1}>{0}</option>'.format(y, selected)

    # Main dashboard HTML
    model.Header = "Statement Audit Dashboard"

    html = '''
<!DOCTYPE html>
<html>
<head>
    <style>
        .dashboard-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        }
        .header-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 15px;
        }
        .year-selector {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .year-selector select {
            padding: 8px 15px;
            font-size: 16px;
            border: 1px solid #ccc;
            border-radius: 4px;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .summary-card {
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .summary-card h3 {
            margin: 0 0 15px 0;
            color: #333;
            font-size: 16px;
            border-bottom: 2px solid #007bff;
            padding-bottom: 10px;
        }
        .summary-card.issues h3 {
            border-bottom-color: #dc3545;
        }
        .stat-table {
            width: 100%;
            border-collapse: collapse;
        }
        .stat-table th, .stat-table td {
            padding: 8px 12px;
            text-align: left;
            border-bottom: 1px solid #eee;
        }
        .stat-table th {
            background: #f8f9fa;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            color: #666;
        }
        .stat-table td.number {
            text-align: right;
            font-weight: 500;
        }
        .stat-table tr.total {
            background: #f8f9fa;
            font-weight: 600;
        }
        .stat-table tr.clickable {
            cursor: pointer;
            transition: background 0.2s;
        }
        .stat-table tr.clickable:hover {
            background: #e3f2fd;
        }
        .stat-table tr.warning {
            background: #fff3cd;
        }
        .stat-table tr.warning:hover {
            background: #ffe69c;
        }
        .badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-success { background: #d4edda; color: #155724; }
        .badge-warning { background: #fff3cd; color: #856404; }
        .badge-danger { background: #f8d7da; color: #721c24; }
        .badge-info { background: #d1ecf1; color: #0c5460; }

        /* Modal styles */
        .modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 1000;
            justify-content: center;
            align-items: center;
        }
        .modal-overlay.active {
            display: flex;
        }
        .modal-content {
            background: #fff;
            border-radius: 8px;
            width: 90%;
            max-width: 1000px;
            max-height: 80vh;
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }
        .modal-header {
            padding: 15px 20px;
            background: #f8f9fa;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .modal-header h4 {
            margin: 0;
            font-size: 18px;
        }
        .modal-close {
            background: none;
            border: none;
            font-size: 24px;
            cursor: pointer;
            color: #666;
        }
        .modal-body {
            padding: 20px;
            overflow-y: auto;
            flex: 1;
        }
        .people-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }
        .people-table th, .people-table td {
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #eee;
        }
        .people-table th {
            background: #f8f9fa;
            font-weight: 600;
            position: sticky;
            top: 0;
        }
        .people-table tr:hover {
            background: #f5f5f5;
        }
        .people-table a {
            color: #007bff;
            text-decoration: none;
        }
        .people-table a:hover {
            text-decoration: underline;
        }
        .loading {
            text-align: center;
            padding: 40px;
            color: #666;
        }
        .no-data {
            text-align: center;
            padding: 40px;
            color: #999;
        }
        .key-issues {
            background: #fff3cd;
            border: 1px solid #ffc107;
            border-radius: 8px;
            padding: 15px 20px;
            margin-bottom: 20px;
        }
        .key-issues h4 {
            margin: 0 0 10px 0;
            color: #856404;
        }
        .key-issues ul {
            margin: 0;
            padding-left: 20px;
        }
        .key-issues li {
            margin-bottom: 5px;
            color: #856404;
        }
        .refresh-btn {
            padding: 8px 15px;
            background: #007bff;
            color: #fff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        .refresh-btn:hover {
            background: #0056b3;
        }
    </style>
</head>
<body>
    <div class="dashboard-container">
        <div class="header-row">
            <h2 style="margin: 0;">Statement Audit Dashboard</h2>
            <div class="year-selector">
                <label for="yearSelect"><strong>Giving Year:</strong></label>
                <select id="yearSelect" onchange="loadData()">
                    ''' + year_options_html + '''
                </select>
                <button class="refresh-btn" onclick="loadData()">Refresh</button>
            </div>
        </div>

        <div id="keyIssues" class="key-issues" style="display: none;">
            <h4>Key Issues to Address</h4>
            <ul id="issuesList"></ul>
        </div>

        <div class="summary-grid">
            <!-- Givers Card -->
            <div class="summary-card">
                <h3 id="giversTitle">Givers - Statement Preferences</h3>
                <table class="stat-table" id="giversTable">
                    <thead>
                        <tr>
                            <th>Statement Type</th>
                            <th>Email Status</th>
                            <th style="text-align:right">Count</th>
                        </tr>
                    </thead>
                    <tbody id="giversBody">
                        <tr><td colspan="3" class="loading">Loading...</td></tr>
                    </tbody>
                </table>
            </div>

            <!-- Non-Givers Card -->
            <div class="summary-card">
                <h3 id="nonGiversTitle">Non-Givers - Statement Preferences</h3>
                <table class="stat-table" id="nonGiversTable">
                    <thead>
                        <tr>
                            <th>Statement Type</th>
                            <th>Email Status</th>
                            <th style="text-align:right">Count</th>
                        </tr>
                    </thead>
                    <tbody id="nonGiversBody">
                        <tr><td colspan="3" class="loading">Loading...</td></tr>
                    </tbody>
                </table>
            </div>

            <!-- Address Issues Card -->
            <div class="summary-card issues">
                <h3 id="addressTitle">Printed Statement - Address Issues (Givers Only)</h3>
                <table class="stat-table" id="addressTable">
                    <thead>
                        <tr>
                            <th>Address Status</th>
                            <th style="text-align:right">Count</th>
                        </tr>
                    </thead>
                    <tbody id="addressBody">
                        <tr><td colspan="2" class="loading">Loading...</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Modal for drill-down -->
    <div class="modal-overlay" id="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h4 id="modalTitle">People List</h4>
                <button class="modal-close" onclick="closeModal()">&times;</button>
            </div>
            <div class="modal-body" id="modalBody">
                <div class="loading">Loading...</div>
            </div>
        </div>
    </div>

    <script>
        var currentYear = ''' + str(current_year) + ''';
        var scriptUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

        function loadData() {
            var year = document.getElementById('yearSelect').value;
            currentYear = year;

            // Update titles
            document.getElementById('giversTitle').textContent = year + ' Givers - Statement Preferences';
            document.getElementById('nonGiversTitle').textContent = year + ' Non-Givers - Statement Preferences';
            document.getElementById('addressTitle').textContent = year + ' Printed Statement - Address Issues (Givers Only)';

            // Show loading
            document.getElementById('giversBody').innerHTML = '<tr><td colspan="3" class="loading">Loading...</td></tr>';
            document.getElementById('nonGiversBody').innerHTML = '<tr><td colspan="3" class="loading">Loading...</td></tr>';
            document.getElementById('addressBody').innerHTML = '<tr><td colspan="2" class="loading">Loading...</td></tr>';
            document.getElementById('keyIssues').style.display = 'none';

            var formData = new FormData();
            formData.append('action', 'get_data');
            formData.append('year', year);

            fetch(scriptUrl, {
                method: 'POST',
                body: formData
            })
            .then(function(response) { return response.text(); })
            .then(function(text) {
                try {
                    var data = JSON.parse(text);
                    if (data.success) {
                        renderGivers(data.givers);
                        renderNonGivers(data.nongivers);
                        renderAddress(data.address);
                        updateKeyIssues(data);
                    } else {
                        alert('Error loading data');
                    }
                } catch(e) {
                    console.error('Parse error:', e, text);
                }
            })
            .catch(function(error) {
                console.error('Fetch error:', error);
            });
        }

        function renderGivers(data) {
            var html = '';
            var total = 0;

            // Electronic with email
            var elecEmail = data['Electronic_HasEmail'] || 0;
            var elecNoEmail = data['Electronic_NoEmail'] || 0;
            var printEmail = data['Printed_HasEmail'] || 0;
            var printNoEmail = data['Printed_NoEmail'] || 0;

            total = elecEmail + elecNoEmail + printEmail + printNoEmail;

            html += '<tr><td>Electronic</td><td><span class="badge badge-success">Has Email</span></td><td class="number">' + elecEmail.toLocaleString() + '</td></tr>';

            if (elecNoEmail > 0) {
                html += '<tr class="warning clickable" onclick="showElectronicNoEmail()"><td>Electronic</td><td><span class="badge badge-danger">No Email</span></td><td class="number">' + elecNoEmail.toLocaleString() + '</td></tr>';
            } else {
                html += '<tr><td>Electronic</td><td><span class="badge badge-info">No Email</span></td><td class="number">0</td></tr>';
            }

            html += '<tr><td>Printed</td><td><span class="badge badge-success">Has Email</span></td><td class="number">' + printEmail.toLocaleString() + '</td></tr>';
            html += '<tr><td>Printed</td><td><span class="badge badge-info">No Email</span></td><td class="number">' + printNoEmail.toLocaleString() + '</td></tr>';
            html += '<tr class="total"><td colspan="2"><strong>Total Givers</strong></td><td class="number"><strong>' + total.toLocaleString() + '</strong></td></tr>';

            document.getElementById('giversBody').innerHTML = html;
        }

        function renderNonGivers(data) {
            var html = '';
            var total = 0;

            var elecEmail = data['Electronic_HasEmail'] || 0;
            var elecNoEmail = data['Electronic_NoEmail'] || 0;
            var printEmail = data['Printed_HasEmail'] || 0;
            var printNoEmail = data['Printed_NoEmail'] || 0;

            total = elecEmail + elecNoEmail + printEmail + printNoEmail;

            html += '<tr><td>Electronic</td><td><span class="badge badge-success">Has Email</span></td><td class="number">' + elecEmail.toLocaleString() + '</td></tr>';
            html += '<tr><td>Electronic</td><td><span class="badge badge-info">No Email</span></td><td class="number">' + elecNoEmail.toLocaleString() + '</td></tr>';
            html += '<tr><td>Printed</td><td><span class="badge badge-success">Has Email</span></td><td class="number">' + printEmail.toLocaleString() + '</td></tr>';
            html += '<tr><td>Printed</td><td><span class="badge badge-info">No Email</span></td><td class="number">' + printNoEmail.toLocaleString() + '</td></tr>';
            html += '<tr class="total"><td colspan="2"><strong>Total Non-Givers</strong></td><td class="number"><strong>' + total.toLocaleString() + '</strong></td></tr>';

            document.getElementById('nonGiversBody').innerHTML = html;
        }

        function renderAddress(data) {
            var html = '';
            var issueTotal = 0;

            var complete = data['Complete'] || 0;
            var missingStreet = data['MissingStreet'] || 0;
            var missingCity = data['MissingCity'] || 0;
            var missingState = data['MissingState'] || 0;
            var missingZip = data['MissingZip'] || 0;

            issueTotal = missingStreet + missingCity + missingState + missingZip;

            html += '<tr><td><span class="badge badge-success">Complete Address</span></td><td class="number">' + complete.toLocaleString() + '</td></tr>';

            if (missingStreet > 0) {
                html += '<tr class="warning clickable" data-issue="MissingStreet" onclick="showAddressIssues(this.dataset.issue)"><td><span class="badge badge-danger">Missing Street Address</span></td><td class="number">' + missingStreet.toLocaleString() + '</td></tr>';
            }
            if (missingCity > 0) {
                html += '<tr class="warning clickable" data-issue="MissingCity" onclick="showAddressIssues(this.dataset.issue)"><td><span class="badge badge-warning">Missing City</span></td><td class="number">' + missingCity.toLocaleString() + '</td></tr>';
            }
            if (missingState > 0) {
                html += '<tr class="warning clickable" data-issue="MissingState" onclick="showAddressIssues(this.dataset.issue)"><td><span class="badge badge-warning">Missing State</span></td><td class="number">' + missingState.toLocaleString() + '</td></tr>';
            }
            if (missingZip > 0) {
                html += '<tr class="warning clickable" data-issue="MissingZip" onclick="showAddressIssues(this.dataset.issue)"><td><span class="badge badge-warning">Missing Zip</span></td><td class="number">' + missingZip.toLocaleString() + '</td></tr>';
            }

            if (issueTotal > 0) {
                html += '<tr class="total clickable" data-issue="all" onclick="showAddressIssues(this.dataset.issue)" style="background:#fff3cd;"><td><strong>Total with Issues</strong></td><td class="number"><strong>' + issueTotal.toLocaleString() + '</strong></td></tr>';
            } else {
                html += '<tr class="total"><td><strong>Total with Issues</strong></td><td class="number"><strong>0</strong></td></tr>';
            }

            document.getElementById('addressBody').innerHTML = html;
        }

        function updateKeyIssues(data) {
            var issues = [];

            var elecNoEmail = data.givers['Electronic_NoEmail'] || 0;
            var missingStreet = data.address['MissingStreet'] || 0;
            var missingCity = data.address['MissingCity'] || 0;
            var missingState = data.address['MissingState'] || 0;
            var missingZip = data.address['MissingZip'] || 0;
            var addressIssues = missingStreet + missingCity + missingState + missingZip;
            var printNoEmail = data.givers['Printed_NoEmail'] || 0;

            if (elecNoEmail > 0) {
                issues.push('<strong>' + elecNoEmail + ' electronic statement givers</strong> have no email on file');
            }
            if (addressIssues > 0) {
                issues.push('<strong>' + addressIssues + ' printed statement givers</strong> have incomplete addresses');
            }
            if (printNoEmail > 0) {
                issues.push('<strong>' + printNoEmail + ' printed statement givers</strong> have no email (cannot easily switch to electronic)');
            }

            if (issues.length > 0) {
                var html = '';
                for (var i = 0; i < issues.length; i++) {
                    html += '<li>' + issues[i] + '</li>';
                }
                document.getElementById('issuesList').innerHTML = html;
                document.getElementById('keyIssues').style.display = 'block';
            }
        }

        function showElectronicNoEmail() {
            document.getElementById('modal').classList.add('active');
            document.getElementById('modalTitle').textContent = 'Electronic Statement - No Email (' + currentYear + ' Givers)';
            document.getElementById('modalBody').innerHTML = '<div class="loading">Loading...</div>';

            var formData = new FormData();
            formData.append('action', 'get_electronic_no_email');
            formData.append('year', currentYear);

            fetch(scriptUrl, {
                method: 'POST',
                body: formData
            })
            .then(function(response) { return response.text(); })
            .then(function(text) {
                try {
                    var data = JSON.parse(text);
                    if (data.success) {
                        renderPeopleTable(data.people, 'electronic');
                    }
                } catch(e) {
                    console.error('Parse error:', e);
                }
            });
        }

        function showAddressIssues(issueType) {
            document.getElementById('modal').classList.add('active');
            var title = 'Address Issues';
            if (issueType === 'MissingStreet') title = 'Missing Street Address';
            else if (issueType === 'MissingCity') title = 'Missing City';
            else if (issueType === 'MissingState') title = 'Missing State';
            else if (issueType === 'MissingZip') title = 'Missing Zip';
            else title = 'All Address Issues';

            document.getElementById('modalTitle').textContent = title + ' (' + currentYear + ' Givers)';
            document.getElementById('modalBody').innerHTML = '<div class="loading">Loading...</div>';

            var formData = new FormData();
            formData.append('action', 'get_address_issues');
            formData.append('year', currentYear);
            formData.append('issue_type', issueType);

            fetch(scriptUrl, {
                method: 'POST',
                body: formData
            })
            .then(function(response) { return response.text(); })
            .then(function(text) {
                try {
                    var data = JSON.parse(text);
                    if (data.success) {
                        renderPeopleTable(data.people, 'address');
                    }
                } catch(e) {
                    console.error('Parse error:', e);
                }
            });
        }

        function renderPeopleTable(people, type) {
            if (people.length === 0) {
                document.getElementById('modalBody').innerHTML = '<div class="no-data">No people found</div>';
                return;
            }

            var html = '<table class="people-table"><thead><tr>';
            html += '<th>Name</th>';

            if (type === 'electronic') {
                html += '<th>Cell Phone</th><th>Home Phone</th><th>Address</th>';
            } else {
                html += '<th>Email</th><th>Cell Phone</th><th>Address</th><th>City</th><th>State</th><th>Zip</th><th>Issue</th>';
            }

            html += '</tr></thead><tbody>';

            for (var i = 0; i < people.length; i++) {
                var p = people[i];
                html += '<tr>';
                html += '<td><a href="/Person2/' + p.id + '" target="_blank">' + p.name + '</a></td>';

                if (type === 'electronic') {
                    html += '<td>' + (p.cell || '-') + '</td>';
                    html += '<td>' + (p.home || '-') + '</td>';
                    var addr = [p.address, p.city, p.state, p.zip].filter(function(x) { return x; }).join(', ');
                    html += '<td>' + (addr || '-') + '</td>';
                } else {
                    html += '<td>' + (p.email || '-') + '</td>';
                    html += '<td>' + (p.cell || '-') + '</td>';
                    html += '<td>' + (p.address || '<span class="badge badge-danger">Missing</span>') + '</td>';
                    html += '<td>' + (p.city || '<span class="badge badge-warning">Missing</span>') + '</td>';
                    html += '<td>' + (p.state || '<span class="badge badge-warning">Missing</span>') + '</td>';
                    html += '<td>' + (p.zip || '<span class="badge badge-warning">Missing</span>') + '</td>';
                    html += '<td><span class="badge badge-danger">' + p.issue + '</span></td>';
                }

                html += '</tr>';
            }

            html += '</tbody></table>';
            html += '<p style="margin-top:15px;color:#666;"><strong>' + people.length + '</strong> people found</p>';

            document.getElementById('modalBody').innerHTML = html;
        }

        function closeModal() {
            document.getElementById('modal').classList.remove('active');
        }

        // Close modal on overlay click
        document.getElementById('modal').addEventListener('click', function(e) {
            if (e.target === this) {
                closeModal();
            }
        });

        // Close modal on Escape key
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                closeModal();
            }
        });

        // Auto-redirect to PyScriptForm if we're on PyScript
        document.addEventListener('DOMContentLoaded', function() {
            if (window.location.pathname.indexOf('/PyScript/') > -1) {
                var newPath = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                window.location.href = newPath + window.location.search;
                return;
            }
            // Load data on page load
            loadData();
        });
    </script>
</body>
</html>
'''

    model.Form = html
