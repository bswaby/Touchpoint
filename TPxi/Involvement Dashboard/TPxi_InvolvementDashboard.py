#Roles=Access

#written by: Ben Swaby
#email:bswaby@fbchtn.org

model.Header = 'Involvement Dashboard'

# Check if this is an AJAX request
is_ajax = hasattr(model.Data, 'ajax') and model.Data.ajax == 'true'

if is_ajax:
    # Handle AJAX requests
    action = model.Data.action if hasattr(model.Data, 'action') else None

    if action == 'get_programs':
        # Get all programs
        sql = "SELECT Id, Name FROM Program ORDER BY Name"
        programs = q.QuerySql(sql)

        import json
        result = [{'id': p.Id, 'name': p.Name} for p in programs]
        print json.dumps(result)

    elif action == 'get_organizations':
        # Get organizations for a program
        prog_id = int(model.Data.prog_id) if hasattr(model.Data, 'prog_id') else 0

        sql = """
            SELECT DISTINCT o.OrganizationId, o.OrganizationName
            FROM Organizations o
            LEFT JOIN Division d ON o.DivisionId = d.Id
            WHERE o.OrganizationStatusId = 30
                AND d.ProgId = {0}
            ORDER BY o.OrganizationName
        """.format(prog_id)

        orgs = q.QuerySql(sql)

        import json
        result = [{'id': o.OrganizationId, 'name': o.OrganizationName} for o in orgs]
        print json.dumps(result)

    elif action == 'get_dashboard':
        import json
        try:
            # Get dashboard data for an organization
            org_id = int(model.Data.org_id) if hasattr(model.Data, 'org_id') else 0

            # Get organization info
            org_sql = """
                SELECT o.OrganizationId, o.OrganizationName,
                       d.Name as DivisionName, p.Name as ProgramName
                FROM Organizations o
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE o.OrganizationId = {0}
            """.format(org_id)

            org_info = list(q.QuerySql(org_sql))

            if not org_info:
                print json.dumps({'error': 'Organization not found'})
            else:
                org = org_info[0]

                # Get demographics
                demo_sql = """
                    SELECT
                        p.GenderId,
                        CASE
                            WHEN p.BirthYear IS NOT NULL AND p.BirthMonth IS NOT NULL AND p.BirthDay IS NOT NULL
                            THEN DATEDIFF(year, DATEFROMPARTS(p.BirthYear, p.BirthMonth, p.BirthDay), GETDATE())
                            ELSE NULL
                        END as Age,
                        mt.Description as MemberType,
                        ms.Description as MaritalStatus,
                        om.EnrollmentDate,
                        DATEDIFF(day, om.EnrollmentDate, GETDATE()) as DaysSinceEnrollment
                    FROM OrganizationMembers om
                    JOIN People p ON om.PeopleId = p.PeopleId
                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                    LEFT JOIN lookup.MaritalStatus ms ON p.MaritalStatusId = ms.Id
                    WHERE om.OrganizationId = {0}
                        AND p.IsDeceased = 0
                """.format(org_id)

                members = list(q.QuerySql(demo_sql))

                # Calculate statistics
                total_members = len(members)
                male_count = len([m for m in members if m.GenderId == 1])
                female_count = len([m for m in members if m.GenderId == 2])

                # Age groups
                age_groups = {'0-12': 0, '13-17': 0, '18-25': 0, '26-35': 0, '36-50': 0, '51-65': 0, '66+': 0, 'Unknown': 0}

                for member in members:
                    age = member.Age if hasattr(member, 'Age') and member.Age else None
                    if age is None:
                        age_groups['Unknown'] += 1
                    elif age <= 12:
                        age_groups['0-12'] += 1
                    elif age <= 17:
                        age_groups['13-17'] += 1
                    elif age <= 25:
                        age_groups['18-25'] += 1
                    elif age <= 35:
                        age_groups['26-35'] += 1
                    elif age <= 50:
                        age_groups['36-50'] += 1
                    elif age <= 65:
                        age_groups['51-65'] += 1
                    else:
                        age_groups['66+'] += 1

                # Member types
                member_types = {}
                for member in members:
                    mtype = member.MemberType if hasattr(member, 'MemberType') and member.MemberType else 'Unknown'
                    member_types[mtype] = member_types.get(mtype, 0) + 1

                # Marital status
                marital_status = {}
                for member in members:
                    status = member.MaritalStatus if hasattr(member, 'MaritalStatus') and member.MaritalStatus else 'Unknown'
                    marital_status[status] = marital_status.get(status, 0) + 1

                # Enrollment timeline (by month)
                enrollment_timeline = {}
                import datetime
                for member in members:
                    if hasattr(member, 'EnrollmentDate') and member.EnrollmentDate:
                        # Group by year-month (IronPython DateTime doesn't have strftime)
                        date_key = "{0:04d}-{1:02d}".format(member.EnrollmentDate.Year, member.EnrollmentDate.Month)
                        enrollment_timeline[date_key] = enrollment_timeline.get(date_key, 0) + 1

                # Sort timeline chronologically and get last 12 months
                sorted_timeline = sorted(enrollment_timeline.items(), key=lambda x: x[0], reverse=True)[:12]
                sorted_timeline.reverse()  # Oldest to newest for chart

                # Get subgroups
                subgroup_sql = """
                    SELECT mt.Name as SubgroupName, COUNT(DISTINCT omt.PeopleId) as MemberCount
                    FROM MemberTags mt
                    LEFT JOIN OrgMemMemTags omt ON mt.Id = omt.MemberTagId AND omt.OrgId = {0}
                    GROUP BY mt.Id, mt.Name
                    HAVING COUNT(DISTINCT omt.PeopleId) > 0
                    ORDER BY mt.Name
                """.format(org_id)

                subgroups = list(q.QuerySql(subgroup_sql))

                # Get transaction data (fees, payments, etc.)
                transaction_sql = """
                    SELECT
                        COUNT(DISTINCT t.Id) as TotalTransactions,
                        SUM(CASE WHEN t.Amt > 0 THEN 1 ELSE 0 END) as CompletedCount,
                        SUM(CASE WHEN t.Amt IS NULL OR t.Amt = 0 THEN 1 ELSE 0 END) as PendingCount,
                        SUM(ISNULL(t.Amt, 0)) as TotalPaid,
                        SUM(ISNULL(t.Amtdue, 0)) as TotalDue,
                        AVG(ISNULL(t.Amt, 0)) as AvgPaid
                    FROM [Transaction] t
                    WHERE t.OrgId = {0}
                """.format(org_id)

                transaction_result = list(q.QuerySql(transaction_sql))
                transactions = transaction_result[0] if transaction_result else None

                result = {
                    'org_name': org.OrganizationName,
                    'program_name': org.ProgramName if hasattr(org, 'ProgramName') and org.ProgramName else 'None',
                    'division_name': org.DivisionName if hasattr(org, 'DivisionName') and org.DivisionName else 'None',
                    'total_members': total_members,
                    'male_count': male_count,
                    'female_count': female_count,
                    'age_groups': age_groups,
                    'member_types': member_types,
                    'marital_status': marital_status,
                    'enrollment_timeline': dict(sorted_timeline),
                    'subgroups': [{'name': s.SubgroupName, 'count': s.MemberCount} for s in subgroups],
                    'transactions': {
                        'total': transactions.TotalTransactions if transactions and transactions.TotalTransactions else 0,
                        'completed': transactions.CompletedCount if transactions and transactions.CompletedCount else 0,
                        'pending': transactions.PendingCount if transactions and transactions.PendingCount else 0,
                        'total_paid': float(transactions.TotalPaid) if transactions and transactions.TotalPaid else 0,
                        'total_due': float(transactions.TotalDue) if transactions and transactions.TotalDue else 0,
                        'avg_paid': float(transactions.AvgPaid) if transactions and transactions.AvgPaid else 0
                    }
                }
                print json.dumps(result)

        except Exception as e:
            import traceback
            print json.dumps({'error': str(e), 'traceback': traceback.format_exc()})

else:
    # Show the main page with selectors
    model.Form = '''
    <style>
        .dashboard-container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        }
        .selector-card {
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .selector-card h2 {
            margin: 0 0 20px 0;
            color: #1e293b;
            font-size: 24px;
        }
        .selector-card select {
            width: 100%;
            padding: 12px;
            font-size: 16px;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            margin-bottom: 15px;
        }
        .dashboard-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 12px;
            margin-bottom: 30px;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
            display: none;
        }
        .dashboard-header h1 {
            margin: 0 0 10px 0;
            font-size: 32px;
        }
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .stat-card {
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        .stat-value {
            font-size: 36px;
            font-weight: bold;
            color: #667eea;
            margin: 10px 0;
        }
        .stat-label {
            color: #64748b;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .section {
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .section-title {
            font-size: 20px;
            font-weight: 600;
            margin: 0 0 20px 0;
            color: #1e293b;
        }
        .chart-bar {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
        }
        .chart-label {
            width: 120px;
            font-size: 14px;
            color: #475569;
        }
        .chart-bar-container {
            flex: 1;
            background: #f1f5f9;
            height: 30px;
            border-radius: 6px;
            overflow: hidden;
            margin: 0 15px;
        }
        .chart-bar-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            display: flex;
            align-items: center;
            justify-content: flex-end;
            padding-right: 10px;
            color: white;
            font-size: 12px;
            font-weight: 600;
        }
        .chart-count {
            width: 60px;
            text-align: right;
            font-weight: 600;
            color: #667eea;
        }
        #dashboard-content {
            display: none;
        }
        .subgroup-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .subgroup-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 15px;
            background: #f8fafc;
            margin-bottom: 8px;
            border-radius: 6px;
            border-left: 3px solid #667eea;
        }
        .subgroup-name {
            font-weight: 500;
            color: #1e293b;
        }
        .subgroup-count {
            background: #667eea;
            color: white;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }
    </style>

    <div class="dashboard-container">
        <div class="selector-card" id="selector-card">
            <h2><i class="fa fa-filter"></i> Select Program & Involvement</h2>
            <div id="selector-content">
                <label>Program:</label>
                <select id="program-select">
                    <option value="">-- Select Program --</option>
                </select>
                <label>Involvement:</label>
                <select id="org-select" disabled>
                    <option value="">-- Select Program First --</option>
                </select>
            </div>
        </div>

        <div id="dashboard-content">
            <div class="dashboard-header" id="dashboard-header">
                <h1 id="org-name"></h1>
                <p id="org-info" style="margin: 0; opacity: 0.9;"></p>
            </div>

            <div class="stats-grid" id="stats-grid"></div>
            <div class="section" id="age-section" style="display:none;">
                <h2 class="section-title"><i class="fa fa-chart-bar"></i> Age Distribution</h2>
                <div id="age-chart"></div>
            </div>
            <div class="section" id="type-section" style="display:none;">
                <h2 class="section-title"><i class="fa fa-users"></i> Member Types</h2>
                <div id="type-chart"></div>
            </div>
            <div class="section" id="marital-section" style="display:none;">
                <h2 class="section-title"><i class="fa fa-heart"></i> Marital Status</h2>
                <div id="marital-chart"></div>
            </div>
            <div class="section" id="timeline-section" style="display:none;">
                <h2 class="section-title"><i class="fa fa-calendar-plus"></i> Enrollment Timeline (Last 12 Months)</h2>
                <div id="timeline-chart"></div>
            </div>
            <div class="section" id="transaction-section" style="display:none;">
                <h2 class="section-title"><i class="fa fa-dollar"></i> Financial Transactions</h2>
                <div id="transaction-summary"></div>
            </div>
            <div class="section" id="subgroup-section" style="display:none;">
                <h2 class="section-title"><i class="fa fa-layer-group"></i> Subgroups</h2>
                <div id="subgroup-list"></div>
            </div>
        </div>
    </div>

    <script>
    // Version: 1.4 - Wait for jQuery to load
    (function() {
        function initDashboard() {
            var scriptUrl = window.location.pathname;

            // Event delegation for toggle selector (works with dynamically added element)
            $(document).on('click', '#toggle-selector', function() {
                $('#selector-content').slideToggle();
            });

            // Load programs on page load
        $.ajax({
            url: scriptUrl,
            type: 'POST',
            data: { ajax: 'true', action: 'get_programs' },
            success: function(response) {
                var programs = JSON.parse(response);
                programs.forEach(function(prog) {
                    $('#program-select').append('<option value="' + prog.id + '">' + prog.name + '</option>');
                });
            }
        });

        // When program is selected, load organizations
        $('#program-select').change(function() {
            var progId = $(this).val();
            $('#org-select').html('<option value="">-- Loading... --</option>').prop('disabled', true);
            $('#dashboard-content').hide();

            if (progId) {
                $.ajax({
                    url: scriptUrl,
                    type: 'POST',
                    data: { ajax: 'true', action: 'get_organizations', prog_id: progId },
                    success: function(response) {
                        var orgs = JSON.parse(response);
                        $('#org-select').html('<option value="">-- Select Involvement --</option>');
                        orgs.forEach(function(org) {
                            $('#org-select').append('<option value="' + org.id + '">' + org.name + '</option>');
                        });
                        $('#org-select').prop('disabled', false);
                    }
                });
            }
        });

        // When organization is selected, load dashboard
        $('#org-select').change(function() {
            var orgId = $(this).val();

            if (orgId) {
                $.ajax({
                    url: scriptUrl,
                    type: 'POST',
                    data: { ajax: 'true', action: 'get_dashboard', org_id: orgId },
                    success: function(response) {
                        var data = JSON.parse(response);

                        if (data.error) {
                            alert('Error: ' + data.error + (data.traceback ? '\\n\\n' + data.traceback : ''));
                            console.error('Dashboard error:', data);
                            return;
                        }

                        // Update header
                        $('#org-name').text(data.org_name);
                        $('#org-info').text(data.program_name + ' - ' + data.division_name);
                        $('#dashboard-header').show();

                        // Minimize the selector card after involvement is selected
                        $('#selector-content').slideUp();
                        $('#selector-card h2').html('<i class="fa fa-filter"></i> Select Program & Involvement <span id="toggle-selector" style="float:right; cursor:pointer;"><i class="fa fa-chevron-down"></i></span>');

                        // Update stats
                        var statsHtml = '';
                        statsHtml += '<div class="stat-card"><div class="stat-label">Total Members</div><div class="stat-value">' + data.total_members + '</div></div>';
                        statsHtml += '<div class="stat-card"><div class="stat-label">Male</div><div class="stat-value">' + data.male_count + '</div></div>';
                        statsHtml += '<div class="stat-card"><div class="stat-label">Female</div><div class="stat-value">' + data.female_count + '</div></div>';
                        statsHtml += '<div class="stat-card"><div class="stat-label">Subgroups</div><div class="stat-value">' + data.subgroups.length + '</div></div>';
                        $('#stats-grid').html(statsHtml);

                        // Age distribution chart
                        var ageHtml = '';
                        var ageOrder = ['0-12', '13-17', '18-25', '26-35', '36-50', '51-65', '66+', 'Unknown'];
                        var maxAge = Math.max.apply(null, Object.values(data.age_groups));
                        ageOrder.forEach(function(ageGroup) {
                            if (data.age_groups[ageGroup] && data.age_groups[ageGroup] > 0) {
                                var count = data.age_groups[ageGroup];
                                var pct = maxAge > 0 ? Math.round((count / maxAge) * 100) : 0;
                                var totalPct = data.total_members > 0 ? Math.round((count / data.total_members) * 100) : 0;
                                ageHtml += '<div class="chart-bar">';
                                ageHtml += '<div class="chart-label">' + ageGroup + '</div>';
                                ageHtml += '<div class="chart-bar-container"><div class="chart-bar-fill" style="width:' + pct + '%;background:linear-gradient(90deg,#667eea,#764ba2)">' + totalPct + '%</div></div>';
                                ageHtml += '<div class="chart-count">' + count + '</div>';
                                ageHtml += '</div>';
                            }
                        });
                        $('#age-chart').html(ageHtml);
                        $('#age-section').show();

                        // Member types chart
                        var typeHtml = '';
                        var maxType = Math.max.apply(null, Object.values(data.member_types));
                        for (var type in data.member_types) {
                            var count = data.member_types[type];
                            var pct = maxType > 0 ? Math.round((count / maxType) * 100) : 0;
                            var totalPct = data.total_members > 0 ? Math.round((count / data.total_members) * 100) : 0;
                            typeHtml += '<div class="chart-bar">';
                            typeHtml += '<div class="chart-label">' + type + '</div>';
                            typeHtml += '<div class="chart-bar-container"><div class="chart-bar-fill" style="width:' + pct + '%;background:linear-gradient(90deg,#667eea,#764ba2)">' + totalPct + '%</div></div>';
                            typeHtml += '<div class="chart-count">' + count + '</div>';
                            typeHtml += '</div>';
                        }
                        $('#type-chart').html(typeHtml);
                        $('#type-section').show();

                        // Marital status chart
                        if (data.marital_status && Object.keys(data.marital_status).length > 0) {
                            var maritalHtml = '';
                            var maxMarital = Math.max.apply(null, Object.values(data.marital_status));
                            for (var status in data.marital_status) {
                                var count = data.marital_status[status];
                                var pct = maxMarital > 0 ? Math.round((count / maxMarital) * 100) : 0;
                                var totalPct = data.total_members > 0 ? Math.round((count / data.total_members) * 100) : 0;
                                maritalHtml += '<div class="chart-bar">';
                                maritalHtml += '<div class="chart-label">' + status + '</div>';
                                maritalHtml += '<div class="chart-bar-container"><div class="chart-bar-fill" style="width:' + pct + '%;background:linear-gradient(90deg,#667eea,#764ba2)">' + totalPct + '%</div></div>';
                                maritalHtml += '<div class="chart-count">' + count + '</div>';
                                maritalHtml += '</div>';
                            }
                            $('#marital-chart').html(maritalHtml);
                            $('#marital-section').show();
                        }

                        // Transaction summary
                        if (data.transactions && data.transactions.total > 0) {
                            var transHtml = '<div class="row" style="margin-bottom: 15px;">';
                            transHtml += '<div class="col-md-4"><div style="background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">';
                            transHtml += '<div style="font-size: 24px; font-weight: bold; color: #667eea;">' + data.transactions.total + '</div>';
                            transHtml += '<div style="font-size: 12px; color: #666;">Total Transactions</div>';
                            transHtml += '</div></div>';
                            transHtml += '<div class="col-md-4"><div style="background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">';
                            transHtml += '<div style="font-size: 24px; font-weight: bold; color: #27ae60;">$' + data.transactions.total_paid.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,') + '</div>';
                            transHtml += '<div style="font-size: 12px; color: #666;">Total Paid</div>';
                            transHtml += '</div></div>';
                            transHtml += '<div class="col-md-4"><div style="background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">';
                            transHtml += '<div style="font-size: 24px; font-weight: bold; color: #e74c3c;">$' + data.transactions.total_due.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,') + '</div>';
                            transHtml += '<div style="font-size: 12px; color: #666;">Total Due</div>';
                            transHtml += '</div></div>';
                            transHtml += '</div>';
                            transHtml += '<div class="row">';
                            transHtml += '<div class="col-md-6"><div style="background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">';
                            transHtml += '<div style="font-size: 20px; font-weight: bold; color: #27ae60;">' + data.transactions.completed + '</div>';
                            transHtml += '<div style="font-size: 12px; color: #666;">Completed</div>';
                            transHtml += '</div></div>';
                            transHtml += '<div class="col-md-6"><div style="background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">';
                            transHtml += '<div style="font-size: 20px; font-weight: bold; color: #f39c12;">' + data.transactions.pending + '</div>';
                            transHtml += '<div style="font-size: 12px; color: #666;">Pending</div>';
                            transHtml += '</div></div>';
                            transHtml += '</div>';
                            $('#transaction-summary').html(transHtml);
                            $('#transaction-section').show();
                        }

                        // Enrollment timeline chart
                        if (data.enrollment_timeline && Object.keys(data.enrollment_timeline).length > 0) {
                            var timelineHtml = '';
                            var maxTimeline = Math.max.apply(null, Object.values(data.enrollment_timeline));
                            for (var month in data.enrollment_timeline) {
                                var count = data.enrollment_timeline[month];
                                var pct = maxTimeline > 0 ? Math.round((count / maxTimeline) * 100) : 0;
                                // Format month (YYYY-MM to MMM YYYY)
                                var parts = month.split('-');
                                var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                                var monthLabel = monthNames[parseInt(parts[1]) - 1] + ' ' + parts[0];
                                timelineHtml += '<div class="chart-bar">';
                                timelineHtml += '<div class="chart-label">' + monthLabel + '</div>';
                                timelineHtml += '<div class="chart-bar-container"><div class="chart-bar-fill" style="width:' + pct + '%;"></div></div>';
                                timelineHtml += '<div class="chart-count">' + count + '</div>';
                                timelineHtml += '</div>';
                            }
                            $('#timeline-chart').html(timelineHtml);
                            $('#timeline-section').show();
                        }

                        // Subgroups list
                        if (data.subgroups && data.subgroups.length > 0) {
                            var subgroupHtml = '<ul class="subgroup-list">';
                            data.subgroups.forEach(function(subgroup) {
                                subgroupHtml += '<li class="subgroup-item">';
                                subgroupHtml += '<span class="subgroup-name">' + subgroup.name + '</span>';
                                subgroupHtml += '<span class="subgroup-count">' + subgroup.count + '</span>';
                                subgroupHtml += '</li>';
                            });
                            subgroupHtml += '</ul>';
                            $('#subgroup-list').html(subgroupHtml);
                            $('#subgroup-section').show();
                        } else {
                            $('#subgroup-section').hide();
                        }

                        $('#dashboard-content').fadeIn();
                    }
                });
            } else {
                $('#dashboard-content').hide();
            }
        });
        }

        // Wait for jQuery to be loaded by TouchPoint
        if (window.jQuery) {
            $(document).ready(initDashboard);
        } else {
            // jQuery not loaded yet, wait for it
            var checkJQuery = setInterval(function() {
                if (window.jQuery) {
                    clearInterval(checkJQuery);
                    $(document).ready(initDashboard);
                }
            }, 50);
        }
    })();
    </script>
    '''
