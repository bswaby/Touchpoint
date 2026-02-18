#roles=Edit

#####################################################################
# TPxi Ministry Deposit Report
#####################################################################
# Purpose: Generate deposit reports for ministries to calculate their
# daily deposits with breakdown by payment type (cash, check, credit card)
# and fund allocation for reconciliation with finance team
#
# Features:
# 1. Date/Time range selection with flexible filtering
# 2. Program, Division, and Organization filters for ministry focus
# 3. Daily breakdown with payment type totals (Cash, Check, Credit Card)
# 4. Fund-by-fund breakdown for each payment type
# 5. Summary totals for Cash and Check deposits (what ministries turn in)
# 6. Credit card totals for reference (processed electronically)
# 7. Export functionality for finance team reconciliation
#
# Upload Instructions:
# 1. Admin > Advanced > Special Content > Python
# 2. New Python Script File
# 3. Name: TPxi_MinistryDepositReport
# 4. Configure funds and payment types below
# 5. Schedule or run manually for deposit periods
#
# Configuration: Update fund IDs and payment types for your church

# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
#####################################################################

# ===== CONFIGURATION SECTION =====
class Config:
    # Report settings
    REPORT_TITLE = "Ministry Deposit Report"
    
    # Default date range (days back from today)
    DEFAULT_DAYS_BACK = 7
    
    # Payment type configuration (TouchPoint ContributionTypeId values)
    PAYMENT_TYPES = {
        'Check': [1],           # ContributionTypeId = 1 (Check)
        'Cash': [2],            # ContributionTypeId = 2 (Cash)  
        'Credit Card': [3, 5],  # ContributionTypeId = 3 (Credit Card), 5 (Online)
        'ACH': [4],             # ContributionTypeId = 4 (ACH)
        'Other': [6, 7, 8, 9]   # Other payment types
    }
    
    # Funds to track (will be dynamically loaded, but you can specify priority funds)
    PRIORITY_FUNDS = [
        # Add your most common fund IDs here, e.g.:
        # 1,  # General Fund
        # 2,  # Building Fund
        # 3,  # Missions Fund
    ]
    
    # Display settings
    SHOW_CREDIT_CARD_DETAILS = True  # Show CC details for reference (not deposited)
    SHOW_FUND_BREAKDOWN = True      # Show breakdown by fund
    SHOW_ZERO_AMOUNTS = False       # Include rows with $0.00
    SHOW_DAILY_FUND_DETAIL = False  # Show detailed daily fund breakdown (can make report very long)
    SHOW_TRANSACTION_DETAILS = True # Show expand option for individual transactions per day
    
    # Security settings
    REQUIRED_ROLES = ["Edit", "Finance", "Admin"]  # Roles that can access this report
    
    # Export settings
    ENABLE_CSV_EXPORT = True
    CSV_FILENAME_PREFIX = "ministry_deposit"

# Set page header
model.Header = Config.REPORT_TITLE

def main():
    """Main execution function"""
    try:
        # Check permissions
        if not check_permissions():
            return
        
        # Get request parameters
        action = getattr(model.Data, 'action', '')
        
        if action == 'generate_report':
            generate_deposit_report()
        elif action == 'export_csv':
            export_csv_data()
        elif action == 'load_divisions':
            load_divisions_ajax()
        elif action == 'load_organizations':
            load_organizations_ajax()
        elif action == 'get_day_details':
            get_day_transaction_details()
        elif action == 'get_person_info':
            get_person_popup_info()
        else:
            show_filter_form()
    
    except Exception as e:
        print '<div class="alert alert-danger">Error: {}</div>'.format(str(e))
        if model.UserIsInRole("Developer"):
            import traceback
            print '<pre>{}</pre>'.format(traceback.format_exc())

def check_permissions():
    """Check if user has required permissions"""
    has_permission = False
    for role in Config.REQUIRED_ROLES:
        if model.UserIsInRole(role):
            has_permission = True
            break
    
    if not has_permission:
        print '''
        <div class="alert alert-danger">
            <h4><i class="fa fa-lock"></i> Access Denied</h4>
            <p>You need one of these roles to access this report: {}</p>
        </div>
        '''.format(", ".join(Config.REQUIRED_ROLES))
        return False
    return True

def show_filter_form():
    """Show the filter form for report parameters"""
    # Get default dates
    end_date = model.DateTime
    start_date_sql = q.QuerySqlScalar("""
        SELECT CONVERT(varchar, DATEADD(day, -{}, GETDATE()), 120)
    """.format(Config.DEFAULT_DAYS_BACK))
    
    # Load programs for dropdown
    programs = q.QuerySql("""
        SELECT DISTINCT p.Id, p.Name
        FROM Program p
        INNER JOIN Division d ON p.Id = d.ProgId
        INNER JOIN Organizations o ON d.Id = o.DivisionId
        WHERE o.OrganizationStatusId = 30
        ORDER BY p.Name
    """)
    
    # Add CSS for better styling
    print '''
    <style>
    .panel-heading {
        background-color: #f5f5f5;
        border-bottom: 1px solid #ddd;
    }
    .table th {
        background-color: #f9f9f9;
    }
    .text-muted {
        color: #777 !important;
    }
    @media print {
        .btn, .page-header .pull-right { display: none; }
    }
    </style>
    '''
    
    print '''
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">
                <div class="page-header">
                    <h1>{} <small>Generate Deposit Reports by Date Range and Ministry</small></h1>
                </div>
            </div>
        </div>
        
        <div class="row">
            <div class="col-md-8">
                <div class="panel panel-primary">
                    <div class="panel-heading">
                        <h3 class="panel-title">Report Filters</h3>
                    </div>
                    <div class="panel-body">
                        <form id="depositReportForm" method="post" onsubmit="console.log('Form submitting to:', this.action); return true;">
                            <input type="hidden" name="action" value="generate_report">
                            
                            <div class="row">
                                <div class="col-md-6">
                                    <div class="form-group">
                                        <label for="startDate">Start Date:</label>
                                        <input type="date" class="form-control" id="startDate" name="start_date" 
                                               value="{}" required>
                                    </div>
                                </div>
                                <div class="col-md-6">
                                    <div class="form-group">
                                        <label for="endDate">End Date:</label>
                                        <input type="date" class="form-control" id="endDate" name="end_date" 
                                               value="{}" required>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="row">
                                <div class="col-md-4">
                                    <div class="form-group">
                                        <label for="programId">Program (Optional):</label>
                                        <select class="form-control" id="programId" name="program_id" 
                                                onchange="loadDivisions()">
                                            <option value="">All Programs</option>
    '''.format(
        Config.REPORT_TITLE,
        start_date_sql[:10] if start_date_sql else '',
        end_date.ToString("yyyy-MM-dd")
    )
    
    # Add program options
    for program in programs:
        print '                            <option value="{}">{}</option>'.format(
            safe_get_value(program, 'Id', ''),
            safe_get_value(program, 'Name', '')
        )
    
    print '''
                                        </select>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="form-group">
                                        <label for="divisionId">Division (Optional):</label>
                                        <select class="form-control" id="divisionId" name="division_id" 
                                                onchange="loadOrganizations()">
                                            <option value="">All Divisions</option>
                                        </select>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="form-group">
                                        <label for="organizationId">Organization (Optional):</label>
                                        <select class="form-control" id="organizationId" name="organization_id">
                                            <option value="">All Organizations</option>
                                        </select>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="row">
                                <div class="col-md-12">
                                    <button type="button" class="btn btn-primary btn-lg" onclick="submitDepositForm()">
                                        <i class="fa fa-calculator"></i> Generate Deposit Report
                                    </button>
                                    <button type="button" class="btn btn-default" onclick="setQuickRange(7)">
                                        Last 7 Days
                                    </button>
                                    <button type="button" class="btn btn-default" onclick="setQuickRange(30)">
                                        Last 30 Days
                                    </button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
            
            <div class="col-md-4">
                <div class="panel panel-info">
                    <div class="panel-heading">
                        <h3 class="panel-title">About This Report</h3>
                    </div>
                    <div class="panel-body">
                        <p><strong>Purpose:</strong> Generate daily deposit reports showing cash and check totals 
                           that ministries need to turn in to the finance team.</p>
                        
                        <p><strong>What's Included:</strong></p>
                        <ul>
                            <li><strong>Cash & Check Totals</strong> - What gets physically deposited</li>
                            <li><strong>Credit Card Reference</strong> - Electronic payments for comparison</li>
                            <li><strong>Fund Breakdown</strong> - How money is allocated by fund</li>
                            <li><strong>Daily Summary</strong> - Each day broken out separately</li>
                        </ul>
                        
                        <p><strong>Filters:</strong></p>
                        <ul>
                            <li><strong>Date Range</strong> - Select the period for the deposit</li>
                            <li><strong>Program/Division/Org</strong> - Focus on specific ministries</li>
                        </ul>
                        
                        <div class="alert alert-warning">
                            <small><strong>Note:</strong> This report focuses on Cash and Check deposits that 
                            require physical handling. Credit card amounts are shown for reference but are 
                            processed electronically.</small>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Helper function to get proper PyScriptForm URL (from CLAUDE.md pattern)
    function getPyScriptAddress() {
        let path = window.location.pathname;
        return path.replace("/PyScript/", "/PyScriptForm/");
    }
    
    function setQuickRange(days) {
        var endDate = new Date();
        var startDate = new Date();
        startDate.setDate(endDate.getDate() - days);
        
        document.getElementById('startDate').value = formatDate(startDate);
        document.getElementById('endDate').value = formatDate(endDate);
    }
    
    function formatDate(date) {
        var year = date.getFullYear();
        var month = ('0' + (date.getMonth() + 1)).slice(-2);
        var day = ('0' + date.getDate()).slice(-2);
        return year + '-' + month + '-' + day;
    }
    
    function loadDivisions() {
        var programId = document.getElementById('programId').value;
        var divisionSelect = document.getElementById('divisionId');
        var orgSelect = document.getElementById('organizationId');
        
        // Clear existing options
        divisionSelect.innerHTML = '<option value="">All Divisions</option>';
        orgSelect.innerHTML = '<option value="">All Organizations</option>';
        
        if (!programId) return;
        
        $.ajax({
            url: getPyScriptAddress(),
            method: 'POST',
            data: { action: 'load_divisions', program_id: programId },
            success: function(data) {
                var divisions = JSON.parse(data);
                divisions.forEach(function(div) {
                    var option = document.createElement('option');
                    option.value = div.Id;
                    option.textContent = div.Name;
                    divisionSelect.appendChild(option);
                });
            }
        });
    }
    
    function loadOrganizations() {
        var divisionId = document.getElementById('divisionId').value;
        var orgSelect = document.getElementById('organizationId');
        
        // Clear existing options
        orgSelect.innerHTML = '<option value="">All Organizations</option>';
        
        if (!divisionId) return;
        
        $.ajax({
            url: getPyScriptAddress(),
            method: 'POST',
            data: { action: 'load_organizations', division_id: divisionId },
            success: function(data) {
                var organizations = JSON.parse(data);
                organizations.forEach(function(org) {
                    var option = document.createElement('option');
                    option.value = org.OrganizationId;
                    option.textContent = org.OrganizationName;
                    orgSelect.appendChild(option);
                });
            }
        });
    }
    
    // Submit form using PyScriptForm URL
    function submitDepositForm() {
        var form = document.getElementById('depositReportForm');
        var formData = new FormData(form);
        
        // Log for debugging
        console.log('Submitting to:', getPyScriptAddress());
        
        // Create a temporary form to submit with correct action
        var tempForm = document.createElement('form');
        tempForm.method = 'POST';
        tempForm.action = getPyScriptAddress();
        
        // Copy all form data
        for (var pair of formData.entries()) {
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = pair[0];
            input.value = pair[1];
            tempForm.appendChild(input);
        }
        
        document.body.appendChild(tempForm);
        tempForm.submit();
    }
    </script>
    '''

def load_divisions_ajax():
    """Load divisions for a selected program via AJAX"""
    # Set content type for JSON response
    model.Header = 'Content-Type: application/json'
    
    program_id = getattr(model.Data, 'program_id', '')
    
    if not program_id:
        print '[]'
        return
    
    divisions = q.QuerySql("""
        SELECT DISTINCT d.Id, d.Name
        FROM Division d
        INNER JOIN Organizations o ON d.Id = o.DivisionId
        WHERE d.ProgId = {} AND o.OrganizationStatusId = 30
        ORDER BY d.Name
    """.format(int(program_id)))
    
    result = []
    for div in divisions:
        result.append({
            'Id': safe_get_value(div, 'Id', ''),
            'Name': safe_get_value(div, 'Name', '')
        })
    
    import json
    print json.dumps(result)

def load_organizations_ajax():
    """Load organizations for a selected division via AJAX"""
    # Set content type for JSON response
    model.Header = 'Content-Type: application/json'
    
    division_id = getattr(model.Data, 'division_id', '')
    
    if not division_id:
        print '[]'
        return
    
    organizations = q.QuerySql("""
        SELECT OrganizationId, OrganizationName
        FROM Organizations
        WHERE DivisionId = {} AND OrganizationStatusId = 30
        ORDER BY OrganizationName
    """.format(int(division_id)))
    
    result = []
    for org in organizations:
        result.append({
            'OrganizationId': safe_get_value(org, 'OrganizationId', ''),
            'OrganizationName': safe_get_value(org, 'OrganizationName', '')
        })
    
    import json
    print json.dumps(result)

def generate_deposit_report():
    """Generate the main deposit report"""
    # Get form parameters
    start_date = getattr(model.Data, 'start_date', '')
    end_date = getattr(model.Data, 'end_date', '')
    program_id = getattr(model.Data, 'program_id', '')
    division_id = getattr(model.Data, 'division_id', '')
    organization_id = getattr(model.Data, 'organization_id', '')
    
    if not start_date or not end_date:
        print '<div class="alert alert-danger">Start date and end date are required.</div>'
        return
    
    # Build filter description
    filter_desc = get_filter_description(program_id, division_id, organization_id)
    
    # Get deposit data
    deposit_data = get_deposit_data(start_date, end_date, program_id, division_id, organization_id)
    
    if not deposit_data:
        print '''
        <div class="alert alert-warning">
            <h4>No deposit data found</h4>
            <p>No contributions found for the selected date range and filters.</p>
            <a href="javascript:history.back()" class="btn btn-default">
                <i class="fa fa-arrow-left"></i> Back to Filters
            </a>
        </div>
        '''
        return
    
    # Render the report with all parameters
    render_deposit_report(deposit_data, start_date, end_date, filter_desc, program_id, division_id, organization_id)

def get_filter_description(program_id, division_id, organization_id):
    """Get human-readable description of applied filters"""
    filters = []
    
    if organization_id:
        org_name = q.QuerySqlScalar("""
            SELECT OrganizationName FROM Organizations WHERE OrganizationId = {}
        """.format(int(organization_id)))
        if org_name:
            filters.append("Organization: {}".format(org_name))
    elif division_id:
        div_name = q.QuerySqlScalar("""
            SELECT Name FROM Division WHERE Id = {}
        """.format(int(division_id)))
        if div_name:
            filters.append("Division: {}".format(div_name))
    elif program_id:
        prog_name = q.QuerySqlScalar("""
            SELECT Name FROM Program WHERE Id = {}
        """.format(int(program_id)))
        if prog_name:
            filters.append("Program: {}".format(prog_name))
    
    return " | ".join(filters) if filters else "All Ministries"

def get_deposit_data(start_date, end_date, program_id, division_id, organization_id):
    """Get transaction data for the deposit report using Transaction table"""
    # Build WHERE clause for filters
    where_clauses = [
        "t.TransactionDate >= '{}'".format(start_date),
        "t.TransactionDate <= '{} 23:59:59.999'".format(end_date),
        "t.amt <> 0",  # Only non-zero payments
        "t.TransactionId IS NOT NULL",
        "t.voided IS NULL"  # Exclude voided transactions
    ]
    
    # Add ministry filters based on organization hierarchy
    if organization_id:
        where_clauses.append("t.OrgId = {}".format(int(organization_id)))
    elif division_id:
        where_clauses.append("o.DivisionId = {}".format(int(division_id)))
    elif program_id:
        where_clauses.append("d.ProgId = {}".format(int(program_id)))
    
    # Using Transaction table with accounting code from Organizations
    sql = """
    WITH TransactionData AS (
        SELECT 
            t.*,
            o.OrganizationId,
            o.OrganizationName,
            o.DivisionId,
            CASE 
                WHEN o.RegAccountCodeId IS NOT NULL THEN CAST(o.RegAccountCodeId AS NVARCHAR(50))
                ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'NVARCHAR(50)')
            END AS AccountingCode
        FROM [Transaction] t
        LEFT JOIN Organizations o ON t.OrgId = o.OrganizationId
        WHERE t.TransactionDate >= '{0}'
          AND t.TransactionDate <= '{1} 23:59:59.999'
          AND t.amt <> 0
          AND t.TransactionId IS NOT NULL
          AND t.voided IS NULL
    )
    SELECT 
        CONVERT(date, td.TransactionDate) AS DepositDate,
        DATENAME(weekday, td.TransactionDate) AS DayOfWeek,
        td.Message,
        ISNULL(td.AccountingCode, '1') AS FundId,
        ISNULL(ac.Description, 'General Fund') AS FundName,
        ISNULL(ac.Code, '') AS FundIncomeAccount,
        COUNT(*) AS TransactionCount,
        SUM(td.amt) AS TotalAmount
    FROM TransactionData td
    LEFT JOIN Division d ON td.DivisionId = d.Id
    LEFT JOIN lookup.AccountCode ac ON ac.Id = td.AccountingCode
    WHERE 1=1
    """.format(start_date, end_date)
    
    # Add additional filters
    if organization_id:
        sql += " AND td.OrganizationId = {}".format(int(organization_id))
    elif division_id:
        sql += " AND td.DivisionId = {}".format(int(division_id))
    elif program_id:
        sql += " AND d.ProgId = {}".format(int(program_id))
    
    sql += """
    GROUP BY 
        CONVERT(date, td.TransactionDate),
        DATENAME(weekday, td.TransactionDate),
        td.Message,
        td.AccountingCode,
        ac.Description,
        ac.Code
    ORDER BY 
        DepositDate DESC,
        ac.Description,
        td.Message
    """
    
    return q.QuerySql(sql)

def render_deposit_report(deposit_data, start_date, end_date, filter_desc, program_id='', division_id='', organization_id=''):
    """Render the complete deposit report"""
    # Add CSS for report styling
    print '''
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css">
    <style>
    body {
        font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
        font-size: 14px;
        line-height: 1.42857143;
        color: #333;
    }
    .container-fluid {
        padding-right: 15px;
        padding-left: 15px;
        margin-right: auto;
        margin-left: auto;
    }
    .page-header {
        padding-bottom: 9px;
        margin: 40px 0 20px;
        border-bottom: 1px solid #eee;
    }
    .page-header h1 {
        margin-top: 0;
        font-size: 36px;
    }
    .page-header h1 small {
        font-size: 60%;
        color: #777;
    }
    h2 {
        font-size: 30px;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    h3 {
        font-size: 24px;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    h4 {
        font-size: 18px;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    .panel {
        border: 1px solid #ddd;
        border-radius: 4px;
        margin-bottom: 20px;
        background-color: #fff;
        box-shadow: 0 1px 1px rgba(0,0,0,.05);
    }
    .panel-heading {
        background-color: #f5f5f5;
        border-bottom: 1px solid #ddd;
        padding: 10px 15px;
        border-top-left-radius: 3px;
        border-top-right-radius: 3px;
    }
    .panel-title {
        margin-top: 0;
        margin-bottom: 0;
        font-size: 16px;
    }
    .panel-body {
        padding: 15px;
    }
    .panel-success {
        border-color: #d6e9c6;
    }
    .panel-success > .panel-heading {
        background-color: #dff0d8;
        border-color: #d6e9c6;
        color: #3c763d;
    }
    .panel-success .panel-body {
        background-color: #f8fff8;
    }
    .panel-primary {
        border-color: #337ab7;
    }
    .panel-primary > .panel-heading {
        background-color: #337ab7;
        border-color: #337ab7;
        color: #fff;
    }
    .panel-primary .panel-body {
        background-color: #f7f9ff;
    }
    .panel-info {
        border-color: #bce8f1;
    }
    .panel-info > .panel-heading {
        background-color: #d9edf7;
        border-color: #bce8f1;
        color: #31708f;
    }
    .panel-info .panel-body {
        background-color: #f6fbff;
    }
    .panel-warning {
        border-color: #faebcc;
    }
    .panel-warning > .panel-heading {
        background-color: #fcf8e3;
        border-color: #faebcc;
        color: #8a6d3b;
    }
    .panel-warning .panel-body {
        background-color: #fffef9;
    }
    .panel-default {
        border-color: #ddd;
    }
    .panel-default > .panel-heading {
        background-color: #f5f5f5;
        border-color: #ddd;
        color: #333;
    }
    .alert {
        padding: 15px;
        margin-bottom: 20px;
        border: 1px solid transparent;
        border-radius: 4px;
    }
    .alert-success {
        color: #3c763d;
        background-color: #dff0d8;
        border-color: #d6e9c6;
    }
    .alert-success h4 {
        color: #3c763d;
    }
    .table {
        width: 100%;
        max-width: 100%;
        margin-bottom: 20px;
        background-color: #fff;
    }
    .table > thead > tr > th,
    .table > tbody > tr > th,
    .table > tfoot > tr > th,
    .table > thead > tr > td,
    .table > tbody > tr > td,
    .table > tfoot > tr > td {
        padding: 8px;
        line-height: 1.42857143;
        vertical-align: top;
        border-top: 1px solid #ddd;
    }
    .table > thead > tr > th {
        vertical-align: bottom;
        border-bottom: 2px solid #ddd;
        background-color: #f9f9f9;
        font-weight: bold;
    }
    .table-striped > tbody > tr:nth-of-type(odd) {
        background-color: #f9f9f9;
    }
    .table-bordered {
        border: 1px solid #ddd;
    }
    .table-bordered > thead > tr > th,
    .table-bordered > tbody > tr > th,
    .table-bordered > tfoot > tr > th,
    .table-bordered > thead > tr > td,
    .table-bordered > tbody > tr > td,
    .table-bordered > tfoot > tr > td {
        border: 1px solid #ddd;
    }
    .table-responsive {
        min-height: .01%;
        overflow-x: auto;
    }
    .table-condensed > thead > tr > th,
    .table-condensed > tbody > tr > th,
    .table-condensed > tfoot > tr > th,
    .table-condensed > thead > tr > td,
    .table-condensed > tbody > tr > td,
    .table-condensed > tfoot > tr > td {
        padding: 5px;
    }
    .well {
        min-height: 20px;
        padding: 19px;
        margin-bottom: 20px;
        background-color: #f5f5f5;
        border: 1px solid #e3e3e3;
        border-radius: 4px;
        box-shadow: inset 0 1px 1px rgba(0,0,0,.05);
    }
    .well-sm {
        padding: 9px;
        border-radius: 3px;
    }
    .btn {
        display: inline-block;
        padding: 6px 12px;
        margin-bottom: 0;
        font-size: 14px;
        font-weight: normal;
        line-height: 1.42857143;
        text-align: center;
        white-space: nowrap;
        vertical-align: middle;
        cursor: pointer;
        border: 1px solid transparent;
        border-radius: 4px;
    }
    .btn-default {
        color: #333;
        background-color: #fff;
        border-color: #ccc;
    }
    .btn-success {
        color: #fff;
        background-color: #5cb85c;
        border-color: #4cae4c;
    }
    .text-muted {
        color: #777 !important;
    }
    .text-center {
        text-align: center;
    }
    .text-right {
        text-align: right;
    }
    .pull-right {
        float: right !important;
    }
    .row {
        margin-right: -15px;
        margin-left: -15px;
    }
    .row:before,
    .row:after {
        display: table;
        content: " ";
    }
    .row:after {
        clear: both;
    }
    .col-md-1, .col-md-2, .col-md-3, .col-md-4, .col-md-5, .col-md-6,
    .col-md-7, .col-md-8, .col-md-9, .col-md-10, .col-md-11, .col-md-12 {
        position: relative;
        min-height: 1px;
        padding-right: 15px;
        padding-left: 15px;
        float: left;
    }
    .col-md-12 { width: 100%; }
    .col-md-11 { width: 91.66666667%; }
    .col-md-10 { width: 83.33333333%; }
    .col-md-9 { width: 75%; }
    .col-md-8 { width: 66.66666667%; }
    .col-md-7 { width: 58.33333333%; }
    .col-md-6 { width: 50%; }
    .col-md-5 { width: 41.66666667%; }
    .col-md-4 { width: 33.33333333%; }
    .col-md-3 { width: 25%; }
    .col-md-2 { width: 16.66666667%; }
    .col-md-1 { width: 8.33333333%; }
    
    /* Print styles - Optimized for compact printing */
    /* Expand/Collapse button styles */
    .expand-btn {
        cursor: pointer;
        color: #337ab7;
        font-size: 12px;
        margin-left: 10px;
    }
    .expand-btn:hover {
        color: #23527c;
        text-decoration: underline;
    }
    .detail-row {
        display: none;
    }
    .detail-content {
        padding: 20px;
        background-color: #f5f7fa;
        border: 3px solid #337ab7;
        border-radius: 8px;
        margin: 10px;
        box-shadow: inset 0 2px 5px rgba(0,0,0,0.1);
    }
    
    /* Sortable table styles */
    .sortable {
        cursor: pointer !important;
        user-select: none;
        position: relative;
        transition: background-color 0.2s;
    }
    .sortable:hover {
        background-color: #e8e8e8 !important;
        text-decoration: underline;
    }
    .sortable i {
        font-size: 12px;
        color: #999;
        margin-left: 5px;
        transition: color 0.2s;
    }
    .sortable:hover i {
        color: #666;
    }
    .sortable.sort-asc i:before {
        content: '\f0de'; /* fa-sort-asc */
        color: #337ab7 !important;
    }
    .sortable.sort-desc i:before {
        content: '\f0dd'; /* fa-sort-desc */
        color: #337ab7 !important;
    }
    /* Ensure sortable headers in detail content are clickable */
    .detail-content th.sortable {
        cursor: pointer !important;
        z-index: 10;
    }
    
    /* Modal styles for person popup */
    .modal {
        display: none;
        position: fixed;
        z-index: 1000;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        overflow: auto;
        background-color: rgba(0,0,0,0.4);
    }
    .modal-content {
        background-color: #fefefe;
        margin: 10% auto;
        padding: 0;
        border: 1px solid #888;
        width: 80%;
        max-width: 600px;
        border-radius: 5px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .modal-header {
        padding: 15px 20px;
        background-color: #f5f5f5;
        border-bottom: 1px solid #ddd;
        border-radius: 5px 5px 0 0;
    }
    .modal-header h2 {
        margin: 0;
        font-size: 24px;
    }
    .modal-body {
        padding: 20px;
    }
    .modal-close {
        color: #aaa;
        float: right;
        font-size: 28px;
        font-weight: bold;
        cursor: pointer;
    }
    .modal-close:hover,
    .modal-close:focus {
        color: #000;
    }
    .person-info-section {
        margin-bottom: 20px;
    }
    .person-info-section h4 {
        margin-bottom: 10px;
        color: #337ab7;
    }
    .info-table {
        width: 100%;
        margin-bottom: 10px;
    }
    .info-table td {
        padding: 5px 10px;
        vertical-align: top;
    }
    .info-table td:first-child {
        font-weight: bold;
        width: 40%;
        color: #666;
    }
    .family-member-item {
        padding: 8px;
        margin: 5px 0;
        background-color: #f9f9f9;
        border-radius: 3px;
        border-left: 3px solid #337ab7;
    }
    
    @media print {
        /* Hide non-printable elements */
        .btn, 
        .page-header .pull-right,
        .no-print,
        .expand-btn { 
            display: none !important; 
        }
        
        /* Make text-muted visible but lighter for printing */
        .text-muted {
            color: #666 !important;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }
        
        /* General print formatting - COMPACT */
        body {
            font-size: 10px !important;
            line-height: 1.2 !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        
        /* Page setup */
        @page {
            margin: 0.5in !important;
            size: letter portrait;
        }
        
        /* Compact header */
        .page-header {
            margin: 0 0 10px 0 !important;
            padding: 0 0 5px 0 !important;
            border-bottom: 1px solid #000 !important;
            page-break-after: avoid;
        }
        
        .page-header h1 {
            font-size: 18px !important;
            margin: 0 !important;
            line-height: 1.2 !important;
        }
        
        .page-header h1 small {
            font-size: 12px !important;
            display: inline !important;
        }
        
        /* Compact panels */
        .panel {
            border: 1px solid #000 !important;
            box-shadow: none !important;
            margin-bottom: 10px !important;
            page-break-inside: avoid;
        }
        
        .panel-heading {
            background-color: #f5f5f5 !important;
            padding: 3px 8px !important;
            border-bottom: 1px solid #000 !important;
        }
        
        .panel-title {
            font-size: 12px !important;
            font-weight: bold !important;
            margin: 0 !important;
        }
        
        .panel-body {
            background-color: #fff !important;
            padding: 8px !important;
        }
        
        /* Compact alert boxes */
        .alert {
            border: 1px solid #000 !important;
            background-color: #f5f5f5 !important;
            color: #000 !important;
            padding: 8px !important;
            margin-bottom: 10px !important;
        }
        
        .alert h3 {
            font-size: 14px !important;
            margin: 0 0 5px 0 !important;
        }
        
        .alert h1 {
            font-size: 18px !important;
            margin: 5px 0 !important;
        }
        
        .alert p {
            margin: 2px 0 !important;
            font-size: 10px !important;
        }
        
        /* Compact tables */
        .table {
            border-collapse: collapse !important;
            width: 100% !important;
            margin-bottom: 10px !important;
            font-size: 9px !important;
        }
        
        .table-condensed > thead > tr > th,
        .table-condensed > tbody > tr > th,
        .table-condensed > tfoot > tr > th,
        .table-condensed > thead > tr > td,
        .table-condensed > tbody > tr > td,
        .table-condensed > tfoot > tr > td {
            padding: 2px 4px !important;
        }
        
        .table > thead > tr > th,
        .table > tbody > tr > th,
        .table > tfoot > tr > th,
        .table > thead > tr > td,
        .table > tbody > tr > td,
        .table > tfoot > tr > td {
            padding: 3px 5px !important;
            border: 1px solid #000 !important;
            line-height: 1.2 !important;
        }
        
        .table > thead > tr > th {
            background-color: #ddd !important;
            font-weight: bold !important;
            font-size: 9px !important;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }
        
        /* Remove striping for cleaner look */
        .table-striped > tbody > tr:nth-of-type(odd) {
            background-color: transparent !important;
        }
        
        /* Compact headers */
        h1 { 
            font-size: 16px !important; 
            margin: 8px 0 !important;
        }
        h2 { 
            font-size: 14px !important; 
            margin: 8px 0 5px 0 !important;
        }
        h3 { 
            font-size: 12px !important; 
            margin: 5px 0 !important;
        }
        h4 { 
            font-size: 11px !important; 
            margin: 5px 0 !important;
        }
        
        /* Ensure full width printing */
        .container-fluid {
            width: 100% !important;
            max-width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        
        .row {
            margin: 0 !important;
            page-break-inside: avoid;
        }
        
        .col-md-1, .col-md-2, .col-md-3, .col-md-4, .col-md-5, .col-md-6,
        .col-md-7, .col-md-8, .col-md-9, .col-md-10, .col-md-11, .col-md-12 {
            padding: 0 5px !important;
        }
        
        /* Well compact */
        .well-sm {
            padding: 5px !important;
            margin: 5px 0 !important;
        }
        
        .well-sm p {
            margin: 0 !important;
            font-size: 9px !important;
        }
        
        /* Force page breaks for sections */
        .daily-breakdown-section {
            page-break-before: auto;
        }
        
        .fund-breakdown-section {
            page-break-before: always;
        }
        
        .daily-fund-section {
            page-break-before: always;
        }
        
        /* Hide icons to save space */
        .fa {
            display: none !important;
        }
        
        /* Compact deposit slip */
        .deposit-slip-table td {
            padding: 2px 5px !important;
            font-size: 10px !important;
        }
        
        /* Handwriting areas - keep larger for printing */
        .handwriting-area {
            min-height: 25px !important;
            border-bottom: 1px solid #000 !important;
            padding: 8px !important;
        }
        
        /* Table layout fixes to ensure all columns print */
        .table {
            table-layout: fixed !important;
        }
        
        /* Adjust column widths for daily breakdown */
        #dailyBreakdown th:nth-child(1) { width: 60px !important; } /* Date */
        #dailyBreakdown th:nth-child(2) { width: 50px !important; } /* Day */
        #dailyBreakdown th:nth-child(3) { width: 50px !important; } /* Cash */
        #dailyBreakdown th:nth-child(4) { width: 50px !important; } /* Checks */
        #dailyBreakdown th:nth-child(5) { width: 60px !important; } /* Physical */
        #dailyBreakdown th:nth-child(6) { width: 50px !important; } /* CC */
        #dailyBreakdown th:nth-child(7) { width: 40px !important; } /* Other */
        #dailyBreakdown th:nth-child(8) { width: 60px !important; } /* Grand Total */
        
        /* Adjust column widths for fund breakdown */
        #fundBreakdown th:nth-child(1) { width: 60px !important; } /* Income Acct */
        #fundBreakdown th:nth-child(2) { width: 100px !important; } /* Fund Name */
        #fundBreakdown th:nth-child(3) { width: 45px !important; } /* Cash */
        #fundBreakdown th:nth-child(4) { width: 45px !important; } /* Checks */
        #fundBreakdown th:nth-child(5) { width: 55px !important; } /* Physical */
        #fundBreakdown th:nth-child(6) { width: 45px !important; } /* CC */
        #fundBreakdown th:nth-child(7) { width: 40px !important; } /* Other */
        #fundBreakdown th:nth-child(8) { width: 55px !important; } /* Fund Total */
        
        /* Text wrapping control */
        td {
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        /* Allow wrapping for fund names only */
        #fundBreakdown td:nth-child(2) {
            white-space: normal !important;
            word-wrap: break-word;
        }
    }
    </style>
    '''
    
    # Group data by date
    dates = {}
    funds = {}  # Changed to dict to store fund info
    
    for row in deposit_data:
        deposit_date = safe_get_value(row, 'DepositDate', '')
        if deposit_date not in dates:
            dates[deposit_date] = {
                'day_of_week': safe_get_value(row, 'DayOfWeek', ''),
                'payments': {},
                'funds': {}
            }
        
        message = safe_get_value(row, 'Message', '')
        fund_id = safe_get_value(row, 'FundId', 0)
        fund_name = safe_get_value(row, 'FundName', 'General Fund')
        fund_income_account = safe_get_value(row, 'FundIncomeAccount', '')
        amount = float(safe_get_value(row, 'TotalAmount', 0))
        count = int(safe_get_value(row, 'TransactionCount', 0))
        
        # Store fund info with income account
        fund_key = "{0}|{1}".format(fund_id, fund_name)  # Composite key
        funds[fund_key] = {
            'id': fund_id, 
            'name': fund_name,
            'income_account': fund_income_account
        }
        
        # Categorize payment type using TouchPoint message pattern
        payment_category = classify_payment_type_by_message(message)
        
        # Track by payment type
        if payment_category not in dates[deposit_date]['payments']:
            dates[deposit_date]['payments'][payment_category] = 0
        dates[deposit_date]['payments'][payment_category] += amount
        
        # Track by fund (using composite key)
        if fund_key not in dates[deposit_date]['funds']:
            dates[deposit_date]['funds'][fund_key] = {}
        if payment_category not in dates[deposit_date]['funds'][fund_key]:
            dates[deposit_date]['funds'][fund_key][payment_category] = 0
        dates[deposit_date]['funds'][fund_key][payment_category] += amount
    
    # Render header
    print '''
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">
                <div class="page-header">
                    <h1>Ministry Deposit Report 
                        <small>{} to {}</small>
                        <div class="pull-right">
                            <a href="javascript:history.back()" class="btn btn-default">
                                <i class="fa fa-arrow-left"></i> Back to Filters
                            </a>
    '''.format(format_date_display(start_date), format_date_display(end_date))
    
    if Config.ENABLE_CSV_EXPORT:
        print '''
                            <button onclick="exportToCSV()" class="btn btn-success">
                                <i class="fa fa-download"></i> Export CSV
                            </button>
        '''
    
    print '''
                            <button onclick="window.print()" class="btn btn-primary">
                                <i class="fa fa-print"></i> Print Report
                            </button>
                        </div>
                    </h1>
                    <p class="text-muted">Filters Applied: {}</p>
                </div>
            </div>
        </div>
    '''.format(filter_desc)
    
    # Calculate and render summary
    render_deposit_summary(dates)
    
    # Render deposit slip section
    render_deposit_slip(dates, filter_desc, start_date, end_date)
    
    # Render daily breakdown
    render_daily_breakdown(dates, funds)
    
    # Render fund breakdown
    if Config.SHOW_FUND_BREAKDOWN:
        render_fund_breakdown(dates, funds)
        
    # Render detailed daily fund breakdown (optional - can be very long)
    if Config.SHOW_DAILY_FUND_DETAIL:
        render_daily_fund_breakdown(dates, funds)
    
    print '</div>'  # Close container
    
    # Add JavaScript for export functionality
    if Config.ENABLE_CSV_EXPORT:
        render_export_javascript(deposit_data, start_date, end_date, program_id, division_id, organization_id)
    
    # Add JavaScript for expand/collapse functionality
    if Config.SHOW_TRANSACTION_DETAILS:
        render_expand_collapse_javascript(start_date, end_date, program_id, division_id, organization_id)


def classify_payment_type_by_message(message):
    """Classify payment type based on transaction message (TouchPoint pattern)"""
    if not message:
        return 'Other'  # Changed from 'Unknown' to 'Other'
        
    # Standard payment type prefixes
    payment_types = {
        'CHK': 'Check',
        'CSH': 'Cash', 
        'CC': 'Credit Card',
        'ACH': 'ACH',
        'ADJ': 'Other',
        'FEE': 'Other',
        'REF': 'Other'
    }
    
    message_upper = str(message).upper()
    
    for prefix, ptype in payment_types.items():
        if message_upper.startswith(prefix):
            return ptype
    
    # Special cases
    if 'RESPONSE' in message_upper or 'CARD' in message_upper:
        return 'Credit Card'
    elif 'COUPON' in message_upper:
        return 'Other'
    elif 'ONLINE' in message_upper:
        return 'Credit Card'
    
    return 'Other'  # Changed from 'Unknown' to 'Other'

def extract_payment_note_from_message(message):
    """Extract the payment note/memo from the message field after the payment type prefix"""
    if not message:
        return ''
    
    # Standard payment type prefixes to remove
    prefixes = ['CHK', 'CSH', 'CC', 'ACH', 'ADJ', 'FEE', 'REF']
    
    message_str = str(message).strip()
    
    # Check if message starts with any prefix
    for prefix in prefixes:
        if message_str.upper().startswith(prefix):
            # Remove the prefix and look for pipe delimiter
            remainder = message_str[len(prefix):].strip()
            # Split by pipe character to get the note portion
            if '|' in remainder:
                parts = remainder.split('|', 1)
                if len(parts) > 1:
                    return parts[1].strip()
                else:
                    return parts[0].strip()
            else:
                # No pipe found, return remainder after cleaning
                while remainder and remainder[0] in ['-', ':', ' ', '_']:
                    remainder = remainder[1:].strip()
                return remainder
    
    # If no standard prefix found, check for pipe delimiter anyway
    if '|' in message_str:
        parts = message_str.split('|', 1)
        if len(parts) > 1:
            return parts[1].strip()
    
    # Return the whole message if no pattern matches
    return message_str

def render_deposit_summary(dates):
    """Render the summary totals section"""
    total_cash = 0.0
    total_check = 0.0
    total_credit_card = 0.0
    total_other = 0.0
    total_days = len(dates)
    check_count = 0
    
    for date_data in dates.values():
        total_cash += float(date_data['payments'].get('Cash', 0))
        total_check += float(date_data['payments'].get('Check', 0))
        total_credit_card += float(date_data['payments'].get('Credit Card', 0))
        total_other += float(date_data['payments'].get('ACH', 0)) + float(date_data['payments'].get('Other', 0))
    
    total_deposit = total_cash + total_check  # What actually gets deposited
    total_all = total_cash + total_check + total_credit_card + total_other
    
    print '''
    <div class="row">
        <div class="col-md-12">
            <div class="alert alert-info" style="background: #e3f2fd; border-color: #1976d2;">
                <h3 style="margin-top: 0;"><i class="fa fa-bank"></i> BANK DEPOSIT REQUIRED</h3>
                <div class="row">
                    <div class="col-md-4">
                        <h1 style="margin: 10px 0; color: #1976d2;">${0:,.2f}</h1>
                        <p style="font-size: 18px; margin: 0;"><strong>Total to Deposit at Bank</strong></p>
                        <p class="text-muted">Cash + Checks</p>
                    </div>
                    <div class="col-md-8">
                        <table class="table table-condensed" style="margin-bottom: 0;">
                            <tr>
                                <td style="border-top: none;"><strong>Cash Amount:</strong></td>
                                <td style="border-top: none; text-align: right;"><strong>${1:,.2f}</strong></td>
                                <td style="border-top: none;" rowspan="3" width="50%">
                                    <div style="padding-left: 20px;">
                                        <p style="margin: 0;"><strong>Electronic Payments (Reference Only):</strong></p>
                                        <p style="margin: 0;">Credit Cards: ${3:,.2f}</p>
                                        <p style="margin: 0;">ACH/Other: ${4:,.2f}</p>
                                        <p style="margin: 0;"><small class="text-muted">These are processed electronically - do not deposit</small></p>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td><strong>Check Amount:</strong></td>
                                <td style="text-align: right;"><strong>${2:,.2f}</strong></td>
                            </tr>
                            <tr>
                                <td><strong>Period:</strong></td>
                                <td style="text-align: right;"><strong>{5} days</strong></td>
                            </tr>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
    '''.format(
        total_deposit,
        total_cash, 
        total_check,
        total_credit_card,
        total_other,
        total_days
    )

def render_deposit_slip(dates, filter_desc, start_date, end_date):
    """Render deposit slip style summary"""
    # Calculate totals
    total_cash = 0.0
    total_check = 0.0
    check_count = 0
    
    for date_data in dates.values():
        total_cash += float(date_data['payments'].get('Cash', 0))
        total_check += float(date_data['payments'].get('Check', 0))
    
    print '''
    <div class="row">
        <div class="col-md-12">
            <div class="panel panel-primary">
                <div class="panel-heading">
                    <h3 class="panel-title"><i class="fa fa-file-text"></i> Deposit Slip Summary</h3>
                </div>
                <div class="panel-body">
                    <div class="row">
                        <div class="col-md-6">
                            <h4>Deposit Information</h4>
                            <table class="table table-condensed">
                                <tr>
                                    <td><strong>Date Range:</strong></td>
                                    <td>{0} to {1}</td>
                                </tr>
                                <tr>
                                    <td><strong>Ministry:</strong></td>
                                    <td>{2}</td>
                                </tr>
                                <tr>
                                    <td style="padding: 10px 8px;"><strong>Prepared By:</strong></td>
                                    <td class="handwriting-area" style="padding: 10px 8px;">
                                        <div style="min-height: 25px;">&nbsp;</div>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 10px 8px;"><strong>Prepared Date:</strong></td>
                                    <td class="handwriting-area" style="padding: 10px 8px;">
                                        <div style="min-height: 25px;">&nbsp;</div>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <div class="col-md-6">
                            <h4>Deposit Breakdown</h4>
                            <table class="table table-bordered">
                                <tr>
                                    <td width="50%"><strong>CASH</strong></td>
                                    <td class="text-right"><strong>${3:,.2f}</strong></td>
                                </tr>
                                <tr>
                                    <td><strong>CHECKS</strong></td>
                                    <td class="text-right"><strong>${4:,.2f}</strong></td>
                                </tr>
                                <tr class="success">
                                    <td><strong>TOTAL DEPOSIT</strong></td>
                                    <td class="text-right"><h4 style="margin: 0;">${5:,.2f}</h4></td>
                                </tr>
                            </table>
                            <div class="well well-sm">
                                <p class="text-center" style="margin: 0;">
                                    <strong>For Bank Use Only</strong><br>
                                    Verified By: _________________ Date: _________
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    '''.format(
        format_date_display(start_date),
        format_date_display(end_date),
        filter_desc,
        total_cash,
        total_check,
        total_cash + total_check
    )

def render_daily_breakdown(dates, funds):
    """Render daily breakdown table"""
    # Build header with optional expand/collapse buttons
    header_html = '<h2>Daily Breakdown'
    if Config.SHOW_TRANSACTION_DETAILS:
        header_html += '''
                <div class="pull-right">
                    <button class="btn btn-xs btn-default" onclick="expandAllDays()">
                        <i class="fa fa-plus"></i> Expand All
                    </button>
                    <button class="btn btn-xs btn-default" onclick="collapseAllDays()">
                        <i class="fa fa-minus"></i> Collapse All
                    </button>
                </div>'''
    header_html += '</h2>'
    
    print '''
    <div class="row">
        <div class="col-md-12">
            {0}
            <div class="table-responsive">
                <table class="table table-striped table-bordered" id="dailyBreakdown">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Day</th>
                            <th class="text-right">Cash</th>
                            <th class="text-right">Checks</th>
                            <th class="text-right">Physical Total</th>
                            <th class="text-right">Credit Card</th>
                            <th class="text-right">Other</th>
                            <th class="text-right">Grand Total</th>
                        </tr>
                    </thead>
                    <tbody>
    '''.format(header_html)
    
    # Sort dates in ascending order (oldest to newest)
    sorted_dates = sorted(dates.keys())
    
    for deposit_date in sorted_dates:
        date_data = dates[deposit_date]
        
        cash = float(date_data['payments'].get('Cash', 0))
        check = float(date_data['payments'].get('Check', 0))
        credit_card = float(date_data['payments'].get('Credit Card', 0))
        ach = float(date_data['payments'].get('ACH', 0))
        other = float(date_data['payments'].get('Other', 0))
        
        physical_total = cash + check
        other_total = ach + other
        grand_total = physical_total + credit_card  # Exclude "Other" from grand total
        
        # Skip rows with zero amounts if configured
        if not Config.SHOW_ZERO_AMOUNTS and grand_total == 0:
            continue
        
        # Format date for ID - ensure we have a proper date string first
        date_str = str(deposit_date)
        if hasattr(deposit_date, 'ToString'):
            date_str = deposit_date.ToString('yyyy-MM-dd')
        # Remove hyphens to create ID
        date_id = date_str.replace('-', '')
        
        # Build the date cell content
        date_cell = format_date_display(str(deposit_date))
        if Config.SHOW_TRANSACTION_DETAILS:
            date_cell += '''
                                <span class="expand-btn" onclick="toggleDayDetails('{0}')" id="expand-btn-{0}" data-date="{1}">
                                    <i class="fa fa-plus-circle"></i> Details
                                </span>
            '''.format(date_id, date_str)
        
        print '''
                        <tr>
                            <td>{0}</td>
                            <td>{1}</td>
                            <td class="text-right">${2:,.2f}</td>
                            <td class="text-right">${3:,.2f}</td>
                            <td class="text-right"><strong>${4:,.2f}</strong></td>
                            <td class="text-right text-muted">${5:,.2f}</td>
                            <td class="text-right text-muted">${6:,.2f}</td>
                            <td class="text-right"><strong>${7:,.2f}</strong></td>
                        </tr>
        '''.format(
            date_cell,
            date_data['day_of_week'],
            cash,
            check,
            physical_total,
            credit_card,
            other_total,
            grand_total
        )
        
        # Only add detail row if transaction details are enabled
        if Config.SHOW_TRANSACTION_DETAILS:
            print '''
                        <tr class="detail-row" id="detail-row-{0}">
                            <td colspan="8" class="detail-content" id="detail-content-{0}">
                                <div class="text-center">
                                    <i class="fa fa-spinner fa-spin"></i> Loading transaction details...
                                </div>
                            </td>
                        </tr>
            '''.format(date_id)
    
    print '''
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    '''

def render_fund_breakdown(dates, funds):
    """Render fund-by-fund breakdown with fund IDs"""
    print '''
    <div class="row">
        <div class="col-md-12">
            <h2>Fund Breakdown</h2>
            <p class="text-muted">Shows how deposits are allocated across different funds</p>
        </div>
    </div>
    '''
    
    # Calculate fund totals across all dates
    fund_totals = {}
    for fund_key in funds:
        fund_totals[fund_key] = {'Cash': 0.0, 'Check': 0.0, 'Credit Card': 0.0, 'Other': 0.0, 'ACH': 0.0}
        
        for date_data in dates.values():
            if fund_key in date_data['funds']:
                for payment_type, amount in date_data['funds'][fund_key].items():
                    if payment_type == 'ACH' or payment_type == 'Other':
                        fund_totals[fund_key]['Other'] += float(amount)
                    elif payment_type in fund_totals[fund_key]:
                        fund_totals[fund_key][payment_type] += float(amount)
    
    print '''
    <div class="row">
        <div class="col-md-12">
            <div class="table-responsive">
                <table class="table table-striped table-bordered" id="fundBreakdown">
                    <thead>
                        <tr>
                            <th>Income Account</th>
                            <th>Fund Name</th>
                            <th class="text-right">Cash</th>
                            <th class="text-right">Checks</th>
                            <th class="text-right">Physical Total</th>
                            <th class="text-right">Credit Card</th>
                            <th class="text-right">Other</th>
                            <th class="text-right">Fund Total</th>
                        </tr>
                    </thead>
                    <tbody>
    '''
    
    # Sort funds by fund name
    sorted_fund_keys = sorted(funds.keys(), key=lambda k: funds[k]['name'])
    
    for fund_key in sorted_fund_keys:
        fund_info = funds[fund_key]
        cash = fund_totals[fund_key]['Cash']
        check = fund_totals[fund_key]['Check']
        credit_card = fund_totals[fund_key]['Credit Card']
        other = fund_totals[fund_key]['Other']
        
        physical_total = cash + check
        fund_total = physical_total + credit_card  # Exclude "Other" from fund total
        
        # Skip funds with zero amounts if configured
        if not Config.SHOW_ZERO_AMOUNTS and fund_total == 0:
            continue
        
        print '''
                        <tr>
                            <td>{0}</td>
                            <td><strong>{1}</strong></td>
                            <td class="text-right">${2:,.2f}</td>
                            <td class="text-right">${3:,.2f}</td>
                            <td class="text-right"><strong>${4:,.2f}</strong></td>
                            <td class="text-right text-muted">${5:,.2f}</td>
                            <td class="text-right text-muted">${6:,.2f}</td>
                            <td class="text-right"><strong>${7:,.2f}</strong></td>
                        </tr>
        '''.format(
            fund_info.get('income_account', ''),
            fund_info['name'],
            cash,
            check,
            physical_total,
            credit_card,
            other,
            fund_total
        )
    
    print '''
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    '''

def render_daily_fund_breakdown(dates, funds):
    """Render compact daily breakdown by fund"""
    print '''
    <div class="row">
        <div class="col-md-12">
            <h2>Daily Fund Detail</h2>
            <p class="text-muted">Compact view of daily deposits by fund</p>
            <div class="table-responsive">
                <table class="table table-bordered table-condensed">
                    <thead>
                        <tr>
                            <th rowspan="2">Date</th>
                            <th rowspan="2">Fund</th>
                            <th rowspan="2">Income Acct</th>
                            <th colspan="2" class="text-center">Physical Deposit</th>
                            <th colspan="2" class="text-center">Electronic</th>
                            <th rowspan="2" class="text-right">Total</th>
                        </tr>
                        <tr>
                            <th class="text-right">Cash</th>
                            <th class="text-right">Checks</th>
                            <th class="text-right">Credit Card</th>
                            <th class="text-right">Other</th>
                        </tr>
                    </thead>
                    <tbody>
    '''
    
    # Sort dates in descending order
    sorted_dates = sorted(dates.keys(), reverse=True)
    
    for deposit_date in sorted_dates:
        date_data = dates[deposit_date]
        
        # Check if this date has any funds
        if not date_data['funds']:
            continue
        
        # Sort fund keys for this date
        sorted_fund_keys = sorted(date_data['funds'].keys(), key=lambda k: funds[k]['name'] if k in funds else k)
        
        first_row = True
        date_row_count = 0
        
        # Count non-zero rows for this date
        for fund_key in sorted_fund_keys:
            if fund_key in funds:
                payments = date_data['funds'][fund_key]
                fund_total = sum(float(amount) for amount in payments.values())
                if Config.SHOW_ZERO_AMOUNTS or fund_total != 0:
                    date_row_count += 1
        
        if date_row_count == 0:
            continue
        
        # Render rows for this date
        for fund_key in sorted_fund_keys:
            if fund_key not in funds:
                continue
                
            fund_info = funds[fund_key]
            payments = date_data['funds'][fund_key]
            
            cash = float(payments.get('Cash', 0))
            check = float(payments.get('Check', 0))
            credit_card = float(payments.get('Credit Card', 0))
            other = float(payments.get('ACH', 0)) + float(payments.get('Other', 0))
            fund_total = cash + check + credit_card  # Exclude "Other" from fund total
            
            if not Config.SHOW_ZERO_AMOUNTS and fund_total == 0:
                continue
            
            print '<tr>'
            
            # Date column (only on first row for each date)
            if first_row:
                print '<td rowspan="{0}" style="vertical-align: middle;"><strong>{1}</strong><br><small>{2}</small></td>'.format(
                    date_row_count,
                    format_date_display(str(deposit_date)),
                    date_data['day_of_week']
                )
                first_row = False
            
            # Fund details
            print '''
                            <td>{0}</td>
                            <td>{1}</td>
                            <td class="text-right">{2}</td>
                            <td class="text-right">{3}</td>
                            <td class="text-right text-muted">{4}</td>
                            <td class="text-right text-muted">{5}</td>
                            <td class="text-right"><strong>${6:,.2f}</strong></td>
                        </tr>
            '''.format(
                fund_info['name'],
                fund_info.get('income_account', ''),
                '${:,.2f}'.format(cash) if cash > 0 else '-',
                '${:,.2f}'.format(check) if check > 0 else '-',
                '${:,.2f}'.format(credit_card) if credit_card > 0 else '-',
                '${:,.2f}'.format(other) if other > 0 else '-',
                fund_total
            )
    
    print '''
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    '''

def render_export_javascript(deposit_data, start_date, end_date, program_id='', division_id='', organization_id=''):
    """Render JavaScript for CSV export functionality"""
    print '''
    <script>
    function exportToCSV() {{
        // Create form to submit export request
        var form = document.createElement('form');
        form.method = 'POST';
        form.action = window.location.href;
        
        var actionInput = document.createElement('input');
        actionInput.type = 'hidden';
        actionInput.name = 'action';
        actionInput.value = 'export_csv';
        form.appendChild(actionInput);
        
        var startDateInput = document.createElement('input');
        startDateInput.type = 'hidden';
        startDateInput.name = 'start_date';
        startDateInput.value = '{0}';
        form.appendChild(startDateInput);
        
        var endDateInput = document.createElement('input');
        endDateInput.type = 'hidden';
        endDateInput.name = 'end_date';
        endDateInput.value = '{1}';
        form.appendChild(endDateInput);
    '''.format(start_date, end_date)
    
    # Add filter parameters if they exist
    if program_id:
        print '''
        var programInput = document.createElement('input');
        programInput.type = 'hidden';
        programInput.name = 'program_id';
        programInput.value = '{0}';
        form.appendChild(programInput);
        '''.format(program_id)
    
    if division_id:
        print '''
        var divisionInput = document.createElement('input');
        divisionInput.type = 'hidden';
        divisionInput.name = 'division_id';
        divisionInput.value = '{0}';
        form.appendChild(divisionInput);
        '''.format(division_id)
    
    if organization_id:
        print '''
        var orgInput = document.createElement('input');
        orgInput.type = 'hidden';
        orgInput.name = 'organization_id';
        orgInput.value = '{0}';
        form.appendChild(orgInput);
        '''.format(organization_id)
    
    print '''
        document.body.appendChild(form);
        form.submit();
        document.body.removeChild(form);
    }
    </script>
    '''

def export_csv_data():
    """Export deposit data as CSV"""
    # Get parameters (same as generate_report)
    start_date = getattr(model.Data, 'start_date', '')
    end_date = getattr(model.Data, 'end_date', '')
    program_id = getattr(model.Data, 'program_id', '')
    division_id = getattr(model.Data, 'division_id', '')
    organization_id = getattr(model.Data, 'organization_id', '')
    
    # Get the data
    deposit_data = get_deposit_data(start_date, end_date, program_id, division_id, organization_id)
    
    if not deposit_data:
        print 'No data to export'
        return
    
    # Set CSV headers
    filename = "{}_{}_to_{}.csv".format(Config.CSV_FILENAME_PREFIX, start_date, end_date)
    model.Header = 'Content-Type: text/csv'
    model.Header = 'Content-Disposition: attachment; filename="{}"'.format(filename)
    
    # Generate CSV content
    print "Date,Day of Week,Payment Type,Fund Name,Transaction Count,Amount"
    
    for row in deposit_data:
        print '"{0}","{1}","{2}","{3}",{4},{5:.2f}'.format(
            safe_get_value(row, 'DepositDate', ''),
            safe_get_value(row, 'DayOfWeek', ''),
            safe_get_value(row, 'PaymentType', ''),
            safe_get_value(row, 'FundName', ''),
            safe_get_value(row, 'TransactionCount', 0),
            float(safe_get_value(row, 'TotalAmount', 0))
        )

def format_date_display(date_str):
    """Format date for display"""
    try:
        # Handle datetime objects directly
        if hasattr(date_str, 'ToString'):
            return date_str.ToString("MMM d, yyyy")
        
        # Handle string dates
        if isinstance(date_str, str):
            # Remove any time portion if present
            if ' ' in date_str:
                date_str = date_str.split(' ')[0]
            
            # Parse and format consistently
            if len(date_str) >= 10:
                date_obj = model.ParseDate(date_str[:10])
                return date_obj.ToString("MMM d, yyyy")
    except:
        pass
    
    # Fallback - return as string
    return str(date_str).split(' ')[0] if ' ' in str(date_str) else str(date_str)

def safe_get_value(obj, attr, default=0):
    """Safely get value from TouchPoint row object"""
    try:
        if hasattr(obj, attr):
            val = getattr(obj, attr)
            if val is not None:
                return val
    except:
        pass
    return default

def get_day_transaction_details():
    """Get detailed transactions for a specific day via AJAX"""
    # Set content type for HTML response
    model.Header = 'Content-Type: text/html'
    
    # Get parameters
    deposit_date = getattr(model.Data, 'deposit_date', '')
    program_id = getattr(model.Data, 'program_id', '')
    division_id = getattr(model.Data, 'division_id', '')
    organization_id = getattr(model.Data, 'organization_id', '')
    
    if not deposit_date:
        print '<div class="alert alert-danger">No date specified</div>'
        return
    
    # Debug logging
    if model.UserIsInRole("Developer"):
        print '<div class="alert alert-info">Debug: Raw deposit_date = {} (type: {})</div>'.format(
            deposit_date, type(deposit_date).__name__)
    
    # Validate the date format
    try:
        # Clean the date string
        deposit_date = str(deposit_date).strip()
        
        # Since JavaScript sends YYYY-MM-DD format, validate it
        if len(deposit_date) != 10 or deposit_date[4] != '-' or deposit_date[7] != '-':
            print '<div class="alert alert-danger">Invalid date format. Expected YYYY-MM-DD, got: {}</div>'.format(deposit_date)
            return
        
        # Parse to validate it's a real date
        parsed_date = model.ParseDate(deposit_date)
        
        if model.UserIsInRole("Developer"):
            print '<div class="alert alert-info">Debug: Input date = {}, Parsed = {}</div>'.format(
                deposit_date, parsed_date.ToString('yyyy-MM-dd'))
        
    except Exception as e:
        print '<div class="alert alert-danger">Date validation error: {} (input: {})</div>'.format(str(e), deposit_date)
        return
    
    # Build WHERE clause for filters
    where_clauses = []
    join_clauses = []
    
    # Add ministry filters
    if organization_id:
        where_clauses.append("t.OrgId = {}".format(int(organization_id)))
    elif division_id:
        # Check if we're joining through TransactionSummary or directly
        where_clauses.append("(o.DivisionId = {} OR (ts.OrganizationId IS NOT NULL AND o2.DivisionId = {}))".format(int(division_id), int(division_id)))
        join_clauses.append("LEFT JOIN Organizations o2 ON ts.OrganizationId = o2.OrganizationId")
    elif program_id:
        # Organizations is already joined as 'o', just need Division
        join_clauses.append("LEFT JOIN Division d ON o.DivisionId = d.Id")
        join_clauses.append("LEFT JOIN Organizations o2 ON ts.OrganizationId = o2.OrganizationId")
        join_clauses.append("LEFT JOIN Division d2 ON o2.DivisionId = d2.Id")
        where_clauses.append("(d.ProgId = {} OR d2.ProgId = {})".format(int(program_id), int(program_id)))
    
    # Get individual transactions for the date
    # Use date range comparison for most reliable results
    start_datetime = deposit_date + ' 00:00:00'
    end_datetime = deposit_date + ' 23:59:59.999'
    
    sql = """
    SELECT 
        t.TransactionId,
        t.TransactionDate,
        t.Name AS PaidBy,
        t.amt AS Amount,
        t.Message,
        t.Description,
        t.OrgId,
        -- Get who the payment was for
        CASE 
            WHEN ts.PeopleId IS NOT NULL AND p.Name2 != t.Name 
                THEN p.Name2
            ELSE NULL
        END AS PaidFor,
        ts.PeopleId AS PaidForId,
        ISNULL(o.OrganizationName, 'General') AS OrganizationName,
        CASE 
            WHEN o.RegAccountCodeId IS NOT NULL THEN CAST(o.RegAccountCodeId AS NVARCHAR(50))
            ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'NVARCHAR(50)')
        END AS AccountingCode,
        ISNULL(ac.Description, 'General Fund') AS FundName,
        ISNULL(ac.Code, '') AS FundIncomeAccount
    FROM [Transaction] t
    LEFT JOIN TransactionSummary ts ON t.OriginalId = ts.RegId AND ts.IsLatestTransaction = 1
    LEFT JOIN People p ON ts.PeopleId = p.PeopleId
    LEFT JOIN Organizations o ON t.OrgId = o.OrganizationId
    LEFT JOIN lookup.AccountCode ac ON ac.Id = CASE 
        WHEN o.RegAccountCodeId IS NOT NULL THEN CAST(o.RegAccountCodeId AS NVARCHAR(50))
        ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'NVARCHAR(50)')
    END
    {0}
    WHERE t.TransactionDate >= CONVERT(datetime, '{1}', 120)
      AND t.TransactionDate <= CONVERT(datetime, '{2}', 120)
      AND t.amt <> 0
      AND t.TransactionId IS NOT NULL
      AND t.voided IS NULL
      {3}
    ORDER BY t.TransactionDate DESC, t.TransactionId DESC
    """.format(
        ' '.join(join_clauses),
        start_datetime,
        end_datetime,
        'AND ' + ' AND '.join(where_clauses) if where_clauses else ''
    )
    
    transactions = q.QuerySql(sql)
    
    if not transactions:
        print '<p class="text-muted">No transactions found for this date.</p>'
        return
    
    # Group transactions by fund
    funds_data = {}
    grand_total = 0.0
    
    for trans in transactions:
        fund_name = safe_get_value(trans, 'FundName', 'General Fund')
        fund_income_account = safe_get_value(trans, 'FundIncomeAccount', '')
        
        # Create fund key with income account for unique identification
        fund_key = "{0}|{1}".format(fund_name, fund_income_account)
        
        if fund_key not in funds_data:
            funds_data[fund_key] = {
                'name': fund_name,
                'income_account': fund_income_account,
                'transactions': [],
                'totals': {'Cash': 0.0, 'Check': 0.0, 'Credit Card': 0.0, 'Other': 0.0}
            }
        
        # Categorize payment type
        message = safe_get_value(trans, 'Message', '')
        payment_type = classify_payment_type_by_message(message)
        amount = float(safe_get_value(trans, 'Amount', 0))
        
        # Add to fund totals
        if payment_type in funds_data[fund_key]['totals']:
            funds_data[fund_key]['totals'][payment_type] += amount
        else:
            funds_data[fund_key]['totals']['Other'] += amount
        
        # Add transaction to fund
        funds_data[fund_key]['transactions'].append(trans)
        grand_total += amount
    
    # Sort funds by name
    sorted_fund_keys = sorted(funds_data.keys(), key=lambda k: funds_data[k]['name'])
    
    # Render summary by fund with enhanced visual styling
    print '''
    <div style="background: #f0f4f8; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <h5 style="margin-top: 0; color: #2c3e50;">Summary by Fund for {0}</h5>
        <table class="table table-bordered table-condensed" style="background: white;">
            <thead style="background: #e8ecf1;">
                <tr>
                    <th>Income Account</th>
                    <th>Fund Name</th>
                    <th class="text-right">Cash</th>
                    <th class="text-right">Checks</th>
                    <th class="text-right" style="background: #d4e6f1;">Physical Total</th>
                    <th class="text-right">Credit Card</th>
                    <th class="text-right">Other</th>
                    <th class="text-right">Fund Total</th>
                </tr>
            </thead>
            <tbody>
    '''.format(parsed_date.ToString('MMMM d, yyyy'))
    
    # Calculate grand totals for summary
    grand_cash = 0.0
    grand_check = 0.0
    grand_cc = 0.0
    grand_other = 0.0
    
    # Display fund summaries
    for fund_key in sorted_fund_keys:
        fund_data = funds_data[fund_key]
        cash_total = fund_data['totals']['Cash']
        check_total = fund_data['totals']['Check']
        cc_total = fund_data['totals']['Credit Card']
        other_total = fund_data['totals']['Other']
        physical_total = cash_total + check_total
        fund_total = physical_total + cc_total  # Exclude "Other" from fund total
        
        # Add to grand totals
        grand_cash += cash_total
        grand_check += check_total
        grand_cc += cc_total
        grand_other += other_total
        
        print '''
                <tr>
                    <td>{0}</td>
                    <td><strong>{1}</strong></td>
                    <td class="text-right">${2:,.2f}</td>
                    <td class="text-right">${3:,.2f}</td>
                    <td class="text-right" style="background: #f0f8ff;"><strong>${4:,.2f}</strong></td>
                    <td class="text-right text-muted">${5:,.2f}</td>
                    <td class="text-right text-muted">${6:,.2f}</td>
                    <td class="text-right"><strong>${7:,.2f}</strong></td>
                </tr>
        '''.format(
            fund_data['income_account'],
            fund_data['name'],
            cash_total,
            check_total,
            physical_total,
            cc_total,
            other_total,
            fund_total
        )
    
    # Calculate grand physical total and grand total (excluding other)
    grand_physical = grand_cash + grand_check
    grand_total = grand_physical + grand_cc  # Exclude "Other" from grand total
    
    print '''
            </tbody>
            <tfoot style="background: #e8ecf1; font-weight: bold;">
                <tr>
                    <th colspan="2" class="text-right">Grand Total:</th>
                    <th class="text-right">${0:,.2f}</th>
                    <th class="text-right">${1:,.2f}</th>
                    <th class="text-right" style="background: #d4e6f1;">${2:,.2f}</th>
                    <th class="text-right text-muted">${3:,.2f}</th>
                    <th class="text-right text-muted">${4:,.2f}</th>
                    <th class="text-right">${5:,.2f}</th>
                </tr>
            </tfoot>
        </table>
        <div style="margin-top: 10px; padding: 10px; background: #ffeaa7; border-radius: 4px;">
            <strong>Note:</strong> "Other" includes miscellaneous adjustment types and is shown for reference only. This column may include actual cash or check payments only when the transaction was not entered through payment manager or was entered manually without the proper CSH or CHK type prefix.
        </div>
    </div>
    '''.format(grand_cash, grand_check, grand_physical, grand_cc, grand_other, grand_total)
    
    # Render detailed transactions by fund
    print '<h5>Transaction Details by Fund</h5>'
    
    for fund_key in sorted_fund_keys:
        fund_data = funds_data[fund_key]
        fund_total = sum(fund_data['totals'].values())
        
        # Generate unique table ID for this fund
        fund_table_id = "trans-table-{0}".format(fund_key.replace('|', '-').replace(' ', '-'))
        
        print '''
        <div style="margin-bottom: 30px;">
            <h6 style="background-color: #f5f5f5; padding: 10px; margin-bottom: 0;">
                {0} - {1} (${2:,.2f})
            </h6>
            <table class="table table-condensed table-hover" id="{3}" style="margin-bottom: 0;">
                <thead>
                    <tr>
                        <th class="sortable" data-sort="time">Time <i class="fa fa-sort"></i></th>
                        <th class="sortable" data-sort="id">Transaction ID <i class="fa fa-sort"></i></th>
                        <th class="sortable" data-sort="paidby">Paid By <i class="fa fa-sort"></i></th>
                        <th class="sortable" data-sort="paidfor">Paid For <i class="fa fa-sort"></i></th>
                        <th>Description</th>
                        <th class="sortable" data-sort="type">Type <i class="fa fa-sort"></i></th>
                        <th>Payment Note</th>
                        <th class="text-right sortable" data-sort="amount">Amount <i class="fa fa-sort"></i></th>
                    </tr>
                </thead>
                <tbody>
        '''.format(fund_data['income_account'], fund_data['name'], fund_total, fund_table_id)
        
        # Sort transactions by time
        sorted_transactions = sorted(fund_data['transactions'], 
                                   key=lambda t: safe_get_value(t, 'TransactionDate', ''))
        
        for trans in sorted_transactions:
            trans_time = safe_get_value(trans, 'TransactionDate', '')
            if trans_time:
                try:
                    trans_time = trans_time.ToString("h:mm tt")
                except:
                    try:
                        trans_time = model.ParseDate(str(trans_time)).ToString("h:mm tt")
                    except:
                        trans_time = str(trans_time)
            
            message = safe_get_value(trans, 'Message', '')
            payment_type = classify_payment_type_by_message(message)
            payment_note = extract_payment_note_from_message(message)
            amount = float(safe_get_value(trans, 'Amount', 0))
            
            # Style based on payment type
            type_class = ''
            if payment_type == 'Cash':
                type_class = 'text-success'
            elif payment_type == 'Check':
                type_class = 'text-primary'
            elif payment_type == 'Credit Card':
                type_class = 'text-warning'
            
            paid_for_value = safe_get_value(trans, 'PaidFor', '')
            paid_for_id = safe_get_value(trans, 'PaidForId', '')
            if paid_for_value and paid_for_id:
                paid_for_display = '<a href="#" onclick="showPersonPopup({0}, \'{1}\'); return false;" style="cursor: pointer; text-decoration: underline;">{1}</a>'.format(
                    paid_for_id, paid_for_value.replace("'", "\\'"))
            elif paid_for_value:
                paid_for_display = paid_for_value
            else:
                paid_for_display = '<span class="text-muted">-</span>'
                paid_for_value = ''  # Empty for sorting
            
            # Get time in sortable format (HHMMSS)
            trans_date_obj = safe_get_value(trans, 'TransactionDate', '')
            time_sort = '000000'
            if trans_date_obj:
                try:
                    time_sort = trans_date_obj.ToString("HHmmss")
                except:
                    time_sort = '000000'
            
            print '''
                <tr data-time="{0}" data-amount="{1}" data-paidby="{2}" data-paidfor="{3}">
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6}</td>
                    <td>{7}</td>
                    <td>{8}</td>
                    <td class="{9}">{10}</td>
                    <td>{11}</td>
                    <td class="text-right"><strong>${12:,.2f}</strong></td>
                </tr>
            '''.format(
                time_sort,
                amount,
                safe_get_value(trans, 'PaidBy', '').replace('"', '&quot;'),
                paid_for_value.replace('"', '&quot;'),
                trans_time,
                safe_get_value(trans, 'TransactionId', ''),
                safe_get_value(trans, 'PaidBy', ''),
                paid_for_display,
                safe_get_value(trans, 'Description', ''),
                type_class,
                payment_type,
                payment_note,
                amount
            )
        
        print '''
                </tbody>
            </table>
        </div>
        '''
    
    print '''
    <div style="margin-top: 10px;">
        <small class="text-muted">
            <i class="fa fa-info-circle"></i> Showing {} transactions for {} grouped by fund. 
            Click on column headers to sort (Time, Transaction ID, Paid By, Paid For, Type, Amount).
        </small>
    </div>
    
    <script>
    // Verify sorting is ready for these tables
    (function() {{
        if (typeof jQuery !== 'undefined') {{
            console.log('jQuery is available in AJAX content');
            
            // Find all sortable headers
            var allHeaders = jQuery('th');
            console.log('Total th elements found:', allHeaders.length);
            
            var sortableHeaders = jQuery('th.sortable');
            console.log('Found ' + sortableHeaders.length + ' sortable headers (th.sortable)');
            
            // Check specifically in tables
            var tables = jQuery('table');
            console.log('Found ' + tables.length + ' tables');
            tables.each(function(idx) {{
                var table = jQuery(this);
                var tableId = table.attr('id') || 'no-id';
                var headersInTable = table.find('th.sortable');
                console.log('Table ' + idx + ' (id: ' + tableId + ') has ' + headersInTable.length + ' sortable headers');
            }});
            
            // Log details of each sortable header
            sortableHeaders.each(function() {{
                var $header = jQuery(this);
                console.log('Sortable header:', $header.text().trim(), 
                           'Sort key:', $header.attr('data-sort'),
                           'Classes:', $header.attr('class'));
            }});
            
            // Make sortable headers visually distinct
            sortableHeaders.css('background-color', '#f0f0f0');
            
        }} else {{
            console.error('jQuery is NOT available in AJAX content!');
        }}
    }})();
    </script>
    '''.format(len(list(transactions)), parsed_date.ToString('MMMM d, yyyy'))

def render_expand_collapse_javascript(start_date, end_date, program_id, division_id, organization_id):
    """Render JavaScript for expand/collapse functionality"""
    print '''
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script type="text/javascript">
    // Wait for jQuery to be ready
    $(document).ready(function() {
        console.log('jQuery ready, initializing Ministry Deposit Report');
        window.expandedDates = {};
        
        // Test if jQuery event delegation is working
        console.log('Setting up sortable header click handler');
        
        // Define getPyScriptAddress if not already defined
        if (typeof window.getPyScriptAddress === "undefined") {
            window.getPyScriptAddress = function() {
                var path = window.location.pathname;
                return path.replace("/PyScript/", "/PyScriptForm/");
            };
        }
        
        // Define toggleDayDetails in global scope
        window.toggleDayDetails = function(dateId) {
            var detailRow = document.getElementById("detail-row-" + dateId);
            var expandBtn = document.getElementById("expand-btn-" + dateId);
            var contentDiv = document.getElementById("detail-content-" + dateId);
            
            if (expandedDates[dateId]) {
                // Collapse
                detailRow.style.display = "none";
                expandBtn.innerHTML = '<i class="fa fa-plus-circle"></i> Details';
                expandedDates[dateId] = false;
            } else {
                // Expand
                detailRow.style.display = "table-row";
                expandBtn.innerHTML = '<i class="fa fa-minus-circle"></i> Hide';
                
                // Load details if not already loaded
                if (!expandedDates[dateId + "_loaded"]) {
                    loadDayDetails(dateId, contentDiv);
                    expandedDates[dateId + "_loaded"] = true;
                }
                
                expandedDates[dateId] = true;
            }
        };
        
        // Define loadDayDetails
        window.loadDayDetails = function(dateId, contentDiv) {
            // Try to get date from data attribute first
            var expandBtn = document.getElementById("expand-btn-" + dateId);
            var dateStr = "";
            
            if (expandBtn && expandBtn.getAttribute("data-date")) {
                dateStr = expandBtn.getAttribute("data-date");
                console.log("Using data-date attribute:", dateStr);
            } else {
                // Fallback: Convert dateId back to date format (YYYYMMDD to YYYY-MM-DD)
                dateId = String(dateId);
                if (dateId.length !== 8) {
                    console.error("Invalid dateId format:", dateId);
                    contentDiv.innerHTML = '<div class="alert alert-danger">Invalid date format.</div>';
                    return;
                }
                dateStr = dateId.substring(0, 4) + "-" + dateId.substring(4, 6) + "-" + dateId.substring(6, 8);
                console.log("Constructed date from ID:", dateStr);
            }
            
            console.log("Loading details for date:", dateStr);
            
            $.ajax({
                url: getPyScriptAddress(),
                method: "POST",
                data: {
                    action: "get_day_details",
                    deposit_date: dateStr,
                    program_id: "''' + str(program_id or '') + '''",
                    division_id: "''' + str(division_id or '') + '''",
                    organization_id: "''' + str(organization_id or '') + '''"
                },
                success: function(html) {
                    contentDiv.innerHTML = html;
                    // Sorting will be handled by the event delegation we set up
                    console.log('Transaction details loaded successfully');
                },
                error: function(xhr, status, error) {
                    console.error("AJAX error:", error);
                    contentDiv.innerHTML = '<div class="alert alert-danger">Error loading transaction details: ' + error + '</div>';
                }
            });
        };
        
        // Table sorting functionality using jQuery event delegation
        // This will work for dynamically loaded content
        $(document).on('click', '.detail-content th.sortable', function(e) {
            console.log('Sortable header clicked in detail content', e);
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            
            var $header = $(this);
            var sortKey = $header.attr('data-sort');
            var $table = $header.closest('table');
            var $tbody = $table.find('tbody');
            var isAscending = !$header.hasClass('sort-asc');
            
            console.log('Sorting by:', sortKey, 'Ascending:', isAscending);
            console.log('Table ID:', $table.attr('id'));
            
            // Remove sort classes from all headers in this table
            $table.find('th.sortable').removeClass('sort-asc sort-desc');
            
            // Add appropriate sort class
            $header.addClass(isAscending ? 'sort-asc' : 'sort-desc');
            
            // Get all rows
            var rows = $tbody.find('tr').toArray();
            
            // Sort rows
            rows.sort(function(a, b) {
                var aValue, bValue;
                var $a = $(a);
                var $b = $(b);
                
                switch(sortKey) {
                    case 'time':
                        aValue = $a.attr('data-time') || '000000';
                        bValue = $b.attr('data-time') || '000000';
                        break;
                    case 'amount':
                        aValue = parseFloat($a.attr('data-amount')) || 0;
                        bValue = parseFloat($b.attr('data-amount')) || 0;
                        break;
                    case 'id':
                        aValue = $a.find('td').eq(1).text().trim();
                        bValue = $b.find('td').eq(1).text().trim();
                        break;
                    case 'paidby':
                        aValue = $a.attr('data-paidby') || '';
                        bValue = $b.attr('data-paidby') || '';
                        break;
                    case 'paidfor':
                        aValue = $a.attr('data-paidfor') || '';
                        bValue = $b.attr('data-paidfor') || '';
                        break;
                    case 'type':
                        aValue = $a.find('td').eq(5).text().trim();
                        bValue = $b.find('td').eq(5).text().trim();
                        break;
                    default:
                        return 0;
                }
                
                if (aValue < bValue) return isAscending ? -1 : 1;
                if (aValue > bValue) return isAscending ? 1 : -1;
                return 0;
            });
            
            // Re-append sorted rows
            $.each(rows, function(index, row) {
                $tbody.append(row);
            });
        });
        
        // Define expand/collapse all functions
        window.expandAllDays = function() {
            var detailRows = document.querySelectorAll(".detail-row");
            detailRows.forEach(function(row) {
                var dateId = row.id.replace("detail-row-", "");
                if (!expandedDates[dateId]) {
                    toggleDayDetails(dateId);
                }
            });
        };
        
        window.collapseAllDays = function() {
            var detailRows = document.querySelectorAll(".detail-row");
            detailRows.forEach(function(row) {
                var dateId = row.id.replace("detail-row-", "");
                if (expandedDates[dateId]) {
                    toggleDayDetails(dateId);
                }
            });
        };
        
        // Person popup functionality
        window.showPersonPopup = function(peopleId, personName) {
            // Create modal if it doesn't exist
            if (!document.getElementById('personModal')) {
                var modalHtml = `
                    <div id="personModal" class="modal">
                        <div class="modal-content">
                            <div class="modal-header">
                                <span class="modal-close" onclick="closePersonModal()">&times;</span>
                                <h2 id="modalPersonName"></h2>
                            </div>
                            <div class="modal-body" id="modalPersonContent">
                                <div class="text-center">
                                    <i class="fa fa-spinner fa-spin"></i> Loading person information...
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                $('body').append(modalHtml);
            }
            
            // Show modal with loading state
            $('#modalPersonName').text(personName);
            $('#personModal').show();
            
            // Load person data
            $.ajax({
                url: getPyScriptAddress(),
                method: "POST",
                data: {
                    action: "get_person_info",
                    people_id: peopleId
                },
                success: function(html) {
                    $('#modalPersonContent').html(html);
                },
                error: function(xhr, status, error) {
                    $('#modalPersonContent').html('<div class="alert alert-danger">Error loading person information: ' + error + '</div>');
                }
            });
        };
        
        window.closePersonModal = function() {
            $('#personModal').hide();
        };
        
        // Close modal when clicking outside
        $(document).on('click', '#personModal', function(e) {
            if (e.target === this) {
                closePersonModal();
            }
        });
    });
    </script>
    '''

def get_person_popup_info():
    """Get person information for popup display"""
    # Set content type for HTML response
    model.Header = 'Content-Type: text/html'
    
    # Get person ID
    people_id = getattr(model.Data, 'people_id', '')
    
    if not people_id:
        print '<div class="alert alert-danger">No person ID specified</div>'
        return
    
    try:
        people_id = int(people_id)
    except:
        print '<div class="alert alert-danger">Invalid person ID</div>'
        return
    
    # Get person information
    person_sql = """
    SELECT 
        p.PeopleId,
        p.Name2 AS Name,
        p.Age,
        g.Description AS Gender,
        m.Description AS MaritalStatus,
        p.EmailAddress,
        p.CellPhone,
        p.HomePhone,
        p.FamilyId,
        ms.Description AS MemberStatus,
        ISNULL(p.PrimaryAddress, '') + 
        CASE WHEN p.PrimaryAddress2 IS NOT NULL AND p.PrimaryAddress2 != '' 
             THEN ', ' + p.PrimaryAddress2 ELSE '' END AS Address,
        ISNULL(p.PrimaryCity, '') + ', ' + ISNULL(p.PrimaryState, '') + ' ' + ISNULL(p.PrimaryZip, '') AS CityStateZip
    FROM People p
    LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
    LEFT JOIN lookup.MaritalStatus m ON p.MaritalStatusId = m.Id
    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
    WHERE p.PeopleId = {0}
    """.format(people_id)
    
    person = q.QuerySqlTop1(person_sql)
    
    if not person:
        print '<div class="alert alert-danger">Person not found</div>'
        return
    
    # Render person information
    print '<div class="person-info-section">'
    print '<h4><i class="fa fa-user"></i> Basic Information</h4>'
    print '<table class="info-table">'
    
    if safe_get_value(person, 'Age', ''):
        print '<tr><td>Age:</td><td>{0}</td></tr>'.format(safe_get_value(person, 'Age', ''))
    
    if safe_get_value(person, 'Gender', ''):
        print '<tr><td>Gender:</td><td>{0}</td></tr>'.format(safe_get_value(person, 'Gender', ''))
    
    if safe_get_value(person, 'MaritalStatus', ''):
        print '<tr><td>Marital Status:</td><td>{0}</td></tr>'.format(safe_get_value(person, 'MaritalStatus', ''))
    
    if safe_get_value(person, 'MemberStatus', ''):
        print '<tr><td>Member Status:</td><td>{0}</td></tr>'.format(safe_get_value(person, 'MemberStatus', ''))
    
    print '</table>'
    print '</div>'
    
    # Contact information
    print '<div class="person-info-section">'
    print '<h4><i class="fa fa-phone"></i> Contact Information</h4>'
    print '<table class="info-table">'
    
    email = safe_get_value(person, 'EmailAddress', '')
    if email:
        print '<tr><td>Email:</td><td><a href="mailto:{0}">{0}</a></td></tr>'.format(email)
    
    cell = safe_get_value(person, 'CellPhone', '')
    if cell:
        print '<tr><td>Cell Phone:</td><td>{0}</td></tr>'.format(format_phone(cell))
    
    home = safe_get_value(person, 'HomePhone', '')
    if home:
        print '<tr><td>Home Phone:</td><td>{0}</td></tr>'.format(format_phone(home))
    
    address = safe_get_value(person, 'Address', '')
    city_state_zip = safe_get_value(person, 'CityStateZip', '')
    if address or city_state_zip:
        full_address = address
        if city_state_zip and city_state_zip != ', ':
            full_address += '<br>' + city_state_zip
        print '<tr><td>Address:</td><td>{0}</td></tr>'.format(full_address)
    
    print '</table>'
    print '</div>'
    
    # Get family members in same organizations
    family_id = safe_get_value(person, 'FamilyId', 0)
    if family_id:
        # First get the organizations this person belongs to
        family_sql = """
        WITH PersonOrgs AS (
            SELECT DISTINCT om.OrganizationId
            FROM OrganizationMembers om
            INNER JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            WHERE om.PeopleId = {0}
            AND o.OrganizationStatusId = 30
        )
        SELECT 
            p.PeopleId,
            p.Name2 AS Name,
            p.Age,
            g.Description AS Gender,
            o.OrganizationName,
            CASE 
                WHEN p.PositionInFamilyId = 10 THEN 'Primary Adult'
                WHEN p.PositionInFamilyId = 20 THEN 'Secondary Adult'
                WHEN p.PositionInFamilyId = 30 THEN 'Child'
                ELSE 'Other'
            END AS FamilyPosition
        FROM People p
        LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
        INNER JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
        INNER JOIN Organizations o ON om.OrganizationId = o.OrganizationId
        WHERE p.FamilyId = {1}
        AND p.PeopleId != {0}
        AND p.IsDeceased = 0
        AND p.ArchivedFlag = 0
        AND om.OrganizationId IN (SELECT OrganizationId FROM PersonOrgs)
        ORDER BY p.PositionInFamilyId, p.Age DESC, p.Name2
        """.format(people_id, family_id)
        
        family_members = q.QuerySql(family_sql)
        
        if family_members:
            print '<div class="person-info-section">'
            print '<h4><i class="fa fa-users"></i> Family Members in Same Organizations</h4>'
            
            # Group by person to avoid duplicates if they're in multiple same orgs
            seen_people = set()
            for member in family_members:
                member_id = safe_get_value(member, 'PeopleId', 0)
                if member_id in seen_people:
                    continue
                seen_people.add(member_id)
                
                print '<div class="family-member-item">'
                print '<strong>{0}</strong>'.format(safe_get_value(member, 'Name', ''))
                
                details = []
                if safe_get_value(member, 'Age', ''):
                    details.append('Age {0}'.format(safe_get_value(member, 'Age', '')))
                if safe_get_value(member, 'Gender', ''):
                    details.append(safe_get_value(member, 'Gender', ''))
                if safe_get_value(member, 'FamilyPosition', ''):
                    details.append(safe_get_value(member, 'FamilyPosition', ''))
                
                if details:
                    print ' ({0})'.format(', '.join(details))
                
                print '<br><small class="text-muted">In: {0}</small>'.format(
                    safe_get_value(member, 'OrganizationName', ''))
                print '</div>'
            
            print '</div>'
    
    # Get last 5 transactions for this person
    recent_trans_sql = """
    SELECT TOP 5
        t.TransactionDate,
        t.TransactionId,
        t.amt AS Amount,
        t.Message,
        t.Description,
        ISNULL(o.OrganizationName, 'General') AS OrganizationName,
        ISNULL(ac.Description, 'General Fund') AS FundName
    FROM TransactionSummary ts
    INNER JOIN [Transaction] t ON ts.RegId = t.OriginalId
    LEFT JOIN Organizations o ON t.OrgId = o.OrganizationId
    LEFT JOIN lookup.AccountCode ac ON ac.Id = CASE 
        WHEN o.RegAccountCodeId IS NOT NULL THEN CAST(o.RegAccountCodeId AS NVARCHAR(50))
        ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'NVARCHAR(50)')
    END
    WHERE ts.PeopleId = {0}
    AND ts.IsLatestTransaction = 1
    AND t.amt <> 0
    AND t.TransactionId IS NOT NULL
    AND t.voided IS NULL
    ORDER BY t.TransactionDate DESC
    """.format(people_id)
    
    recent_transactions = q.QuerySql(recent_trans_sql)
    
    if recent_transactions:
        print '<div class="person-info-section">'
        print '<h4><i class="fa fa-history"></i> Recent Transactions</h4>'
        print '<table class="table table-condensed" style="font-size: 12px;">'
        print '<thead>'
        print '<tr>'
        print '<th>Date</th>'
        print '<th>Type</th>'
        print '<th>Fund</th>'
        print '<th class="text-right">Amount</th>'
        print '</tr>'
        print '</thead>'
        print '<tbody>'
        
        for trans in recent_transactions:
            trans_date = safe_get_value(trans, 'TransactionDate', '')
            if trans_date:
                try:
                    trans_date = trans_date.ToString("MM/dd/yy")
                except:
                    trans_date = str(trans_date)[:10]
            
            message = safe_get_value(trans, 'Message', '')
            payment_type = classify_payment_type_by_message(message)
            amount = float(safe_get_value(trans, 'Amount', 0))
            
            print '<tr>'
            print '<td>{0}</td>'.format(trans_date)
            print '<td>{0}</td>'.format(payment_type)
            print '<td>{0}</td>'.format(safe_get_value(trans, 'FundName', ''))
            print '<td class="text-right">${0:,.2f}</td>'.format(amount)
            print '</tr>'
        
        print '</tbody>'
        print '</table>'
        print '</div>'
    
    # Link to profile
    print '<div class="person-info-section" style="text-align: center; margin-top: 20px;">'
    print '<a href="/Person2/{0}" target="_blank" class="btn btn-primary">'.format(people_id)
    print '<i class="fa fa-external-link"></i> View Full Profile</a>'
    print '</div>'

def format_phone(phone):
    """Format phone number for display"""
    if not phone:
        return ''
    
    # Remove non-digits
    digits = ''.join(c for c in str(phone) if c.isdigit())
    
    if len(digits) == 10:
        return '({0}) {1}-{2}'.format(digits[:3], digits[3:6], digits[6:])
    elif len(digits) == 7:
        return '{0}-{1}'.format(digits[:3], digits[3:])
    
    return phone

# Execute main function
main()
