'''
Purpose: Finance tool to report on involvements (organizations) with fees
Shows creation date, creator, fee counts and amounts from both old XML and new JSON formats

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File  
3. Name the Python script "TPxi_InvWithFees" and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Written By: Ben Swaby
Email: bswaby@fbchtn.org
'''
#roles=Finance,Admin

import json
from datetime import datetime, timedelta
import sys
import os

# ===== CONFIGURATION SECTION =====
class Config:
    # === Display Settings ===
    PAGE_TITLE = "Involvements with Fees Report"
    PAGE_SIZE = 100
    MAX_EXPORT_ROWS = 5000
    
    # === Feature Flags ===
    ENABLE_DEBUG = False  # Enable debug mode to see SQL queries
    ENABLE_EXPORT = True
    ENABLE_EMAIL_AUTOMATION = False  # Set True when ready to implement
    
    # === Security Settings ===
    REQUIRED_ROLE = "Finance"
    ADMIN_FEATURES_ROLE = "Admin"
    
    # === Date Settings ===
    DEFAULT_DAYS_BACK = 30  # Default to last 30 days
    DATE_FORMAT = "%m/%d/%Y"

    # === Filter Settings ===
    INCLUDE_ACTIVE_FEE_COLLECTORS = True  # Include involvements collecting fees in date range (not just newly created)
    
    # === Email Settings ===
    FINANCE_EMAIL_LIST = "bswaby@fbchtn.org"  # Update with actual email
    EMAIL_SEND_TIME = "08:00"  # 8 AM daily
    
    # === Display preferences ===
    CURRENCY_FORMAT = "${:,.2f}"
    
    # === Database Compatibility ===
    USE_LOOKUP_SCHEMA = True  # Set False if lookup. prefix fails

# Set page header from config
model.Header = Config.PAGE_TITLE

# Global variable for dynamic schema handling
USE_LOOKUP_SCHEMA = Config.USE_LOOKUP_SCHEMA

# ===== MAIN CONTROLLER =====
def main():
    try:
        # Check permissions first
        if not check_permissions():
            return
            
        # Get routing parameters safely
        action = getattr(model.Data, 'action', '')
        
        # Route to appropriate handler
        if action == 'export':
            handle_export()
        elif action == 'test':
            run_diagnostic_test()
        else:
            show_default_view()
            
    except Exception as e:
        print_error("Main Controller", e)

def run_diagnostic_test():
    """Run a simple diagnostic test"""
    print "<h2>Diagnostic Test</h2>"
    
    # Test 0: Check what's actually in the XML fields
    print "<h3>Test 0: Raw XML Field Values</h3>"
    try:
        sql = """SELECT TOP 10 
                 o.OrganizationId,
                 o.OrganizationName,
                 o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(MAX)') AS RawAccountingCode,
                 o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(MAX)') AS RawDonationFundId,
                 LEN(o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(MAX)')) AS AcctCodeLength,
                 LEN(o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(MAX)')) AS FundIdLength
                 FROM Organizations o 
                 WHERE o.RegSettingXML.exist('(/Settings/Fees/AccountingCode)[1]') = 1
                    OR o.RegSettingXML.exist('(/Settings/Fees/DonationFundId)[1]') = 1
                 ORDER BY o.CreatedDate DESC"""
        
        results = q.QuerySql(sql)
        print "<table class='table table-bordered'>"
        print "<tr><th>ID</th><th>Name</th><th>Raw AcctCode</th><th>Raw FundId</th><th>AcctCode Len</th><th>FundId Len</th></tr>"
        for row in results:
            print "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>".format(
                row.OrganizationId, row.OrganizationName, 
                row.RawAccountingCode or 'NULL', row.RawDonationFundId or 'NULL',
                row.AcctCodeLength or 0, row.FundIdLength or 0)
        print "</table>"
    except Exception as e:
        print "<div class='alert alert-danger'>Error checking raw XML values: {0}</div>".format(str(e))
    
    # Test 0.5: Check for problematic RegAccountCodeId and RegFundId values
    print "<h3>Test 0.5: Check Organization Registration Fields</h3>"
    try:
        sql = """SELECT TOP 10 
                 o.OrganizationId,
                 o.OrganizationName,
                 o.RegAccountCodeId,
                 o.RegFundId,
                 ISNUMERIC(o.RegAccountCodeId) AS AcctCodeIsNumeric,
                 ISNUMERIC(o.RegFundId) AS FundIdIsNumeric,
                 LEN(o.RegAccountCodeId) AS AcctCodeLen,
                 LEN(o.RegFundId) AS FundIdLen
                 FROM Organizations o 
                 WHERE o.RegistrationTypeId = 26
                   AND (o.RegAccountCodeId IS NOT NULL OR o.RegFundId IS NOT NULL)
                 ORDER BY o.CreatedDate DESC"""
        
        results = q.QuerySql(sql)
        print "<table class='table table-bordered'>"
        print "<tr><th>ID</th><th>Name</th><th>RegAccountCodeId</th><th>RegFundId</th><th>AcctNumeric?</th><th>FundNumeric?</th></tr>"
        for row in results:
            print "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>".format(
                row.OrganizationId, row.OrganizationName, 
                row.RegAccountCodeId or 'NULL', row.RegFundId or 'NULL',
                row.AcctCodeIsNumeric, row.FundIdIsNumeric)
        print "</table>"
    except Exception as e:
        print "<div class='alert alert-danger'>Error checking registration fields: {0}</div>".format(str(e))
    
    # Test 0.6: Check ALL fields that might contain '[]'
    print "<h3>Test 0.6: Comprehensive Check for '[]' Values in All Numeric Fields</h3>"
    try:
        sql = """SELECT TOP 10 
                 o.OrganizationId,
                 o.OrganizationName,
                 CASE WHEN o.CreatedBy = '[]' THEN 'HAS []' ELSE 'OK' END AS CreatedByCheck,
                 CASE WHEN o.RegAccountCodeId = '[]' THEN 'HAS []' ELSE 'OK' END AS RegAccountCodeIdCheck,
                 CASE WHEN o.RegFundId = '[]' THEN 'HAS []' ELSE 'OK' END AS RegFundIdCheck,
                 CASE WHEN CAST(o.OrganizationStatusId AS VARCHAR) = '[]' THEN 'HAS []' ELSE 'OK' END AS OrgStatusIdCheck,
                 CASE WHEN CAST(o.DivisionId AS VARCHAR) = '[]' THEN 'HAS []' ELSE 'OK' END AS DivisionIdCheck,
                 o.CreatedBy,
                 o.RegAccountCodeId,
                 o.RegFundId,
                 o.OrganizationStatusId,
                 o.DivisionId
                 FROM Organizations o 
                 WHERE o.CreatedBy = '[]' 
                    OR o.RegAccountCodeId = '[]' 
                    OR o.RegFundId = '[]'
                    OR CAST(o.OrganizationStatusId AS VARCHAR) = '[]'
                    OR CAST(o.DivisionId AS VARCHAR) = '[]'"""
        
        results = q.QuerySql(sql)
        if results:
            print "<table class='table table-bordered'>"
            print "<tr><th>ID</th><th>Name</th><th>CreatedBy</th><th>RegAcct</th><th>RegFund</th><th>OrgStatus</th><th>Division</th></tr>"
            for row in results:
                print "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>".format(
                    row.OrganizationId, row.OrganizationName, 
                    row.CreatedByCheck, row.RegAccountCodeIdCheck, row.RegFundIdCheck,
                    row.OrgStatusIdCheck, row.DivisionIdCheck)
            print "</table>"
        else:
            print "<div class='alert alert-success'>No organizations found with '[]' values in checked fields</div>"
    except Exception as e:
        print "<div class='alert alert-danger'>Error in comprehensive check: {0}</div>".format(str(e))
    
    # Test for specific organization with JSON fees
    print "<h3>Test for Organization 3019 (Student Sounds Game)</h3>"
    try:
        # First check the organization details and date
        sql = """SELECT 
                 o.OrganizationId,
                 o.OrganizationName,
                 o.RegistrationTypeId,
                 o.CreatedDate,
                 CONVERT(varchar, o.CreatedDate, 101) AS CreatedDateFormatted,
                 o.RegFeePerPerson,
                 o.RegFeeChange,
                 o.RegMaxFee,
                 o.RegInvolvementFees,
                 o.RegDepositAmount
                 FROM Organizations o
                 WHERE o.OrganizationId = 3019"""
        
        org_result = q.QuerySqlTop1(sql)
        if org_result:
            print "<div class='alert alert-info'>"
            print "<strong>Organization Details:</strong><br/>"
            print "ID: {0}<br/>".format(org_result.OrganizationId)
            print "Name: {0}<br/>".format(org_result.OrganizationName)
            print "Registration Type: {0}<br/>".format(org_result.RegistrationTypeId)
            print "Created Date: {0} (Raw: {1})<br/>".format(org_result.CreatedDateFormatted, org_result.CreatedDate)
            print "<strong>Note: Created date is {0}, which is outside the default 30-day search range!</strong><br/>".format(org_result.CreatedDateFormatted)
            print "RegFeePerPerson: {0}<br/>".format(org_result.RegFeePerPerson or 'NULL')
            print "RegFeeChange: {0}<br/>".format(org_result.RegFeeChange or 'NULL')
            print "RegMaxFee: {0}<br/>".format(org_result.RegMaxFee or 'NULL')
            print "RegInvolvementFees: {0}<br/>".format(org_result.RegInvolvementFees or 'NULL')
            print "RegDepositAmount: {0}<br/>".format(org_result.RegDepositAmount or 'NULL')
            print "</div>"
        
        # Now check RegQuestion
        sql = """SELECT 
                 rq.QuestionId,
                 rq.Options,
                 LEN(rq.Options) AS OptionsLength
                 FROM RegQuestion rq
                 WHERE rq.OrganizationId = 3019"""
        
        results = q.QuerySql(sql)
        if results:
            print "<table class='table table-bordered'>"
            print "<tr><th>QuestionId</th><th>Has Options?</th><th>Options Length</th></tr>"
            for row in results:
                print "<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>".format(
                    row.QuestionId or 'NULL',
                    'YES' if row.Options else 'NO',
                    row.OptionsLength or 0)
                if row.Options:
                    print "<tr><td colspan='3'><strong>Options Content:</strong><br/><pre>{0}</pre></td></tr>".format(row.Options[:500] + '...' if len(row.Options) > 500 else row.Options)
                    # Try to parse the fees
                    try:
                        fees = parse_json_fees(row.Options)
                        if fees:
                            print "<tr><td colspan='3'><strong>Parsed Fees:</strong><ul>"
                            for fee in fees:
                                print "<li>{0}: ${1:.2f}</li>".format(fee['Description'], fee['Amount'])
                            print "</ul></td></tr>"
                    except:
                        pass
            print "</table>"
        else:
            print "<p>No RegQuestion record found for Org 3019</p>"
            
        # Test if it would appear in the main query without date filter
        test_sql = """
        SELECT 
            o.OrganizationId,
            o.OrganizationName,
            'JSON' AS FeeSource,
            rq.Options
        FROM Organizations o
        INNER JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
        WHERE o.OrganizationId = 3019
          AND o.RegistrationTypeId = 26
          AND rq.Options LIKE '%"Fee":%' 
          AND rq.Options NOT LIKE '%"Fee":null%'
          AND CHARINDEX('"Fee":0', rq.Options) = 0
        """
        
        test_result = q.QuerySqlTop1(test_sql)
        if test_result:
            print "<div class='alert alert-success'>Organization 3019 WOULD appear in results if date range included April 2025</div>"
        else:
            print "<div class='alert alert-warning'>Organization 3019 would NOT appear even without date filter - check the fee criteria</div>"
            
    except Exception as e:
        print "<div class='alert alert-danger'>Error checking org 3019: {0}</div>".format(str(e))
    
    # Test 1: Simple query for any organization with XML fees
    print "<h3>Test 1: Organizations with XML Fees (No Date Filter)</h3>"
    try:
        sql = """SELECT TOP 10 
                 o.OrganizationId, 
                 o.OrganizationName,
                 CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1 
                      THEN o.RegSettingXML.value('(/Settings/Fees/Fee)[1]', 'money') 
                      ELSE NULL END AS Fee,
                 CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1 
                      THEN o.RegSettingXML.value('(/Settings/Fees/Deposit)[1]', 'money') 
                      ELSE NULL END AS Deposit,
                 CASE 
                     WHEN o.RegSettingXML.exist('(/Settings/Fees/AccountingCode)[1]') = 1 
                     THEN 
                         CASE 
                             WHEN o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(50)') = '[]' 
                             THEN NULL
                             ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(50)')
                         END
                     ELSE NULL 
                 END AS AccountingCode,
                 CASE 
                     WHEN o.RegSettingXML.exist('(/Settings/Fees/DonationFundId)[1]') = 1 
                     THEN 
                         CASE 
                             WHEN o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(50)') = '[]' 
                             THEN NULL
                             ELSE o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(50)')
                         END
                     ELSE NULL 
                 END AS DonationFundId,
                 CONVERT(varchar, o.CreatedDate, 101) AS CreatedDate
                 FROM Organizations o 
                 WHERE (o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1
                    OR o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1)
                 ORDER BY o.CreatedDate DESC"""
        
        results = q.QuerySql(sql)
        count = 0
        print "<table class='table table-bordered'>"
        print "<tr><th>ID</th><th>Name</th><th>Fee</th><th>Deposit</th><th>Acct Code</th><th>Fund ID</th><th>Created</th></tr>"
        for row in results:
            print "<tr><td>{0}</td><td>{1}</td><td>${2}</td><td>${3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>".format(
                row.OrganizationId, row.OrganizationName, 
                row.Fee or 0, row.Deposit or 0,
                row.AccountingCode or '', row.DonationFundId or '',
                row.CreatedDate)
            count += 1
        print "</table>"
        print "<p>Found {0} organizations with XML fees</p>".format(count)
    except Exception as e:
        print "<div class='alert alert-danger'>Error: {0}</div>".format(str(e))
    
    print "<hr>"
    print "<a href='/PyScript/{0}' class='btn btn-primary'>Back to Main View</a>".format(model.ScriptName)

# ===== PERMISSION CHECKING =====
def check_permissions():
    """Check if user has required permissions"""
    if not model.UserIsInRole(Config.REQUIRED_ROLE):
        print """
        <div class="alert alert-danger">
            <h4><i class="fa fa-lock"></i> Access Denied</h4>
            <p>You need the "{0}" role to access this tool.</p>
        </div>
        """.format(Config.REQUIRED_ROLE)
        return False
    return True

# ===== DATA ACCESS FUNCTIONS =====
def get_involvements_with_fees(start_date=None, end_date=None, include_active=None):
    """Get all involvements with fees from both old and new formats

    Args:
        start_date: Start date for filtering (default: Config.DEFAULT_DAYS_BACK)
        end_date: End date for filtering (default: today)
        include_active: Include involvements that collected fees in date range, not just newly created (default: Config.INCLUDE_ACTIVE_FEE_COLLECTORS)
    """

    # Default include_active from config
    if include_active is None:
        include_active = Config.INCLUDE_ACTIVE_FEE_COLLECTORS

    # Default date range if not provided
    if not start_date:
        start_date = (datetime.now() - timedelta(days=Config.DEFAULT_DAYS_BACK)).strftime('%Y-%m-%d')
    else:
        # If start_date is already a string, use it as-is
        if hasattr(start_date, 'strftime'):
            start_date = start_date.strftime('%Y-%m-%d')

    if not end_date:
        end_date = datetime.now().strftime('%Y-%m-%d')
    else:
        # If end_date is already a string, use it as-is
        if hasattr(end_date, 'strftime'):
            end_date = end_date.strftime('%Y-%m-%d')
    
    # Debug: Print the date range being used
    if Config.ENABLE_DEBUG:
        print """<div class="alert alert-info">
            <strong>Debug:</strong> Searching for involvements from {0} to {1}<br>
            Include Active: {2}<br>
            Start Date Type: {3}, End Date Type: {4}
        </div>""".format(start_date, end_date, include_active, type(start_date).__name__, type(end_date).__name__)
    
    try:
        # First, let's run a simple test query to see if we have any orgs with fees
        test_sql = """
        SELECT TOP 5 
            o.OrganizationId,
            o.OrganizationName,
            CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1 
                 THEN o.RegSettingXML.value('(/Settings/Fees/Fee)[1]', 'money') 
                 ELSE NULL END AS XMLFee,
            CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1 
                 THEN o.RegSettingXML.value('(/Settings/Fees/Deposit)[1]', 'money') 
                 ELSE NULL END AS XMLDeposit
        FROM Organizations o
        WHERE (o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1
           OR o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1)
        """
        
        if Config.ENABLE_DEBUG:
            print """<div class="alert alert-warning">
                <strong>Debug - Test Query for XML Fees:</strong>
            </div>"""
            
            try:
                test_results = q.QuerySql(test_sql)
                count = 0
                for row in test_results:
                    count += 1
                    print """<div class="alert alert-success">
                        Found Org: {0} (ID: {1}) with Fee: ${2}, Deposit: ${3}
                    </div>""".format(row.OrganizationName, row.OrganizationId, row.XMLFee or 0, row.XMLDeposit or 0)
                if count == 0:
                    print """<div class="alert alert-warning">
                        No organizations found with XML fees
                    </div>"""
            except Exception as e:
                print """<div class="alert alert-danger">
                    Error running XML test query: {0}
                </div>""".format(str(e))
        
        # Now test for JSON format fees
        # First check if RegQuestion table exists
        try:
            check_table = q.QuerySql("SELECT TOP 1 * FROM RegQuestion")
            has_regquestion = True
        except:
            has_regquestion = False
            if Config.ENABLE_DEBUG:
                print """<div class="alert alert-warning">
                    RegQuestion table not found - JSON fees not supported
                </div>"""
        
        if has_regquestion:
            test_json_sql = """
            SELECT TOP 5 
                o.OrganizationId,
                o.OrganizationName,
                rq.Options
            FROM Organizations o
            INNER JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
            WHERE rq.Options LIKE '%"Fee":%' AND rq.Options NOT LIKE '%"Fee":null%'
            """
        
            if Config.ENABLE_DEBUG:
                print """<div class="alert alert-warning">
                    <strong>Debug - Test Query for JSON Fees:</strong>
                </div>"""
                
                try:
                    test_json_results = q.QuerySql(test_json_sql)
                    count = 0
                    for row in test_json_results:
                        count += 1
                        print """<div class="alert alert-info">
                            Found Org: {0} (ID: {1}) with JSON Fee Data
                        </div>""".format(row.OrganizationName, row.OrganizationId)
                    if count == 0:
                        print """<div class="alert alert-warning">
                            No organizations found with JSON fees
                        </div>"""
                except Exception as e:
                    print """<div class="alert alert-danger">
                        Error running JSON test query: {0}
                    </div>""".format(str(e))
        
        # Check if we should include RegQuestion in the query
        include_json = has_regquestion if 'has_regquestion' in locals() else False
        
        # Build date filter clause - used for created date filter within CTEs
        # When include_active is True, we don't filter by creation date in the CTE
        # Instead we apply OR logic in final WHERE (created OR has transactions)
        if include_active:
            date_filter_created = ""  # No filter in CTE, we'll filter in final WHERE
        else:
            date_filter_created = "AND o.CreatedDate >= '{0}' AND o.CreatedDate <= DATEADD(day, 1, '{1}')".format(start_date, end_date)

        # Query for organizations with fees
        if include_json:
            sql = """
        WITH OrgFees AS (
            -- Old XML format fees (newly created in date range)
            SELECT 
                o.OrganizationId,
                o.OrganizationName,
                o.CreatedDate,
                CASE 
                    WHEN CAST(o.CreatedBy AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR)) = 0 THEN NULL
                    ELSE o.CreatedBy 
                END AS CreatedBy,
                'XML' AS FeeSource,
                o.RegistrationTypeId,
                CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1 
                     THEN o.RegSettingXML.value('(/Settings/Fees/Fee)[1]', 'money') 
                     ELSE NULL END AS Fee,
                CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1 
                     THEN o.RegSettingXML.value('(/Settings/Fees/Deposit)[1]', 'money') 
                     ELSE NULL END AS Deposit,
                CASE 
                    WHEN o.RegSettingXML.exist('(/Settings/Fees/AccountingCode)[1]') = 1 
                    THEN 
                        CASE 
                            WHEN o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(50)') = '[]' 
                            THEN NULL
                            ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(50)')
                        END
                    ELSE NULL 
                END AS AccountingCode,
                CASE 
                    WHEN o.RegSettingXML.exist('(/Settings/Fees/DonationFundId)[1]') = 1 
                    THEN 
                        CASE 
                            WHEN o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(50)') = '[]' 
                            THEN NULL
                            ELSE o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(50)')
                        END
                    ELSE NULL 
                END AS DonationFundId,
                NULL AS RegFeePerPerson,
                NULL AS RegFeeChange,
                NULL AS RegMaxFee,
                NULL AS RegInvolvementFees,
                NULL AS RegDepositAmount,
                NULL AS RegAccountCodeId,
                NULL AS RegFundId,
                CAST(o.RegSettingXML AS VARCHAR(MAX)) AS FeeData
            FROM Organizations o
            WHERE (o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1
               OR o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1)
              {1}
            
            UNION ALL
            
            -- Organizations with RegistrationTypeId = 26 (New registration system)
            SELECT 
                o.OrganizationId,
                o.OrganizationName,
                o.CreatedDate,
                CASE 
                    WHEN CAST(o.CreatedBy AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR)) = 0 THEN NULL
                    ELSE o.CreatedBy 
                END AS CreatedBy,
                'Registration' AS FeeSource,
                o.RegistrationTypeId,
                NULL AS Fee,
                NULL AS Deposit,
                NULL AS AccountingCode,
                NULL AS DonationFundId,
                o.RegFeePerPerson,
                o.RegFeeChange,
                o.RegMaxFee,
                o.RegInvolvementFees,
                o.RegDepositAmount,
                CASE 
                    WHEN CAST(o.RegAccountCodeId AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.RegAccountCodeId AS VARCHAR)) = 0 THEN NULL
                    ELSE o.RegAccountCodeId 
                END AS RegAccountCodeId,
                CASE 
                    WHEN CAST(o.RegFundId AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.RegFundId AS VARCHAR)) = 0 THEN NULL
                    ELSE o.RegFundId 
                END AS RegFundId,
                NULL AS FeeData
            FROM Organizations o
            WHERE ISNUMERIC(CAST(o.RegistrationTypeId AS VARCHAR)) = 1 AND o.RegistrationTypeId = 26
              AND (ISNUMERIC(CAST(o.RegFeePerPerson AS VARCHAR)) = 1 AND CAST(o.RegFeePerPerson AS VARCHAR) != '[]' AND o.RegFeePerPerson > 0
                   OR ISNUMERIC(CAST(o.RegFeeChange AS VARCHAR)) = 1 AND CAST(o.RegFeeChange AS VARCHAR) != '[]' AND o.RegFeeChange > 0
                   OR ISNUMERIC(CAST(o.RegMaxFee AS VARCHAR)) = 1 AND CAST(o.RegMaxFee AS VARCHAR) != '[]' AND o.RegMaxFee > 0
                   OR ISNUMERIC(CAST(o.RegInvolvementFees AS VARCHAR)) = 1 AND CAST(o.RegInvolvementFees AS VARCHAR) != '[]' AND o.RegInvolvementFees > 0
                   OR ISNUMERIC(CAST(o.RegDepositAmount AS VARCHAR)) = 1 AND CAST(o.RegDepositAmount AS VARCHAR) != '[]' AND o.RegDepositAmount > 0)
              {1}
            
            UNION ALL
            
            -- JSON format fees from RegQuestion (for RegistrationTypeId = 26)
            SELECT 
                o.OrganizationId,
                o.OrganizationName,
                o.CreatedDate,
                CASE 
                    WHEN CAST(o.CreatedBy AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR)) = 0 THEN NULL
                    ELSE o.CreatedBy 
                END AS CreatedBy,
                'JSON' AS FeeSource,
                o.RegistrationTypeId,
                NULL AS Fee,
                NULL AS Deposit,
                NULL AS AccountingCode,
                NULL AS DonationFundId,
                o.RegFeePerPerson,
                o.RegFeeChange,
                o.RegMaxFee,
                o.RegInvolvementFees,
                o.RegDepositAmount,
                CASE 
                    WHEN CAST(o.RegAccountCodeId AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.RegAccountCodeId AS VARCHAR)) = 0 THEN NULL
                    ELSE o.RegAccountCodeId 
                END AS RegAccountCodeId,
                CASE 
                    WHEN CAST(o.RegFundId AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.RegFundId AS VARCHAR)) = 0 THEN NULL
                    ELSE o.RegFundId 
                END AS RegFundId,
                rq.Options AS FeeData
            FROM Organizations o
            INNER JOIN RegQuestion rq ON CASE WHEN ISNUMERIC(CAST(o.OrganizationId AS VARCHAR)) = 1 THEN o.OrganizationId ELSE 0 END = rq.OrganizationId
            WHERE ISNUMERIC(CAST(o.RegistrationTypeId AS VARCHAR)) = 1 AND o.RegistrationTypeId = 26
              AND rq.Options LIKE '%"Fee":%'
              -- Check if there's at least one non-null, non-zero fee
              AND (rq.Options LIKE '%"Fee":1%'
                   OR rq.Options LIKE '%"Fee":2%'
                   OR rq.Options LIKE '%"Fee":3%'
                   OR rq.Options LIKE '%"Fee":4%'
                   OR rq.Options LIKE '%"Fee":5%'
                   OR rq.Options LIKE '%"Fee":6%'
                   OR rq.Options LIKE '%"Fee":7%'
                   OR rq.Options LIKE '%"Fee":8%'
                   OR rq.Options LIKE '%"Fee":9%')
              {1}
        )
        SELECT
            orgfees.OrganizationId,
            orgfees.OrganizationName,
            CONVERT(varchar, orgfees.CreatedDate, 101) AS CreatedDate,
            orgfees.CreatedBy,
            ISNULL(u.Name, 'Unknown') AS CreatorName,
            orgfees.FeeSource,
            orgfees.Fee,
            orgfees.Deposit,
            orgfees.AccountingCode,
            orgfees.DonationFundId,
            orgfees.RegFeePerPerson,
            orgfees.RegFeeChange,
            orgfees.RegMaxFee,
            orgfees.RegInvolvementFees,
            orgfees.RegDepositAmount,
            orgfees.RegAccountCodeId,
            orgfees.RegFundId,
            orgfees.FeeData,
            o.OrganizationStatusId,
            ISNULL(os.OrgStatus, 'Active') AS StatusDescription,
            ISNULL(p.Name, 'No Program') AS ProgramName,
            ISNULL(d.Name, 'No Division') AS DivisionName,
            o.RegistrationClosed,
            o.RegistrationTypeId,
            (SELECT COUNT(*) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId) AS MemberCount,
            (SELECT COUNT(*) FROM RegistrationData rd WHERE rd.OrganizationId = o.OrganizationId) AS RegistrationCount,
            (SELECT ISNULL(SUM(t.Amt), 0) FROM [Transaction] t WHERE t.OrgId = o.OrganizationId AND t.TransactionDate >= '{3}' AND t.TransactionDate <= DATEADD(day, 1, '{4}')) AS TotalCollected,
            (SELECT MAX(t.TransactionDate) FROM [Transaction] t WHERE t.OrgId = o.OrganizationId AND t.TransactionDate >= '{3}' AND t.TransactionDate <= DATEADD(day, 1, '{4}')) AS LastTransactionDate,
            orgfees.CreatedDate AS CreatedDateSort
        FROM OrgFees orgfees
        INNER JOIN Organizations o ON orgfees.OrganizationId = o.OrganizationId
        LEFT JOIN Users u ON orgfees.CreatedBy = u.UserId
        LEFT JOIN OrganizationStructure os ON CASE WHEN ISNUMERIC(o.OrganizationStatusId) = 1 THEN o.OrganizationStatusId ELSE NULL END = os.OrgId
        LEFT JOIN Division d ON CASE WHEN ISNUMERIC(o.DivisionId) = 1 THEN o.DivisionId ELSE NULL END = d.Id
        LEFT JOIN Program p ON CASE WHEN ISNUMERIC(d.ProgId) = 1 THEN d.ProgId ELSE NULL END = p.Id
        WHERE 1=1
            {2}
        ORDER BY orgfees.CreatedDate DESC
        """.format('lookup.' if USE_LOOKUP_SCHEMA else '', date_filter_created,
                  "AND (orgfees.CreatedDate >= '{0}' AND orgfees.CreatedDate <= DATEADD(day, 1, '{1}') OR EXISTS (SELECT 1 FROM [Transaction] t WHERE t.OrgId = o.OrganizationId AND t.TransactionDate >= '{0}' AND t.TransactionDate <= DATEADD(day, 1, '{1}')))".format(start_date, end_date) if include_active else "",
                  start_date, end_date)
        else:
            # Simpler query without RegQuestion table
            sql = """
            WITH OrgFees AS (
                -- Old XML format fees
                SELECT 
                    o.OrganizationId,
                    o.OrganizationName,
                    o.CreatedDate,
                    CASE 
                    WHEN CAST(o.CreatedBy AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR)) = 0 THEN NULL
                    ELSE o.CreatedBy 
                END AS CreatedBy,
                    'XML' AS FeeSource,
                    o.RegistrationTypeId,
                    CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1 
                         THEN o.RegSettingXML.value('(/Settings/Fees/Fee)[1]', 'money') 
                         ELSE NULL END AS Fee,
                    CASE WHEN o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1 
                         THEN o.RegSettingXML.value('(/Settings/Fees/Deposit)[1]', 'money') 
                         ELSE NULL END AS Deposit,
                    CASE 
                        WHEN o.RegSettingXML.exist('(/Settings/Fees/AccountingCode)[1]') = 1 
                        THEN 
                            CASE 
                                WHEN o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(50)') = '[]' 
                                THEN NULL
                                ELSE o.RegSettingXML.value('(/Settings/Fees/AccountingCode)[1]', 'VARCHAR(50)')
                            END
                        ELSE NULL 
                    END AS AccountingCode,
                    CASE 
                        WHEN o.RegSettingXML.exist('(/Settings/Fees/DonationFundId)[1]') = 1 
                        THEN 
                            CASE 
                                WHEN o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(50)') = '[]' 
                                THEN NULL
                                ELSE o.RegSettingXML.value('(/Settings/Fees/DonationFundId)[1]', 'VARCHAR(50)')
                            END
                        ELSE NULL 
                    END AS DonationFundId,
                    NULL AS RegFeePerPerson,
                    NULL AS RegFeeChange,
                    NULL AS RegMaxFee,
                    NULL AS RegInvolvementFees,
                    NULL AS RegDepositAmount,
                    NULL AS RegAccountCodeId,
                    NULL AS RegFundId,
                    CAST(o.RegSettingXML AS VARCHAR(MAX)) AS FeeData
                FROM Organizations o
                WHERE (o.RegSettingXML.exist('(/Settings/Fees/Fee)[1]') = 1
                   OR o.RegSettingXML.exist('(/Settings/Fees/Deposit)[1]') = 1)
                  {1}
                
                UNION ALL
                
                -- Organizations with RegistrationTypeId = 26 (New registration system)
                SELECT 
                    o.OrganizationId,
                    o.OrganizationName,
                    o.CreatedDate,
                    CASE 
                    WHEN CAST(o.CreatedBy AS VARCHAR) = '[]' THEN NULL
                    WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR)) = 0 THEN NULL
                    ELSE o.CreatedBy 
                END AS CreatedBy,
                    'Registration' AS FeeSource,
                    o.RegistrationTypeId,
                    NULL AS Fee,
                    NULL AS Deposit,
                    NULL AS AccountingCode,
                    NULL AS DonationFundId,
                    o.RegFeePerPerson,
                    o.RegFeeChange,
                    o.RegMaxFee,
                    o.RegInvolvementFees,
                    o.RegDepositAmount,
                    CASE 
                        WHEN CAST(o.RegAccountCodeId AS VARCHAR) = '[]' THEN NULL
                        WHEN ISNUMERIC(CAST(o.RegAccountCodeId AS VARCHAR)) = 0 THEN NULL
                        ELSE o.RegAccountCodeId 
                    END AS RegAccountCodeId,
                    CASE 
                        WHEN CAST(o.RegFundId AS VARCHAR) = '[]' THEN NULL
                        WHEN ISNUMERIC(CAST(o.RegFundId AS VARCHAR)) = 0 THEN NULL
                        ELSE o.RegFundId 
                    END AS RegFundId,
                    NULL AS FeeData
                FROM Organizations o
                WHERE ISNUMERIC(CAST(o.RegistrationTypeId AS VARCHAR)) = 1 AND o.RegistrationTypeId = 26
                  AND (ISNUMERIC(CAST(o.RegFeePerPerson AS VARCHAR)) = 1 AND CAST(o.RegFeePerPerson AS VARCHAR) != '[]' AND o.RegFeePerPerson > 0
                       OR ISNUMERIC(CAST(o.RegFeeChange AS VARCHAR)) = 1 AND CAST(o.RegFeeChange AS VARCHAR) != '[]' AND o.RegFeeChange > 0
                       OR ISNUMERIC(CAST(o.RegMaxFee AS VARCHAR)) = 1 AND CAST(o.RegMaxFee AS VARCHAR) != '[]' AND o.RegMaxFee > 0
                       OR ISNUMERIC(CAST(o.RegInvolvementFees AS VARCHAR)) = 1 AND CAST(o.RegInvolvementFees AS VARCHAR) != '[]' AND o.RegInvolvementFees > 0
                       OR ISNUMERIC(CAST(o.RegDepositAmount AS VARCHAR)) = 1 AND CAST(o.RegDepositAmount AS VARCHAR) != '[]' AND o.RegDepositAmount > 0)
                  {1}
            )
            SELECT
                orgfees.OrganizationId,
                orgfees.OrganizationName,
                CONVERT(varchar, orgfees.CreatedDate, 101) AS CreatedDate,
                orgfees.CreatedBy,
                ISNULL(u.Name, 'Unknown') AS CreatorName,
                orgfees.FeeSource,
                orgfees.Fee,
                orgfees.Deposit,
                orgfees.AccountingCode,
                orgfees.DonationFundId,
                orgfees.RegFeePerPerson,
                orgfees.RegFeeChange,
                orgfees.RegMaxFee,
                orgfees.RegInvolvementFees,
                orgfees.RegDepositAmount,
                orgfees.RegAccountCodeId,
                orgfees.RegFundId,
                orgfees.FeeData,
                o.OrganizationStatusId,
                ISNULL(os.OrgStatus, 'Active') AS StatusDescription,
                ISNULL(p.Name, 'No Program') AS ProgramName,
                ISNULL(d.Name, 'No Division') AS DivisionName,
                o.RegistrationClosed,
                o.RegistrationTypeId,
                ISNULL((SELECT COUNT(*) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId), 0) AS MemberCount,
                ISNULL((SELECT COUNT(*) FROM RegistrationData rd WHERE rd.OrganizationId = o.OrganizationId), 0) AS RegistrationCount,
                (SELECT ISNULL(SUM(t.Amt), 0) FROM [Transaction] t WHERE t.OrgId = o.OrganizationId AND t.TransactionDate >= '{3}' AND t.TransactionDate <= DATEADD(day, 1, '{4}')) AS TotalCollected,
                (SELECT MAX(t.TransactionDate) FROM [Transaction] t WHERE t.OrgId = o.OrganizationId AND t.TransactionDate >= '{3}' AND t.TransactionDate <= DATEADD(day, 1, '{4}')) AS LastTransactionDate,
                orgfees.CreatedDate AS CreatedDateSort
            FROM OrgFees orgfees
            INNER JOIN Organizations o ON orgfees.OrganizationId = o.OrganizationId
            LEFT JOIN Users u ON orgfees.CreatedBy = u.UserId
            LEFT JOIN OrganizationStructure os ON CASE WHEN ISNUMERIC(CAST(o.OrganizationStatusId AS VARCHAR)) = 1 THEN o.OrganizationStatusId ELSE NULL END = os.OrgId
            LEFT JOIN Division d ON CASE WHEN ISNUMERIC(CAST(o.DivisionId AS VARCHAR)) = 1 THEN o.DivisionId ELSE NULL END = d.Id
            LEFT JOIN Program p ON CASE WHEN ISNUMERIC(CAST(d.ProgId AS VARCHAR)) = 1 THEN d.ProgId ELSE NULL END = p.Id
            WHERE 1=1
                {2}
            ORDER BY orgfees.CreatedDate DESC
            """.format('lookup.' if USE_LOOKUP_SCHEMA else '', date_filter_created,
                      "AND (orgfees.CreatedDate >= '{0}' AND orgfees.CreatedDate <= DATEADD(day, 1, '{1}') OR EXISTS (SELECT 1 FROM [Transaction] t WHERE t.OrgId = o.OrganizationId AND t.TransactionDate >= '{0}' AND t.TransactionDate <= DATEADD(day, 1, '{1}')))".format(start_date, end_date) if include_active else "",
                      start_date, end_date)
        
        # Debug: Print the SQL query
        if Config.ENABLE_DEBUG:
            # Extract just the transaction subquery portion for clarity
            trans_query_sample = "TotalCollected subquery: WHERE t.TransactionDate >= '{0}' AND t.TransactionDate <= DATEADD(day, 1, '{1}')".format(start_date, end_date)
            print """<div class="alert alert-warning">
                <strong>Debug SQL Query Parameters:</strong><br>
                start_date: {1} (type: {2})<br>
                end_date: {3} (type: {4})<br>
                include_active: {5}<br>
                {6}<br><br>
                <details>
                    <summary>Click to see full SQL</summary>
                    <pre>{0}</pre>
                </details>
            </div>""".format(sql.replace('<', '&lt;').replace('>', '&gt;'),
                           start_date, type(start_date).__name__,
                           end_date, type(end_date).__name__,
                           include_active, trans_query_sample)
        
        # Try to identify the specific issue
        if Config.ENABLE_DEBUG:
            try:
                # First, let's check if the issue is with the ISNUMERIC function
                test_isnumeric = """
                SELECT TOP 5
                    CAST(OrganizationId AS VARCHAR(50)) AS OrgIdStr,
                    CASE WHEN CAST(OrganizationId AS VARCHAR(50)) = '[]' THEN 'YES' ELSE 'NO' END AS OrgIdIsBrackets,
                    RegAccountCodeId,
                    RegFundId,
                    CASE WHEN RegAccountCodeId = '[]' THEN 'JSON Array' ELSE 'Other' END AS AcctType,
                    CASE WHEN RegFundId = '[]' THEN 'JSON Array' ELSE 'Other' END AS FundType
                FROM Organizations
                WHERE RegAccountCodeId = '[]' OR RegFundId = '[]' OR CAST(OrganizationId AS VARCHAR(50)) = '[]'
                """
                test_results = q.QuerySql(test_isnumeric)
                if test_results:
                    print "<div class='alert alert-warning'>Found organizations with '[]' values in RegAccountCodeId, RegFundId, or OrganizationId fields</div>"
                    for r in test_results:
                        print "<div>OrgId: {0}, Is '[]': {1}</div>".format(r.OrgIdStr, r.OrgIdIsBrackets)
            except Exception as e:
                print "<div class='alert alert-info'>No '[]' values found or error: {0}</div>".format(str(e))
            
        # Additional test to check the problematic joins
        if Config.ENABLE_DEBUG:
            try:
                test_joins = """
            SELECT TOP 1 
                o.OrganizationId,
                o.OrganizationName,
                o.OrganizationStatusId,
                o.DivisionId,
                d.ProgId
            FROM Organizations o
            LEFT JOIN Division d ON o.DivisionId = d.Id
            WHERE o.OrganizationStatusId = '[]' 
               OR o.DivisionId = '[]'
               OR d.ProgId = '[]'
            """
                join_test_results = q.QuerySql(test_joins)
                if join_test_results:
                    print "<div class='alert alert-danger'>Found '[]' values in join fields!</div>"
            except Exception as e:
                print "<div class='alert alert-info'>Join test error or no issues: {0}</div>".format(str(e))
            
        # Test OrganizationMembers and RegistrationData tables
        if Config.ENABLE_DEBUG:
            try:
                test_tables = """
            SELECT TOP 5 
                'OrganizationMembers' AS TableName,
                CAST(OrganizationId AS VARCHAR(50)) AS OrgId,
                CASE WHEN CAST(OrganizationId AS VARCHAR(50)) = '[]' THEN 'YES' ELSE 'NO' END AS HasBrackets
            FROM OrganizationMembers
            WHERE CAST(OrganizationId AS VARCHAR(50)) = '[]'
            UNION ALL
            SELECT TOP 5 
                'RegistrationData' AS TableName,
                CAST(OrganizationId AS VARCHAR(50)) AS OrgId,
                CASE WHEN CAST(OrganizationId AS VARCHAR(50)) = '[]' THEN 'YES' ELSE 'NO' END AS HasBrackets
            FROM RegistrationData
            WHERE CAST(OrganizationId AS VARCHAR(50)) = '[]'
            """
                table_test_results = q.QuerySql(test_tables)
                if table_test_results:
                    print "<div class='alert alert-danger'>Found '[]' values in OrganizationMembers or RegistrationData!</div>"
                    for r in table_test_results:
                        print "<div>Table: {0}, OrgId: {1}</div>".format(r.TableName, r.OrgId)
            except Exception as e:
                print "<div class='alert alert-info'>Table test error or no issues: {0}</div>".format(str(e))
            
        # Test RegistrationTypeId
        if Config.ENABLE_DEBUG:
            try:
                test_regtype = """
            SELECT TOP 5
                OrganizationId,
                OrganizationName,
                CAST(RegistrationTypeId AS VARCHAR(50)) AS RegTypeStr,
                CASE WHEN CAST(RegistrationTypeId AS VARCHAR(50)) = '[]' THEN 'YES' ELSE 'NO' END AS IsBrackets
            FROM Organizations
            WHERE CAST(RegistrationTypeId AS VARCHAR(50)) = '[]'
            """
                regtype_results = q.QuerySql(test_regtype)
                if regtype_results:
                    print "<div class='alert alert-danger'>Found '[]' values in RegistrationTypeId!</div>"
                    for r in regtype_results:
                        print "<div>OrgId: {0}, RegType: {1}</div>".format(r.OrganizationId, r.RegTypeStr)
            except Exception as e:
                print "<div class='alert alert-info'>RegType test error or no issues: {0}</div>".format(str(e))
        
        # Last test: check numeric fields
        if Config.ENABLE_DEBUG:
            try:
                test_numeric = """
            SELECT TOP 5
                OrganizationId,
                CASE WHEN ISNUMERIC(CAST(RegFeePerPerson AS VARCHAR(50))) = 0 THEN 'BAD' ELSE 'OK' END AS FeePerPersonOK,
                CASE WHEN ISNUMERIC(CAST(RegFeeChange AS VARCHAR(50))) = 0 THEN 'BAD' ELSE 'OK' END AS FeeChangeOK,
                CASE WHEN ISNUMERIC(CAST(RegMaxFee AS VARCHAR(50))) = 0 THEN 'BAD' ELSE 'OK' END AS MaxFeeOK,
                CASE WHEN ISNUMERIC(CAST(RegInvolvementFees AS VARCHAR(50))) = 0 THEN 'BAD' ELSE 'OK' END AS InvFeesOK,
                CASE WHEN ISNUMERIC(CAST(RegDepositAmount AS VARCHAR(50))) = 0 THEN 'BAD' ELSE 'OK' END AS DepositOK
            FROM Organizations
            WHERE RegistrationTypeId = 26
            """
                numeric_results = q.QuerySql(test_numeric)
                if numeric_results:
                    for r in numeric_results:
                        if r.FeePerPersonOK == 'BAD' or r.FeeChangeOK == 'BAD' or r.MaxFeeOK == 'BAD' or r.InvFeesOK == 'BAD' or r.DepositOK == 'BAD':
                            print "<div class='alert alert-warning'>Found non-numeric values in fee fields for OrgId: {0}</div>".format(r.OrganizationId)
            except Exception as e:
                print "<div class='alert alert-info'>Numeric test error: {0}</div>".format(str(e))
        
        results = q.QuerySql(sql)
        
        # Debug: Check if we got any results
        if Config.ENABLE_DEBUG:
            print """<div class="alert alert-info">
                <strong>Debug:</strong> Query returned {0} raw results
            </div>""".format(len(list(results)) if results else 0)
        
        # Re-run the query since we consumed it for counting
        if Config.ENABLE_DEBUG:
            print "<!-- DEBUG: Executing main SQL query -->"
            
            # Test specifically for Org 3019
            test_3019_sql = """
            SELECT 
                o.OrganizationId,
                o.OrganizationName,
                o.RegistrationTypeId,
                o.CreatedDate,
                CONVERT(varchar, o.CreatedDate, 101) AS CreatedDateFormatted,
                rq.Options AS FeeData,
                CASE WHEN rq.Options LIKE '%"Fee":%' THEN 'Has Fee' ELSE 'No Fee' END AS FeeCheck,
                CASE WHEN rq.Options NOT LIKE '%"Fee":null%' THEN 'Not Null' ELSE 'Is Null' END AS NullCheck,
                CASE WHEN CHARINDEX('"Fee":0', rq.Options) = 0 THEN 'Not Zero' ELSE 'Is Zero' END AS ZeroCheck
            FROM Organizations o
            LEFT JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
            WHERE o.OrganizationId = 3019
            """
            
            try:
                test_result = q.QuerySqlTop1(test_3019_sql)
                if test_result:
                    print """<div class="alert alert-warning">
                        <strong>Debug - Org 3019 Status:</strong><br/>
                        Created: {0}<br/>
                        RegType: {1}<br/>
                        Fee Check: {2}<br/>
                        Null Check: {3}<br/>
                        Zero Check: {4}<br/>
                        Date Range: {5} to {6}<br/>
                        In Range: {7}
                    </div>""".format(
                        test_result.CreatedDateFormatted,
                        test_result.RegistrationTypeId,
                        test_result.FeeCheck,
                        test_result.NullCheck,
                        test_result.ZeroCheck,
                        start_date,
                        end_date,
                        'YES' if test_result.CreatedDate >= start_date and test_result.CreatedDate <= end_date else 'NO'
                    )
            except Exception as e:
                print "<div class='alert alert-danger'>Error testing Org 3019: {0}</div>".format(str(e))
                
            # Test the exact JSON matching conditions
            test_json_match_sql = """
            SELECT 
                o.OrganizationId,
                o.OrganizationName,
                CASE WHEN rq.Options LIKE '%"Fee":%' THEN 1 ELSE 0 END AS HasFeeColon,
                CASE WHEN rq.Options LIKE '%"Fee":null%' THEN 1 ELSE 0 END AS HasFeeNull,
                CASE WHEN (rq.Options LIKE '%"Fee":1%' 
                           OR rq.Options LIKE '%"Fee":2%' 
                           OR rq.Options LIKE '%"Fee":3%' 
                           OR rq.Options LIKE '%"Fee":4%' 
                           OR rq.Options LIKE '%"Fee":5%' 
                           OR rq.Options LIKE '%"Fee":6%' 
                           OR rq.Options LIKE '%"Fee":7%' 
                           OR rq.Options LIKE '%"Fee":8%' 
                           OR rq.Options LIKE '%"Fee":9%') THEN 1 ELSE 0 END AS HasPositiveFee,
                CASE WHEN o.CreatedDate >= '{0}' THEN 1 ELSE 0 END AS AfterStartDate,
                CASE WHEN o.CreatedDate <= DATEADD(day, 1, '{1}') THEN 1 ELSE 0 END AS BeforeEndDate,
                CASE WHEN o.RegistrationTypeId = 26 THEN 1 ELSE 0 END AS IsRegType26,
                LEN(rq.Options) AS OptionsLength
            FROM Organizations o
            INNER JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
            WHERE o.OrganizationId = 3019
            """.format(start_date, end_date)
            
            try:
                match_result = q.QuerySqlTop1(test_json_match_sql)
                if match_result:
                    print """<div class="alert alert-info">
                        <strong>JSON Match Test for Org 3019:</strong><br/>
                        Has "Fee": = {0}<br/>
                        Has "Fee":null = {1}<br/>
                        Has Positive Fee = {2}<br/>
                        After Start Date = {3}<br/>
                        Before End Date = {4}<br/>
                        Is RegType 26 = {5}<br/>
                        Options Length: {6} chars<br/>
                        ALL MATCH = {7}
                    </div>""".format(
                        match_result.HasFeeColon,
                        match_result.HasFeeNull,
                        match_result.HasPositiveFee,
                        match_result.AfterStartDate,
                        match_result.BeforeEndDate,
                        match_result.IsRegType26,
                        match_result.OptionsLength,
                        'YES' if all([match_result.HasFeeColon, match_result.HasPositiveFee, 
                                     match_result.AfterStartDate, match_result.BeforeEndDate, match_result.IsRegType26]) else 'NO'
                    )
            except Exception as e:
                print "<div class='alert alert-danger'>Error in JSON match test: {0}</div>".format(str(e))
        
        results = q.QuerySql(sql)
        
        # Convert to list to check count
        results_list = list(results)
        
        if Config.ENABLE_DEBUG:
            print "<div class='alert alert-info'>Query returned {0} raw results</div>".format(len(results_list))
        
        # Process results to extract fee details - aggregate by OrganizationId
        org_dict = {}  # Dictionary to track organizations by ID
        
        for row in results_list:
            org_id = row.OrganizationId
            
            # Check if we've already seen this organization
            if org_id not in org_dict:
                # First time seeing this org, create the entry
                org_dict[org_id] = {
                    'OrganizationId': row.OrganizationId,
                    'OrganizationName': row.OrganizationName,
                    'CreatedDate': row.CreatedDate,
                    'CreatedBy': row.CreatedBy,
                    'CreatorName': row.CreatorName or 'Unknown',
                    'StatusDescription': row.StatusDescription,
                    'ProgramName': row.ProgramName or 'No Program',
                    'DivisionName': row.DivisionName or 'No Division',
                    'RegistrationClosed': row.RegistrationClosed,
                    'MemberCount': row.MemberCount,
                    'RegistrationCount': row.RegistrationCount,
                    'TotalCollected': getattr(row, 'TotalCollected', 0),
                    'LastTransactionDate': getattr(row, 'LastTransactionDate', None),
                    'Fees': [],
                    'FeeSources': set()  # Track which sources we've seen
                }
            
            # Get the org data
            org_data = org_dict[org_id]
            
            # Extract fees based on source
            if row.FeeSource == 'XML' and 'XML' not in org_data['FeeSources']:
                # Handle both Fee and Deposit amounts
                try:
                    if hasattr(row, 'Fee') and row.Fee:
                        fee_amount = float(row.Fee)
                        if fee_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Registration Fee',
                                'Amount': fee_amount
                            })
                except:
                    pass
                
                try:
                    if hasattr(row, 'Deposit') and row.Deposit:
                        deposit_amount = float(row.Deposit)
                        if deposit_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Deposit',
                                'Amount': deposit_amount
                            })
                except:
                    pass
                
                # Store AccountingCode and DonationFundId if present
                if hasattr(row, 'AccountingCode') and row.AccountingCode:
                    org_data['AccountingCode'] = row.AccountingCode
                if hasattr(row, 'DonationFundId') and row.DonationFundId:
                    org_data['DonationFundId'] = row.DonationFundId
                    
                if org_data['Fees']:  # Only mark as having XML source if we found fees
                    org_data['FeeSources'].add('XML')
                    
            elif row.FeeSource == 'Registration' and 'Registration' not in org_data['FeeSources']:
                # Handle RegistrationTypeId = 26 fees
                try:
                    if hasattr(row, 'RegFeePerPerson') and row.RegFeePerPerson:
                        fee_amount = float(row.RegFeePerPerson)
                        if fee_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Fee Per Person',
                                'Amount': fee_amount
                            })
                except:
                    pass
                
                try:
                    if hasattr(row, 'RegFeeChange') and row.RegFeeChange:
                        fee_amount = float(row.RegFeeChange)
                        if fee_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Fee Change',
                                'Amount': fee_amount
                            })
                except:
                    pass
                
                try:
                    if hasattr(row, 'RegMaxFee') and row.RegMaxFee:
                        fee_amount = float(row.RegMaxFee)
                        if fee_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Maximum Fee',
                                'Amount': fee_amount
                            })
                except:
                    pass
                
                try:
                    if hasattr(row, 'RegInvolvementFees') and row.RegInvolvementFees:
                        fee_amount = float(row.RegInvolvementFees)
                        if fee_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Involvement Fees',
                                'Amount': fee_amount
                            })
                except:
                    pass
                
                try:
                    if hasattr(row, 'RegDepositAmount') and row.RegDepositAmount:
                        fee_amount = float(row.RegDepositAmount)
                        if fee_amount > 0:
                            org_data['Fees'].append({
                                'Description': 'Deposit Amount',
                                'Amount': fee_amount
                            })
                except:
                    pass
                
                # Store AccountingCode and FundId from Registration fields
                if hasattr(row, 'RegAccountCodeId') and row.RegAccountCodeId:
                    org_data['AccountingCode'] = row.RegAccountCodeId
                if hasattr(row, 'RegFundId') and row.RegFundId:
                    org_data['DonationFundId'] = row.RegFundId
                    
                if org_data['Fees']:  # Only mark as having Registration source if we found fees
                    org_data['FeeSources'].add('Registration')
            elif row.FeeSource == 'JSON' and row.FeeData and row.FeeData.strip():
                # Parse JSON fees - these might be different for each row
                # For RegistrationTypeId = 26, JSON represents selectable fee options
                fees = parse_json_fees(row.FeeData)
                if fees:  # Only add if we got valid fees
                    org_data['Fees'].extend(fees)
                    org_data['FeeSources'].add('JSON')
                    
                # For RegistrationTypeId = 26, also use Registration fields for AccountingCode and FundId
                if hasattr(row, 'RegistrationTypeId') and row.RegistrationTypeId == 26:
                    if hasattr(row, 'RegAccountCodeId') and row.RegAccountCodeId:
                        org_data['AccountingCode'] = row.RegAccountCodeId
                    if hasattr(row, 'RegFundId') and row.RegFundId:
                        org_data['DonationFundId'] = row.RegFundId
        
        # Convert dictionary to list and calculate totals
        processed_results = []
        for org_id, org_data in org_dict.items():
            # Remove the FeeSources tracking set before final output
            del org_data['FeeSources']
            
            # Calculate totals
            org_data['FeeCount'] = len(org_data['Fees'])
            org_data['TotalFees'] = sum(fee['Amount'] for fee in org_data['Fees'])
            
            # Determine primary fee source for display
            if org_data['Fees']:
                # Check what types of fees we have
                fee_descriptions = [f['Description'] for f in org_data['Fees']]
                has_xml = any(d in ['Registration Fee', 'Deposit'] for d in fee_descriptions)
                has_registration = any(d in ['Fee Per Person', 'Fee Change', 'Maximum Fee', 'Involvement Fees', 'Deposit Amount'] for d in fee_descriptions)
                has_json = any(d not in ['Registration Fee', 'Deposit', 'Fee Per Person', 'Fee Change', 'Maximum Fee', 'Involvement Fees', 'Deposit Amount'] for d in fee_descriptions)
                
                # Determine the source
                source_count = sum([has_xml, has_registration, has_json])
                if source_count > 1:
                    org_data['FeeSource'] = 'Multiple'
                elif has_registration:
                    org_data['FeeSource'] = 'Registration'
                elif has_json:
                    org_data['FeeSource'] = 'JSON'
                else:
                    org_data['FeeSource'] = 'XML'
            
            if org_data['FeeCount'] > 0:  # Only include if has fees
                processed_results.append(org_data)
        
        # Sort by CreatedDate descending (latest first)
        # Need to parse the date string to sort properly
        from datetime import datetime as dt
        def parse_date(date_str):
            try:
                # Parse MM/DD/YYYY format
                return dt.strptime(date_str, '%m/%d/%Y')
            except:
                return dt.min
        
        processed_results.sort(key=lambda x: parse_date(x['CreatedDate']), reverse=True)
        
        if Config.ENABLE_DEBUG:
            print "<div class='alert alert-info'>Processed {0} organizations with fees</div>".format(len(processed_results))
        
        return processed_results
        
    except Exception as e:
        if Config.ENABLE_DEBUG:
            import traceback
            print "<div class='alert alert-danger'>Error in get_involvements_with_fees: {0}</div>".format(str(e))
            print "<pre>{0}</pre>".format(traceback.format_exc())
        # Try simpler query without lookup schema
        global USE_LOOKUP_SCHEMA
        if USE_LOOKUP_SCHEMA:
            USE_LOOKUP_SCHEMA = False
            return get_involvements_with_fees(start_date, end_date)
        else:
            # Return empty list if all fails
            return []

def parse_json_fees(json_data):
    """Parse fee information from JSON Options field
    
    For RegistrationTypeId = 26, JSON fees represent selectable fee options
    where typically only one option is selected during registration.
    """
    fees = []
    try:
        if json_data:
            # Handle the JSON data
            options = json.loads(json_data)
            if isinstance(options, list):
                for option in options:
                    if isinstance(option, dict) and 'Fee' in option:
                        # Check if Fee is not null and has a numeric value
                        fee_value = option.get('Fee')
                        if fee_value is not None and fee_value != 'null':
                            try:
                                fee_amount = float(fee_value)
                                if fee_amount > 0:
                                    # Get the description from Text field or use default
                                    description = option.get('Text', option.get('Description', 'Fee Option'))
                                    fees.append({
                                        'Description': description,
                                        'Amount': fee_amount
                                    })
                            except (ValueError, TypeError):
                                # Skip if can't convert to float
                                pass
    except Exception as e:
        if Config.ENABLE_DEBUG:
            print "<!-- Error parsing JSON: {0} -->".format(str(e))
    
    return fees

def get_summary_statistics(involvements):
    """Calculate summary statistics"""
    total_involvements = len(involvements)
    total_fees = sum(inv['TotalFees'] for inv in involvements)
    total_collected = sum(inv.get('TotalCollected', 0) for inv in involvements)
    total_registrations = sum(inv['RegistrationCount'] for inv in involvements)
    
    # Group by creator
    by_creator = {}
    for inv in involvements:
        creator = inv['CreatorName']
        if creator not in by_creator:
            by_creator[creator] = {'count': 0, 'total': 0}
        by_creator[creator]['count'] += 1
        by_creator[creator]['total'] += inv['TotalFees']
    
    # Group by source type
    xml_count = len([inv for inv in involvements if inv['FeeSource'] == 'XML'])
    json_count = len([inv for inv in involvements if inv['FeeSource'] == 'JSON'])
    registration_count = len([inv for inv in involvements if inv['FeeSource'] == 'Registration'])
    multiple_count = len([inv for inv in involvements if inv['FeeSource'] == 'Multiple'])
    
    return {
        'total_involvements': total_involvements,
        'total_fees': total_fees,
        'total_collected': total_collected,
        'total_registrations': total_registrations,
        'by_creator': by_creator,
        'xml_count': xml_count,
        'json_count': json_count,
        'registration_count': registration_count,
        'multiple_count': multiple_count,
        'avg_fee': total_fees / total_involvements if total_involvements > 0 else 0
    }

# ===== VIEW FUNCTIONS =====
def show_default_view():
    """Show main dashboard view"""
    # Get date range from form or use defaults
    start_date = getattr(model.Data, 'start_date', '')
    end_date = getattr(model.Data, 'end_date', '')

    # Check if this is a form submission (has start_date) or initial load
    is_form_submit = bool(start_date)

    if is_form_submit:
        # Form submitted - read checkbox value
        include_active = getattr(model.Data, 'include_active', '') == 'on'
    else:
        # Initial page load - use config default
        include_active = Config.INCLUDE_ACTIVE_FEE_COLLECTORS

    # Use defaults if empty
    if not start_date:
        start_date = (datetime.now() - timedelta(days=Config.DEFAULT_DAYS_BACK)).strftime('%Y-%m-%d')
    if not end_date:
        end_date = datetime.now().strftime('%Y-%m-%d')
    
    # Get data
    involvements = get_involvements_with_fees(start_date, end_date, include_active)
    stats = get_summary_statistics(involvements)
    
    # Debug: Show if we have any results
    if Config.ENABLE_DEBUG:
        print "<div class='alert alert-info'>Found {0} involvements with fees for display</div>".format(len(involvements))
    
    # Show help message if no results
    if not involvements and Config.ENABLE_DEBUG:
        print """
        <div class="alert alert-warning">
            <h4><i class="fa fa-info-circle"></i> No Results Found - Debugging Tips</h4>
            <p>If you're not seeing any results, try the following:</p>
            <ol>
                <li>Click the <strong>"Run Diagnostic"</strong> button to check if any organizations have fees in your database</li>
                <li>Expand your date range - try looking back 90 days or more</li>
                <li>Check that organizations have fees in either:
                    <ul>
                        <li>The XML format: RegSettingXML with Fee attribute</li>
                        <li>The JSON format: RegQuestion table with Fee in Options field</li>
                    </ul>
                </li>
                <li>The debug messages above show what queries were run and any errors encountered</li>
            </ol>
        </div>
        """
    
    # Render page
    print """
    <style>
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .kpi-card {{
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }}
        .kpi-value {{
            font-size: 32px;
            font-weight: bold;
            margin: 10px 0;
            color: #337ab7;
        }}
        .kpi-label {{
            color: #666;
            font-size: 14px;
            text-transform: uppercase;
        }}
        .fee-table {{
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        .fee-details {{
            font-size: 12px;
            color: #666;
        }}
        .fee-list {{
            margin: 5px 0;
            padding-left: 20px;
        }}
        .creator-stats {{
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-top: 20px;
        }}
    </style>
    
    <div>
        <h2><i class="fa fa-dollar"></i> {0}</h2>
        
        <!-- Date Filter Form -->
        <div class="well">
            <form method="get" action="/PyScript/{7}" class="form-inline">
                <div class="form-group">
                    <label>Start Date:</label>
                    <input type="date" name="start_date" value="{1}" class="form-control">
                </div>
                <div class="form-group">
                    <label>End Date:</label>
                    <input type="date" name="end_date" value="{2}" class="form-control">
                </div>
                <div class="form-group" style="margin-left: 15px;">
                    <label>
                        <input type="checkbox" name="include_active" {8}>
                        Include involvements collecting fees in date range
                    </label>
                </div>
                <button type="submit" class="btn btn-primary">
                    <i class="fa fa-filter"></i> Update
                </button>
                <button type="button" onclick="exportData()" class="btn btn-success">
                    <i class="fa fa-download"></i> Export
                </button>
                <a href="/PyScript/{7}?action=test" class="btn btn-warning">
                    <i class="fa fa-wrench"></i> Run Diagnostic
                </a>
            </form>
            <div class="alert alert-info" style="margin-top: 10px;">
                <i class="fa fa-info-circle"></i> <strong>Note:</strong> Fee totals can include sum from multiple options from a drop down.<br>
                <strong>Include active:</strong> When checked, shows involvements created in date range OR that collected fees (via Transaction table) in date range.
            </div>
        </div>
        
        <!-- KPI Cards -->
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-label">Filtered Involvements w/Fees</div>
                <div class="kpi-value">{3}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Total Fees</div>
                <div class="kpi-value">${4}</div>
            </div>
            <div class="kpi-card" style="border: 2px solid #28a745;">
                <div class="kpi-label">Total Collected</div>
                <div class="kpi-value" style="color: #28a745;">${9}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Registrations Examined</div>
                <div class="kpi-value">{5}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Average Fee</div>
                <div class="kpi-value">${6}</div>
            </div>
        </div>
    """.format(
        Config.PAGE_TITLE,
        start_date,
        end_date,
        stats['total_involvements'],
        '{0:,.2f}'.format(float(stats['total_fees'])),
        stats['total_registrations'],
        '{0:.2f}'.format(float(stats['avg_fee'])),
        model.ScriptName,
        'checked' if include_active else '',
        '{0:,.2f}'.format(float(stats['total_collected']))
    )
    
    # Creator Statistics
    if stats['by_creator']:
        print """
        <div class="creator-stats">
            <h4>By Creator</h4>
            <div class="row">
        """
        for creator, data in sorted(stats['by_creator'].items(), 
                                   key=lambda x: x[1]['total'], reverse=True):
            print """
            <div class="col-md-3">
                <strong>{0}</strong><br>
                {1} involvement{2} - ${3}
            </div>
            """.format(
                creator,
                data['count'],
                's' if data['count'] != 1 else '',
                '{0:.2f}'.format(float(data['total']))
            )
        print "</div></div>"
    
    # Involvements Table
    print """
    <div class="fee-table">
        <h3>Involvements with Fees</h3>
        <div class="table-responsive">
            <table class="table table-striped table-hover">
                <thead>
                    <tr>
                        <th>Organization</th>
                        <th>Program/Division</th>
                        <th>Created</th>
                        <th>Creator</th>
                        <th>Fee Source</th>
                        <th>Fees</th>
                        <th>Total</th>
                        <th>Collected</th>
                        <th>Last Trans</th>
                        <th>Acct Code</th>
                        <th>Fund ID</th>
                        <th>Members</th>
                        <th>Status</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    if not involvements:
        print """
        <tr>
            <td colspan="14" class="text-center">
                <div class="alert alert-warning" style="margin: 20px;">
                    <h4>No Results Found</h4>
                    <p>No involvements with fees were found for the selected date range.</p>
                    <p>Try adjusting your date range or check the debug messages above.</p>
                </div>
            </td>
        </tr>
        """
    
    for inv in involvements:
        # Format fee details
        fee_details = '<div class="fee-list">'
        for fee in inv['Fees']:
            fee_details += '<div>{0}: ${1}</div>'.format(
                fee['Description'], '{0:.2f}'.format(float(fee['Amount'])))
        fee_details += '</div>'
        
        # Status indicator
        status_class = 'success' if not inv['RegistrationClosed'] else 'default'
        status_text = 'Open' if not inv['RegistrationClosed'] else 'Closed'

        # Format collected amount and last transaction date
        total_collected = inv.get('TotalCollected', 0)
        last_trans_date = inv.get('LastTransactionDate', '')
        if last_trans_date:
            try:
                # Format datetime to just date
                if hasattr(last_trans_date, 'strftime'):
                    last_trans_display = last_trans_date.strftime('%m/%d/%Y')
                else:
                    last_trans_display = str(last_trans_date).split(' ')[0] if last_trans_date else ''
            except:
                last_trans_display = str(last_trans_date)
        else:
            last_trans_display = '-'

        print """
        <tr>
            <td>
                <strong>{0}</strong><br>
                <span class="text-muted">#{1}</span>
            </td>
            <td>{2}<br><small>{3}</small></td>
            <td>{4}</td>
            <td>{5}</td>
            <td>
                <span class="label label-{6}">
                    {7}
                </span>
            </td>
            <td class="fee-details">
                {8} fee{9}
                {10}
            </td>
            <td><strong>${11}</strong></td>
            <td><strong style="color: #28a745;">${12}</strong></td>
            <td><small>{13}</small></td>
            <td>{14}</td>
            <td>{15}</td>
            <td>{16}</td>
            <td>
                <span class="label label-{17}">{18}</span>
            </td>
            <td>
                <a href="/Organization/{1}" target="_blank" class="btn btn-xs btn-default">
                    <i class="fa fa-eye"></i> View
                </a>
            </td>
        </tr>
        """.format(
            inv['OrganizationName'],
            inv['OrganizationId'],
            inv['ProgramName'],
            inv['DivisionName'],
            inv['CreatedDate'],
            inv['CreatorName'],
            'success' if inv['FeeSource'] == 'Multiple' else ('primary' if inv['FeeSource'] == 'Registration' else ('info' if inv['FeeSource'] == 'JSON' else 'warning')),
            inv['FeeSource'],
            inv['FeeCount'],
            's' if inv['FeeCount'] != 1 else '',
            fee_details,
            '{0:.2f}'.format(float(inv['TotalFees'])),
            '{0:.2f}'.format(float(total_collected)),
            last_trans_display,
            inv.get('AccountingCode', ''),
            inv.get('DonationFundId', ''),
            inv['MemberCount'],
            status_class,
            status_text
        )
    
    print """
                </tbody>
            </table>
        </div>
    </div>
    
    <script>
    function exportData() {{
        window.location.href = '/PyScript/{0}?action=export&start_date={1}&end_date={2}';
    }}
    </script>
    </div>
    """.format(model.ScriptName, start_date, end_date)
    
    # Show admin features if applicable
    if model.UserIsInRole(Config.ADMIN_FEATURES_ROLE):
        show_admin_panel(involvements)

def show_admin_panel(involvements):
    """Show admin panel with additional features"""
    print """
    <div class="panel panel-warning" style="margin-top: 30px;">
        <div class="panel-heading">
            <h3 class="panel-title">
                <i class="fa fa-cog"></i> Admin Features
            </h3>
        </div>
        <div class="panel-body">
            <h4>Email Automation Setup</h4>
            <p>Email automation is currently: <strong>{0}</strong></p>
            
            <div class="alert alert-info">
                <p><strong>When enabled, this will send daily emails to:</strong> {1}</p>
                <p><strong>Send time:</strong> {2} daily</p>
                <p><strong>Email will include:</strong></p>
                <ul>
                    <li>All involvements created in the last 24 hours with fees</li>
                    <li>Summary statistics</li>
                    <li>Direct links to each involvement</li>
                </ul>
            </div>
            
            <h4>Fee Source Statistics</h4>
            <p>XML Format (Old): {3} involvements</p>
            <p>Registration Type 26: {4} involvements</p>
            <p>JSON Format (Options): {5} involvements</p>
            <p>Multiple Sources: {6} involvements</p>
        </div>
    </div>
    """.format(
        'ENABLED' if Config.ENABLE_EMAIL_AUTOMATION else 'DISABLED',
        Config.FINANCE_EMAIL_LIST,
        Config.EMAIL_SEND_TIME,
        len([i for i in involvements if i['FeeSource'] == 'XML']),
        len([i for i in involvements if i['FeeSource'] == 'Registration']),
        len([i for i in involvements if i['FeeSource'] == 'JSON']),
        len([i for i in involvements if i['FeeSource'] == 'Multiple'])
    )

def handle_export():
    """Handle export to CSV"""
    # Get parameters
    start_date = getattr(model.Data, 'start_date', None)
    end_date = getattr(model.Data, 'end_date', None)
    
    # Get data
    involvements = get_involvements_with_fees(start_date, end_date)
    
    # Generate CSV
    import csv
    try:
        from StringIO import StringIO
    except ImportError:
        from io import StringIO
    
    output = StringIO()
    writer = csv.writer(output)
    
    # Headers
    writer.writerow([
        'Organization ID',
        'Organization Name',
        'Program',
        'Division',
        'Created Date',
        'Creator',
        'Fee Source',
        'Fee Count',
        'Total Fees',
        'Total Collected',
        'Last Transaction Date',
        'Accounting Code',
        'Donation Fund ID',
        'Member Count',
        'Registration Count',
        'Status',
        'Fee Details'
    ])
    
    # Data rows
    for inv in involvements:
        fee_details = '; '.join(['{0}: ${1}'.format(f['Description'], '{0:.2f}'.format(float(f['Amount'])))
                                for f in inv['Fees']])

        # Format last transaction date for CSV
        last_trans_date = inv.get('LastTransactionDate', '')
        if last_trans_date:
            try:
                if hasattr(last_trans_date, 'strftime'):
                    last_trans_csv = last_trans_date.strftime('%m/%d/%Y')
                else:
                    last_trans_csv = str(last_trans_date).split(' ')[0] if last_trans_date else ''
            except:
                last_trans_csv = str(last_trans_date)
        else:
            last_trans_csv = ''

        writer.writerow([
            inv['OrganizationId'],
            inv['OrganizationName'],
            inv['ProgramName'],
            inv['DivisionName'],
            inv['CreatedDate'],
            inv['CreatorName'],
            inv['FeeSource'],
            inv['FeeCount'],
            inv['TotalFees'],
            inv.get('TotalCollected', 0),
            last_trans_csv,
            inv.get('AccountingCode', ''),
            inv.get('DonationFundId', ''),
            inv['MemberCount'],
            inv['RegistrationCount'],
            'Open' if not inv['RegistrationClosed'] else 'Closed',
            fee_details
        ])
    
    # Output CSV
    csv_content = output.getvalue()
    output.close()
    
    print csv_content
    
    # Exit cleanly for download
    import sys
    import os
    sys.stdout.flush()
    os._exit(0)

# ===== UTILITY FUNCTIONS =====
def print_error(section, error):
    """Standardized error display"""
    import traceback
    
    # Basic error message
    print """
    <div class="alert alert-danger">
        <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Error in {0}</h4>
        <p>{1}</p>
    """.format(section, str(error))
    
    # Show detailed traceback if debug is enabled
    if Config.ENABLE_DEBUG:
        print """
        <h5>Debug Details:</h5>
        <pre style="background: #f5f5f5; padding: 10px; border: 1px solid #ddd;">
{0}
        </pre>
        """.format(traceback.format_exc())
    
    print "</div>"

def send_daily_email_report():
    """Send daily email report to finance team"""
    if not Config.ENABLE_EMAIL_AUTOMATION:
        return
    
    # Get yesterday's involvements
    yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    today = datetime.now().strftime('%Y-%m-%d')
    
    involvements = get_involvements_with_fees(yesterday, today)
    
    if not involvements:
        return  # No new involvements to report
    
    # Build email content
    stats = get_summary_statistics(involvements)
    
    subject = "Daily Report: {0} New Involvements with Fees".format(stats['total_involvements'])
    
    body = """
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2>Daily Involvements with Fees Report</h2>
        <p>The following involvements with fees were created yesterday:</p>
        
        <table style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;">
                <th style="border: 1px solid #ddd; padding: 8px;">Organization</th>
                <th style="border: 1px solid #ddd; padding: 8px;">Creator</th>
                <th style="border: 1px solid #ddd; padding: 8px;">Total Fees</th>
                <th style="border: 1px solid #ddd; padding: 8px;">Link</th>
            </tr>
    """
    
    for inv in involvements:
        body += """
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px;">{0}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">{1}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">${2}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">
                    <a href="{3}/Organization/{4}">View</a>
                </td>
            </tr>
        """.format(
            inv['OrganizationName'],
            inv['CreatorName'],
            '{0:.2f}'.format(float(inv['TotalFees'])),
            model.ServerLink(""),
            inv['OrganizationId']
        )
    
    body += """
        </table>
        
        <h3>Summary</h3>
        <ul>
            <li>Total Involvements: {0}</li>
            <li>Total Fees: ${1}</li>
        </ul>
        
        <p style="color: #666; font-size: 12px;">
            This is an automated report from TouchPoint. 
            To modify recipients or disable, contact your administrator.
        </p>
    </body>
    </html>
    """.format(stats['total_involvements'], '{0:.2f}'.format(float(stats['total_fees'])))
    
    # Send email
    model.Email(
        Config.FINANCE_EMAIL_LIST,
        model.Setting("DefaultFromEmail"),
        subject,
        body
    )

# ===== ENTRY POINT =====
# Call main function
main()
