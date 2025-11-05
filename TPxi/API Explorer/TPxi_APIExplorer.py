#####################################################################
#### TouchPoint API Explorer 
#
# written by: Ben Swaby
# email: bswaby@fbchtn.org
#
# Warning: This tool will cause issues with your database.  Know 
#          what you are doing before using it

#####################################################################

import datetime
import json

# Initialize execution output
execution_output = ""
debug_info = ""

model.Header = "🚀 Touchpoint API Explorer"

# Check if we're executing code via AJAX
try:
    # First try Data (for AJAX requests)
    if 'model' in globals() and model and hasattr(model, 'Data') and hasattr(model.Data, 'code'):
        code_to_execute = model.Data.code
        
        import sys
        from StringIO import StringIO
        old_stdout = sys.stdout
        sys.stdout = StringIO()
        
        try:
            exec_globals = {'datetime': datetime}
            if 'model' in globals():
                exec_globals['model'] = model
            if 'q' in globals():
                exec_globals['q'] = q
                
            exec(code_to_execute, exec_globals)
            output = sys.stdout.getvalue()
            if not output:
                output = ""
        except Exception as e:
            output = "Error: " + str(e)
        finally:
            sys.stdout = old_stdout
            
        # Return ONLY the output for AJAX requests
        print output
        print "\n<!-- AJAX_COMPLETE -->"
        
except:
    pass

# COMPLETE TouchPoint API Documentation from https://docs.touchpointsoftware.com
method_docs = {
    # ===== CONTRIBUTIONS =====
    'model.AddContribution': {
        'category': 'Contributions',
        'desc': 'Create new contribution and bundle detail record',
        'params': [
            ('date', 'date', 'Contribution date'),
            ('fundId', 'integer', 'Fund ID'),
            ('amount', 'string', 'Amount (can contain $ and commas)'),
            ('checkNo', 'string', 'Check number'),
            ('description', 'string', 'Description'),
            ('peopleId', 'integer', 'Contributor person ID'),
            ('contributionTypeId', 'integer', 'Contribution type ID (optional)')
        ],
        'returns': 'BundleDetail object',
        'example': 'bundleDetail = model.AddContribution(datetime.datetime.now(), 1, "$100.00", "1234", "Tithe", 828)'
    },
    'model.AddContributionDetail': {
        'category': 'Contributions',
        'desc': 'Add contribution with bank routing/account for importing',
        'params': [
            ('date', 'date', 'Contribution date'),
            ('fundId', 'integer', 'Fund ID'),
            ('amount', 'string', 'Amount (can contain $ and commas)'),
            ('checkNo', 'string', 'Check number'),
            ('routing', 'string', 'Bank routing number'),
            ('account', 'string', 'Account number'),
            ('contributionType', 'integer', 'Contribution type ID (optional)')
        ],
        'returns': 'BundleDetail object',
        'example': 'bundleDetail = model.AddContributionDetail(datetime.datetime.now(), 1, "100.00", "1234", "123456789", "987654321")'
    },
    'model.CloseBundle': {
        'category': 'Contributions',
        'desc': 'Close bundle preventing further modifications',
        'params': [('header', 'BundleHeader', 'Bundle header object')],
        'returns': 'Boolean (True if closed successfully)',
        'example': 'success = model.CloseBundle(bundleHeader)'
    },
    'model.CreateContributionTag': {
        'category': 'Contributions',
        'desc': 'Create contribution tag for dashboard reports',
        'params': [
            ('name', 'string', 'Tag name (use common prefix for organization)'),
            ('dyndata', 'DynamicData', 'Dynamic data with search parameters')
        ],
        'returns': 'JSON string of search parameters',
        'example': 'data = model.DynamicData()\ndata.MinDate = datetime.datetime(2024,1,1)\ndata.MaxDate = datetime.datetime(2024,12,31)\njson = model.CreateContributionTag("2024Donors", data)'
    },
    'model.CreateContributionTagFromSql': {
        'category': 'Contributions',
        'desc': 'Create contribution tag using SQL query',
        'params': [
            ('name', 'string', 'Tag name (use common prefix)'),
            ('dyndata', 'DynamicData', 'Dynamic data with SQL parameters'),
            ('sql', 'string', 'SQL code to identify contributions')
        ],
        'returns': 'JSON string of SQL parameters',
        'example': 'data = model.DynamicData()\ndata.MinAmount = 1000\njson = model.CreateContributionTagFromSql("LargeDonors", data, "SELECT ContributionId FROM Contribution WHERE Amount > @MinAmount")'
    },
    'model.DeleteContribution': {
        'category': 'Contributions',
        'desc': 'Delete contribution and its bundle detail and tags',
        'params': [('cid', 'integer', 'Contribution ID to delete')],
        'returns': 'None',
        'example': 'model.DeleteContribution(123)'
    },
    'model.DeleteContributionTags': {
        'category': 'Contributions',
        'desc': 'Delete contribution tags matching pattern',
        'params': [('namelike', 'string', 'Tag name or pattern with % wildcard')],
        'returns': 'None',
        'example': 'model.DeleteContributionTags("2023%")  # Delete all tags starting with 2023'
    },
    'model.FetchOrCreateFund': {
        'category': 'Contributions',
        'desc': 'Fetch existing fund or create new one',
        'params': [('description', 'string', 'Fund description')],
        'returns': 'ContributionFund object',
        'example': 'fund = model.FetchOrCreateFund("Building Fund")'
    },
    'model.FindBundleHeader': {
        'category': 'Contributions',
        'desc': 'Find existing bundle header (multiple overloads)',
        'params': [
            ('bundleType', 'string', 'Bundle type description (e.g., "Online")'),
            ('referenceId', 'string', 'Reference ID'),
            ('date', 'date', 'Bundle date (optional)'),
            ('referenceIdType', 'integer', 'Reference ID type (optional)')
        ],
        'returns': 'BundleHeader or None',
        'example': 'bundle = model.FindBundleHeader("Online", "REF123")'
    },
    'model.CreateContributionTagFromSql': {
        'category': 'Contributions',
        'desc': 'Create contribution tag from SQL query',
        'params': [
            ('name', 'string', 'Tag name'),
            ('sql', 'string', 'SQL query returning contribution IDs')
        ],
        'returns': 'Tag ID',
        'example': 'tagId = model.CreateContributionTagFromSql("LargeDonors", "SELECT ContributionId FROM Contribution WHERE Amount > 1000")'
    },
    'model.DeleteContributionTags': {
        'category': 'Contributions',
        'desc': 'Delete contribution tags',
        'params': [('tagIds', 'list', 'List of tag IDs to delete')],
        'returns': 'None',
        'example': 'model.DeleteContributionTags([123, 124, 125])'
    },
    'model.FindContribution': {
        'category': 'Contributions',
        'desc': 'Find contribution by various criteria',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('date', 'datetime', 'Contribution date'),
            ('amount', 'decimal', 'Amount')
        ],
        'returns': 'Contribution object or None',
        'example': 'contrib = model.FindContribution(828, datetime.datetime(2024, 1, 1), 100.00)'
    },
    'model.FindOrCreateBundleHeader': {
        'category': 'Contributions',
        'desc': 'Find or create bundle header',
        'params': [
            ('date', 'datetime', 'Bundle date'),
            ('type', 'string', 'Bundle type')
        ],
        'returns': 'Bundle header',
        'example': 'bundle = model.FindOrCreateBundleHeader(datetime.datetime.now(), "Online")'
    },
    'model.FinishBundle': {
        'category': 'Contributions',
        'desc': 'Finish and close a bundle',
        'params': [('bundleId', 'integer', 'Bundle ID')],
        'returns': 'None',
        'example': 'model.FinishBundle(123)'
    },
    'model.FirstFundId': {
        'category': 'Contributions',
        'desc': 'Get first active fund ID',
        'params': [],
        'returns': 'Fund ID',
        'example': 'fundId = model.FirstFundId()'
    },
    'model.GetBundleHeader': {
        'category': 'Contributions',
        'desc': 'Get bundle header by ID',
        'params': [('bundleId', 'integer', 'Bundle ID')],
        'returns': 'Bundle header object',
        'example': 'bundle = model.GetBundleHeader(123)'
    },
    'model.MoveFundIdToExistingFundId': {
        'category': 'Contributions',
        'desc': 'Move contributions from one fund to another existing fund',
        'params': [
            ('fromFundId', 'integer', 'Source fund ID'),
            ('toFundId', 'integer', 'Destination fund ID')
        ],
        'returns': 'None',
        'example': 'model.MoveFundIdToExistingFundId(10, 20)'
    },
    'model.MoveFundIdToNewFundId': {
        'category': 'Contributions',
        'desc': 'Move contributions to a new fund',
        'params': [
            ('fromFundId', 'integer', 'Source fund ID'),
            ('newFundName', 'string', 'New fund name')
        ],
        'returns': 'New fund ID',
        'example': 'newFundId = model.MoveFundIdToNewFundId(10, "New Building Fund")'
    },
    'model.QueryContributionIds': {
        'category': 'Contributions',
        'desc': 'Query contribution IDs by criteria',
        'params': [('sql', 'string', 'SQL WHERE clause')],
        'returns': 'List of contribution IDs',
        'example': 'ids = model.QueryContributionIds("Amount > 500 AND YEAR(ContributionDate) = 2024")'
    },
    'model.FindContribution': {
        'category': 'Contributions',
        'desc': 'Find existing contribution by various criteria',
        'params': [
            ('peopleId', 'integer', 'People ID (optional)'),
            ('checkNo', 'string', 'Check number (optional)'),
            ('metaInfo', 'string', 'Meta information (optional)'),
            ('fundId', 'integer', 'Fund ID (optional)'),
            ('amount', 'decimal', 'Contribution amount (optional)'),
            ('date', 'date', 'Contribution date (optional)')
        ],
        'returns': 'Contribution object or None',
        'example': 'contrib = model.FindContribution(828, "1234", None, 1, 100.00, datetime.datetime.now())'
    },
    'model.FindOrCreateBundleHeader': {
        'category': 'Contributions',
        'desc': 'Find bundle header or create new if not exists',
        'params': [
            ('date', 'date', 'Bundle date (required)'),
            ('bundleType', 'string', 'Bundle type (e.g., "Online") (optional)'),
            ('referenceId', 'string', 'Reference ID (optional)'),
            ('referenceIdType', 'integer', 'Reference ID type (optional)')
        ],
        'returns': 'BundleHeader object',
        'example': 'header = model.FindOrCreateBundleHeader(datetime.datetime.now(), "Online", "REF123")'
    },
    'model.FinishBundle': {
        'category': 'Contributions',
        'desc': 'Sum bundle totals and submit changes',
        'params': [('header', 'BundleHeader', 'Bundle header to finish')],
        'returns': 'None',
        'example': 'model.FinishBundle(bundleHeader)'
    },
    'model.FirstFundId': {
        'category': 'Contributions',
        'desc': 'Get lowest numbered active fund ID (default fund)',
        'params': [],
        'returns': 'Fund ID (integer)',
        'example': 'defaultFundId = model.FirstFundId()'
    },
    'model.GetBundleHeader': {
        'category': 'Contributions',
        'desc': 'Create new bundle header',
        'params': [
            ('date', 'date', 'Bundle date (typically Sunday)'),
            ('now', 'date', 'When bundle was posted'),
            ('bundleType', 'integer', 'Bundle type ID (optional, defaults to ChecksAndCash)')
        ],
        'returns': 'New BundleHeader object',
        'example': 'header = model.GetBundleHeader(datetime.datetime.now(), datetime.datetime.now())'
    },
    'model.MoveFundIdToExistingFundId': {
        'category': 'Contributions',
        'desc': 'Move all contributions from one fund to existing fund',
        'params': [
            ('fromId', 'integer', 'Source fund ID'),
            ('toId', 'integer', 'Target fund ID'),
            ('name', 'string', 'Optional name for target fund')
        ],
        'returns': 'None',
        'example': 'model.MoveFundIdToExistingFundId(10, 20)'
    },
    'model.MoveFundIdToNewFundId': {
        'category': 'Contributions',
        'desc': 'Move all contributions to new fund',
        'params': [
            ('fromId', 'integer', 'Source fund ID'),
            ('toId', 'integer', 'New fund ID (must not exist)'),
            ('name', 'string', 'Optional name for new fund')
        ],
        'returns': 'None',
        'example': 'model.MoveFundIdToNewFundId(10, 99, "New Building Fund")'
    },
    'model.QueryContributionIds': {
        'category': 'Contributions',
        'desc': 'Execute SQL query returning contribution IDs',
        'params': [
            ('sql', 'string', 'SQL query to execute'),
            ('declarations', 'object', 'Parameters for SQL query')
        ],
        'returns': 'List of contribution IDs',
        'example': 'ids = model.QueryContributionIds("SELECT ContributionId FROM Contribution WHERE Amount > @amt", {"amt": 100})'
    },
    'model.ResolveFund': {
        'category': 'Contributions',
        'desc': 'Resolve fund by name',
        'params': [('name', 'string', 'Fund name')],
        'returns': 'ContributionFund object or None',
        'example': 'fund = model.ResolveFund("General Fund")'
    },
    'model.ResolveFundId': {
        'category': 'Contributions',
        'desc': 'Get valid fund ID by name or return first active',
        'params': [('fundName', 'string', 'Fund name or ID')],
        'returns': 'Fund ID (integer)',
        'example': 'fundId = model.ResolveFundId("General Fund")'
    },
    
    # ===== DATES =====
    'model.AddDays': {
        'category': 'Dates',
        'desc': 'Add days to date and return as string',
        'params': [
            ('date', 'datetime', 'The existing date'),
            ('days', 'integer', 'Days to add (negative for before)')
        ],
        'returns': 'String in short date format',
        'example': 'dateStr = model.AddDays(datetime.datetime.now(), 7)'
    },
    'model.DateTime': {
        'category': 'Dates',
        'desc': 'Current date and time property (readonly)',
        'params': [],
        'returns': 'Current datetime',
        'example': 'now = model.DateTime\nprint "Current time:", now'
    },
    'model.DayOfWeek': {
        'category': 'Dates',
        'desc': 'Current day of week as number (0=Sunday, 6=Saturday)',
        'params': [],
        'returns': 'Integer (0-6)',
        'example': 'if model.DayOfWeek == 0:\n    print "It\'s Sunday!"'
    },
    'model.DateAddDays': {
        'category': 'Dates',
        'desc': 'Add days to date and return datetime',
        'params': [
            ('date', 'datetime', 'The existing date'),
            ('days', 'integer', 'Days to add (negative for before)')
        ],
        'returns': 'New datetime object',
        'example': 'futureDate = model.DateAddDays(datetime.datetime.now(), 30)'
    },
    'model.DateAddHours': {
        'category': 'Dates',
        'desc': 'Add hours to date/time',
        'params': [
            ('date', 'datetime', 'The existing date and time'),
            ('hours', 'integer', 'Hours to add (negative for before)')
        ],
        'returns': 'New datetime object',
        'example': 'later = model.DateAddHours(datetime.datetime.now(), 2)'
    },
    'model.DateDiffDays': {
        'category': 'Dates',
        'desc': 'Calculate days between two dates',
        'params': [
            ('date1', 'datetime', 'Beginning date'),
            ('date2', 'datetime', 'Ending date')
        ],
        'returns': 'Integer days (negative if date2 < date1)',
        'example': 'days = model.DateDiffDays(startDate, endDate)'
    },
    'model.FormatDate': {
        'category': 'Dates',
        'desc': 'Format date as short date string',
        'params': [('date', 'datetime', 'Date to format')],
        'returns': 'Formatted date string',
        'example': 'dateStr = model.FormatDate(datetime.datetime.now())'
    },
    'model.MostRecentAttendedSunday': {
        'category': 'Dates',
        'desc': 'Get most recent Sunday with attendance for program',
        'params': [('progId', 'integer', 'Program ID')],
        'returns': 'Date of most recent attendance',
        'example': 'lastSunday = model.MostRecentAttendedSunday(1)'
    },
    'model.ParseDate': {
        'category': 'Dates',
        'desc': 'Parse string to datetime using various formats',
        'params': [('dt', 'string', 'Date string to parse')],
        'returns': 'Datetime object or None if invalid',
        'example': 'date = model.ParseDate("5/30/2024")\nif date:\n    print "Parsed:", date'
    },
    'model.ParseDateUTC': {
        'category': 'Dates',
        'desc': 'Parse UTC date string and convert to local timezone',
        'params': [('dt', 'string', 'UTC date string to parse')],
        'returns': 'Datetime in local timezone',
        'example': 'localDate = model.ParseDateUTC("2024-01-15T10:00:00Z")'
    },
    'model.ResetToday': {
        'category': 'Dates',
        'desc': 'Reset simulated date back to actual date (debugging)',
        'params': [],
        'returns': 'None',
        'example': 'model.ResetToday()  # Revert from SetToday'
    },
    'model.ScheduledTime': {
        'category': 'Dates',
        'desc': '15-minute time frame for TaskScheduler (readonly)',
        'params': [],
        'returns': 'String in HHmm format (e.g., "0930")',
        'example': 'if model.DayOfWeek == 2 and model.ScheduledTime == "1800":\n    model.CallScript("EmailReports")'
    },
    'model.SetToday': {
        'category': 'Dates',
        'desc': 'Simulate different date for debugging',
        'params': [('dt', 'datetime', 'Date to simulate')],
        'returns': 'None',
        'example': 'model.SetToday(datetime.datetime(2024, 12, 25))  # Test as if it\'s Christmas'
    },
    'model.SundayForDate': {
        'category': 'Dates',
        'desc': 'Get Sunday of the week containing given date',
        'params': [('dt', 'any', 'Date (string or datetime)')],
        'returns': 'Sunday datetime',
        'example': 'sunday = model.SundayForDate(datetime.datetime.now())'
    },
    'model.SundayForWeek': {
        'category': 'Dates',
        'desc': 'Get Sunday for numbered week of year',
        'params': [
            ('year', 'integer', 'Year (e.g., 2024)'),
            ('week', 'integer', 'Week number (1-52)')
        ],
        'returns': 'Sunday datetime',
        'example': 'firstSunday = model.SundayForWeek(2024, 1)'
    },
    'model.WeekNumber': {
        'category': 'Dates',
        'desc': 'Get week number of year for date',
        'params': [('dt', 'datetime', 'Date to check')],
        'returns': 'Week number (1-52)',
        'example': 'week = model.WeekNumber(datetime.datetime.now())\nprint "Week", week, "of the year"'
    },
    'model.WeekOfMonth': {
        'category': 'Dates',
        'desc': 'Get week number of month (1st Sunday = week 1)',
        'params': [('dt', 'datetime', 'Date to check')],
        'returns': 'Week number in month',
        'example': 'weekOfMonth = model.WeekOfMonth(datetime.datetime.now())\nif weekOfMonth == 1:\n    print "First week of month"'
    },
    
    # ===== EMAIL =====
    'model.TestEmail': {
        'category': 'Email',
        'desc': 'Property to redirect all emails to current user (for testing)',
        'params': [],
        'returns': 'Boolean (set to True to enable)',
        'example': 'model.TestEmail = True  # All emails will go to your address\n# ... send test emails ...\nmodel.TestEmail = False'
    },
    'model.Transactional': {
        'category': 'Email',
        'desc': 'Property to prevent sending email notifications about emails',
        'params': [],
        'returns': 'Boolean (set to True to enable)',
        'example': 'model.Transactional = True  # Suppress email sent notices'
    },
    'model.ContentForDate': {
        'category': 'Email',
        'desc': 'Get HTML content for specific date (for daily devotionals)',
        'params': [
            ('file', 'string', 'Name of HTML content file'),
            ('date', 'date', 'Date to search for content')
        ],
        'returns': 'HTML content string',
        'example': 'content = model.ContentForDate("DailyDevotional", datetime.datetime.now())\n# Looks for <h1>m/d/yyyy======</h1> markers'
    },
    'model.Email': {
        'category': 'Email',
        'desc': 'Send email to query results (overload 1)',
        'params': [
            ('query', 'string', 'Search query for recipients'),
            ('queuedBy', 'integer', 'Coordinator person ID'),
            ('fromEmail', 'string', 'From email address'),
            ('fromName', 'string', 'Sender name'),
            ('subject', 'string', 'Email subject'),
            ('body', 'string', 'Email body HTML'),
            ('ccList', 'string', 'Optional comma-separated CC emails')
        ],
        'returns': 'None',
        'example': 'model.Email("age > 18", 1, "info@church.org", "Church", "News", "<p>Content</p>", "admin@church.org")'
    },
    'model.EmailContent': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Send email to query using saved content. Use model.EmailContentWithSubject or newer methods.',
        'params': [
            ('query', 'string', 'Search query for recipients'),
            ('queuedBy', 'integer', 'Coordinator person ID'),
            ('fromEmail', 'string', 'From email address'),
            ('fromName', 'string', 'Sender name'),
            ('file', 'string', 'Special content file name')
        ],
        'returns': 'None',
        'example': '# DEPRECATED - Use newer email methods\n# model.EmailContent("age > 18", 1, "info@church.org", "Church", "WelcomeEmail")'
    },
    'model.EmailContentWithSubject': {
        'category': 'Email',
        'desc': 'Send email with custom subject overriding saved content',
        'params': [
            ('query', 'string', 'Search query for recipients'),
            ('queuedBy', 'integer', 'Coordinator person ID'),
            ('fromEmail', 'string', 'From email address'),
            ('fromName', 'string', 'Sender name'),
            ('subject', 'string', 'Subject (overrides content)'),
            ('file', 'string', 'Special content file name'),
            ('ccList', 'string', 'Optional comma-separated CC emails')
        ],
        'returns': 'None',
        'example': 'model.EmailContentWithSubject("age > 18", 1, "info@church.org", "Church", "Special Announcement", "Template1")'
    },
    'model.EmailReminders': {
        'category': 'Email',
        'desc': 'Send reminders for organization (volunteers or events)',
        'params': [('orgId', 'integer', 'Organization ID')],
        'returns': 'None',
        'example': 'model.EmailReminders(30)  # Sends volunteer/event reminders'
    },
    'model.EmailReport': {
        'category': 'Email',
        'desc': 'Email report from Python script (2 overloads)',
        'params': [
            ('query', 'string', 'Search query for recipients'),
            ('queuedBy', 'integer', 'Coordinator person ID'),
            ('fromEmail', 'string', 'From email address'),
            ('fromName', 'string', 'Sender name'),
            ('subject', 'string', 'Email subject'),
            ('report', 'string', 'Python script name for body'),
            ('query2', 'string', 'Optional secondary query'),
            ('query2Title', 'string', 'Optional title for query2')
        ],
        'returns': 'None',
        'example': 'model.EmailReport("Leaders", 1, "reports@church.org", "Church", "Weekly Stats", "StatsReport")\n# With custom query:\nmodel.EmailReport("Leaders", 1, "reports@church.org", "Church", "Stats", "GenericReport", "age > 50", "Senior Members")'
    },
    'model.EmailStr': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Process email replacements for testing. Use model.RenderTemplate for newer template processing.',
        'params': [
            ('body', 'string', 'Email content to process'),
            ('peopleId', 'integer', 'Optional person ID for replacements')
        ],
        'returns': 'Processed email string',
        'example': '# DEPRECATED - Use model.RenderTemplate instead\n# processed = model.EmailStr("Hello {{FirstName}}", 828)'
    },
    
    # ===== EXTRA VALUES =====
    'model.AddExtraValueBool': {
        'category': 'ExtraValues',
        'desc': 'Add or update boolean extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('value', 'boolean', 'True/False value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueBool(828, "IsVolunteer", True)'
    },
    'model.AddExtraValueCode': {
        'category': 'ExtraValues',
        'desc': 'Add or update code extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('value', 'string', 'Code value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueCode(828, "Status", "ACTIVE")'
    },
    'model.AddExtraValueDate': {
        'category': 'ExtraValues',
        'desc': 'Add or update date extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('value', 'datetime', 'Date value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueDate(828, "BaptismDate", datetime.datetime(2020, 1, 15))'
    },
    'model.AddExtraValueInt': {
        'category': 'ExtraValues',
        'desc': 'Add or update integer extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('value', 'integer', 'Integer value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueInt(828, "YearsAtChurch", 5)'
    },
    'model.AddExtraValueText': {
        'category': 'ExtraValues',
        'desc': 'Add or update text extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('value', 'string', 'Text value'),
            ('type', 'string', 'Type: text/code/date/bit/int')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueText(828, "Notes", "Important info", "text")'
    },
    'model.ExtraValue': {
        'category': 'ExtraValues',
        'desc': 'Get extra value for person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Extra value or None',
        'example': 'value = model.ExtraValue(828, "CustomField")'
    },
    'model.DeleteExtraValue': {
        'category': 'ExtraValues',
        'desc': 'Delete an extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'None',
        'example': 'model.DeleteExtraValue(828, "OldField")'
    },
    'model.AddExtraValueAttributes': {
        'category': 'ExtraValues',
        'desc': 'Add JSON attributes extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('attributes', 'string', 'JSON string')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueAttributes(828, "Preferences", "{\\"color\\": \\"blue\\"}")'
    },
    'model.ExtraValueAttributes': {
        'category': 'ExtraValues',
        'desc': 'Get JSON attributes extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'JSON string',
        'example': 'attrs = model.ExtraValueAttributes(828, "Preferences")'
    },
    'model.ExtraValueBit': {
        'category': 'ExtraValues',
        'desc': 'Get boolean extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Boolean',
        'example': 'isActive = model.ExtraValueBit(828, "IsActive")'
    },
    'model.ExtraValueCode': {
        'category': 'ExtraValues',
        'desc': 'Get code extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Code string',
        'example': 'status = model.ExtraValueCode(828, "Status")'
    },
    'model.ExtraValueDate': {
        'category': 'ExtraValues',
        'desc': 'Get date extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Datetime',
        'example': 'baptismDate = model.ExtraValueDate(828, "BaptismDate")'
    },
    'model.ExtraValueInt': {
        'category': 'ExtraValues',
        'desc': 'Get integer extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Integer',
        'example': 'years = model.ExtraValueInt(828, "YearsAtChurch")'
    },
    'model.ExtraValueText': {
        'category': 'ExtraValues',
        'desc': 'Get text extra value',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Text string',
        'example': 'notes = model.ExtraValueText(828, "Notes")'
    },
    'model.DeleteAllExtraValueLike': {
        'category': 'ExtraValues',
        'desc': 'Delete all extra values matching pattern',
        'params': [('pattern', 'string', 'Field name pattern')],
        'returns': 'None',
        'example': 'model.DeleteAllExtraValueLike("Temp%")'
    },
    
    # ===== ORGANIZATION EXTRA VALUES =====
    'model.AddExtraValueBoolOrg': {
        'category': 'ExtraValuesOrg',
        'desc': 'Add boolean extra value to organization',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('field', 'string', 'Field name'),
            ('value', 'boolean', 'True/False value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueBoolOrg(30, "IsActive", True)'
    },
    'model.AddExtraValueCodeOrg': {
        'category': 'ExtraValuesOrg',
        'desc': 'Add code extra value to organization',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('field', 'string', 'Field name'),
            ('value', 'string', 'Code value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueCodeOrg(30, "ClassType", "ADULT")'
    },
    'model.ExtraValueOrg': {
        'category': 'ExtraValuesOrg',
        'desc': 'Get organization extra value',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('field', 'string', 'Field name')
        ],
        'returns': 'Extra value or None',
        'example': 'value = model.ExtraValueOrg(30, "CustomField")'
    },
    
    # ===== MEETINGS =====
    'model.GetMeeting': {
        'category': 'Meetings',
        'desc': 'Get meeting by ID',
        'params': [('meetingId', 'integer', 'Meeting ID')],
        'returns': 'Meeting object',
        'example': 'meeting = model.GetMeeting(123)'
    },
    'model.GetMeetingIdByDateTime': {
        'category': 'Meetings',
        'desc': 'Get or create meeting ID by org and datetime',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('meetingtime_or_date', 'timestamp|datetime', 'Unix timestamp or DateTime object'),
            ('createIfNotExists', 'boolean', 'Create meeting if not exists (default True)')
        ],
        'returns': 'Meeting ID or None if not found and createIfNotExists=False',
        'example': '# With datetime:\nmeetingId = model.GetMeetingIdByDateTime(30, datetime.datetime.now())\n# With timestamp:\nmeetingId = model.GetMeetingIdByDateTime(30, 1709251200, False)'
    },
    'model.EditPersonAttendance': {
        'category': 'Meetings',
        'desc': 'Edit attendance for a person for a specific meeting',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('peopleId', 'integer', 'Person ID'),
            ('attended', 'boolean', 'True if attended, False if not')
        ],
        'returns': 'Description of attendance action taken',
        'example': 'result = model.EditPersonAttendance(123, 828, True)'
    },
    'model.UpdateMeetingDate': {
        'category': 'Meetings',
        'desc': 'Update meeting date/time',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('dateTimeStart', 'object', 'New date/time start')
        ],
        'returns': 'Updated meeting date or None',
        'example': 'model.UpdateMeetingDate(123, datetime.datetime(2024, 3, 1, 9, 0))'
    },
    'model.AddExtraValueMeeting': {
        'category': 'Meetings',
        'desc': 'Add or edit extra value for a meeting',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('field', 'string', 'Extra value field name'),
            ('value', 'string', 'Extra value field value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueMeeting(123, "Location", "Room 101")'
    },
    'model.DeleteExtraValueMeeting': {
        'category': 'Meetings',
        'desc': 'Delete extra value from a meeting',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('field', 'string', 'Extra value field name to remove')
        ],
        'returns': 'Number of extra values deleted',
        'example': 'deleted = model.DeleteExtraValueMeeting(123, "Location")'
    },
    'model.EditCommitment': {
        'category': 'Meetings',
        'desc': 'Edit commitment status for a person in a meeting',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('peopleId', 'integer', 'Person ID'),
            ('commitment', 'string', 'Commitment status: Attending, Regrets, Find Sub, Sub Found, Substitute, Uncommitted')
        ],
        'returns': 'None',
        'example': 'model.EditCommitment(123, 828, "Attending")'
    },
    'model.ExtraValueMeeting': {
        'category': 'Meetings',
        'desc': 'Get extra value from a meeting',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('field', 'string', 'Extra value field name')
        ],
        'returns': 'Extra value string',
        'example': 'location = model.ExtraValueMeeting(123, "Location")'
    },
    'model.GetMeetingIdsByDateTime': {
        'category': 'Meetings',
        'desc': 'Get meeting IDs by datetime for all or specific org',
        'params': [
            ('meetingDate', 'object', 'Meeting date/time or timestamp'),
            ('orgId', 'integer', 'Organization ID (None for all)'),
            ('exactStart', 'boolean', 'True for exact start time match')
        ],
        'returns': 'List of meeting IDs',
        'example': 'meetingIds = model.GetMeetingIdsByDateTime(datetime.datetime.now(), 30)'
    },
    'model.GetMeetingsByDateTime': {
        'category': 'Meetings',
        'desc': 'Get meetings by datetime for org(s)',
        'params': [
            ('meetingDate', 'object', 'Meeting date/time or timestamp'),
            ('orgId_or_orgIds', 'integer|list', 'Org ID or list of org IDs'),
            ('exactStart', 'boolean', 'True for exact start time match')
        ],
        'returns': 'List of Meeting objects',
        'example': 'meetings = model.GetMeetingsByDateTime(datetime.datetime.now(), [30, 31])'
    },
    'model.GetMeetingsByDateTimeRange': {
        'category': 'Meetings',
        'desc': 'Get meetings within datetime range for org(s)',
        'params': [
            ('rangeStart', 'object', 'Range start date/time or timestamp'),
            ('rangeEnd', 'object', 'Range end date/time or timestamp'),
            ('orgId_or_orgIds', 'integer|list', 'Org ID or list of org IDs (None for all)'),
            ('exactStartEnd', 'boolean', 'True for exact start/end match')
        ],
        'returns': 'List of Meeting objects',
        'example': 'meetings = model.GetMeetingsByDateTimeRange(start_date, end_date, 30)'
    },
    'model.GetNextScheduledMeetingDate': {
        'category': 'Meetings',
        'desc': 'Get next scheduled meeting date for an org',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('referenceDate', 'object', 'Reference date to check from (None for today)')
        ],
        'returns': 'DateTime of next meeting or None',
        'example': 'next_meeting = model.GetNextScheduledMeetingDate(30)'
    },
    'model.MeetingDidNotMeet': {
        'category': 'Meetings',
        'desc': 'Get or set meeting DidNotMeet flag',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('didNotMeet', 'boolean', 'True if meeting did not occur (None to get current value)')
        ],
        'returns': 'Boolean value or None if meeting not found',
        'example': 'model.MeetingDidNotMeet(123, True)'
    },
    'model.MeetingIsAllDay': {
        'category': 'Meetings',
        'desc': 'Get or set meeting all-day flag',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('isAllDay', 'boolean', 'True for all-day event (None to get current value)')
        ],
        'returns': 'Boolean value or None if meeting not found',
        'example': 'is_all_day = model.MeetingIsAllDay(123)'
    },
    'model.MeetingShowAsBusy': {
        'category': 'Meetings',
        'desc': 'Get or set meeting show-as-busy flag',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('showAsBusy', 'boolean', 'True to show as busy on calendars (None to get current value)')
        ],
        'returns': 'Boolean value or None if meeting not found',
        'example': 'model.MeetingShowAsBusy(123, True)'
    },
    'model.UpdateMeetingDuration': {
        'category': 'Meetings',
        'desc': 'Update meeting duration in minutes',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('duration', 'integer', 'Duration in minutes')
        ],
        'returns': 'None',
        'example': 'model.UpdateMeetingDuration(123, 90)'
    },
    'model.UpdateMeetingEnd': {
        'category': 'Meetings',
        'desc': 'Update meeting end date/time',
        'params': [
            ('meetingId', 'integer', 'Meeting ID'),
            ('dateTimeEnd', 'object', 'New end date/time')
        ],
        'returns': 'Updated end date or None',
        'example': 'model.UpdateMeetingEnd(123, datetime.datetime(2024, 3, 1, 10, 30))'
    },
    
    # ===== ORGANIZATIONS =====
    'model.GetOrganization': {
        'category': 'Organizations',
        'desc': 'Get organization by ID',
        'params': [('orgId', 'integer', 'Organization ID')],
        'returns': 'Organization object',
        'example': 'org = model.GetOrganization(30)\nif org:\n    print org.OrganizationName'
    },
    'model.AddOrganization': {
        'category': 'Organizations',
        'desc': 'Create new organization with program/division or from template',
        'params': [
            ('name', 'string', 'Organization name'),
            ('program_or_templateid', 'string|integer', 'Program name or template org ID'),
            ('division_or_copysettings', 'string|boolean', 'Division name or copy settings flag')
        ],
        'returns': 'New organization ID',
        'example': '# Create with program/division:\norgId = model.AddOrganization("Bible Study", "Adult Ed", "Main")\n# From template:\norgId = model.AddOrganization("New Group", 30, True)'
    },
    'model.AddMembersToOrg': {
        'category': 'Organizations',
        'desc': 'Add multiple members to organization from query',
        'params': [
            ('query', 'string', 'Query to select people'),
            ('orgid', 'integer', 'Organization ID')
        ],
        'returns': 'None',
        'example': 'model.AddMembersToOrg("Age >= 18", 30)'
    },
    'model.AddMemberToOrg': {
        'category': 'Organizations',
        'desc': 'Add single member to organization',
        'params': [
            ('pid', 'integer', 'Person ID'),
            ('orgid', 'integer', 'Organization ID')
        ],
        'returns': 'None',
        'example': 'model.AddMemberToOrg(828, 30)'
    },
    'model.AddOrganizationMember': {
        'category': 'Organizations',
        'desc': 'Add member with full options',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('peopleId', 'integer', 'Person ID'),
            ('memberTypeId', 'integer', 'Member type: 220=Member, 230=Inactive'),
            ('enrollmentDate', 'datetime', 'Date enrolled'),
            ('inactiveDate', 'datetime', 'Inactive date (optional)'),
            ('pending', 'boolean', 'Is pending')
        ],
        'returns': 'None',
        'example': 'model.AddOrganizationMember(30, 828, 220, datetime.datetime.now(), None, False)'
    },
    'model.DropOrgMember': {
        'category': 'Organizations',
        'desc': 'Remove person from organization',
        'params': [
            ('orgId', 'integer', 'Organization ID'),
            ('peopleId', 'integer', 'Person ID')
        ],
        'returns': 'None',
        'example': 'model.DropOrgMember(30, 828)'
    },
    'model.InOrg': {
        'category': 'Organizations',
        'desc': 'Check if person is in organization',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID')
        ],
        'returns': 'Boolean',
        'example': 'if model.InOrg(828, 30):\n    print "Person is member"'
    },
    'model.InSubGroup': {
        'category': 'Organizations',
        'desc': 'Check if person is in sub-group',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID'),
            ('groupName', 'string', 'Sub-group name')
        ],
        'returns': 'Boolean',
        'example': 'if model.InSubGroup(828, 30, "Team A"):\n    print "In sub-group"'
    },
    'model.AddSubGroup': {
        'category': 'Organizations',
        'desc': 'Add person to sub-group',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID'),
            ('groupName', 'string', 'Sub-group name')
        ],
        'returns': 'None',
        'example': 'model.AddSubGroup(828, 30, "Team A")'
    },
    'model.RemoveSubGroup': {
        'category': 'Organizations',
        'desc': 'Remove person from sub-group',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID'),
            ('groupName', 'string', 'Sub-group name')
        ],
        'returns': 'None',
        'example': 'model.RemoveSubGroup(828, 30, "Team A")'
    },
    'model.MoveToOrg': {
        'category': 'Organizations',
        'desc': 'Move person to different organization',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('fromOrgId', 'integer', 'From organization'),
            ('toOrgId', 'integer', 'To organization')
        ],
        'returns': 'None',
        'example': 'model.MoveToOrg(828, 30, 31)'
    },
    'model.UpdateMainFellowship': {
        'category': 'Organizations',
        'desc': 'Update MainFellowship flag and MemberCounts',
        'params': [('orgid', 'integer', 'Organization ID')],
        'returns': 'None',
        'example': 'model.UpdateMainFellowship(30)'
    },
    'model.JoinOrg': {
        'category': 'Organizations',
        'desc': 'Add person to organization as regular member',
        'params': [
            ('orgid', 'integer', 'Organization ID'),
            ('person', 'integer|Person', 'Person object or PeopleId')
        ],
        'returns': 'None',
        'example': 'model.JoinOrg(30, 828)\n# or with Person object:\nmodel.JoinOrg(30, person)'
    },
    'model.AddSubGroupFromQuery': {
        'category': 'Organizations',
        'desc': 'Add multiple org members to sub-group using query',
        'params': [
            ('query', 'string', 'Query to select people'),
            ('orgid', 'integer', 'Organization ID'),
            ('group', 'string', 'Sub-group name')
        ],
        'returns': 'None',
        'example': 'model.AddSubGroupFromQuery("IsMemberOf = 30", 30, "Team Leaders")'
    },
    'model.AddTransaction': {
        'category': 'Organizations',
        'desc': 'Add transaction for org member',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID'),
            ('payment', 'decimal', 'Payment amount'),
            ('description', 'string', 'Transaction description')
        ],
        'returns': 'Transaction ID or None',
        'example': 'transId = model.AddTransaction(828, 30, 50.00, "Registration payment")'
    },
    'model.AdjustFee': {
        'category': 'Organizations',
        'desc': 'Adjust fee for org member',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID'),
            ('adjustmentAmount', 'decimal', 'Adjustment amount (negative increases fee)'),
            ('description', 'string', 'Adjustment description')
        ],
        'returns': 'Transaction ID or None',
        'example': 'transId = model.AdjustFee(828, 30, -25.00, "Scholarship applied")'
    },
    'model.DeleteOrg': {
        'category': 'Organizations',
        'desc': 'Delete organization (requires developer role)',
        'params': [
            ('name', 'string', 'Organization name'),
            ('program', 'string', 'Program name'),
            ('division', 'string', 'Division name')
        ],
        'returns': 'None',
        'example': 'model.DeleteOrg("Old Study Group", "Adult Education", "Main Campus")'
    },
    'model.GetPayLink': {
        'category': 'Organizations',
        'desc': 'Get payment link for org member',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('orgId', 'integer', 'Organization ID')
        ],
        'returns': 'Payment URL or None',
        'example': 'paylink = model.GetPayLink(828, 30)'
    },
    'model.OnlineRegPerson': {
        'category': 'Organizations',
        'desc': 'Create online registration model from member or XML',
        'params': [
            ('om_or_xml', 'OrganizationMember|string', 'Org member object or XML string')
        ],
        'returns': 'OnlineRegPersonModel0 object',
        'example': '# From org member:\nregModel = model.OnlineRegPerson(orgMember)\n# From XML:\nregModel = model.OnlineRegPerson(xmlString)'
    },
    'model.OrganizationIds': {
        'category': 'Organizations',
        'desc': 'Get organization IDs by program and division',
        'params': [
            ('progid', 'integer', 'Program ID (0 for any)'),
            ('divid', 'integer', 'Division ID (0 for any)'),
            ('includeinactive', 'boolean', 'Include inactive orgs (default False)')
        ],
        'returns': 'List of organization IDs',
        'example': 'orgIds = model.OrganizationIds(5, 10)\n# Include inactive:\norgIds = model.OrganizationIds(5, 10, True)'
    },
    'model.SendAttendanceReminders': {
        'category': 'Organizations',
        'desc': 'Send attendance reminders for scheduled meetings',
        'params': [
            ('dt', 'object', 'Date/time for reminders')
        ],
        'returns': 'None',
        'example': 'model.SendAttendanceReminders(datetime.datetime.now())'
    },
    'model.SetMemberType': {
        'category': 'Organizations',
        'desc': 'Set member type for person in org',
        'params': [
            ('peopleid', 'integer', 'Person ID'),
            ('orgid', 'integer', 'Organization ID'),
            ('type', 'string', 'Member type name')
        ],
        'returns': 'None',
        'example': 'model.SetMemberType(828, 30, "Leader")'
    },
    'model.CurrentOrgId': {
        'category': 'Organizations',
        'desc': 'Get or set current organization ID for email replacements',
        'params': [],
        'returns': 'Organization ID',
        'example': 'model.CurrentOrgId = 30\norgId = model.CurrentOrgId'
    },
    
    # ===== PERSON =====
    'model.GetPerson': {
        'category': 'Person',
        'desc': 'Get person by ID',
        'params': [('peopleId', 'integer', 'Person ID')],
        'returns': 'Person object',
        'example': 'person = model.GetPerson(828)\nif person:\n    print person.Name'
    },
    'model.AddPerson': {
        'category': 'Person',
        'desc': 'Add new person to database',
        'params': [
            ('familyId', 'integer', 'Family ID (0 creates new)'),
            ('position', 'integer', '10=Head, 20=Spouse, 30=Child'),
            ('title', 'string', 'Title'),
            ('firstName', 'string', 'First name'),
            ('nickName', 'string', 'Nickname'),
            ('lastName', 'string', 'Last name'),
            ('dob', 'string', 'DOB MM/DD/YYYY'),
            ('married', 'boolean', 'Marital status'),
            ('gender', 'integer', '1=Male, 2=Female'),
            ('primaryAddress', 'string', 'Address'),
            ('primaryAddress2', 'string', 'Address 2'),
            ('primaryCity', 'string', 'City'),
            ('primaryState', 'string', 'State'),
            ('primaryZip', 'string', 'Zip'),
            ('homePhone', 'string', 'Home phone'),
            ('email', 'string', 'Email'),
            ('cellPhone', 'string', 'Cell phone')
        ],
        'returns': 'New person ID',
        'example': 'newId = model.AddPerson(0, 10, "", "John", "", "Smith", "01/01/1980", False, 1, "", "", "", "", "", "", "john@email.com", "")'
    },
    'model.UpdatePerson': {
        'category': 'Person',
        'desc': 'Update person record',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('field', 'string', 'Field name'),
            ('value', 'any', 'New value')
        ],
        'returns': 'None',
        'example': 'model.UpdatePerson(828, "EmailAddress", "new@email.com")'
    },
    'model.AddRole': {
        'category': 'Person',
        'desc': 'Add role to person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('roleName', 'string', 'Role name')
        ],
        'returns': 'None',
        'example': 'model.AddRole(828, "Leader")'
    },
    'model.RemoveRole': {
        'category': 'Person',
        'desc': 'Remove role from person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('roleName', 'string', 'Role name')
        ],
        'returns': 'None',
        'example': 'model.RemoveRole(828, "Leader")'
    },
    'model.AddTag': {
        'category': 'Person',
        'desc': 'Add tag to person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('tagName', 'string', 'Tag name')
        ],
        'returns': 'None',
        'example': 'model.AddTag(828, "Volunteer")'
    },
    'model.GetSpouse': {
        'category': 'Person',
        'desc': 'Get spouse of person',
        'params': [('peopleId', 'integer', 'Person ID')],
        'returns': 'Spouse person object or None',
        'example': 'spouse = model.GetSpouse(828)\nif spouse:\n    print spouse.Name'
    },
    'model.UpdateCampus': {
        'category': 'Person',
        'desc': 'Update campus for person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('campusId', 'integer', 'Campus ID')
        ],
        'returns': 'None',
        'example': 'model.UpdateCampus(828, 2)'
    },
    'model.UpdateMemberStatus': {
        'category': 'Person',
        'desc': 'Update member status',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('statusId', 'integer', 'Status ID')
        ],
        'returns': 'None',
        'example': 'model.UpdateMemberStatus(828, 10)'
    },
    'model.FindPersonId': {
        'category': 'Person',
        'desc': 'Find person ID by email',
        'params': [('email', 'string', 'Email address')],
        'returns': 'Person ID or None',
        'example': 'peopleId = model.FindPersonId("john@email.com")'
    },
    'model.FindAddPerson': {
        'category': 'Person',
        'desc': 'Find or add person',
        'params': [
            ('firstName', 'string', 'First name'),
            ('lastName', 'string', 'Last name'),
            ('email', 'string', 'Email')
        ],
        'returns': 'Person ID',
        'example': 'peopleId = model.FindAddPerson("John", "Smith", "john@email.com")'
    },
    'model.DeletePeople': {
        'category': 'Person',
        'desc': 'Permanently delete people matching query (requires developer role)',
        'params': [('query', 'string', 'Search query')],
        'returns': 'None',
        'example': 'model.DeletePeople("DeceasedDate IS NOT NULL")'
    },
    'model.AddBackgroundCheck': {
        'category': 'Person',
        'desc': 'Create and submit background check for person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('code', 'string', 'Background check package code'),
            ('submitType', 'string', 'Submission type (default: "PII")'),
            ('type', 'integer', 'Background check type (default: 1)'),
            ('label', 'integer', 'Background check label (default: 0)'),
            ('sSSN', 'string', 'Optional social security number'),
            ('sDLN', 'string', 'Optional driver license number'),
            ('sPlusCounty', 'string', 'Optional county for county-level checks'),
            ('sPlusState', 'string', 'Optional state for state-level checks'),
            ('sPackageName', 'string', 'Optional package name')
        ],
        'returns': 'Background check ID',
        'example': 'checkId = model.AddBackgroundCheck(828, "VOL", "PII", 1, 0, None, None, None, None, None)'
    },
    'model.AddRole': {
        'category': 'Person',
        'desc': 'Add role to people matching query',
        'params': [
            ('query', 'string', 'Search query'),
            ('role', 'string', 'Role name to add')
        ],
        'returns': 'None',
        'example': 'model.AddRole("age > 18", "Adult")'
    },
    'model.RemoveRole': {
        'category': 'Person',
        'desc': 'Remove role from people matching query',
        'params': [
            ('query', 'string', 'Search query'),
            ('role', 'string', 'Role name to remove')
        ],
        'returns': 'None',
        'example': 'model.RemoveRole("age < 18", "Adult")'
    },
    'model.AddTag': {
        'category': 'Person',
        'desc': 'Add tag to people matching query',
        'params': [
            ('query', 'string', 'Search query'),
            ('tagName', 'string', 'Tag name'),
            ('ownerId', 'integer', 'Tag owner person ID'),
            ('clear', 'boolean', 'Clear existing tag first (default: False)')
        ],
        'returns': 'None',
        'example': 'model.AddTag("age > 18", "Adults", 1, False)'
    },
    'model.ClearTag': {
        'category': 'Person',
        'desc': 'Remove all people from a tag',
        'params': [
            ('tagName', 'string', 'Tag name'),
            ('ownerId', 'integer', 'Tag owner person ID')
        ],
        'returns': 'None',
        'example': 'model.ClearTag("TempTag", 1)'
    },
    'model.AgeInMonths': {
        'category': 'Person',
        'desc': 'Calculate age in months between two dates',
        'params': [
            ('birthdate', 'date', 'Birth date'),
            ('asof', 'date', 'As of date')
        ],
        'returns': 'Age in months',
        'example': 'months = model.AgeInMonths(datetime.datetime(2000, 1, 1), datetime.datetime.now())'
    },
    'model.ArchiveRecords': {
        'category': 'Person',
        'desc': 'Set ArchivedFlag to true for people matching query',
        'params': [('query', 'string', 'Search query')],
        'returns': 'None',
        'example': 'model.ArchiveRecords("LastAttended < 365")'
    },
    'model.UnArchiveRecords': {
        'category': 'Person',
        'desc': 'Set ArchivedFlag to false for people matching query',
        'params': [('query', 'string', 'Search query')],
        'returns': 'None',
        'example': 'model.UnArchiveRecords("RecentInterest = 1")'
    },
    'model.FindAddPeopleId': {
        'category': 'Person',
        'desc': 'Find person or create if not exists, return ID',
        'params': [
            ('first', 'string', 'First name'),
            ('last', 'string', 'Last name'),
            ('dob', 'string', 'Date of birth'),
            ('email', 'string', 'Email address'),
            ('phone', 'string', 'Phone number'),
            ('firstLastMatch', 'boolean', 'Require exact first/last name match (default: False)')
        ],
        'returns': 'Person ID',
        'example': 'pid = model.FindAddPeopleId("John", "Smith", "01/01/1980", "john@email.com", "555-1234", False)'
    },
    'model.FindPersonId': {
        'category': 'Person',
        'desc': 'Find person ID by matching information',
        'params': [
            ('first', 'string', 'First name'),
            ('last', 'string', 'Last name'),
            ('dob', 'string', 'Date of birth'),
            ('email', 'string', 'Email address'),
            ('phone', 'string', 'Phone number')
        ],
        'returns': 'Person ID or None',
        'example': 'pid = model.FindPersonId("John", "Smith", "01/01/1980", "john@email.com", "555-1234")'
    },
    'model.FindPersonIdExtraValue': {
        'category': 'Person',
        'desc': 'Find person ID by matching extra value',
        'params': [
            ('extraKey', 'string', 'Extra value field name'),
            ('extraValue', 'string', 'Extra value to search for')
        ],
        'returns': 'Person ID or None',
        'example': 'pid = model.FindPersonIdExtraValue("EmployeeId", "12345")'
    },
    'model.FindPersonIdExtraValueInt': {
        'category': 'Person',
        'desc': 'Find person ID by matching integer extra value',
        'params': [
            ('extraKey', 'string', 'Extra value field name'),
            ('extraValue', 'integer', 'Integer extra value to search for')
        ],
        'returns': 'Person ID or None',
        'example': 'pid = model.FindPersonIdExtraValueInt("BadgeNumber", 100)'
    },
    'model.GetAuthenticatedUrl': {
        'category': 'Person',
        'desc': 'Create authenticated URL that auto-logs in specified person',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('url', 'string', 'URL to authenticate'),
            ('shorten', 'boolean', 'Whether to shorten URL (default: True)')
        ],
        'returns': 'Authenticated URL string',
        'example': 'authUrl = model.GetAuthenticatedUrl(828, "/Person2/Current", True)'
    },
    'model.GetSpouse': {
        'category': 'Person',
        'desc': 'Get spouse of person',
        'params': [('peopleId', 'integer', 'Person ID')],
        'returns': 'Spouse person object or None',
        'example': 'spouse = model.GetSpouse(828)\nif spouse:\n    print spouse.Name'
    },
    'model.UpdateAllSpouseId': {
        'category': 'Person',
        'desc': 'Update SpouseId field for all people (requires developer role)',
        'params': [],
        'returns': 'None',
        'example': 'model.UpdateAllSpouseId()'
    },
    'model.UpdateCampus': {
        'category': 'Person',
        'desc': 'Update campus for people matching query',
        'params': [
            ('query', 'string', 'Search query'),
            ('campus', 'any', 'Campus ID or campus name')
        ],
        'returns': 'None',
        'example': 'model.UpdateCampus("age > 18", "Main Campus")'
    },
    'model.UpdateContributionOption': {
        'category': 'Person',
        'desc': 'Update contribution statement option (Joint/Individual/None)',
        'params': [
            ('query', 'string', 'Search query'),
            ('option', 'any', 'Option ID or name')
        ],
        'returns': 'None',
        'example': 'model.UpdateContributionOption("MaritalStatus = Married", "Joint")'
    },
    'model.UpdateElectronicStatement': {
        'category': 'Person',
        'desc': 'Update electronic statement preference',
        'params': [
            ('query', 'string', 'Search query'),
            ('truefalse', 'boolean', 'True for electronic, False for printed')
        ],
        'returns': 'None',
        'example': 'model.UpdateElectronicStatement("HasEmail = 1", True)'
    },
    'model.UpdateEnvelopeOption': {
        'category': 'Person',
        'desc': 'Update envelope option (Joint/Individual/None)',
        'params': [
            ('query', 'string', 'Search query'),
            ('option', 'any', 'Option ID or name')
        ],
        'returns': 'None',
        'example': 'model.UpdateEnvelopeOption("WantsEnvelopes = 1", "Individual")'
    },
    'model.UpdateField': {
        'category': 'Person',
        'desc': 'Update specific field for a person',
        'params': [
            ('person', 'Person', 'Person object'),
            ('field', 'string', 'Field name to update'),
            ('value', 'any', 'New value')
        ],
        'returns': 'None',
        'example': 'model.UpdateField(person, "EmailAddress", "new@email.com")'
    },
    'model.UpdateMemberStatus': {
        'category': 'Person',
        'desc': 'Update member status for people matching query',
        'params': [
            ('query', 'string', 'Search query'),
            ('status', 'any', 'Status ID or status description')
        ],
        'returns': 'None',
        'example': 'model.UpdateMemberStatus("JoinDate > 365", "Member")'
    },
    'model.UpdateNamedField': {
        'category': 'Person',
        'desc': 'Update named field for people matching query',
        'params': [
            ('query', 'string', 'Search query'),
            ('field', 'string', 'Field name'),
            ('value', 'any', 'New value')
        ],
        'returns': 'None',
        'example': 'model.UpdateNamedField("age > 18", "MemberStatusId", 10)'
    },
    'model.UpdateNewMemberClassDate': {
        'category': 'Person',
        'desc': 'Update new member class date for people',
        'params': [
            ('query', 'string', 'Search query'),
            ('date', 'any', 'Date value')
        ],
        'returns': 'None',
        'example': 'model.UpdateNewMemberClassDate("RecentJoin = 1", datetime.datetime.now())'
    },
    'model.UpdateNewMemberClassDateIfNullForLastAttended': {
        'category': 'Person',
        'desc': 'Update new member class date based on last attendance if null',
        'params': [
            ('query', 'string', 'Search query'),
            ('orgId', 'integer', 'Organization ID for last attendance')
        ],
        'returns': 'None',
        'example': 'model.UpdateNewMemberClassDateIfNullForLastAttended("NewMemberClassDate IS NULL", 30)'
    },
    'model.UpdateNewMemberClassStatus': {
        'category': 'Person',
        'desc': 'Update new member class status',
        'params': [
            ('query', 'string', 'Search query'),
            ('status', 'any', 'Status ID or description')
        ],
        'returns': 'None',
        'example': 'model.UpdateNewMemberClassStatus("ClassAttendance = 100", "Completed")'
    },
    'model.UpdatePerson': {
        'category': 'Person',
        'desc': 'Update multiple fields for a person at once',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('data', 'dict', 'Dictionary of field names and values')
        ],
        'returns': 'None',
        'example': 'model.UpdatePerson(828, {"EmailAddress": "new@email.com", "CellPhone": "555-1234"})'
    },
    
    # ===== SMS =====
    'model.SendSms': {
        'category': 'Sms',
        'desc': 'Queue SMS text message to people from query',
        'params': [
            ('query', 'string', 'Query to select recipients'),
            ('iSendGroup', 'integer', 'SMS sending group ID from SMSGroups table'),
            ('sTitle', 'string', 'Subject/title (max 150 chars, not sent in SMS)'),
            ('sMessage', 'string', 'SMS message text (max 160 chars)')
        ],
        'returns': 'None',
        'example': 'model.SendSms("IsMemberOf = 30", 1, "Meeting Reminder", "Bible study tonight at 7pm")'
    },
    'model.CreateTinyUrl': {
        'category': 'Sms',
        'desc': 'Generate shortened URL using configured shortener service',
        'params': [('url', 'string', 'URL to shorten')],
        'returns': 'Shortened URL or original if service fails',
        'example': 'shortUrl = model.CreateTinyUrl("https://church.org/very/long/url/to/event")'
    },
    
    # ===== TASKS & NOTES =====
    'model.CreateTaskNote': {
        'category': 'Tasks & Notes',
        'desc': 'Create task or note with full options',
        'params': [
            ('ownerId', 'integer', 'Owner PeopleId (required)'),
            ('aboutPersonId', 'integer', 'About PeopleId (required)'),
            ('assigneeId', 'integer', 'Assignee PeopleId (optional)'),
            ('roleId', 'integer', 'Role ID to limit to (optional)'),
            ('isNote', 'string|boolean', 'True/False or "True"/"False" or 1/0'),
            ('instructions', 'string', 'Task/Note instructions'),
            ('notes', 'string', 'Task/Note notes'),
            ('dueDate', 'date', 'Due date (optional)'),
            ('keywordIdList', 'list', 'Keyword IDs list (optional)'),
            ('sendEmails', 'boolean', 'Send notification email (default True)')
        ],
        'returns': 'Task/Note ID',
        'example': 'taskId = model.CreateTaskNote(1, 828, 2, None, False, "Follow up", "Call about visit", datetime.datetime.now() + datetime.timedelta(days=7), [1,2,3], True)'
    },
    'model.GetTaskNote': {
        'category': 'Tasks & Notes',
        'desc': 'Get specific task/note by ID',
        'params': [('taskNoteId', 'integer', 'Task/Note ID')],
        'returns': 'Task/Note object',
        'example': 'task = model.GetTaskNote(123)'
    },
    'model.GetTasksNotes': {
        'category': 'Tasks & Notes',
        'desc': 'Get tasks/notes with filters',
        'params': [
            ('numOfTaskNotes', 'integer', 'Number to return (optional)'),
            ('notesOnly', 'boolean', 'Only return notes (optional)'),
            ('orgId', 'integer', 'Organization ID filter (optional)'),
            ('statuses', 'list', 'Status IDs to filter (optional)'),
            ('peopleId', 'integer', 'Person ID filter (optional)'),
            ('filter', 'string', 'Search filter string (optional)')
        ],
        'returns': 'List of Task/Note objects',
        'example': 'tasks = model.GetTasksNotes(10, False, 30, [1,2], 828, "follow")'
    },
    'model.GetExtraQuestions': {
        'category': 'Tasks & Notes',
        'desc': 'Get extra questions for task/note',
        'params': [('taskNoteId', 'integer', 'Task/Note ID')],
        'returns': 'List of extra questions',
        'example': 'questions = model.GetExtraQuestions(123)'
    },
    'model.GetExtraQuestionsByKeywords': {
        'category': 'Tasks & Notes',
        'desc': 'Get extra questions by keyword IDs',
        'params': [('keywords', 'list', 'List of keyword IDs')],
        'returns': 'List of extra questions',
        'example': 'questions = model.GetExtraQuestionsByKeywords([1, 2, 3])'
    },
    'model.GetExtraQuestionsWithAnswers': {
        'category': 'Tasks & Notes',
        'desc': 'Get extra questions with answers for task/note',
        'params': [('taskNoteId', 'integer', 'Task/Note ID')],
        'returns': 'List of questions with answers',
        'example': 'qa = model.GetExtraQuestionsWithAnswers(123)'
    },
    'model.GetTaskNoteKeywordIds': {
        'category': 'Tasks & Notes',
        'desc': 'Get keyword IDs for task/note',
        'params': [('taskNoteId', 'integer', 'Task/Note ID')],
        'returns': 'List of keyword IDs',
        'example': 'keywords = model.GetTaskNoteKeywordIds(123)'
    },
    'model.SetTaskNoteKeywordIds': {
        'category': 'Tasks & Notes',
        'desc': 'Set keyword IDs for task/note (replaces existing)',
        'params': [
            ('taskNoteId', 'integer', 'Task/Note ID'),
            ('keywordIds', 'list', 'Array of keyword IDs')
        ],
        'returns': 'None',
        'example': 'model.SetTaskNoteKeywordIds(123, [1, 2, 3])'
    },
    'model.SetTaskNoteStatus': {
        'category': 'Tasks & Notes',
        'desc': 'Set task/note status',
        'params': [
            ('taskNoteId', 'integer', 'Task/Note ID'),
            ('status', 'integer|string', 'Status: Complete/Pending/Accepted/Declined/Archived/Note or ID')
        ],
        'returns': 'None',
        'example': 'model.SetTaskNoteStatus(123, "Complete")'
    },
    'model.TaskNoteAccept': {
        'category': 'Tasks & Notes',
        'desc': 'Accept a task',
        'params': [('taskNoteId', 'integer', 'Task/Note ID')],
        'returns': 'None',
        'example': 'model.TaskNoteAccept(123)'
    },
    'model.TaskNoteComplete': {
        'category': 'Tasks & Notes',
        'desc': 'Complete a task with note and extra questions',
        'params': [
            ('taskNoteId', 'integer', 'Task/Note ID'),
            ('note', 'string', 'Completion note'),
            ('extraQuestions', 'list', 'Extra questions with answers')
        ],
        'returns': 'None',
        'example': 'model.TaskNoteComplete(123, "Completed successfully", [])'
    },
    'model.TaskNoteDecline': {
        'category': 'Tasks & Notes',
        'desc': 'Decline a task with reason',
        'params': [
            ('taskNoteId', 'integer', 'Task/Note ID'),
            ('reason', 'string', 'Reason for declining')
        ],
        'returns': 'None',
        'example': 'model.TaskNoteDecline(123, "Not available")'
    },
    'model.TaskNoteEdit': {
        'category': 'Tasks & Notes',
        'desc': 'Edit task/note with view model',
        'params': [('viewModel', 'object', 'Task/Note view model with updated values')],
        'returns': 'None',
        'example': 'model.TaskNoteEdit(updatedViewModel)'
    },
    'model.TaskNoteMassAddKeywords': {
        'category': 'Tasks & Notes',
        'desc': 'Mass add keywords to tasks/notes',
        'params': [
            ('taskNoteIdList', 'list', 'List of Task/Note IDs'),
            ('keywordsArray', 'list', 'List of keyword IDs to add')
        ],
        'returns': 'None',
        'example': 'model.TaskNoteMassAddKeywords([123, 124, 125], [1, 2])'
    },
    'model.TaskNoteMassArchive': {
        'category': 'Tasks & Notes',
        'desc': 'Mass archive tasks/notes',
        'params': [('taskNoteIdList', 'list', 'List of Task/Note IDs')],
        'returns': 'None',
        'example': 'model.TaskNoteMassArchive([123, 124, 125])'
    },
    'model.TaskNoteMassAssign': {
        'category': 'Tasks & Notes',
        'desc': 'Mass assign tasks to person',
        'params': [
            ('taskNoteIdList', 'list', 'List of Task/Note IDs'),
            ('assigneeId', 'integer', 'Assignee PeopleId')
        ],
        'returns': 'None',
        'example': 'model.TaskNoteMassAssign([123, 124, 125], 828)'
    },
    'model.TaskNoteMassComplete': {
        'category': 'Tasks & Notes',
        'desc': 'Mass complete tasks',
        'params': [
            ('taskNoteIdList', 'list', 'List of Task/Note IDs'),
            ('completionDate', 'date', 'Completion date'),
            ('note', 'string', 'Completion note')
        ],
        'returns': 'None',
        'example': 'model.TaskNoteMassComplete([123, 124], datetime.datetime.now(), "All done")'
    },
    'model.TaskNoteMassDelete': {
        'category': 'Tasks & Notes',
        'desc': 'Mass delete tasks/notes',
        'params': [('taskNoteIdList', 'list', 'List of Task/Note IDs')],
        'returns': 'None',
        'example': 'model.TaskNoteMassDelete([123, 124, 125])'
    },
    'model.TaskNoteMassUnarchive': {
        'category': 'Tasks & Notes',
        'desc': 'Mass unarchive tasks/notes',
        'params': [('taskNoteIdList', 'list', 'List of Task/Note IDs')],
        'returns': 'None',
        'example': 'model.TaskNoteMassUnarchive([123, 124, 125])'
    },
    
    # ===== MISCELLANEOUS =====
    'model.CallScript': {
        'category': 'Misc',
        'desc': 'Call another Python script',
        'params': [('scriptName', 'string', 'Script name')],
        'returns': 'Script output',
        'example': 'result = model.CallScript("MyOtherScript")'
    },
    'model.FmtPhone': {
        'category': 'Misc',
        'desc': 'Format phone number',
        'params': [('phone', 'string', 'Phone number')],
        'returns': 'Formatted phone',
        'example': 'formatted = model.FmtPhone("5551234567")'
    },
    'model.FmtZip': {
        'category': 'Misc',
        'desc': 'Format zip code',
        'params': [('zip', 'string', 'Zip code')],
        'returns': 'Formatted zip',
        'example': 'formatted = model.FmtZip("123456789")'
    },
    'model.HtmlContent': {
        'category': 'Misc',
        'desc': 'Get saved HTML content',
        'params': [('name', 'string', 'Content name')],
        'returns': 'HTML string',
        'example': 'html = model.HtmlContent("WelcomeMessage")'
    },
    'model.TextContent': {
        'category': 'Misc',
        'desc': 'Get saved text content',
        'params': [('name', 'string', 'Content name')],
        'returns': 'Text string',
        'example': 'text = model.TextContent("EmailSignature")'
    },
    'model.SqlContent': {
        'category': 'Misc',
        'desc': 'Get saved SQL content',
        'params': [('name', 'string', 'SQL script name')],
        'returns': 'SQL string',
        'example': 'sql = model.SqlContent("CustomQuery")'
    },
    'model.Setting': {
        'category': 'Misc',
        'desc': 'Get system setting',
        'params': [('name', 'string', 'Setting name')],
        'returns': 'Setting value',
        'example': 'value = model.Setting("ChurchName")'
    },
    'model.SetSetting': {
        'category': 'Misc',
        'desc': 'Set system setting',
        'params': [
            ('name', 'string', 'Setting name'),
            ('value', 'string', 'Setting value')
        ],
        'returns': 'None',
        'example': 'model.SetSetting("CustomSetting", "Value")'
    },
    'model.RestGet': {
        'category': 'Misc',
        'desc': 'Make REST GET request',
        'params': [
            ('url', 'string', 'URL'),
            ('headers', 'dict', 'Headers (optional)')
        ],
        'returns': 'Response text',
        'example': 'response = model.RestGet("https://api.example.com/data")'
    },
    'model.RestPost': {
        'category': 'Misc',
        'desc': 'Make REST POST request',
        'params': [
            ('url', 'string', 'API URL'),
            ('headers', 'dict', 'Headers dictionary'),
            ('obj', 'object', 'POST body'),
            ('user', 'string', 'Username (optional)'),
            ('password', 'string', 'Password (optional)')
        ],
        'returns': 'Response text',
        'example': 'response = model.RestPost("https://api.example.com/submit", {}, "data")'
    },
    'model.AppendIfBoth': {
        'category': 'Misc',
        'desc': 'Concatenate strings with separator if both have values',
        'params': [
            ('s1', 'string', 'First string'),
            ('join', 'string', 'Join text'),
            ('s2', 'string', 'Second string')
        ],
        'returns': 'Concatenated string or just s1',
        'example': 'result = model.AppendIfBoth("First", " - ", "Second")'
    },
    'model.Content': {
        'category': 'Misc',
        'desc': 'Get content file by name',
        'params': [
            ('name', 'string', 'Content file name')
        ],
        'returns': 'Content string',
        'example': 'content = model.Content("MyContent")'
    },
    'model.CreateCustomView': {
        'category': 'Misc',
        'desc': 'Create custom database view (requires developer and admin roles)',
        'params': [
            ('view', 'string', 'View name (alphanumeric)'),
            ('sql', 'string', 'SQL query defining the view')
        ],
        'returns': 'None',
        'example': 'model.CreateCustomView("MyView", "SELECT * FROM People WHERE Age > 18")'
    },
    'model.CreateQueryTag': {
        'category': 'Misc',
        'desc': 'Create QueryTag from Search Builder code',
        'params': [
            ('name', 'string', 'QueryTag name'),
            ('code', 'string', 'Search Builder code')
        ],
        'returns': 'Count of people in tag',
        'example': 'count = model.CreateQueryTag("ActiveMembers", "IsMemberOf = 1[True]")'
    },
    'model.CsvReader': {
        'category': 'Misc',
        'desc': 'Create CSV reader from text',
        'params': [
            ('text', 'string', 'CSV text content'),
            ('delimiter', 'string', 'Delimiter character (optional)')
        ],
        'returns': 'CsvReader object',
        'example': 'csv = model.CsvReader(model.Data.file)\nwhile csv.Read():\n    date = csv["Date"]'
    },
    'model.CsvReaderNoHeader': {
        'category': 'Misc',
        'desc': 'Create CSV reader without headers',
        'params': [
            ('text', 'string', 'CSV text content')
        ],
        'returns': 'CsvReader object',
        'example': 'csv = model.CsvReaderNoHeader(text)\nwhile csv.Read():\n    date = csv[0]'
    },
    'model.CustomStatementsFundIdList': {
        'category': 'Misc',
        'desc': 'Get comma-separated list of fund IDs',
        'params': [
            ('name', 'string', 'FundList name')
        ],
        'returns': 'Comma-separated fund IDs string',
        'example': 'fundIds = model.CustomStatementsFundIdList("GeneralFunds")'
    },
    'model.CmsHost': {
        'category': 'Misc',
        'desc': 'Get church database host URL',
        'params': [],
        'returns': 'Host URL string',
        'example': 'url = model.CmsHost + "/Person2/" + str(peopleId)'
    },
    'model.DataHas': {
        'category': 'Misc',
        'desc': 'Check if Data has property',
        'params': [
            ('key', 'string', 'Property key (case sensitive)')
        ],
        'returns': 'Boolean',
        'example': 'if model.DataHas("orgId"):\n    orgId = model.Data.orgId'
    },
    'model.DeleteQueryTags': {
        'category': 'Misc',
        'desc': 'Delete QueryTags by pattern',
        'params': [
            ('namelike', 'string', 'Name pattern with % wildcards')
        ],
        'returns': 'None',
        'example': 'model.DeleteQueryTags("Project_%")'
    },
    'model.DynamicDataFromJson': {
        'category': 'Misc',
        'desc': 'Create DynamicData from JSON string',
        'params': [
            ('json', 'string', 'JSON string')
        ],
        'returns': 'DynamicData object',
        'example': 'data = model.DynamicDataFromJson(\'{"name":"John","age":30}\')'
    },
    'model.DynamicDataFromJsonArray': {
        'category': 'Misc',
        'desc': 'Create list of DynamicData from JSON array',
        'params': [
            ('json', 'string', 'JSON array string')
        ],
        'returns': 'List of DynamicData objects',
        'example': 'dataList = model.DynamicDataFromJsonArray(\'[{"id":1},{"id":2}]\')'
    },
    'model.ElementList': {
        'category': 'Misc',
        'desc': 'Extract property values from DynamicData collection',
        'params': [
            ('array', 'IEnumerable<DynamicData>', 'Collection of DynamicData'),
            ('name', 'string', 'Property name to extract')
        ],
        'returns': 'List of string values',
        'example': 'names = model.ElementList(people, "Name")'
    },
    'model.FormatJson': {
        'category': 'Misc',
        'desc': 'Format JSON string or object with indentation',
        'params': [
            ('json_or_data', 'string|object', 'JSON string or Python object')
        ],
        'returns': 'Formatted JSON string',
        'example': 'formatted = model.FormatJson(data)'
    },
    'model.GetCacheVariable': {
        'category': 'Misc',
        'desc': 'Get cached variable value',
        'params': [
            ('name', 'string', 'Cache variable name')
        ],
        'returns': 'Cached value or empty string',
        'example': 'value = model.GetCacheVariable("TempData")'
    },
    'model.JsonDeserialize': {
        'category': 'Misc',
        'desc': 'Deserialize JSON to dynamic object',
        'params': [
            ('jsontext', 'string', 'JSON text')
        ],
        'returns': 'Dynamic object',
        'example': 'obj = model.JsonDeserialize(jsonString)'
    },
    'model.JsonDeserialize2': {
        'category': 'Misc',
        'desc': 'Deserialize JSON array to list of dictionaries',
        'params': [
            ('jsontext', 'string', 'JSON array text')
        ],
        'returns': 'List of dictionary objects',
        'example': 'items = model.JsonDeserialize2(jsonArrayString)'
    },
    'model.JsonSerialize': {
        'category': 'Misc',
        'desc': 'Serialize object to JSON string',
        'params': [
            ('obj', 'object', 'Object to serialize')
        ],
        'returns': 'JSON string',
        'example': 'json = model.JsonSerialize(data)'
    },
    'model.Markdown': {
        'category': 'Misc',
        'desc': 'Convert markdown to HTML',
        'params': [
            ('text', 'string', 'Markdown text')
        ],
        'returns': 'HTML string',
        'example': 'html = model.Markdown("# Heading\\n**Bold text**")'
    },
    'model.Md5Hash': {
        'category': 'Misc',
        'desc': 'Generate MD5 hash of text',
        'params': [
            ('text', 'string', 'Text to hash')
        ],
        'returns': 'MD5 hash string',
        'example': 'hash = model.Md5Hash("password123")'
    },
    'model.RegexMatch': {
        'category': 'Misc',
        'desc': 'Match regex pattern in string',
        'params': [
            ('s', 'string', 'Target string'),
            ('regex', 'string', 'Regex pattern')
        ],
        'returns': 'Matched text or None',
        'example': 'match = model.RegexMatch("abc123", r"\\d+")'
    },
    'model.Replace': {
        'category': 'Misc',
        'desc': 'Replace all occurrences in text',
        'params': [
            ('text', 'string', 'Text to modify'),
            ('pattern', 'string', 'Pattern to replace'),
            ('replacement', 'string', 'Replacement text')
        ],
        'returns': 'Modified string',
        'example': 'result = model.Replace("Hello World", "World", "Python")'
    },
    'model.RestDelete': {
        'category': 'Misc',
        'desc': 'Make REST DELETE request',
        'params': [
            ('url', 'string', 'API URL'),
            ('headers', 'dict', 'Headers dictionary'),
            ('user', 'string', 'Username (optional)'),
            ('password', 'string', 'Password (optional)')
        ],
        'returns': 'Response text',
        'example': 'response = model.RestDelete("https://api.example.com/item/123", {})'
    },
    'model.RestPostJson': {
        'category': 'Misc',
        'desc': 'Make REST POST request with JSON body',
        'params': [
            ('url', 'string', 'API URL'),
            ('headers', 'dict', 'Headers dictionary'),
            ('obj', 'object', 'Object to serialize as JSON'),
            ('user', 'string', 'Username (optional)'),
            ('password', 'string', 'Password (optional)')
        ],
        'returns': 'Response text',
        'example': 'response = model.RestPostJson(url, {}, data)'
    },
    'model.RestPostXml': {
        'category': 'Misc',
        'desc': 'Make REST POST request with XML body',
        'params': [
            ('url', 'string', 'API URL'),
            ('headers', 'dict', 'Headers dictionary'),
            ('body', 'string', 'XML body'),
            ('user', 'string', 'Username (optional)'),
            ('password', 'string', 'Password (optional)')
        ],
        'returns': 'Response text',
        'example': 'response = model.RestPostXml(url, {}, xmlString)'
    },
    'model.SetCacheVariable': {
        'category': 'Misc',
        'desc': 'Set cache variable with 1-minute expiration',
        'params': [
            ('name', 'string', 'Cache variable name'),
            ('value', 'string', 'Value to cache')
        ],
        'returns': 'None',
        'example': 'model.SetCacheVariable("TempData", "value")'
    },
    'model.SpaceCamelCase': {
        'category': 'Misc',
        'desc': 'Add spaces to CamelCase string',
        'params': [
            ('s', 'string', 'CamelCase string')
        ],
        'returns': 'Space-separated string',
        'example': 'spaced = model.SpaceCamelCase("MyVariableName")'
    },
    'model.TitleContent': {
        'category': 'Misc',
        'desc': 'Get title attribute of content',
        'params': [
            ('name', 'string', 'Content name')
        ],
        'returns': 'Title string',
        'example': 'title = model.TitleContent("EmailTemplate")'
    },
    'model.Trim': {
        'category': 'Misc',
        'desc': 'Remove leading/trailing whitespace',
        'params': [
            ('s', 'string', 'String to trim')
        ],
        'returns': 'Trimmed string',
        'example': 'trimmed = model.Trim("  text  ")'
    },
    'model.UrlEncode': {
        'category': 'Misc',
        'desc': 'URL-encode a string',
        'params': [
            ('s', 'string', 'String to encode')
        ],
        'returns': 'URL-encoded string',
        'example': 'encoded = model.UrlEncode("hello world")'
    },
    'model.UserPeopleId': {
        'category': 'Misc',
        'desc': 'Get current user PeopleId',
        'params': [],
        'returns': 'PeopleId or None',
        'example': 'peopleId = model.UserPeopleId'
    },
    'model.UserIsInRole': {
        'category': 'Misc',
        'desc': 'Check if user has role',
        'params': [
            ('role', 'string', 'Role name to check')
        ],
        'returns': 'Boolean',
        'example': 'if model.UserIsInRole("Admin"):\n    print "User is admin"'
    },
    'model.WriteContentHtml': {
        'category': 'Misc',
        'desc': 'Write HTML content to Special Content. NOTE: If submitting via web form, HTML must be entity-encoded to bypass ASP.NET validation',
        'params': [
            ('name', 'string', 'File name'),
            ('text', 'string', 'HTML content (entity-encode if from web form)'),
            ('keyword', 'string', 'Optional keyword')
        ],
        'returns': 'None',
        'example': '# Direct Python call:\nmodel.WriteContentHtml("NewPage", "<h1>Title</h1>", "")\n\n# Direct call with IronPython 2.7 HTML encoding:\nimport cgi\nhtml_content = "<h1>Title</h1><p>Content with & special chars</p>"\nescaped = cgi.escape(html_content)  # Encodes < > & \nmodel.WriteContentHtml("Page", html_content, "")  # Use original, not escaped\n\n# From web form (JavaScript encoding):\n# JS: content.replace(/</g, "&lt;").replace(/>/g, "&gt;")\n# Then decode in Python:\n# content = content.replace("&lt;", "<").replace("&gt;", ">")'
    },
    'model.WriteContentPython': {
        'category': 'Misc',
        'desc': 'Write Python script to Special Content',
        'params': [
            ('name', 'string', 'Script name'),
            ('script', 'string', 'Python script'),
            ('keyword', 'string', 'Optional keyword')
        ],
        'returns': 'None',
        'example': 'model.WriteContentPython("NewScript", "print \'Hello\'")'
    },
    'model.WriteContentSql': {
        'category': 'Misc',
        'desc': 'Write SQL script to Special Content',
        'params': [
            ('name', 'string', 'Script name'),
            ('sql', 'string', 'SQL code'),
            ('keyword', 'string', 'Optional keyword')
        ],
        'returns': 'None',
        'example': 'model.WriteContentSql("Query", "SELECT * FROM People")'
    },
    'model.WriteContentText': {
        'category': 'Misc',
        'desc': 'Write text file to Special Content',
        'params': [
            ('name', 'string', 'File name'),
            ('text', 'string', 'Text content'),
            ('keyword', 'string', 'Optional keyword')
        ],
        'returns': 'None',
        'example': 'model.WriteContentText("Notes", "My notes here")'
    },
    
    # ===== JSON DOCUMENT RECORDS =====
    'model.AddUpdateJsonRecord': {
        'category': 'JsonDocumentRecords',
        'desc': 'Add or update a JSON document record using JSON string or DynamicData object',
        'params': [
            ('json_or_data', 'string|DynamicData', 'JSON string or DynamicData object to store'),
            ('section', 'string', 'Section identifier for the record'),
            ('pk1', 'object', 'Primary key value 1 (required)'),
            ('pk2', 'object', 'Primary key value 2 (optional)'),
            ('pk3', 'object', 'Primary key value 3 (optional)'),
            ('pk4', 'object', 'Primary key value 4 (optional)')
        ],
        'returns': 'None',
        'example': '# With JSON string:\nmodel.AddUpdateJsonRecord(\'{"name":"John","age":30}\', "users", 123)\n# With DynamicData:\ndata = model.DynamicData()\ndata.name = "John"\ndata.age = 30\nmodel.AddUpdateJsonRecord(data, "users", 123)'
    },
    'model.DeleteJsonRecord': {
        'category': 'JsonDocumentRecords',
        'desc': 'Delete a JSON document record',
        'params': [
            ('section', 'string', 'Section identifier for the record'),
            ('pk1', 'object', 'Primary key value 1 (required)'),
            ('pk2', 'object', 'Primary key value 2 (optional)'),
            ('pk3', 'object', 'Primary key value 3 (optional)'),
            ('pk4', 'object', 'Primary key value 4 (optional)')
        ],
        'returns': 'None',
        'example': 'model.DeleteJsonRecord("users", 123)'
    },
    'model.DeleteJsonRecordSection': {
        'category': 'JsonDocumentRecords',
        'desc': 'Delete all JSON document records in a section',
        'params': [
            ('section', 'string', 'Section identifier for the records to delete')
        ],
        'returns': 'None',
        'example': 'model.DeleteJsonRecordSection("temp_data")'
    },
    
    # ===== SQL DYNAMIC DATA =====
    'model.SqlListDynamicData': {
        'category': 'SqlDynamicData',
        'desc': 'Execute SQL query and return results as list of DynamicData objects',
        'params': [
            ('sql', 'string', 'SQL query to execute'),
            ('metadata', 'DynamicData', 'Optional metadata for type conversion')
        ],
        'returns': 'List of DynamicData objects',
        'example': 'results = model.SqlListDynamicData("SELECT * FROM People WHERE Age > 18")'
    },
    'model.SqlTop1DynamicData': {
        'category': 'SqlDynamicData',
        'desc': 'Execute SQL query and return first row as DynamicData object',
        'params': [
            ('sql', 'string', 'SQL query to execute'),
            ('metadata', 'DynamicData', 'Optional metadata for type conversion')
        ],
        'returns': 'DynamicData object or None',
        'example': 'person = model.SqlTop1DynamicData("SELECT * FROM People WHERE PeopleId = 123")'
    },
    'model.SqlList': {
        'category': 'SqlDynamicData',
        'desc': 'Execute SQL query and return results as list of dynamic objects',
        'params': [
            ('sql', 'string', 'SQL query to execute')
        ],
        'returns': 'List of dynamic objects',
        'example': 'results = model.SqlList("SELECT PeopleId, Name FROM People")'
    },
    'model.SqlGrid': {
        'category': 'SqlDynamicData',
        'desc': 'Execute SQL query and return results as HTML table with formatting',
        'params': [
            ('sql', 'string', 'SQL query to execute and display as grid')
        ],
        'returns': 'HTML string with formatted table',
        'example': 'sql = """\n    SELECT TOP 10\n        p.PeopleId,\n        p.Name,\n        p.Age,\n        o.OrganizationId,\n        o.OrganizationName\n    FROM People p\n    JOIN OrganizationMembers om ON om.PeopleId = p.PeopleId\n    JOIN Organizations o ON o.OrganizationId = om.OrganizationId\n    WHERE p.Age > 18\n    ORDER BY p.Name\n"""\nprint model.SqlGrid(sql)'
    },
    
    # ===== DYNAMICDATA =====
    'model.DynamicData': {
        'category': 'DynamicData',
        'desc': 'Create dynamic data object (empty or from dictionary)',
        'params': [
            ('datadict', 'dict', 'Optional dictionary to initialize with')
        ],
        'returns': 'New DynamicData object',
        'example': 'data = model.DynamicData()\ndata.Name = "John"\ndata.Age = 30\n# Or from dict:\ndata2 = model.DynamicData({"Name": "Jane", "Age": 25})'
    },
    'DynamicData.GetValue': {
        'category': 'DynamicData',
        'desc': 'Get value from DynamicData by key',
        'params': [('key', 'string', 'Dictionary key to look up')],
        'returns': 'Value or None if key not found',
        'example': 'data = model.DynamicData()\ndata.Name = "John"\nvalue = data.GetValue("Name")'
    },
    'DynamicData.Remove': {
        'category': 'DynamicData',
        'desc': 'Remove entry from DynamicData by key if exists',
        'params': [('name', 'string', 'Dictionary key to remove')],
        'returns': 'None',
        'example': 'data.Remove("TempField")'
    },
    'DynamicData.AddValue': {
        'category': 'DynamicData',
        'desc': 'Add or update value in DynamicData dictionary',
        'params': [
            ('name', 'string', 'Dictionary key to add or update'),
            ('value', 'any', 'Value to store')
        ],
        'returns': 'None',
        'example': 'data.AddValue("Status", "Active")'
    },
    'DynamicData.SetValue': {
        'category': 'DynamicData',
        'desc': 'Set value with automatic type conversion from string',
        'params': [
            ('name', 'string', 'Dictionary key to add or update'),
            ('value', 'string', 'String value (converts to numeric if possible)')
        ],
        'returns': 'None',
        'example': 'data.SetValue("Count", "123")  # Automatically converts to integer'
    },
    'DynamicData.ToString': {
        'category': 'DynamicData',
        'desc': 'Get formatted JSON string representation',
        'params': [],
        'returns': 'JSON string representation',
        'example': 'jsonStr = data.ToString()\nprint jsonStr'
    },
    'DynamicData.ToFlatString': {
        'category': 'DynamicData',
        'desc': 'Get flat JSON with empty values removed and quotes escaped',
        'params': [],
        'returns': 'Flat JSON string',
        'example': 'flatJson = data.ToFlatString()'
    },
    'DynamicData.Keys': {
        'category': 'DynamicData',
        'desc': 'Get list of keys in DynamicData dictionary',
        'params': [('metadata', 'DynamicData', 'Optional metadata for filtering keys')],
        'returns': 'List of key strings',
        'example': 'keys = data.Keys()\nfor key in keys:\n    print key, "=", data.GetValue(key)'
    },
    'DynamicData.SpecialKeys': {
        'category': 'DynamicData',
        'desc': 'Get keys marked as special in metadata',
        'params': [('metadata', 'DynamicData', 'Metadata for identifying special keys')],
        'returns': 'List of special key strings',
        'example': 'meta = model.DynamicData()\nmeta.Special = ["Priority", "Status"]\nspecialKeys = data.SpecialKeys(meta)'
    },
    
    # ===== QUERY FUNCTIONS - SQL =====
    'q.QuerySql': {
        'category': 'Query SQL',
        'desc': 'Execute SQL query returning multiple rows',
        'params': [
            ('sql', 'string', 'SQL query'),
            ('p1', 'any', 'Optional parameter @p1'),
            ('declarations', 'dict', 'Optional named parameters')
        ],
        'returns': 'List of results',
        'example': 'results = q.QuerySql("SELECT * FROM People WHERE LastName = \'Smith\'")'
    },
    'q.QuerySqlTop1': {
        'category': 'Query SQL',
        'desc': 'Get first row as dynamic object',
        'params': [
            ('sql', 'string', 'SQL query'),
            ('p1', 'any', 'Optional parameter @p1'),
            ('declarations', 'dict', 'Optional named parameters')
        ],
        'returns': 'Single result or None',
        'example': 'person = q.QuerySqlTop1("SELECT * FROM People WHERE PeopleId = 828")'
    },
    'q.QuerySqlInt': {
        'category': 'Query SQL',
        'desc': 'Get integer result',
        'params': [('sql', 'string', 'SQL query')],
        'returns': 'Integer value',
        'example': 'count = q.QuerySqlInt("SELECT COUNT(*) FROM People")'
    },
    'q.QuerySqlList': {
        'category': 'Query SQL',
        'desc': 'Get list from first column',
        'params': [('sql', 'string', 'SQL query')],
        'returns': 'List of values',
        'example': 'ids = q.QuerySqlList("SELECT PeopleId FROM People WHERE Age > 18")'
    },
    # ===== QUERY FUNCTIONS - PEOPLE =====
    'q.QueryCount': {
        'category': 'Query People',
        'desc': 'Count people matching query',
        'params': [('query', 'string', 'Query parameter or saved search name')],
        'returns': 'Count of people',
        'example': 'count = q.QueryCount("IsMemberOf = 30")'
    },
    'q.QueryList': {
        'category': 'Query People',
        'desc': 'Get list of people matching query (max 1000)',
        'params': [
            ('query', 'string', 'Query parameter'),
            ('sort', 'string', 'Sort parameter (optional)')
        ],
        'returns': 'List of Person objects',
        'example': 'people = q.QueryList("IsMemberOf = 30", "name")'
    },
    'q.QueryPeopleIds': {
        'category': 'Query People',
        'desc': 'Get list of PeopleIds matching query',
        'params': [('query', 'string', 'Query parameter')],
        'returns': 'List of PeopleIds',
        'example': 'peopleIds = q.QueryPeopleIds("Age > 18")'
    },
    'q.SqlPeopleIdsToQuery': {
        'category': 'Query People',
        'desc': 'Convert SQL result to query string',
        'params': [('sql', 'string', 'SQL returning PeopleId column')],
        'returns': 'Query string like peopleids=1001,1002,1003',
        'example': 'query = q.SqlPeopleIdsToQuery("SELECT PeopleId FROM People WHERE Age > 65")'
    },
    'q.QuerySqlPeopleIds': {
        'category': 'Query People',
        'desc': 'Get PeopleIds from SQL query',
        'params': [('sql', 'string', 'SQL returning single column of PeopleIds')],
        'returns': 'List of PeopleIds',
        'example': 'peopleIds = q.QuerySqlPeopleIds("SELECT PeopleId FROM People WHERE City = \'Memphis\'")'
    },
    'q.QuerySqlInts': {
        'category': 'Query People',
        'desc': 'Get integers from SQL query',
        'params': [('sql', 'string', 'SQL returning single column of integers')],
        'returns': 'List of integers',
        'example': 'ids = q.QuerySqlInts("SELECT DISTINCT FamilyId FROM People")'
    },
    'q.QuerySqlScalar': {
        'category': 'Query People',
        'desc': 'Get single string value from SQL',
        'params': [('sql', 'string', 'SQL returning single string value')],
        'returns': 'String value',
        'example': 'name = q.QuerySqlScalar("SELECT Name FROM People WHERE PeopleId = 828")'
    },
    'q.SqlNameCountArray': {
        'category': 'Query People',
        'desc': 'Get name/count array for Google Charts',
        'params': [
            ('title', 'string', 'Title for data points'),
            ('sql', 'string', 'SQL returning Name, Cnt columns')
        ],
        'returns': 'JSON array for Google Charts',
        'example': 'data = q.SqlNameCountArray("Members", "SELECT MemberStatus, COUNT(*) FROM People GROUP BY MemberStatus")'
    },
    'q.StatusCount': {
        'category': 'Query People',
        'desc': 'Count people with status flags',
        'params': [('flags', 'string', 'Comma-separated status flags (F01-F99)')],
        'returns': 'Count of people with all flags',
        'example': 'count = q.StatusCount("F01,F02")'
    },
    'q.TagCount': {
        'category': 'Query People',
        'desc': 'Count people in tag',
        'params': [('tagid', 'integer', 'Tag ID')],
        'returns': 'Count of people in tag',
        'example': 'count = q.TagCount(tagId)'
    },
    'q.TagQueryList': {
        'category': 'Query People',
        'desc': 'Create temporary tag from query',
        'params': [('query', 'string', 'Query parameter')],
        'returns': 'Tag ID',
        'example': 'tagId = q.TagQueryList("IsMemberOf = 30")'
    },
    'q.TagSqlPeopleIds': {
        'category': 'Query People',
        'desc': 'Create temporary tag from SQL',
        'params': [('sql', 'string', 'SQL returning PeopleIds')],
        'returns': 'Tag ID',
        'example': 'tagId = q.TagSqlPeopleIds("SELECT PeopleId FROM People WHERE Age > 65")'
    },
    'q.GetWhereClause': {
        'category': 'Query People',
        'desc': 'Get SQL WHERE clause from search builder code',
        'params': [('code', 'string', 'Search builder code')],
        'returns': 'SQL WHERE clause string',
        'example': 'where = q.GetWhereClause("IsMemberOf = 1[True]")'
    },
    'q.SqlNameValues': {
        'category': 'Query People',
        'desc': 'Create DynamicData from SQL name/value columns',
        'params': [
            ('sql', 'string', 'SQL query'),
            ('namecol', 'string', 'Column name for property names'),
            ('valuecol', 'string', 'Column name for property values')
        ],
        'returns': 'DynamicData object',
        'example': 'data = q.SqlNameValues("SELECT Status, COUNT(*) FROM People GROUP BY Status", "Status", "COUNT")'
    },
    'q.SqlFirstColumnRowKey': {
        'category': 'Query People',
        'desc': 'Create DynamicData with first column as property names',
        'params': [
            ('sql', 'string', 'SQL query'),
            ('declarations', 'dict', 'Optional parameters (optional)')
        ],
        'returns': 'DynamicData object',
        'example': 'data = q.SqlFirstColumnRowKey("SELECT MemberType, COUNT(*) as Cnt FROM People GROUP BY MemberType")'
    },
    'q.BlueToolbarReport': {
        'category': 'Query People',
        'desc': 'Get people for Blue Toolbar report (max 1000)',
        'params': [
            ('sort', 'string', 'Sort parameter (optional)')
        ],
        'returns': 'List of Person objects',
        'example': 'people = q.BlueToolbarReport("name")'
    },
    
    # ===== QUERY FUNCTIONS - ATTENDANCE =====
    'q.LastSunday': {
        'category': 'Query Attendance',
        'desc': 'Most recent Sunday with recorded attendance',
        'params': [],
        'returns': 'DateTime of last Sunday with attendance',
        'example': 'lastSunday = q.LastSunday\nprint "Last Sunday with attendance: " + str(lastSunday)'
    },
    'q.AttendanceTypeCountDateRange': {
        'category': 'Query Attendance',
        'desc': 'Count attendances by type in date range',
        'params': [
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)'),
            ('attendtype', 'string', 'Comma-separated attendance types (names or IDs)'),
            ('startdt', 'datetime', 'Start date'),
            ('days', 'integer', 'Days to look forward')
        ],
        'returns': 'Count of attendances',
        'example': 'visitors = q.AttendanceTypeCountDateRange(0, 0, 0, "New Guest, Recent Guest", datetime.datetime(2024,1,1), 30)'
    },
    'q.AttendCountAsOf': {
        'category': 'Query Attendance',
        'desc': 'Unique count of people attending in date range',
        'params': [
            ('startdt', 'datetime', 'Start date (inclusive)'),
            ('enddt', 'datetime', 'End date (exclusive)'),
            ('guestonly', 'boolean', 'Only include guests'),
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)')
        ],
        'returns': 'Unique count of people',
        'example': 'people = q.AttendCountAsOf(datetime.datetime(2024,1,1), datetime.datetime(2024,2,1), False, 0, 0, 30)'
    },
    'q.AttendMemberTypeCountAsOf': {
        'category': 'Query Attendance',
        'desc': 'Count attendances by member type',
        'params': [
            ('startdt', 'datetime', 'Start date (inclusive)'),
            ('enddt', 'datetime', 'End date (exclusive)'),
            ('membertypes', 'string', 'Include these member types (blank for all)'),
            ('notmembertypes', 'string', 'Exclude these member types'),
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)')
        ],
        'returns': 'Total count of attendances',
        'example': 'members = q.AttendMemberTypeCountAsOf(datetime.datetime(2024,1,1), datetime.datetime(2024,2,1), "Member", "", 101, 0, 0)'
    },
    'q.LastWeekAttendance': {
        'category': 'Query Attendance',
        'desc': 'Count attendances offset from last Sunday',
        'params': [
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('starthour', 'integer', 'Hours offset from Sunday midnight'),
            ('endhour', 'integer', 'Hours offset from Sunday midnight')
        ],
        'returns': 'Count of attendances',
        'example': 'lastWeek = q.LastWeekAttendance(101, 201, -96, 72)  # Wed to Wed'
    },
    'q.MeetingCount': {
        'category': 'Query Attendance',
        'desc': 'Count meetings in past days',
        'params': [
            ('days', 'integer', 'Days to look back'),
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)')
        ],
        'returns': 'Number of meetings',
        'example': 'meetings = q.MeetingCount(7, 0, 0, 30)'
    },
    'q.MeetingCountDateHours': {
        'category': 'Query Attendance',
        'desc': 'Count meetings from date for hours',
        'params': [
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)'),
            ('startdt', 'datetime', 'Start date/time'),
            ('hours', 'integer', 'Hours to look forward')
        ],
        'returns': 'Number of meetings',
        'example': 'meetings = q.MeetingCountDateHours(101, 0, 0, q.LastSunday, 168)'
    },
    'q.NumPresent': {
        'category': 'Query Attendance',
        'desc': 'Count attendees (uses max of headcount or present)',
        'params': [
            ('days', 'integer', 'Days to look back'),
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)')
        ],
        'returns': 'Number of attendees',
        'example': 'present = q.NumPresent(7, 0, 0, 30)'
    },
    'q.NumPresentDateRange': {
        'category': 'Query Attendance',
        'desc': 'Count attendees from date (uses max of headcount or present)',
        'params': [
            ('progid', 'integer', 'Program ID (0 for all)'),
            ('divid', 'integer', 'Division ID (0 for all)'),
            ('orgid', 'integer', 'Organization ID (0 for all)'),
            ('startdt', 'datetime', 'Start date'),
            ('days', 'integer', 'Days to look forward')
        ],
        'returns': 'Number of attendees',
        'example': 'present = q.NumPresentDateRange(0, 0, 30, datetime.datetime(2024,1,1), 30)'
    },
    
    # ===== QUERY FUNCTIONS - CONTRIBUTIONS =====
    'q.ContributionTotals': {
        'category': 'Query Contributions',
        'desc': 'Get total dollar amount of contributions',
        'params': [
            ('days1', 'integer', 'Days before today for start date'),
            ('days2', 'integer', 'Days before today for end date'),
            ('funds', 'string|integer', 'Fund ID(s): 0=all, comma-separated, negative to exclude')
        ],
        'returns': 'Total contribution amount (float)',
        'example': 'total = q.ContributionTotals(30, 0, 0)  # Last 30 days, all funds\ntotal = q.ContributionTotals(365, 0, "1001,1002,-1003")  # Multiple funds, exclude 1003'
    },
    'q.ContributionCount': {
        'category': 'Query Contributions',
        'desc': 'Get count of contributions',
        'params': [
            ('days1', 'integer', 'Days before today for start date'),
            ('days2', 'integer', 'Days before today for end date'),
            ('funds', 'string|integer', 'Fund ID(s): 0=all, comma-separated, negative to exclude')
        ],
        'returns': 'Number of contributions (int)',
        'example': 'count = q.ContributionCount(30, 0, 0)  # Count last 30 days'
    },
    'q.DateRangeForContributionTotals': {
        'category': 'Query Contributions',
        'desc': 'Get date range string for contribution period',
        'params': [
            ('days1', 'integer', 'Days before today for start date'),
            ('days2', 'integer', 'Days before today for end date')
        ],
        'returns': 'String showing from/to dates',
        'example': 'dateRange = q.DateRangeForContributionTotals(30, 0)\nprint "Contributions for: " + dateRange'
    },
    
    # ===== GLOBAL PROPERTIES =====
    'model.PeopleId': {
        'category': 'Global Properties',
        'desc': 'Current user PeopleId',
        'params': [],
        'returns': 'Integer',
        'example': 'print "User ID: " + str(model.PeopleId)'
    },
    'model.UserName': {
        'category': 'Global Properties',
        'desc': 'Current username',
        'params': [],
        'returns': 'String',
        'example': 'print "Username: " + model.UserName'
    },
    'model.UserId': {
        'category': 'Global Properties',
        'desc': 'Current user ID',
        'params': [],
        'returns': 'Integer',
        'example': 'print "User ID: " + str(model.UserId)'
    },
    'model.FromMorningBatch': {
        'category': 'Global Properties',
        'desc': 'Running from morning batch',
        'params': [],
        'returns': 'Boolean',
        'example': 'if model.FromMorningBatch:\n    print "Batch mode"'
    },
    'model.Data': {
        'category': 'Global Properties',
        'desc': 'Data passed to script',
        'params': [],
        'returns': 'Data object',
        'example': 'if hasattr(model.Data, "orgId"):\n    orgId = model.Data.orgId'
    },
    
    # ===== UPLOAD =====
    'model.UploadExcelFromSqlToDropBox': {
        'category': 'Upload',
        'desc': 'Upload Excel from SQL to DropBox (requires DropBoxAccessToken in settings)',
        'params': [
            ('query_or_sql', 'string', 'Query for BlueToolbar filter (with @qtagid) OR SQL without filter'),
            ('sql', 'string', 'SQL to execute (omit if first param is SQL)'),
            ('targetpath', 'string', 'DropBox directory path'),
            ('filename', 'string', 'File name on DropBox')
        ],
        'returns': 'None',
        'example': '# With BlueToolbar filter:\nmodel.UploadExcelFromSqlToDropBox("IsMemberOf = 30", "SELECT * FROM People WHERE PeopleId IN (SELECT PeopleId FROM TagPerson WHERE Id = @qtagid)", "/Reports", "members.xlsx")\n# Without filter:\nmodel.UploadExcelFromSqlToDropBox("SELECT * FROM People", "/Reports", "all_people.xlsx")'
    },
    'model.UploadExcelFromSqlToFtp': {
        'category': 'Upload',
        'desc': 'Upload Excel from SQL to FTP server',
        'params': [
            ('sql', 'string', 'SQL to execute'),
            ('username', 'string', 'FTP username'),
            ('password', 'string', 'FTP password'),
            ('targetpath', 'string', 'FTP directory path'),
            ('filename', 'string', 'File name on FTP')
        ],
        'returns': 'None',
        'example': 'model.UploadExcelFromSqlToFtp("SELECT * FROM People", "ftpuser", "password", "/uploads", "people_export.xlsx")'
    },
    'model.BuildDisplay': {
        'category': 'Forms',
        'desc': 'Builds an HTML display table for the provided DynamicData object',
        'params': [
            ('name', 'string', 'The name of the display'),
            ('dd', 'DynamicData', 'The DynamicData object containing the data to display'),
            ('edit', 'string', 'Optional HTML for edit controls'),
            ('add', 'string', 'Optional HTML for add controls'),
            ('metadata', 'DynamicData', 'Optional metadata for controlling display formatting')
        ],
        'returns': 'HTML for a formatted display table',
        'example': 'html = model.BuildDisplay("UserList", dd, editBtn, addBtn, meta)'
    },
    'model.BuildDisplayRows': {
        'category': 'Forms',
        'desc': 'Builds the HTML rows for a display table based on the provided DynamicData object',
        'params': [
            ('dd', 'DynamicData', 'The DynamicData object containing the data to display'),
            ('metadata', 'DynamicData', 'Optional metadata for controlling display formatting')
        ],
        'returns': 'HTML for the rows of a display table',
        'example': 'rows = model.BuildDisplayRows(dd, metadata)'
    },
    'model.BuildForm': {
        'category': 'Forms',
        'desc': 'Builds an HTML form for the provided DynamicData object',
        'params': [
            ('name', 'string', 'The name of the form'),
            ('dd', 'DynamicData', 'The DynamicData object containing the form data'),
            ('buttons', 'string', 'Optional HTML for form buttons'),
            ('metadata', 'DynamicData', 'Optional metadata for controlling form formatting')
        ],
        'returns': 'HTML for a formatted form',
        'example': 'form = model.BuildForm("UserForm", dd, submitBtn, meta)'
    },
    'model.BuildFormRows': {
        'category': 'Forms',
        'desc': 'Builds the HTML rows for a form based on the provided DynamicData object',
        'params': [
            ('dd', 'DynamicData', 'The DynamicData object containing the form data'),
            ('metadata', 'DynamicData', 'Optional metadata for controlling form formatting')
        ],
        'returns': 'HTML for the rows of a form',
        'example': 'rows = model.BuildFormRows(dd, metadata)'
    },
    'model.HttpMethod': {
        'category': 'Forms',
        'desc': 'Readonly property that determines action based on whether it is the initial page load (GET) or the AJAX call from JavaScript (POST)',
        'params': [],
        'returns': 'Either "get" or "post"',
        'example': 'if model.HttpMethod == "post": # Handle form submission'
    },
    'model.Script': {
        'category': 'Forms',
        'desc': 'Set this to the JavaScript you want to run on your page. Loaded into your page by the /PyScriptForm/YourScript GET request URL',
        'params': [],
        'returns': 'JavaScript string',
        'example': 'model.Script = "function submitForm() { /* code */ }"'
    },
    'model.BirthdayList': {
        'category': 'People',
        'desc': 'Returns a list of upcoming birthdays. The source of the list and who is included is determined by the user settings for the Home Page Birthdays feature',
        'params': [],
        'returns': 'List of birthday objects containing PeopleId, Name, and Birthday fields',
        'example': 'birthdays = model.BirthdayList()\nfor person in birthdays:\n    print person.Name + " - " + person.Birthday'
    },
    'model.CreateTask': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Creates a new task for a person. Use model.CreateTaskNote for newer functionality.',
        'params': [
            ('ownerId', 'integer', 'The PeopleId of the task owner/assignee'),
            ('person', 'Person', 'The Person object the task is about'),
            ('description', 'string', 'Task description/notes')
        ],
        'returns': 'Task object',
        'example': '# DEPRECATED - Use model.CreateTaskNote instead\n# model.CreateTask(819918, person, "Please Contact about Small Group")'
    },
    'model.DatabaseName': {
        'category': 'System',
        'desc': 'Returns the name of the current database',
        'params': [],
        'returns': 'String containing the database name',
        'example': 'print model.DatabaseName'
    },
    'model.AddExtraValueBoolFamily': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a boolean extra value for all members of a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name'),
            ('truefalse', 'boolean', 'The boolean value to set')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueBoolFamily(12345, "HasPets", True)'
    },
    'model.AddExtraValueCodeFamily': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a code extra value for all members of a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name'),
            ('code', 'string', 'The code value to set')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueCodeFamily(12345, "MembershipType", "GOLD")'
    },
    'model.AddExtraValueDateFamily': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a date extra value for all members of a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name'),
            ('date', 'date/string', 'The date value to set')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueDateFamily(12345, "JoinDate", "2024-01-15")'
    },
    'model.AddExtraValueIntFamily': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates an integer extra value for all members of a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name'),
            ('number', 'integer', 'The integer value to set')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueIntFamily(12345, "FamilySize", 4)'
    },
    'model.AddExtraValueTextFamily': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a text extra value for all members of a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name'),
            ('text', 'string', 'The text value to set')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueTextFamily(12345, "Notes", "Special dietary requirements")'
    },
    'model.DeleteAllExtraValueLikeFamily': {
        'category': 'ExtraValues',
        'desc': 'Deletes all extra values matching a pattern for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('pattern', 'string', 'SQL LIKE pattern to match field names')
        ],
        'returns': 'None',
        'example': 'model.DeleteAllExtraValueLikeFamily(12345, "BP:%")'
    },
    'model.DeleteExtraValueFamily': {
        'category': 'ExtraValues',
        'desc': 'Deletes a specific extra value for all members of a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name to delete')
        ],
        'returns': 'None',
        'example': 'model.DeleteExtraValueFamily(12345, "TempValue")'
    },
    'model.ExtraValueBitFamily': {
        'category': 'ExtraValues',
        'desc': 'Gets a boolean extra value for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Boolean value',
        'example': 'hasPets = model.ExtraValueBitFamily(12345, "HasPets")'
    },
    'model.ExtraValueCodeFamily': {
        'category': 'ExtraValues',
        'desc': 'Gets a code extra value for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'String - The code value',
        'example': 'memberType = model.ExtraValueCodeFamily(12345, "MembershipType")'
    },
    'model.ExtraValueDateFamily': {
        'category': 'ExtraValues',
        'desc': 'Gets a date extra value for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Date value',
        'example': 'joinDate = model.ExtraValueDateFamily(12345, "JoinDate")'
    },
    'model.ExtraValueFamily': {
        'category': 'ExtraValues',
        'desc': 'Gets an extra value as a string for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'String - The extra value',
        'example': 'value = model.ExtraValueFamily(12345, "CustomField")'
    },
    'model.ExtraValueIntFamily': {
        'category': 'ExtraValues',
        'desc': 'Gets an integer extra value for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Integer value',
        'example': 'size = model.ExtraValueIntFamily(12345, "FamilySize")'
    },
    'model.ExtraValueTextFamily': {
        'category': 'ExtraValues',
        'desc': 'Gets a text extra value for a family',
        'params': [
            ('familyId', 'integer', 'The family ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'String - The text value',
        'example': 'notes = model.ExtraValueTextFamily(12345, "Notes")'
    },
    'model.FromJson': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Converts JSON string to a DynamicData object. Use model.DynamicDataFromJson instead.',
        'params': [
            ('json', 'string', 'JSON string to parse')
        ],
        'returns': 'DynamicData object',
        'example': '# DEPRECATED - Use model.DynamicDataFromJson\n# data = model.FromJson(\'{"name": "John", "age": 30}\')'
    },
    'model.HtmlContent': {
        'category': 'Content',
        'desc': 'Gets HTML content from Special Content',
        'params': [
            ('name', 'string', 'Name of the content item')
        ],
        'returns': 'String - The HTML content body',
        'example': 'html = model.HtmlContent("EmailTemplate")'
    },
    'model.PeopleIds': {
        'category': 'People',
        'desc': 'Returns a list of PeopleIds from a query',
        'params': [
            ('query', 'string', 'Search Builder query code')
        ],
        'returns': 'Enumerable list of integers',
        'example': 'pids = model.PeopleIds("IsMember=1[TRUE]")\nfor pid in pids:\n    print pid'
    },
    'model.RenderTemplate': {
        'category': 'Templates',
        'desc': 'Renders a Handlebars template with data',
        'params': [
            ('template', 'string', 'Handlebars template string'),
            ('data', 'object', 'Optional data object for template (defaults to Data)')
        ],
        'returns': 'String - The rendered template',
        'example': 'Data.user = model.UserName\nprint model.RenderTemplate("Hi {{user}}")'
    },
    'model.SqlContent': {
        'category': 'Content',
        'desc': 'Gets SQL content from Special Content SQL Scripts',
        'params': [
            ('name', 'string', 'Name of the SQL script')
        ],
        'returns': 'String - The SQL script content',
        'example': 'sql = model.SqlContent("CustomReport")'
    },
    'model.TextContent': {
        'category': 'Content',
        'desc': 'Gets text content from Special Content',
        'params': [
            ('name', 'string', 'Name of the content item')
        ],
        'returns': 'String - The text content',
        'example': 'text = model.TextContent("Instructions")'
    },
    'model.TitleContent': {
        'category': 'Content',
        'desc': 'Gets the title attribute of content in Special Content',
        'params': [
            ('name', 'string', 'Name of the content item')
        ],
        'returns': 'String - The title/subject',
        'example': 'subject = model.TitleContent("EmailTemplate")'
    },
    'model.WriteContentHtml': {
        'category': 'Content',
        'desc': 'Writes HTML content to Special Content. IMPORTANT: When submitting HTML via web forms, content must be entity-encoded to bypass ASP.NET request validation',
        'params': [
            ('name', 'string', 'Name for the content item'),
            ('text', 'string', 'HTML content to write (must be entity-encoded if from web form)'),
            ('keyword', 'string', 'Optional keyword/category (use empty string "" for default)')
        ],
        'returns': 'None',
        'example': '# Direct call:\nmodel.WriteContentHtml("NewTemplate", "<h1>Hello</h1>", "")\n\n# IronPython 2.7 HTML encoding (for reference):\nimport cgi\nhtml = "<h1>Hello</h1><p>Text & more</p>"\n# cgi.escape() encodes < > & for display\n# But WriteContentHtml needs raw HTML, not escaped\nmodel.WriteContentHtml("Template", html, "")\n\n# From AJAX/form (encode in JS, decode in Python):\n# JS: content.replace(/</g, "&lt;").replace(/>/g, "&gt;")\n# PY: content.replace("&lt;", "<").replace("&gt;", ">")\nmodel.WriteContentHtml("Template", decoded_content, "")'
    },
    'model.WriteContentPython': {
        'category': 'Content',
        'desc': 'Writes Python script to Special Content',
        'params': [
            ('name', 'string', 'Name for the script'),
            ('script', 'string', 'Python code to write'),
            ('keyword', 'string', 'Optional keyword/category')
        ],
        'returns': 'None',
        'example': 'model.WriteContentPython("CustomScript", "print \'Hello\'")'
    },
    'model.WriteContentSql': {
        'category': 'Content',
        'desc': 'Writes SQL script to Special Content',
        'params': [
            ('name', 'string', 'Name for the SQL script'),
            ('sql', 'string', 'SQL code to write'),
            ('keyword', 'string', 'Optional keyword/category')
        ],
        'returns': 'None',
        'example': 'model.WriteContentSql("Query", "SELECT * FROM People")'
    },
    'model.WriteContentText': {
        'category': 'Content',
        'desc': 'Writes text content to Special Content',
        'params': [
            ('name', 'string', 'Name for the content item'),
            ('text', 'string', 'Text content to write'),
            ('keyword', 'string', 'Optional keyword/category')
        ],
        'returns': 'None',
        'example': 'model.WriteContentText("Data", "Sample text content")'
    },
    'q.SqlFirstColumnRowKey': {
        'category': 'SQL',
        'desc': 'Returns DynamicData with first column values as property names',
        'params': [
            ('sql', 'string', 'SQL query to execute'),
            ('declarations', 'string', 'Optional parameter declarations')
        ],
        'returns': 'DynamicData object with properties from first column',
        'example': 'data = q.SqlFirstColumnRowKey("SELECT Status, COUNT(*) FROM People GROUP BY Status")'
    },
    'model.AddExtraValueBoolOrg': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a boolean extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name'),
            ('truefalse', 'boolean', 'The boolean value (True or False)')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueBoolOrg(12345, "IsActive", True)'
    },
    'model.AddExtraValueCodeOrg': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a code extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name'),
            ('code', 'string', 'The code value (text)')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueCodeOrg(12345, "Status", "ACTIVE")'
    },
    'model.AddExtraValueDateOrg': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a date extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name'),
            ('date', 'date/string', 'The date value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueDateOrg(12345, "StartDate", "2024-01-15")'
    },
    'model.AddExtraValueIntOrg': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates an integer extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name'),
            ('number', 'integer', 'The integer value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueIntOrg(12345, "MaxCapacity", 100)'
    },
    'model.AddExtraValueTextOrg': {
        'category': 'ExtraValues',
        'desc': 'Adds or updates a text extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name'),
            ('text', 'string', 'The text value')
        ],
        'returns': 'None',
        'example': 'model.AddExtraValueTextOrg(12345, "Description", "Youth group meeting room")'
    },
    'model.ExtraValueBitOrg': {
        'category': 'ExtraValues',
        'desc': 'Gets a boolean extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Boolean - The bool value for the organization',
        'example': 'isActive = model.ExtraValueBitOrg(12345, "IsActive")'
    },
    'model.ExtraValueCodeOrg': {
        'category': 'ExtraValues',
        'desc': 'Gets a code extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'String - The code value for the organization',
        'example': 'status = model.ExtraValueCodeOrg(12345, "Status")'
    },
    'model.ExtraValueDateOrg': {
        'category': 'ExtraValues',
        'desc': 'Gets a date extra value for an organization. Returns 1/1/01 if missing',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Date - The date value (1/1/01 if missing)',
        'example': 'startDate = model.ExtraValueDateOrg(12345, "StartDate")'
    },
    'model.ExtraValueDateOrgNull': {
        'category': 'ExtraValues',
        'desc': 'Gets a date extra value for an organization. Returns None if missing',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Date or None - The date value or None if missing',
        'example': 'startDate = model.ExtraValueDateOrgNull(12345, "StartDate")\nif startDate:\n    print startDate'
    },
    'model.ExtraValueIntOrg': {
        'category': 'ExtraValues',
        'desc': 'Gets an integer extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'Integer - The int value for the organization',
        'example': 'capacity = model.ExtraValueIntOrg(12345, "MaxCapacity")'
    },
    'model.ExtraValueOrg': {
        'category': 'ExtraValues',
        'desc': 'Gets an extra value as a string for an organization without specifying type',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'String - The extra value as a string',
        'example': 'value = model.ExtraValueOrg(12345, "CustomField")'
    },
    'model.ExtraValueTextOrg': {
        'category': 'ExtraValues',
        'desc': 'Gets a text extra value for an organization',
        'params': [
            ('orgid', 'integer', 'The organization ID'),
            ('name', 'string', 'The extra value field name')
        ],
        'returns': 'String - The text value for the organization',
        'example': 'description = model.ExtraValueTextOrg(12345, "Description")'
    },
    
    # ===== MISCELLANEOUS UTILITIES =====
    'model.WriteFile': {
        'category': 'Misc',
        'desc': '⚠️ DEBUG ONLY - Writes text content to a file. Will throw exception in production.',
        'params': [
            ('path', 'string', 'The file path to write to'),
            ('text', 'string', 'The text content to write')
        ],
        'returns': 'None',
        'example': '# DEBUG MODE ONLY\nmodel.WriteFile("/tmp/debug.txt", "Debug output text")'
    },
    'model.ReadFile': {
        'category': 'Misc',
        'desc': '⚠️ DEBUG ONLY - Reads content from a file. Will throw exception in production.',
        'params': [
            ('name', 'string', 'The file path to read')
        ],
        'returns': 'String - The content of the file',
        'example': '# DEBUG MODE ONLY\ncontent = model.ReadFile("/tmp/debug.txt")\nprint content'
    },
    'model.ExecuteSql': {
        'category': 'Misc',
        'desc': '⚠️ DEBUG ONLY - Executes SQL statement directly. Will throw exception in production.',
        'params': [
            ('sql', 'string', 'The SQL statement to execute')
        ],
        'returns': 'None',
        'example': '# DEBUG MODE ONLY - BE VERY CAREFUL!\nmodel.ExecuteSql("UPDATE People SET LastName = \'Smith\' WHERE PeopleId = 828")'
    },
    'model.UpdateStatusFlag': {
        'category': 'Misc',
        'desc': 'Updates a status flag using results from specified query',
        'params': [
            ('flagid', 'string', 'The ID of the status flag to update'),
            ('encodedguid', 'string', 'The encoded GUID of the query')
        ],
        'returns': 'None',
        'example': 'model.UpdateStatusFlag("active-members", encoded_query_guid)'
    },
    'model.StatusFlagList': {
        'category': 'Misc',
        'desc': 'Returns a list of all status flags defined in the system',
        'params': [],
        'returns': 'List<StatusFlag> - List of all status flags with their properties',
        'example': '# Get all status flags\nflags = model.StatusFlagList()\nfor flag in flags:\n    print "Flag: " + flag.Name + " (ID: " + flag.Id + ")"\n    print "  Description: " + flag.Description\n    print "  Count: " + str(flag.Count)\n\n# Find a specific flag\nflags = model.StatusFlagList()\ntarget_flag = next((f for f in flags if f.Id == "F01"), None)\nif target_flag:\n    print "Found flag: " + target_flag.Name'
    },
    'model.TagLastQuery': {
        'category': 'Misc',
        'desc': 'Tags the results of the last executed query',
        'params': [
            ('defaultcode', 'string', 'Default query code if running from batch')
        ],
        'returns': 'Integer - The ID of the created tag',
        'example': 'tagId = model.TagLastQuery("IsMember=1[True]")\nprint "Created tag with ID:", tagId'
    },
    'model.StatusFlagDictionary': {
        'category': 'Misc',
        'desc': 'Returns dictionary of status flags, optionally filtered',
        'params': [
            ('flags', 'string', 'Optional comma-separated list of flag names to filter')
        ],
        'returns': 'Dictionary<string, StatusFlag> - Status flag dictionary',
        'example': 'flags = model.StatusFlagDictionary("active,inactive,pending")\nfor flag_name, flag_obj in flags.items():\n    print flag_name, flag_obj.Count'
    },
    'model.ReplaceQueryFromCode': {
        'category': 'Misc',
        'desc': 'Replaces an existing query with new query code',
        'params': [
            ('encodedguid', 'string', 'The encoded GUID of query to replace'),
            ('code', 'string', 'The new query code')
        ],
        'returns': 'None',
        'example': 'model.ReplaceQueryFromCode(existing_guid, "IsMember=1[True] AND Age>50")'
    },
    'model.ReplaceCodeStr': {
        'category': 'Misc',
        'desc': 'Replaces codes in text using provided mapping',
        'params': [
            ('text', 'string', 'The text to perform replacements on'),
            ('codes', 'string', 'Replacements in format "code1=val1,code2=val2"')
        ],
        'returns': 'String - Text with replacements applied',
        'example': 'result = model.ReplaceCodeStr("Hello {name}, your {item} is ready", "name=John,item=order")\nprint result  # "Hello John, your order is ready"'
    },
    'model.PythonContent': {
        'category': 'Misc',
        'desc': 'Retrieves Python script content from Special Content',
        'params': [
            ('name', 'string', 'The name of the Python script')
        ],
        'returns': 'String - The Python script content',
        'example': 'script = model.PythonContent("CustomReport")\nprint script  # Shows the Python code'
    },
    'model.Draft': {
        'category': 'Misc',
        'desc': 'Retrieves body content of a saved draft from Special Content',
        'params': [
            ('name', 'string', 'The name of the saved draft')
        ],
        'returns': 'String - The body content of the draft',
        'example': 'draft_body = model.Draft("WelcomeEmailDraft")\nprint draft_body'
    },
    'model.DraftTitle': {
        'category': 'Misc',
        'desc': 'Retrieves title of a saved draft from Special Content',
        'params': [
            ('name', 'string', 'The name of the saved draft')
        ],
        'returns': 'String - The title of the draft',
        'example': 'draft_title = model.DraftTitle("WelcomeEmailDraft")\nprint draft_title'
    },
    'model.Dictionary': {
        'category': 'Misc',
        'desc': 'Retrieves value from script dictionary using specified key',
        'params': [
            ('s', 'string', 'The key to look up')
        ],
        'returns': 'String - The value or empty string if not found',
        'example': 'value = model.Dictionary("setting_name")\nif value:\n    print "Setting value:", value'
    },
    'model.DictionaryAdd': {
        'category': 'Misc',
        'desc': 'Adds key-value pair to script dictionary',
        'params': [
            ('key', 'string', 'The key to add'),
            ('value', 'object', 'The value (string or object)')
        ],
        'returns': 'None',
        'example': 'model.DictionaryAdd("processed_count", 100)\nmodel.DictionaryAdd("last_run", datetime.datetime.now())\n# Retrieve later:\ncount = model.Dictionary("processed_count")'
    },
    'model.DebugPrint': {
        'category': 'Misc',
        'desc': 'Outputs string to debug console for debugging',
        'params': [
            ('s', 'string', 'The string to output')
        ],
        'returns': 'None',
        'example': 'model.DebugPrint("Processing person ID: " + str(person.PeopleId))\nmodel.DebugPrint("Current status: " + status)'
    },
    'model.LogToContent': {
        'category': 'Misc',
        'desc': 'Logs text to a Special Content file (text or Python). Creates file if it doesn\'t exist, appends if it does. Useful for persistent logging across script runs.',
        'params': [
            ('file', 'string', 'The name of the Special Content file to log to (without path)'),
            ('text', 'string', 'The text to append to the file')
        ],
        'returns': 'None',
        'example': '# Log to a text file in Special Content\nmodel.LogToContent("integration_log.txt", "API call completed at " + str(datetime.datetime.now()))\n\n# Log errors to an error log\nmodel.LogToContent("error_log.txt", "Error: " + str(error_message) + "\\n")\n\n# Create an audit trail\nmodel.LogToContent("audit_trail.py", "# User " + str(user_id) + " performed action at " + str(datetime.datetime.now()))'
    },
    'model.DeleteFile': {
        'category': 'Misc',
        'desc': '⚠️ DEBUG ONLY - Deletes a file at the specified path. Will throw exception in production.',
        'params': [
            ('path', 'string', 'The file path to delete')
        ],
        'returns': 'None',
        'example': '# DEBUG MODE ONLY\nmodel.DeleteFile("/tmp/debug.txt")\n\n# Delete temporary files after processing\nmodel.DeleteFile("/tmp/export_" + str(batch_id) + ".csv")'
    },
    
    # ===== DOCUSIGN =====
    'model.docusign': {
        'category': 'DocuSign',
        'desc': 'Access DocuSign electronic signature services through the ApiDocuSign object',
        'params': [],
        'returns': 'ApiDocuSign object with DocuSign API functionality',
        'example': '''# Create a DocuSign envelope for electronic signatures
envelope = model.docusign.EnvelopeDefinition()
envelope.EmailSubject = "Please sign this document"
envelope.Status = "sent"  # Set to "created" to save as draft

# Create a document
document = model.docusign.Document()
document.DocumentBase64 = base64_encoded_pdf  # Your PDF in base64
document.Name = "Agreement.pdf"
document.DocumentId = "1"
envelope.Documents = [document]

# Create a signer recipient
signer = model.docusign.Signer()
signer.Email = person.EmailAddress
signer.Name = person.Name
signer.RecipientId = "1"
signer.RoutingOrder = "1"

# Create signature tab (where to sign)
signHere = model.docusign.SignHere()
signHere.DocumentId = "1"
signHere.PageNumber = "1"
signHere.XPosition = "100"
signHere.YPosition = "100"

# Add tabs to signer
tabs = model.docusign.Tabs()
tabs.SignHereTabs = [signHere]
signer.Tabs = tabs

# Add signer to envelope
recipients = model.docusign.Recipients()
recipients.Signers = [signer]
envelope.Recipients = recipients

# Send the envelope
# result = api.CreateEnvelope(envelope)'''
    },
    'model.docusign.EnvelopeDefinition': {
        'category': 'DocuSign',
        'desc': 'Create a DocuSign envelope definition for documents to be signed',
        'params': [],
        'returns': 'EnvelopeDefinition object',
        'example': '''envelope = model.docusign.EnvelopeDefinition()
envelope.EmailSubject = "Contract for Review and Signature"
envelope.EmailBlurb = "Please review and sign the attached contract"
envelope.Status = "sent"  # or "created" for draft'''
    },
    'model.docusign.Document': {
        'category': 'DocuSign',
        'desc': 'Create a document object for DocuSign envelope',
        'params': [],
        'returns': 'Document object',
        'example': '''document = model.docusign.Document()
document.DocumentBase64 = base64_pdf_content
document.Name = "Contract.pdf"
document.FileExtension = "pdf"
document.DocumentId = "1"'''
    },
    'model.docusign.Signer': {
        'category': 'DocuSign',
        'desc': 'Create a signer recipient for DocuSign envelope',
        'params': [],
        'returns': 'Signer object',
        'example': '''signer = model.docusign.Signer()
signer.Email = "john.doe@example.com"
signer.Name = "John Doe"
signer.RecipientId = "1"
signer.RoutingOrder = "1"
signer.RoleName = "Signer"  # Optional role'''
    },
    'model.docusign.SignHere': {
        'category': 'DocuSign',
        'desc': 'Create a signature tab indicating where recipient should sign',
        'params': [],
        'returns': 'SignHere tab object',
        'example': '''signHere = model.docusign.SignHere()
signHere.DocumentId = "1"
signHere.PageNumber = "1"
signHere.XPosition = "100"  # Distance from left in pixels
signHere.YPosition = "150"  # Distance from top in pixels
signHere.TabLabel = "Sign Here"'''
    },
    'model.docusign.DateSigned': {
        'category': 'DocuSign',
        'desc': 'Create a date signed tab that automatically fills with signing date',
        'params': [],
        'returns': 'DateSigned tab object',
        'example': '''dateSigned = model.docusign.DateSigned()
dateSigned.DocumentId = "1"
dateSigned.PageNumber = "1"
dateSigned.XPosition = "100"
dateSigned.YPosition = "200"'''
    },
    'model.docusign.InitialHere': {
        'category': 'DocuSign',
        'desc': 'Create an initial tab for recipient initials',
        'params': [],
        'returns': 'InitialHere tab object',
        'example': '''initialHere = model.docusign.InitialHere()
initialHere.DocumentId = "1"
initialHere.PageNumber = "1"
initialHere.XPosition = "50"
initialHere.YPosition = "100"'''
    },
    'model.docusign.Text': {
        'category': 'DocuSign',
        'desc': 'Create a text field tab for recipient to enter text',
        'params': [],
        'returns': 'Text tab object',
        'example': '''textField = model.docusign.Text()
textField.DocumentId = "1"
textField.PageNumber = "1"
textField.XPosition = "100"
textField.YPosition = "250"
textField.Width = "200"
textField.Height = "20"
textField.TabLabel = "Phone Number"
textField.Required = "true"'''
    },
    'model.docusign.Checkbox': {
        'category': 'DocuSign',
        'desc': 'Create a checkbox tab for recipient to check',
        'params': [],
        'returns': 'Checkbox tab object',
        'example': '''checkbox = model.docusign.Checkbox()
checkbox.DocumentId = "1"
checkbox.PageNumber = "1"
checkbox.XPosition = "100"
checkbox.YPosition = "300"
checkbox.TabLabel = "I Agree"
checkbox.Selected = "false"'''
    },
    'model.docusign.Tabs': {
        'category': 'DocuSign',
        'desc': 'Container for all tabs (signature fields, text fields, etc.)',
        'params': [],
        'returns': 'Tabs container object',
        'example': '''tabs = model.docusign.Tabs()
tabs.SignHereTabs = [signHere1, signHere2]
tabs.DateSignedTabs = [dateSigned]
tabs.TextTabs = [textField1, textField2]
tabs.CheckboxTabs = [checkbox1]
tabs.InitialHereTabs = [initialHere]'''
    },
    'model.docusign.Recipients': {
        'category': 'DocuSign',
        'desc': 'Container for all envelope recipients',
        'params': [],
        'returns': 'Recipients container object',
        'example': '''recipients = model.docusign.Recipients()
recipients.Signers = [signer1, signer2]
recipients.CarbonCopies = [cc1]  # Optional CC recipients
envelope.Recipients = recipients'''
    },
    'model.docusign.CarbonCopy': {
        'category': 'DocuSign',
        'desc': 'Create a carbon copy recipient who receives a copy but does not sign',
        'params': [],
        'returns': 'CarbonCopy recipient object',
        'example': '''cc = model.docusign.CarbonCopy()
cc.Email = "manager@example.com"
cc.Name = "Manager Name"
cc.RecipientId = "2"
cc.RoutingOrder = "2"  # Receives after signers'''
    },
    
    # ===== DEPRECATED FUNCTIONS =====
    # ⚠️ WARNING: These functions are deprecated and should NOT be used in new code
    
    'model.Email2': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Sends email to people in a query using GUID identifier. Use newer email methods instead.',
        'params': [
            ('qid', 'string', 'GUID query identifier'),
            ('queuedBy', 'integer', 'Person ID who queued email'),
            ('fromAddr', 'string', 'From email address'),
            ('fromName', 'string', 'From name'),
            ('subject', 'string', 'Email subject'),
            ('body', 'string', 'Email body HTML')
        ],
        'returns': 'None',
        'example': '# DEPRECATED - Do not use in new code\n# model.Email2(qid, 1, "info@church.org", "Church", "Subject", "<p>Body</p>")'
    },
    'model.EmailContent2': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Sends email using content with GUID query identifier. Use newer email content methods.',
        'params': [
            ('qid', 'string', 'GUID query identifier'),
            ('queuedBy', 'integer', 'Person ID who queued email'),
            ('fromAddr', 'string', 'From email address'),
            ('fromName', 'string', 'From name'),
            ('contentName', 'string', 'Special content name'),
            ('subject', 'string', 'Optional subject (overrides content)')
        ],
        'returns': 'None',
        'example': '# DEPRECATED - Do not use\n# model.EmailContent2(qid, 1, "info@church.org", "Church", "Template")'
    },
    'model.OrgMembersQuery': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Creates query for organization members. Use q.QuerySql or Search Builder instead.',
        'params': [
            ('progid', 'integer', 'Program ID'),
            ('divid', 'integer', 'Division ID'),
            ('orgid', 'integer', 'Organization ID'),
            ('memberTypes', 'string', 'Member type codes')
        ],
        'returns': 'Query string',
        'example': '# DEPRECATED - Use Search Builder or q.QuerySql instead\n# query = model.OrgMembersQuery(1, 2, 3, "220")'
    },
    'model.DisableRecurringGiving': {
        'category': 'Deprecated',
        'desc': '⚠️ DEPRECATED - Disables recurring giving for a person and fund. This function is no longer functional.',
        'params': [
            ('peopleId', 'integer', 'Person ID'),
            ('fundId', 'integer', 'Fund ID')
        ],
        'returns': 'None (no longer functional)',
        'example': '# DEPRECATED - No longer functional\n# model.DisableRecurringGiving(828, 1)'
    },
    
    # ===== FORMS =====
    'model.Header': {
        'category': 'Forms',
        'desc': 'Property that sets the title/header of your page. Displayed at the top of the page when rendered.',
        'params': [],
        'returns': 'String - The current header text when accessed; None when set',
        'example': '# Set the page header\nmodel.Header = "My Custom Report"\n\n# Or with emojis\nmodel.Header = "📊 Sales Dashboard"\n\n# Access current header\ncurrent_title = model.Header\nprint "Current page title:", current_title'
    }
}

# Continue with the HTML interface...
print """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>TouchPoint API Explorer</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/theme/monokai.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/dialog/dialog.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/show-hint.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/lint/lint.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        html, body { 
            margin: 0;
            padding-bottom: 50px;
        }
        
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            display: flex;
            flex-direction: column;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 10px 20px;
            flex-shrink: 0;
            height: 50px;
        }
        
        .header h1 { font-size: 16px; margin: 0; }
        .header p { font-size: 11px; opacity: 0.9; margin: 2px 0 0 0; }
        
        .main-container {
            display: flex;
            flex: 1;
            min-height: 0;
            width: 100%;
            height: calc(100vh - 50px);
            overflow: hidden;
        }
        
        .sidebar {
            width: 280px;
            min-width: 50px;
            background: #f8f9fa;
            border-right: 1px solid #dee2e6;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            position: relative;
            transition: width 0.3s ease;
            flex-shrink: 0;
        }
        
        
        .sidebar-resize-handle {
            position: absolute;
            right: -3px;
            top: 0;
            bottom: 0;
            width: 6px;
            background: transparent;
            cursor: col-resize;
            z-index: 10;
        }
        
        .sidebar-resize-handle:hover {
            background: #007bff;
            opacity: 0.5;
        }
        
        .search-box {
            padding: 10px;
            border-bottom: 1px solid #dee2e6;
            flex-shrink: 0;
        }
        
        .search-box input {
            width: 100%;
            padding: 5px;
            border: 1px solid #ced4da;
            border-radius: 3px;
            font-size: 12px;
        }
        
        .method-list {
            flex: 1;
            overflow-y: auto;
            overflow-x: hidden;
            padding: 10px;
        }
        
        .category-section {
            margin-bottom: 15px;
        }
        
        .category-header {
            font-size: 11px;
            text-transform: uppercase;
            color: #6c757d;
            margin-bottom: 5px;
            font-weight: 600;
            padding: 5px;
            background: #e9ecef;
            border-radius: 3px;
            cursor: pointer;
            user-select: none;
            position: relative;
            padding-left: 20px;
        }
        
        .category-header::before {
            content: '▼';
            position: absolute;
            left: 5px;
            transition: transform 0.2s;
        }
        
        .category-header.collapsed::before {
            transform: rotate(-90deg);
        }
        
        .category-content {
            display: block;
            overflow: hidden;
            transition: max-height 0.3s ease-out;
            max-height: 1000px;
        }
        
        .category-content.collapsed {
            max-height: 0;
        }
        
        .category-content.collapsed {
            display: none;
        }
        
        .method-item {
            padding: 4px 8px;
            cursor: pointer;
            font-size: 11px;
            font-family: monospace;
            border-radius: 3px;
            margin-bottom: 2px;
            border-left: 2px solid transparent;
        }
        
        .method-item:hover {
            background: #e9ecef;
            border-left-color: #667eea;
        }
        
        .method-item.active {
            background: #667eea;
            color: white;
            border-left-color: #764ba2;
        }
        
        .method-item.documented {
            font-weight: bold;
        }
        
        .method-item.undocumented {
            opacity: 0.8;
            font-style: italic;
        }
        
        .doc-panel {
            width: 380px;
            min-width: 50px;
            padding: 15px;
            background: white;
            border-right: 1px solid #dee2e6;
            overflow-y: auto;
            flex-shrink: 0;
            position: relative;
            transition: width 0.3s ease;
        }
        
        
        .doc-resize-handle {
            position: absolute;
            right: -3px;
            top: 0;
            bottom: 0;
            width: 6px;
            background: transparent;
            cursor: col-resize;
            z-index: 10;
        }
        
        .doc-resize-handle:hover {
            background: #007bff;
            opacity: 0.5;
        }
        
        .doc-title {
            font-size: 14px;
            font-weight: bold;
            margin-bottom: 5px;
            font-family: monospace;
            color: #2c3e50;
        }
        
        .doc-category {
            font-size: 10px;
            color: #667eea;
            text-transform: uppercase;
            margin-bottom: 5px;
        }
        
        .doc-desc {
            font-size: 12px;
            color: #6c757d;
            margin-bottom: 15px;
        }
        
        .param {
            background: #f8f9fa;
            padding: 6px;
            margin-bottom: 4px;
            border-left: 2px solid #667eea;
            font-size: 11px;
        }
        
        .param-name {
            font-weight: bold;
            font-family: monospace;
            color: #2c3e50;
        }
        
        .param-type {
            color: #667eea;
            font-size: 10px;
            margin-left: 5px;
        }
        
        .example-box {
            background: #2d2d30;
            color: #d4d4d4;
            padding: 10px;
            border-radius: 3px;
            font-family: monospace;
            font-size: 11px;
            white-space: pre;
            overflow-x: auto;
        }
        
        .editor-panel {
            flex: 1 1 auto;
            display: flex;
            flex-direction: column;
            min-width: 300px;
            width: 100%;
            min-height: 0;  /* Critical for flexbox children */
            overflow: hidden;
        }
        
        .editor-toolbar {
            padding: 10px;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            flex-shrink: 0;
        }
        
        .btn {
            padding: 5px 12px;
            border: none;
            border-radius: 3px;
            font-size: 12px;
            cursor: pointer;
        }
        
        .btn-primary {
            background: #667eea;
            color: white;
        }
        
        .btn-secondary {
            background: #6c757d;
            color: white;
        }
        
        .editor-split {
            flex: 1 1 auto;
            display: flex;
            flex-direction: column;
            min-height: 0;  /* Critical for proper flex sizing */
            overflow: hidden;
        }
        
        .output-panel {
            flex: 0 0 200px;
            min-height: 100px;
            max-height: 400px;
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 10px;
            font-family: monospace;
            font-size: 11px;
            overflow-y: auto;
            overflow-x: auto;  /* Allow horizontal scroll */
            white-space: pre;   /* Don't wrap output text */
            border-bottom: 2px solid #495057;
            resize: vertical;
            position: relative;
        }
        
        .resize-handle {
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: #495057;
            cursor: ns-resize;
        }
        
        .resize-handle:hover {
            background: #667eea;
        }
        
        .code-container {
            flex: 1 1 auto;
            min-height: 0;  /* Critical - allows container to shrink */
            position: relative;
            overflow: hidden;  /* Let CodeMirror handle scrolling */
        }
        
        .CodeMirror {
            height: 100% !important;
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
        }
        
        .CodeMirror-scroll {
            height: 100% !important;
            overflow: auto !important;
        }
        
        .CodeMirror-sizer {
            padding-bottom: 100px !important;  /* Extra space at bottom */
        }
        
        .CodeMirror-vscrollbar, .CodeMirror-hscrollbar {
            display: block !important;
        }
        
        .stats {
            padding: 10px;
            background: #f8f9fa;
            font-size: 11px;
            color: #6c757d;
            border-top: 1px solid #dee2e6;
            flex-shrink: 0;
        }
        
        /* Modal Styles */
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }
        
        .modal-content {
            background-color: #fefefe;
            margin: 10% auto;
            padding: 0;
            border-radius: 8px;
            width: 600px;
            max-width: 90%;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        .modal-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 20px;
            border-radius: 8px 8px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .modal-title {
            font-size: 18px;
            font-weight: 600;
        }
        
        .close {
            color: white;
            font-size: 28px;
            font-weight: bold;
            border: none;
            background: none;
            cursor: pointer;
            padding: 0;
            width: 30px;
            height: 30px;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        .close:hover,
        .close:focus {
            opacity: 0.8;
        }
        
        .shortcut-list {
            display: grid;
            gap: 10px;
            padding: 20px;
        }
        
        .shortcut-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            border-bottom: 1px solid #f0f0f0;
            font-size: 14px;
        }
        
        .shortcut-key {
            font-family: monospace;
            background: #667eea;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 12px;
            margin: 0 2px;
        }
        
        /* Custom linting styles for dark theme */
        .CodeMirror-lint-tooltip {
            background: #2d2d30;
            color: #f0f0f0;
            border: 1px solid #454545;
            border-radius: 4px;
            font-size: 12px;
            padding: 4px 6px;
            z-index: 10000;
        }
        
        .CodeMirror-lint-mark-error {
            border-bottom: 2px solid #ff5555;
        }
        
        .CodeMirror-lint-mark-warning {
            border-bottom: 2px solid #f1fa8c;
        }
        
        .CodeMirror-lint-mark-info {
            border-bottom: 1px dotted #50fa7b;
        }
        
        .CodeMirror-lint-marker-error {
            color: #ff5555;
        }
        
        .CodeMirror-lint-marker-warning {
            color: #f1fa8c;
        }
        
        .CodeMirror-lint-marker-info {
            color: #50fa7b;
        }
        
        .CodeMirror-gutters {
            background-color: #282a36;
            border-right: 1px solid #44475a;
        }
        
        /* Enhanced autocomplete styles */
        .CodeMirror-hints {
            background: #2d2d30;
            border: 1px solid #454545;
            border-radius: 4px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 13px;
            max-height: 300px;
            overflow-y: auto;
        }
        
        .CodeMirror-hint {
            color: #f0f0f0;
            padding: 4px 8px;
            line-height: 1.4;
            cursor: pointer;
        }
        
        .CodeMirror-hint-active {
            background: #094771;
            color: #ffffff;
        }
        
        /* Custom hint styles */
        .hint-method {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .hint-name {
            font-weight: bold;
            color: #4EC9B0;
        }
        
        .hint-category {
            font-size: 10px;
            background: #667eea;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            margin-left: 10px;
        }
        
        .hint-desc {
            font-size: 11px;
            color: #a0a0a0;
            margin-top: 2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        .CodeMirror-hint.touchpoint-hint {
            padding: 6px 8px;
        }
        
        .CodeMirror-hint.object-hint .hint-name {
            color: #569CD6;
        }
        
        .CodeMirror-hint.keyword-hint .hint-name {
            color: #C586C0;
        }
        
        .CodeMirror-hint.method-hint .hint-name {
            color: #DCDCAA;
        }
        
        /* Tooltip for parameter info */
        .parameter-hint {
            position: absolute;
            background: #2d2d30;
            border: 1px solid #454545;
            padding: 8px;
            border-radius: 4px;
            font-size: 12px;
            color: #f0f0f0;
            z-index: 10001;
            max-width: 400px;
        }
        
        .parameter-hint .param-name {
            color: #9CDCFE;
            font-weight: bold;
        }
        
        .parameter-hint .param-type {
            color: #4EC9B0;
        }
    </style>
</head>
<body>
    <!-- Keyboard Shortcuts Modal -->
    <div id="shortcutsModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <div class="modal-title">⌨️ Keyboard Shortcuts</div>
                <button class="close" onclick="closeShortcutsModal()">&times;</button>
            </div>
            <div class="shortcut-list">
                <div class="shortcut-item">
                    <span>Run all code</span>
                    <span><span class="shortcut-key">Ctrl+Enter</span> or <span class="shortcut-key">Cmd+Enter</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Run selected code (or all if none selected)</span>
                    <span><span class="shortcut-key">Shift+Enter</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Clear editor</span>
                    <span><span class="shortcut-key">Ctrl+L</span> or <span class="shortcut-key">Cmd+L</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Indent / Insert tab (4 spaces)</span>
                    <span><span class="shortcut-key">Tab</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Unindent</span>
                    <span><span class="shortcut-key">Shift+Tab</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Toggle comment</span>
                    <span><span class="shortcut-key">Ctrl+/</span> or <span class="shortcut-key">Cmd+/</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Autocomplete</span>
                    <span><span class="shortcut-key">Ctrl+Space</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Find</span>
                    <span><span class="shortcut-key">Ctrl+F</span> or <span class="shortcut-key">Cmd+F</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Replace</span>
                    <span><span class="shortcut-key">Ctrl+H</span> or <span class="shortcut-key">Cmd+H</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Undo</span>
                    <span><span class="shortcut-key">Ctrl+Z</span> or <span class="shortcut-key">Cmd+Z</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Redo</span>
                    <span><span class="shortcut-key">Ctrl+Y</span> or <span class="shortcut-key">Cmd+Shift+Z</span></span>
                </div>
            </div>
        </div>
    </div>
    
    <div class="header">
        <h1>🚀 TouchPoint API Explorer</h1>
        <p>Swaby's reference with """ + str(len(method_docs)) + """ documented methods</p>
    </div>
    
    <div class="main-container">
        <!-- Sidebar -->
        <div class="sidebar" id="sidebar">
            <div class="sidebar-resize-handle" id="sidebarResizeHandle"></div>
            <div class="search-box">
                <input type="text" placeholder="Search methods..." onkeyup="filterMethods(this.value)">
            </div>
            
            <div class="method-list">
"""

# Group methods by category
categories = {}
for method_name, doc in method_docs.items():
    category = doc.get('category', 'Other')
    if category not in categories:
        categories[category] = []
    categories[category].append(method_name)

# Display methods by category
for category in sorted(categories.keys()):
    print '                <div class="category-section">'
    print '                    <div class="category-header" onclick="toggleCategory(this)">' + category + ' (' + str(len(categories[category])) + ')</div>'
    print '                    <div class="category-content">'
    for method in sorted(categories[category]):
        print '                        <div class="method-item documented" onclick="showMethod(\'' + method + '\')">' + method + '</div>'
    print '                    </div>'
    print '                </div>'

# Add undocumented methods section
# Count undocumented methods first
undoc_model_count = 0
undoc_q_count = 0

try:
    if 'model' in globals():
        documented_model = [m.split('.')[1] for m in method_docs.keys() if m.startswith('model.')]
        all_model = [m for m in dir(model) if not m.startswith('_')]
        undoc_model = [m for m in all_model if m not in documented_model]
        undoc_model_count = len(undoc_model)
except:
    undoc_model = []

try:
    if 'q' in globals():
        documented_q = [m.split('.')[1] for m in method_docs.keys() if m.startswith('q.')]
        all_q = [m for m in dir(q) if not m.startswith('_')]
        undoc_q = [m for m in all_q if m not in documented_q]
        undoc_q_count = len(undoc_q)
except:
    undoc_q = []

total_undoc = undoc_model_count + undoc_q_count

print '                <div class="category-section">'
print '                    <div class="category-header" onclick="toggleCategory(this)">🔍 Undocumented Methods (' + str(total_undoc) + ')</div>'
print '                    <div class="category-content">'  # Removed 'collapsed' class

# List undocumented model methods
if undoc_model_count > 0:
    print '                        <div style="padding: 5px 10px; font-size: 11px; color: #6c757d; font-weight: bold;">Model Methods (' + str(undoc_model_count) + ')</div>'
    for method in sorted(undoc_model):
        print '                        <div class="method-item undocumented" onclick="showMethod(\'model.' + method + '\')">model.' + method + '</div>'

# List undocumented query methods  
if undoc_q_count > 0:
    print '                        <div style="padding: 5px 10px; font-size: 11px; color: #6c757d; font-weight: bold;">Query Methods (' + str(undoc_q_count) + ')</div>'
    for method in sorted(undoc_q):
        print '                        <div class="method-item undocumented" onclick="showMethod(\'q.' + method + '\')">q.' + method + '</div>'

if total_undoc == 0:
    print '                        <div style="padding: 10px; font-size: 12px; color: #6c757d; font-style: italic;">No undocumented methods found</div>'

print '                    </div>'
print '                </div>'

print """            </div>
            
            <div class="stats">
                Documented: """ + str(len(method_docs)) + """ methods<br>
                Categories: """ + str(len(categories)) + """
            </div>
        </div>
        
        <!-- Documentation Panel -->
        <div class="doc-panel" id="docPanel">
            <div class="doc-resize-handle" id="docResizeHandle"></div>
            <div class="doc-category" id="methodCategory">WELCOME</div>
            <div class="doc-title" id="methodName">TouchPoint API Explorer</div>
            <div class="doc-desc" id="methodDesc">
                <strong>About This Tool:</strong><br>
                The TouchPoint API Explorer is an interactive development environment for testing and exploring TouchPoint's Python API. It provides real-time code execution with access to the complete TouchPoint model and query objects, along with comprehensive documentation for all available methods.<br><br>
                
                <div style="background: #fff3cd; border: 1px solid #ffc107; padding: 10px; border-radius: 4px; margin: 10px 0;">
                    <strong style="color: #856404;">⚠️ WARNING:</strong><br>
                    <span style="color: #856404;">This tool executes code directly against your TouchPoint database. Improper use can cause irreversible data changes, deletions, or system issues. Only use this tool if you understand the implications of your code. Always test queries and operations carefully before executing them in production.</span>
                </div>
                
                <strong>Getting Started:</strong><br>
                • Browse methods in the sidebar by category<br>
                • Click any method to view its documentation and example<br>
                • Write your code in the editor and click "Run Code" or press Ctrl/Cmd+Enter<br>
                • Use "Quick Test" to verify your environment setup<br>
                • Use "Discover" to explore available objects and methods
            </div>
            
            <div id="paramsSection" style="display: none;">
                <h4 style="font-size: 12px; margin: 10px 0;">Parameters</h4>
                <div id="paramsList"></div>
            </div>
            
            <div id="returnsSection" style="display: none;">
                <h4 style="font-size: 12px; margin: 10px 0;">Returns</h4>
                <div class="doc-desc" id="returnsInfo"></div>
            </div>
            
            <div id="exampleSection" style="display: none;">
                <h4 style="font-size: 12px; margin: 10px 0;">Example</h4>
                <div class="example-box" id="exampleCode"></div>
            </div>
        </div>
        
        <!-- Editor Panel -->
        <div class="editor-panel">
            <div class="editor-toolbar">
                <button class="btn btn-primary" onclick="executeCode()">▶ Run Code</button>
                <button class="btn btn-secondary" onclick="clearEditor()">Clear</button>
                <button class="btn btn-secondary" onclick="useExample()">Use Example</button>
                <button class="btn btn-secondary" onclick="discover()">🔍 Discover</button>
                <button class="btn btn-secondary" onclick="quickTest()">🧪 Quick Test</button>
                <button class="btn btn-secondary" onclick="showShortcuts()">⌨ Shortcuts</button>
            </div>
            
            <div class="editor-split">
                <div class="output-panel" id="output">Ready to execute code...<div class="resize-handle" id="outputResizeHandle"></div></div>
                
                <div class="code-container">
                    <textarea id="codeEditor" style="display: none;"></textarea>
                </div>
            </div>
        </div>
    </div>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/mode/python/python.min.js"></script>
    
    <!-- Search functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/searchcursor.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/search.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/replace.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/dialog/dialog.min.js"></script>
    
    <!-- Comment functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/comment/comment.min.js"></script>
    
    <!-- Autocomplete functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/show-hint.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/python-hint.min.js"></script>
    
    <!-- Edit functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/edit/closebrackets.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/edit/matchbrackets.min.js"></script>
    
    <!-- Linting functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/lint/lint.min.js"></script>
    
    <script>
        var methodDocs = """ + json.dumps(method_docs) + """;
        
        var currentMethod = null;
        var codeEditor = null;
        
        // Function to properly size CodeMirror
        function resizeCodeMirror() {
            if (codeEditor) {
                var container = document.querySelector('.code-container');
                if (container) {
                    var rect = container.getBoundingClientRect();
                    // Set explicit height to match container
                    codeEditor.setSize(null, rect.height);
                }
                codeEditor.refresh();
            }
        }
        
        // Helper function for calculating edit distance (for typo suggestions)
        function levenshteinDistance(a, b) {
            if (a.length === 0) return b.length;
            if (b.length === 0) return a.length;
            
            var matrix = [];
            for (var i = 0; i <= b.length; i++) {
                matrix[i] = [i];
            }
            for (var j = 0; j <= a.length; j++) {
                matrix[0][j] = j;
            }
            
            for (var i = 1; i <= b.length; i++) {
                for (var j = 1; j <= a.length; j++) {
                    if (b.charAt(i - 1) === a.charAt(j - 1)) {
                        matrix[i][j] = matrix[i - 1][j - 1];
                    } else {
                        matrix[i][j] = Math.min(
                            matrix[i - 1][j - 1] + 1,
                            matrix[i][j - 1] + 1,
                            matrix[i - 1][j] + 1
                        );
                    }
                }
            }
            
            return matrix[b.length][a.length];
        }
        
        // Custom TouchPoint Python Linter
        function touchpointPythonLinter(text, options) {
            var errors = [];
            var lines = text.split('\\n');
            
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                
                // Skip comment lines
                if (/^\\s*#/.test(line)) continue;
                
                // Check for TouchPoint API method calls and validate them
                var methodMatch = /(model|q)\\.([a-zA-Z_][a-zA-Z0-9_]*)\\s*\\(/g;
                var match;
                while ((match = methodMatch.exec(line)) !== null) {
                    var fullMethod = match[1] + '.' + match[2];
                    var methodName = match[2];
                    
                    // Check if method exists in documentation
                    if (methodDocs[fullMethod]) {
                        // Method is documented - check parameter count
                        var methodInfo = methodDocs[fullMethod];
                        var paramsText = line.substring(match.index + match[0].length);
                        var paramCount = 0;
                        
                        // Simple parameter counting (not perfect but good enough)
                        if (paramsText.indexOf(')') > -1) {
                            var paramSection = paramsText.substring(0, paramsText.indexOf(')'));
                            if (paramSection.trim().length > 0) {
                                // Count commas not inside strings or parentheses
                                var depth = 0;
                                var inString = false;
                                var stringChar = '';
                                paramCount = 1;
                                
                                for (var j = 0; j < paramSection.length; j++) {
                                    var char = paramSection[j];
                                    if (!inString && (char === '"' || char === "'")) {
                                        inString = true;
                                        stringChar = char;
                                    } else if (inString && char === stringChar && paramSection[j-1] !== '\\\\') {
                                        inString = false;
                                    } else if (!inString) {
                                        if (char === '(' || char === '[' || char === '{') depth++;
                                        else if (char === ')' || char === ']' || char === '}') depth--;
                                        else if (char === ',' && depth === 0) paramCount++;
                                    }
                                }
                            } else {
                                paramCount = 0;
                            }
                        }
                        
                        // Check required parameters
                        var requiredParams = methodInfo.params ? methodInfo.params.filter(function(p) {
                            return !p[2].toLowerCase().includes('optional');
                        }).length : 0;
                        
                        if (paramCount < requiredParams) {
                            errors.push({
                                from: CodeMirror.Pos(i, match.index),
                                to: CodeMirror.Pos(i, match.index + match[0].length),
                                message: fullMethod + " requires " + requiredParams + " parameters, got " + paramCount,
                                severity: "error"
                            });
                        }
                    } else if (match[1] === 'model' || match[1] === 'q') {
                        // Method not in documentation - check if it's a common typo
                        var suggestions = [];
                        for (var docMethod in methodDocs) {
                            if (docMethod.startsWith(match[1] + '.')) {
                                var docMethodName = docMethod.substring(match[1].length + 1);
                                if (levenshteinDistance(methodName.toLowerCase(), docMethodName.toLowerCase()) <= 2) {
                                    suggestions.push(docMethod);
                                }
                            }
                        }
                        
                        if (suggestions.length > 0) {
                            errors.push({
                                from: CodeMirror.Pos(i, match.index),
                                to: CodeMirror.Pos(i, match.index + match[0].length),
                                message: fullMethod + " not found. Did you mean: " + suggestions.join(", ") + "?",
                                severity: "warning"
                            });
                        } else {
                            errors.push({
                                from: CodeMirror.Pos(i, match.index),
                                to: CodeMirror.Pos(i, match.index + match[0].length),
                                message: fullMethod + " is not a documented TouchPoint method",
                                severity: "info"
                            });
                        }
                    }
                }
                
                // Check for Python 3 print function (should be print statement in Python 2)
                if (/print\\s*\\(/.test(line)) {
                    // Allow if it's a function call like print(something())
                    if (!/print\\s*\\([^)]*\\(\\)/.test(line)) {
                        errors.push({
                            from: CodeMirror.Pos(i, line.indexOf('print(')),
                            to: CodeMirror.Pos(i, line.indexOf('print(') + 6),
                            message: "TouchPoint uses Python 2.7 - use 'print' statement without parentheses",
                            severity: "warning"
                        });
                    }
                }
                
                // Check for f-strings (Python 3.6+)
                if (/f["']/.test(line) && !/^\\s*#/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.search(/f["']/)),
                        to: CodeMirror.Pos(i, line.search(/f["']/) + 2),
                        message: "f-strings not supported in Python 2.7 - use .format() or % formatting",
                        severity: "error"
                    });
                }
                
                // Check for // integer division (Python 3)
                if (/\\/\\//.test(line) && !/^\\s*#/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('//')),
                        to: CodeMirror.Pos(i, line.indexOf('//') + 2),
                        message: "// operator not available in Python 2.7 - use int(a/b) instead",
                        severity: "error"
                    });
                }
                
                // Check for dangerous TouchPoint operations
                if (/model\\.ExecuteSql/.test(line) && !/^\\s*#/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('model.ExecuteSql')),
                        to: CodeMirror.Pos(i, line.indexOf('model.ExecuteSql') + 16),
                        message: "ExecuteSql is DEBUG only - will fail in production",
                        severity: "warning"
                    });
                }
                
                // Check for deprecated methods
                if (/model\\.Email2\\(/.test(line) && !/^\\s*#/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('model.Email2')),
                        to: CodeMirror.Pos(i, line.indexOf('model.Email2') + 12),
                        message: "Email2 is deprecated - use model.Email() instead",
                        severity: "warning"
                    });
                }
                
                if (/model\\.EmailContent2\\(/.test(line) && !/^\\s*#/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('model.EmailContent2')),
                        to: CodeMirror.Pos(i, line.indexOf('model.EmailContent2') + 19),
                        message: "EmailContent2 is deprecated - use model.EmailContent() instead",
                        severity: "warning"
                    });
                }
                
                // Check for common SQL injection risks
                if (/QuerySql.*\\+.*["']/.test(line) && !/^\\s*#/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, 0),
                        to: CodeMirror.Pos(i, line.length),
                        message: "Potential SQL injection - consider parameterizing query",
                        severity: "warning"
                    });
                }
                
                // Check for missing error handling on database operations
                if (/model\\.(AddPerson|UpdatePerson|DeletePerson|AddOrganization)/.test(line) && !/^\\s*#/.test(line)) {
                    // Check if this line is inside a try block
                    var inTry = false;
                    for (var j = i - 1; j >= 0 && j > i - 10; j--) {
                        if (/^\\s*try:/.test(lines[j])) {
                            inTry = true;
                            break;
                        }
                    }
                    if (!inTry) {
                        errors.push({
                            from: CodeMirror.Pos(i, 0),
                            to: CodeMirror.Pos(i, line.length),
                            message: "Consider wrapping database operations in try/except",
                            severity: "info"
                        });
                    }
                }
                
                // Check for Python 3 style type hints
                if (/def\\s+\\w+\\s*\\([^)]*:/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, 0),
                        to: CodeMirror.Pos(i, line.length),
                        message: "Type hints not supported in Python 2.7",
                        severity: "error"
                    });
                }
                
                // Check for common TouchPoint mistakes
                if (/model\\.PeopleId\\s*=/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('model.PeopleId')),
                        to: CodeMirror.Pos(i, line.indexOf('model.PeopleId') + 14),
                        message: "model.PeopleId is read-only - use model.GetPerson() to work with different people",
                        severity: "error"
                    });
                }
                
                // Check for missing imports
                if (/datetime\\./.test(line) && text.indexOf('import datetime') === -1) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('datetime')),
                        to: CodeMirror.Pos(i, line.indexOf('datetime') + 8),
                        message: "datetime is not imported - add 'import datetime' at the top",
                        severity: "error"
                    });
                }
                
                // Suggest using Data object for form data
                if (/request\\./.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('request')),
                        to: CodeMirror.Pos(i, line.indexOf('request') + 7),
                        message: "Use 'model.Data' or 'Data' to access form/request data in TouchPoint",
                        severity: "warning"
                    });
                }
                
                // Check for common attribute errors
                if (/model\\.data/.test(line)) {
                    errors.push({
                        from: CodeMirror.Pos(i, line.indexOf('model.data')),
                        to: CodeMirror.Pos(i, line.indexOf('model.data') + 10),
                        message: "Did you mean 'model.Data' (capital D) or just 'Data'?",
                        severity: "warning"
                    });
                }
            }
            
            return errors;
        }
        
        // Register the linter
        CodeMirror.registerHelper("lint", "python", touchpointPythonLinter);
        
        // Custom TouchPoint autocomplete hint provider
        function touchpointHint(editor, options) {
            var cur = editor.getCursor();
            var token = editor.getTokenAt(cur);
            var start = token.start;
            var end = cur.ch;
            var line = editor.getLine(cur.line);
            var currentWord = token.string;
            
            // Handle the case where we just typed a dot
            if (token.string === '.' && cur.ch > 0) {
                var beforeToken = editor.getTokenAt(CodeMirror.Pos(cur.line, cur.ch - 1));
                if (beforeToken.string === 'model' || beforeToken.string === 'q') {
                    currentWord = '';
                    start = cur.ch;
                }
            }
            
            // Check what we're trying to complete
            var beforeDot = "";
            if (line.charAt(start - 1) === '.') {
                // Get the object before the dot
                var beforeToken = editor.getTokenAt(CodeMirror.Pos(cur.line, start - 1));
                beforeDot = beforeToken.string;
            } else if (currentWord === '' && cur.ch > 0) {
                // Check if we're right after a dot
                var prevChar = line.charAt(cur.ch - 1);
                if (prevChar === '.') {
                    var beforeToken = editor.getTokenAt(CodeMirror.Pos(cur.line, cur.ch - 1));
                    // Get the token before the dot
                    var checkPos = cur.ch - 2;
                    while (checkPos >= 0 && line.charAt(checkPos).match(/[a-zA-Z_]/)) {
                        checkPos--;
                    }
                    beforeDot = line.substring(checkPos + 1, cur.ch - 1);
                }
            }
            
            var list = [];
            
            // If we're after "model." or "q.", show their methods
            if (beforeDot === 'model' || beforeDot === 'q') {
                for (var method in methodDocs) {
                    if (method.startsWith(beforeDot + '.')) {
                        var methodName = method.substring(beforeDot.length + 1);
                        var methodInfo = methodDocs[method];
                        
                        // Build parameter list
                        var params = [];
                        if (methodInfo.params) {
                            params = methodInfo.params.map(function(p) {
                                return p[0] + ':' + p[1];
                            });
                        }
                        
                        // Create hint object
                        var hint = {
                            text: methodName + '(',
                            displayText: methodName + '(' + params.join(', ') + ')',
                            className: 'touchpoint-hint',
                            render: function(element, self, data) {
                                var methodKey = beforeDot + '.' + data.text.replace('(', '');
                                var info = methodDocs[methodKey];
                                if (info) {
                                    var html = '<div class="hint-method">';
                                    html += '<span class="hint-name">' + data.text.replace('(', '') + '</span>';
                                    html += '<span class="hint-category">' + info.category + '</span>';
                                    html += '</div>';
                                    html += '<div class="hint-desc">' + info.desc.substring(0, 100) + '</div>';
                                    element.innerHTML = html;
                                } else {
                                    element.textContent = data.displayText;
                                }
                            }
                        };
                        
                        // Filter by current word (show all if empty)
                        if (!currentWord || methodName.toLowerCase().indexOf(currentWord.toLowerCase()) === 0) {
                            list.push(hint);
                        }
                    }
                }
            } else if (currentWord.length > 0) {
                // Suggest objects (model, q) and common keywords
                var suggestions = [
                    { text: 'model', displayText: 'model - TouchPoint API object', className: 'object-hint' },
                    { text: 'q', displayText: 'q - Query object', className: 'object-hint' },
                    { text: 'Data', displayText: 'Data - Form/request data', className: 'object-hint' },
                    { text: 'datetime', displayText: 'datetime - Date/time module', className: 'object-hint' },
                    { text: 'print ', displayText: 'print - Output text (Python 2.7)', className: 'keyword-hint' },
                    { text: 'import ', displayText: 'import - Import module', className: 'keyword-hint' },
                    { text: 'try:', displayText: 'try: - Error handling', className: 'keyword-hint' },
                    { text: 'except:', displayText: 'except: - Catch errors', className: 'keyword-hint' },
                    { text: 'if ', displayText: 'if - Conditional', className: 'keyword-hint' },
                    { text: 'for ', displayText: 'for - Loop', className: 'keyword-hint' },
                    { text: 'def ', displayText: 'def - Define function', className: 'keyword-hint' }
                ];
                
                // Also suggest method names without prefix if user is typing them
                for (var method in methodDocs) {
                    var fullName = method;
                    var shortName = method.split('.')[1];
                    if (shortName && shortName.toLowerCase().indexOf(currentWord.toLowerCase()) === 0) {
                        var info = methodDocs[method];
                        suggestions.push({
                            text: fullName + '(',
                            displayText: fullName + ' - ' + info.desc.substring(0, 50),
                            className: 'method-hint'
                        });
                    }
                }
                
                // Filter suggestions
                suggestions.forEach(function(s) {
                    if (s.text.toLowerCase().indexOf(currentWord.toLowerCase()) === 0) {
                        list.push(s);
                    }
                });
            }
            
            // Sort alphabetically
            list.sort(function(a, b) {
                return a.text.localeCompare(b.text);
            });
            
            return {
                list: list,
                from: CodeMirror.Pos(cur.line, start),
                to: CodeMirror.Pos(cur.line, end)
            };
        }
        
        // Override Python hint with our custom TouchPoint hint
        CodeMirror.registerHelper('hint', 'python', touchpointHint);
        
        // Initialize CodeMirror
        document.addEventListener('DOMContentLoaded', function() {
            codeEditor = CodeMirror.fromTextArea(document.getElementById('codeEditor'), {
                mode: 'python',
                theme: 'monokai',
                lineNumbers: true,
                lineWrapping: false,  /* Disable line wrapping */
                indentUnit: 4,
                tabSize: 4,
                matchBrackets: true,
                autoCloseBrackets: true,
                gutters: ["CodeMirror-lint-markers"],
                lint: true,
                viewportMargin: 10,  /* Render 10 lines above and below viewport */
                scrollbarStyle: "native",  /* Use native scrollbars */
                lintOnChange: true,
                hintOptions: {
                    completeSingle: false
                },
                extraKeys: {
                    'Ctrl-Enter': function() { executeCode(); },
                    'Cmd-Enter': function() { executeCode(); },
                    'Shift-Enter': function(cm) {
                        var selected = cm.getSelection();
                        if (selected) {
                            // Store the current selection positions
                            var selectionStart = cm.getCursor("from");
                            var selectionEnd = cm.getCursor("to");
                            var tempCode = codeEditor.getValue();
                            
                            // Execute selected code
                            codeEditor.setValue(selected);
                            executeCode();
                            
                            // Restore original code and selection
                            setTimeout(function() { 
                                codeEditor.setValue(tempCode);
                                // Re-select the same text
                                codeEditor.setSelection(selectionStart, selectionEnd);
                                codeEditor.focus();
                            }, 100);
                        } else {
                            executeCode();
                        }
                    },
                    'Tab': function(cm) {
                        if (cm.somethingSelected()) {
                            cm.indentSelection("add");
                        } else {
                            cm.replaceSelection('    ', 'end');
                        }
                    },
                    'Shift-Tab': function(cm) {
                        cm.indentSelection("subtract");
                    },
                    'Ctrl-L': function() { clearEditor(); },
                    'Cmd-L': function() { clearEditor(); },
                    'F5': function() { executeCode(); },
                    'Ctrl-/': 'toggleComment',
                    'Cmd-/': 'toggleComment',
                    'Ctrl-Space': function(cm) { cm.showHint(); },
                    'Ctrl-Z': 'undo',
                    'Cmd-Z': 'undo',
                    'Ctrl-Y': 'redo',
                    'Cmd-Y': 'redo',
                    'Ctrl-Shift-Z': 'redo',
                    'Cmd-Shift-Z': 'redo'
                }
            });
            codeEditor.setSize('100%', '100%');
            codeEditor.setValue('# Enter Python code here\\n# Available: model, q\\n\\nprint "Hello TouchPoint!"');
            
            // Auto-trigger autocomplete on dot after model or q
            codeEditor.on('inputRead', function(cm, change) {
                if (change.text && change.text.length === 1 && change.text[0] === '.') {
                    var cursor = cm.getCursor();
                    var token = cm.getTokenAt(CodeMirror.Pos(cursor.line, cursor.ch - 1));
                    if (token.string === 'model' || token.string === 'q') {
                        setTimeout(function() { cm.showHint(); }, 100);
                    }
                }
            });
            
            // Initialize resizable panels
            initResizablePanels();
            
            // Force refresh CodeMirror after a short delay to ensure proper sizing
            setTimeout(function() {
                resizeCodeMirror();
                // Additional delayed refresh to catch any layout shifts
                setTimeout(function() {
                    if (codeEditor) {
                        codeEditor.refresh();
                        // Force recalculation of scroll area
                        var totalLines = codeEditor.lineCount();
                        if (totalLines > 0) {
                            codeEditor.scrollTo(0, 0);
                            var lastLine = {line: totalLines - 1, ch: 0};
                            var coords = codeEditor.charCoords(lastLine, "local");
                            // Ensure the editor knows about all content
                            codeEditor.setSize(null, null);
                        }
                    }
                }, 300);
            }, 100);
            
            // Refresh CodeMirror on window resize
            window.addEventListener('resize', function() {
                resizeCodeMirror();
            });
            
            // Also refresh when output panel is resized
            var outputResizeObserver = new ResizeObserver(function() {
                resizeCodeMirror();
            });
            var outputPanel = document.getElementById('output');
            if (outputPanel) {
                outputResizeObserver.observe(outputPanel);
            }
        });
        
        // Helper function to update output while preserving resize handle
        function updateOutput(text) {
            var outputPanel = document.getElementById('output');
            var resizeHandle = document.getElementById('outputResizeHandle');
            
            // Clear and update content while preserving the resize handle
            outputPanel.innerHTML = '';
            outputPanel.appendChild(document.createTextNode(text));
            
            // Re-add the resize handle if it exists
            if (!resizeHandle) {
                resizeHandle = document.createElement('div');
                resizeHandle.className = 'resize-handle';
                resizeHandle.id = 'outputResizeHandle';
            }
            outputPanel.appendChild(resizeHandle);
        }
        
        // Make panels resizable
        function initResizablePanels() {
            var outputPanel = document.getElementById('output');
            var resizeHandle = document.getElementById('outputResizeHandle');
            var isResizing = false;
            var startY = 0;
            var startHeight = 0;
            
            resizeHandle.addEventListener('mousedown', function(e) {
                isResizing = true;
                startY = e.pageY;
                startHeight = outputPanel.offsetHeight;
                document.body.style.cursor = 'ns-resize';
                e.preventDefault();
            });
            
            document.addEventListener('mousemove', function(e) {
                if (!isResizing) return;
                
                var deltaY = e.pageY - startY;
                var newHeight = startHeight + deltaY;
                
                // Limit the height
                newHeight = Math.max(100, Math.min(400, newHeight));
                
                outputPanel.style.flex = '0 0 ' + newHeight + 'px';
                
                // Refresh CodeMirror after resize
                if (codeEditor) {
                    setTimeout(function() {
                        codeEditor.refresh();
                    }, 0);
                }
            });
            
            document.addEventListener('mouseup', function() {
                if (isResizing) {
                    isResizing = false;
                    document.body.style.cursor = '';
                }
            });
        }
        
        function toggleCategory(header) {
            var content = header.nextElementSibling;
            content.classList.toggle('collapsed');
        }
        
        function showMethod(methodName) {
            document.querySelectorAll('.method-item').forEach(function(item) {
                item.classList.remove('active');
                if (item.textContent === methodName) {
                    item.classList.add('active');
                }
            });
            
            currentMethod = methodName;
            var doc = methodDocs[methodName];
            
            document.getElementById('methodName').textContent = methodName;
            
            if (doc) {
                document.getElementById('methodCategory').textContent = doc.category || 'OTHER';
                document.getElementById('methodDesc').textContent = doc.desc;
                
                // Show parameters
                if (doc.params && doc.params.length > 0) {
                    var html = '';
                    doc.params.forEach(function(p) {
                        html += '<div class="param">';
                        html += '<span class="param-name">' + p[0] + '</span>';
                        html += '<span class="param-type">' + p[1] + '</span>';
                        html += ': ' + p[2];
                        html += '</div>';
                    });
                    document.getElementById('paramsList').innerHTML = html;
                    document.getElementById('paramsSection').style.display = 'block';
                } else {
                    document.getElementById('paramsSection').style.display = 'none';
                }
                
                // Show returns
                if (doc.returns) {
                    document.getElementById('returnsInfo').textContent = doc.returns;
                    document.getElementById('returnsSection').style.display = 'block';
                } else {
                    document.getElementById('returnsSection').style.display = 'none';
                }
                
                // Show example
                if (doc.example) {
                    document.getElementById('exampleCode').textContent = doc.example;
                    document.getElementById('exampleSection').style.display = 'block';
                } else {
                    document.getElementById('exampleSection').style.display = 'none';
                }
            } else {
                document.getElementById('methodCategory').textContent = 'UNDOCUMENTED';
                document.getElementById('methodDesc').textContent = 'No documentation available. Click "Discover" to explore.';
                document.getElementById('paramsSection').style.display = 'none';
                document.getElementById('returnsSection').style.display = 'none';
                document.getElementById('exampleSection').style.display = 'none';
            }
        }
        
        function discover() {
            if (!currentMethod) {
                alert('Please select a method first');
                return;
            }
            
            var parts = currentMethod.split('.');
            var objName = parts[0];
            var methodName = parts[1];
            
            var code = '# Discovering ' + currentMethod + '\\n' +
'import inspect\\n' +
'\\n' +
'try:\\n' +
'    method = ' + objName + '.' + methodName + '\\n' +
'    print "=" * 60\\n' +
'    print "Method: ' + currentMethod + '"\\n' +
'    print "=" * 60\\n' +
'    \\n' +
'    # Get type\\n' +
'    print "\\\\nType:", type(method)\\n' +
'    \\n' +
'    # Try to get docstring\\n' +
'    try:\\n' +
'        doc = method.__doc__\\n' +
'        if doc:\\n' +
'            print "\\\\nDocumentation:"\\n' +
'            print doc[:500]\\n' +
'    except:\\n' +
'        pass\\n' +
'    \\n' +
'    # Try to get signature\\n' +
'    try:\\n' +
'        if hasattr(inspect, \\'getargspec\\'):\\n' +
'            spec = inspect.getargspec(method)\\n' +
'            print "\\\\nArguments:", spec.args\\n' +
'            if spec.defaults:\\n' +
'                print "Defaults:", spec.defaults\\n' +
'    except:\\n' +
'        pass\\n' +
'    \\n' +
'    # Try calling with no args\\n' +
'    print "\\\\nTest call with no arguments:"\\n' +
'    try:\\n' +
'        result = method()\\n' +
'        print "  Success! Returns:", type(result)\\n' +
'    except TypeError as e:\\n' +
'        print "  TypeError:", str(e)\\n' +
'    except Exception as e:\\n' +
'        print "  Error:", str(e)\\n' +
'        \\n' +
'except Exception as e:\\n' +
'    print "Error:", str(e)';
            
            if (codeEditor) {
                codeEditor.setValue(code);
            }
            executeCode();
        }
        
        function quickTest() {
            var code = '# Quick System Test\\n' +
'print "=" * 50\\n' +
'print "TouchPoint API Explorer - System Test"\\n' +
'print "=" * 50\\n' +
'\\n' +
'# Check user\\n' +
'if hasattr(model, "PeopleId"):\\n' +
'    print "\\\\nUser ID:", model.PeopleId\\n' +
'if hasattr(model, "UserName"):\\n' +
'    print "Username:", model.UserName\\n' +
'\\n' +
'# Count available methods\\n' +
'if "model" in globals():\\n' +
'    model_methods = [m for m in dir(model) if not m.startswith("_")]\\n' +
'    print "\\\\nModel methods available:", len(model_methods)\\n' +
'if "q" in globals():\\n' +
'    q_methods = [m for m in dir(q) if not m.startswith("_")]\\n' +
'    print "Query methods available:", len(q_methods)';
    
            if (codeEditor) {
                codeEditor.setValue(code);
            }
            executeCode();
        }
        
        function executeCode() {
            var code = codeEditor ? codeEditor.getValue() : document.getElementById('codeEditor').value;
            
            if (!code.trim()) {
                updateOutput('Please enter code to execute');
                return;
            }
            
            updateOutput('Executing...');
            
            var scriptPath = window.location.pathname;
            var formPath = scriptPath.replace('/PyScript/', '/PyScriptForm/');
            
            var xhr = new XMLHttpRequest();
            xhr.open('POST', formPath, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        var response = xhr.responseText;
                        // Extract only the output
                        var ajaxMarker = '<!-- AJAX_COMPLETE -->';
                        if (response.indexOf(ajaxMarker) > -1) {
                            response = response.substring(0, response.indexOf(ajaxMarker));
                        }
                        updateOutput(response.trim());
                    } else {
                        updateOutput('Error: Request failed (status ' + xhr.status + ')');
                    }
                }
            };
            
            xhr.send('code=' + encodeURIComponent(code));
        }
        
        function clearEditor() {
            if (codeEditor) {
                codeEditor.setValue('');
            }
        }
        
        function useExample() {
            if (currentMethod && methodDocs[currentMethod] && methodDocs[currentMethod].example) {
                if (codeEditor) {
                    codeEditor.setValue(methodDocs[currentMethod].example);
                }
            }
        }
        
        function filterMethods(term) {
            var search = term.toLowerCase();
            document.querySelectorAll('.method-item').forEach(function(item) {
                item.style.display = item.textContent.toLowerCase().includes(search) ? 'block' : 'none';
            });
            
            // Hide empty categories
            document.querySelectorAll('.category-section').forEach(function(section) {
                var hasVisible = false;
                section.querySelectorAll('.method-item').forEach(function(item) {
                    if (item.style.display !== 'none') {
                        hasVisible = true;
                    }
                });
                section.style.display = hasVisible ? 'block' : 'none';
            });
        }
        
        function showShortcuts() {
            document.getElementById('shortcutsModal').style.display = 'block';
        }
        
        function closeShortcutsModal() {
            document.getElementById('shortcutsModal').style.display = 'none';
        }
        
        // Close modal when clicking outside
        window.onclick = function(event) {
            var modal = document.getElementById('shortcutsModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }
        
        // Add collapsible functionality to category headers
        function toggleCategory(header) {
            header.classList.toggle('collapsed');
            var content = header.nextElementSibling;
            if (content && content.classList.contains('category-content')) {
                content.classList.toggle('collapsed');
            }
        }
        
        // Initialize resizable panels
        function initResizablePanels() {
            // Sidebar resize
            var sidebarHandle = document.getElementById('sidebarResizeHandle');
            var sidebar = document.getElementById('sidebar');
            var isResizingSidebar = false;
            
            if (sidebarHandle && sidebar) {
                sidebarHandle.addEventListener('mousedown', function(e) {
                    isResizingSidebar = true;
                    document.body.style.cursor = 'col-resize';
                    e.preventDefault();
                });
                
                document.addEventListener('mousemove', function(e) {
                    if (!isResizingSidebar) return;
                    var newWidth = e.clientX;
                    if (newWidth > 150 && newWidth < 500) {
                        sidebar.style.width = newWidth + 'px';
                        sidebar.style.flexShrink = '0';
                    }
                });
                
                document.addEventListener('mouseup', function() {
                    if (isResizingSidebar) {
                        isResizingSidebar = false;
                        document.body.style.cursor = 'default';
                        // Refresh CodeMirror after resize
                        if (typeof resizeCodeMirror === 'function') {
                            resizeCodeMirror();
                        }
                    }
                });
            }
            
            // Doc panel resize
            var docHandle = document.getElementById('docResizeHandle');
            var docPanel = document.getElementById('docPanel');
            var isResizingDoc = false;
            var docStartX = 0;
            var docStartWidth = 0;
            
            if (docHandle && docPanel) {
                docHandle.addEventListener('mousedown', function(e) {
                    isResizingDoc = true;
                    docStartX = e.clientX;
                    docStartWidth = parseInt(window.getComputedStyle(docPanel).width, 10);
                    document.body.style.cursor = 'col-resize';
                    e.preventDefault();
                });
                
                document.addEventListener('mousemove', function(e) {
                    if (!isResizingDoc) return;
                    // Fix: for right panel, dragging left should make it bigger
                    var diff = e.clientX - docStartX;  // positive when moving right
                    var newWidth = docStartWidth - diff;  // dragging left (negative diff) = bigger width
                    if (newWidth > 150 && newWidth < 600) {
                        docPanel.style.width = newWidth + 'px';
                        docPanel.style.flexShrink = '0';
                    }
                });
                
                document.addEventListener('mouseup', function() {
                    if (isResizingDoc) {
                        isResizingDoc = false;
                        document.body.style.cursor = 'default';
                        // Refresh CodeMirror after resize
                        if (typeof resizeCodeMirror === 'function') {
                            resizeCodeMirror();
                        }
                    }
                });
            }
        }
        
        // Initialize collapsible categories on page load
        window.addEventListener('load', function() {
            // Initialize resizable panels
            initResizablePanels();
            // Make category headers collapsible
            document.querySelectorAll('.category-header').forEach(function(header) {
                header.onclick = function() {
                    toggleCategory(this);
                };
            });
            
            // Add a "Collapse All" / "Expand All" button to the sidebar
            var searchBox = document.querySelector('.search-box');
            if (searchBox && !document.getElementById('toggleAllBtn')) {
                var toggleBtn = document.createElement('button');
                toggleBtn.id = 'toggleAllBtn';
                toggleBtn.textContent = 'Collapse All';
                toggleBtn.style.cssText = 'width: 100%; margin-top: 5px; padding: 5px; font-size: 11px; background: #007bff; color: white; border: none; border-radius: 3px; cursor: pointer;';
                toggleBtn.onclick = function() {
                    var headers = document.querySelectorAll('.category-header');
                    var allCollapsed = Array.from(headers).every(h => h.classList.contains('collapsed'));
                    
                    headers.forEach(function(header) {
                        var content = header.nextElementSibling;
                        if (allCollapsed) {
                            header.classList.remove('collapsed');
                            if (content) content.classList.remove('collapsed');
                        } else {
                            header.classList.add('collapsed');
                            if (content) content.classList.add('collapsed');
                        }
                    });
                    
                    toggleBtn.textContent = allCollapsed ? 'Collapse All' : 'Expand All';
                };
                searchBox.appendChild(toggleBtn);
            }
        });
    </script>
</body>
</html>"""
