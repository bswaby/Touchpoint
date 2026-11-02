#roles=Admin

#####################################################################
# TouchPoint SQL Query Explorer & Editor - Multi-Tab Version
#####################################################################

# Features:
# 1. Multiple editor tabs
# 2. Database schema explorer with lazy loading
# 3. SQL editor with syntax highlighting
# 4. Enhanced SQL formatting
# 5. Saved queries functionality
# 6. Collapsible schema panel
# 7. Keyboard shortcuts (Ctrl+T, Ctrl+Tab, etc.)
#
# Version: 2.1

#Written By: Ben Swaby (TPxi Software, LLC)
#Email: bswaby@fbchtn.org
#Website: https://tpxisoftware.com
#GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)
#----------------------------------------------------------------
#These tools are free because they should be.
#If they've saved you time or helped your team, and you want to
#support continued development, check out:
#
#DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)
#https://displaycache.com
#
#TPxi Go(TM) - your church contacts, wherever you work.
#Look up anyone in TouchPoint(R), log calls and emails from Outlook
#or your phone. No tab switching, no lost context.
#https://tpxigo.com
#----------------------------------------------------------------

#
# RELEASE NOTES
# 2.0 (November 2026)
# - Row limit: queries stop at 1,000 rows unless you pick more (10,000,
#   or all up to 50,000). The page says when there were more rows.
# - Warning before an unfiltered query on one of TouchPoint's largest
#   tables (TagPerson, ActivityLog, Attend and others).
# - Full values: long text is no longer cut at 1,000 characters or
#   flattened onto one line. The grid shortens long cells (click to see
#   all of it), and Export CSV and Copy always use the full values.
# - Columns come back in the order the query returns them, including
#   CTEs, subqueries and SELECT *. Two columns with the same name both
#   show (the second becomes Name_2).
# - Dates show as 2026-11-02 14:30 instead of the US text form, so they
#   sort correctly.
# - Results grid: click a header to sort, type under a header to filter,
#   NULL shows as NULL, double-click a cell to copy it, click a row
#   number to copy the row, Copy grid copies everything for Excel.
# - Links: PeopleId, OrganizationId, FamilyId, MeetingId and BundleId
#   open their page. Columns named Id are no longer guessed to be people.
# - Saved queries are now private to you (stored on your own record),
#   instead of in Special Content, which anyone can read without signing
#   in. Queries saved by earlier versions still show and can be copied
#   to My queries. Other SQL in Special Content shows in its own group.
# - Query history: the last 50 queries you ran, under the History tab.
# - Open tabs and their SQL come back after a reload.
# - This page only reads data. Writes to permanent tables are refused
#   for everyone (temp tables and table variables still work).
# - Error tips no longer say table names are case sensitive.
# - Accented names no longer break the results.
# - The hide-panel button sits at the end of the tab row instead of over
#   the tabs, and the hidden panel leaves a slim strip to bring it back.
#
# 2.1 (November 2026)
# - Find and replace in the editor (Ctrl+F, Ctrl+H; Cmd on a Mac).
# - Ctrl+/ comments the selected lines on and off.
# - Run statement (Ctrl+Shift+Enter) runs just the statement the cursor
#   is in, so a long scratch file can hold many queries.
# - Format SQL uses a real T-SQL formatter (CTEs, CASE, subqueries).
#   Select part of the SQL to format only that part.
# - Parameters: write @StartDate (or any @Name you do not DECLARE) and a
#   box for it appears above the editor. Values are sent as real query
#   parameters and remembered. @UserPeopleId and @CurrentOrgId are filled
#   in by TouchPoint; @BlueToolbarTagId takes a tag Id to test with.
# - Schema search also finds columns: type CampusId to see every table
#   and view that has it.
# - Results: a totals row (sum and average for number columns), a column
#   profile (distinct values, NULLs, most common values), a quick bar
#   chart, and Excel export that keeps numbers and dates as numbers and
#   dates. Export, Excel, Copy grid and Chart now sit above the results.
# - New look: the page uses its own class names, so TouchPoint's styles
#   no longer reorder the toolbar or box in the left panel. Run and Run
#   statement share one button, the page fills the window, and the grid,
#   tabs, schema list, errors and shortcuts popup were tidied to match.
#####################################################################

#written by: Ben Swaby
#email:bswaby@fbchtn.org

import json
import re
from datetime import datetime

# JSON for AJAX replies. json.dumps breaks on accented names in IronPython,
# so this walks the data and writes every non-ASCII character as \\uXXXX.
def _json_escape_string(s):
    if not isinstance(s, basestring):
        s = unicode(s)
    parts = ['"']
    for ch in s:
        try:
            code = ord(ch)
        except Exception:
            parts.append('\\ufffd')
            continue
        if code == 0x22:
            parts.append('\\"')
        elif code == 0x5C:
            parts.append('\\\\')
        elif code == 0x0A:
            parts.append('\\n')
        elif code == 0x0D:
            parts.append('\\r')
        elif code == 0x09:
            parts.append('\\t')
        elif code < 0x20 or code >= 0x7F:
            if code > 0xFFFF:
                code -= 0x10000
                parts.append('\\u%04x\\u%04x' % (0xD800 + (code >> 10), 0xDC00 + (code & 0x3FF)))
            else:
                parts.append('\\u%04x' % code)
        else:
            parts.append(chr(code))
    parts.append('"')
    return ''.join(parts)

def _json_encode(obj):
    if obj is None:
        return 'null'
    if obj is True:
        return 'true'
    if obj is False:
        return 'false'
    if isinstance(obj, (int, long)):
        return str(obj)
    if isinstance(obj, float):
        return 'null' if obj != obj else repr(obj)
    if isinstance(obj, basestring):
        return _json_escape_string(obj)
    if isinstance(obj, dict):
        return '{' + ','.join(_json_escape_string(unicode(k)) + ':' + _json_encode(v)
                              for k, v in obj.items()) + '}'
    if isinstance(obj, (list, tuple)):
        return '[' + ','.join(_json_encode(x) for x in obj) + ']'
    return _json_escape_string(unicode(obj))

def safe_json(obj):
    try:
        return _json_encode(obj)
    except Exception as e:
        return '{"success":false,"error":' + _json_escape_string('safe_json: ' + repr(e)) + '}'

def get_param(name, default=''):
    try:
        v = getattr(model.Data, name, None)
    except Exception:
        v = None
    if v is None:
        return default
    return unicode(v)

# Row limits. MAX_ROWS_ALL is the ceiling for "All rows".
DEFAULT_MAX_ROWS = 1000
MAX_ROWS_ALL = 50000
# Long values are sent in full up to this size, so exports are complete.
MAX_VALUE_CHARS = 100000
# Private saved queries live on the user's own record (PeopleExtra), not in
# Special Content, which /v1/Contents serves to anyone without signing in.
SAVED_FIELD = 'SQE_SavedQueries'
SQE_KEYWORD = 'SQL Query Explorer'

model.Header = "SQL Query Explorer"

class QueryExplorer:
    def __init__(self):
        self.user_id = model.UserPeopleId
        self.is_developer = model.UserIsInRole("Developer")
        self.is_admin = model.UserIsInRole("Admin")
        
        # Safety check
        self.has_permission = self.is_developer or self.is_admin
    
    def get_database_objects(self):
        """Get tables and views organized by type"""
        objects_sql = """
        SELECT 
            t.TABLE_SCHEMA,
            t.TABLE_NAME,
            t.TABLE_TYPE,
            COUNT(c.COLUMN_NAME) as COLUMN_COUNT
        FROM INFORMATION_SCHEMA.TABLES t
        LEFT JOIN INFORMATION_SCHEMA.COLUMNS c 
            ON t.TABLE_SCHEMA = c.TABLE_SCHEMA 
            AND t.TABLE_NAME = c.TABLE_NAME
        WHERE t.TABLE_TYPE IN ('BASE TABLE', 'VIEW')
            AND t.TABLE_SCHEMA NOT IN ('sys', 'INFORMATION_SCHEMA', 'guest', 'db_owner', 'db_accessadmin', 'db_securityadmin', 'db_ddladmin', 'db_backupoperator', 'db_datareader', 'db_datawriter', 'db_denydatareader', 'db_denydatawriter')
        GROUP BY t.TABLE_SCHEMA, t.TABLE_NAME, t.TABLE_TYPE
        ORDER BY t.TABLE_TYPE, t.TABLE_SCHEMA, t.TABLE_NAME
        """
        
        return q.QuerySql(objects_sql)
    
    def get_full_schema(self):
        """Get complete schema with all tables and their columns in one query"""
        schema_sql = """
        SELECT 
            t.TABLE_SCHEMA,
            t.TABLE_NAME,
            t.TABLE_TYPE,
            c.COLUMN_NAME,
            c.DATA_TYPE,
            c.ORDINAL_POSITION
        FROM INFORMATION_SCHEMA.TABLES t
        LEFT JOIN INFORMATION_SCHEMA.COLUMNS c 
            ON t.TABLE_SCHEMA = c.TABLE_SCHEMA 
            AND t.TABLE_NAME = c.TABLE_NAME
        WHERE t.TABLE_TYPE IN ('BASE TABLE', 'VIEW')
            AND t.TABLE_SCHEMA NOT IN ('sys', 'INFORMATION_SCHEMA', 'guest', 'db_owner', 'db_accessadmin', 'db_securityadmin', 'db_ddladmin', 'db_backupoperator', 'db_datareader', 'db_datawriter', 'db_denydatareader', 'db_denydatawriter')
        ORDER BY t.TABLE_SCHEMA, t.TABLE_NAME, c.ORDINAL_POSITION
        """
        
        results = q.QuerySql(schema_sql)
        
        # Organize into a structured format
        schema = {
            'tables': {},
            'views': {}
        }
        
        # Track some stats for debugging
        lookup_table_count = 0
        lookup_column_count = 0
        
        for row in results:
            # Build the key for storage - use schema.table for non-dbo, just table for dbo
            if row.TABLE_SCHEMA == 'dbo':
                storage_key = row.TABLE_NAME
            else:
                storage_key = row.TABLE_SCHEMA + '.' + row.TABLE_NAME
            
            target = schema['tables'] if row.TABLE_TYPE == 'BASE TABLE' else schema['views']
            
            if storage_key not in target:
                target[storage_key] = {
                    'schema': row.TABLE_SCHEMA,
                    'name': row.TABLE_NAME,
                    'columns': []
                }
                if row.TABLE_SCHEMA == 'lookup':
                    lookup_table_count += 1
            
            if row.COLUMN_NAME:
                target[storage_key]['columns'].append(row.COLUMN_NAME)
                if row.TABLE_SCHEMA == 'lookup':
                    lookup_column_count += 1
        
        # Debug: Log lookup tables found
        if lookup_table_count > 0:
            model.DebugPrint("Found %d lookup tables with %d total columns" % (lookup_table_count, lookup_column_count))
        
        return schema
    
    def get_table_columns(self, schema_name, table_name):
        """Get columns for a specific table"""
        columns_sql = """
        SELECT 
            c.COLUMN_NAME,
            c.DATA_TYPE,
            c.CHARACTER_MAXIMUM_LENGTH,
            c.IS_NULLABLE,
            CASE 
                WHEN pk.COLUMN_NAME IS NOT NULL THEN 'PK'
                WHEN fk.COLUMN_NAME IS NOT NULL THEN 'FK'
                ELSE ''
            END as KEY_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS c
        LEFT JOIN (
            SELECT ku.COLUMN_NAME
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
            JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
            WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
                AND tc.TABLE_SCHEMA = '{0}'
                AND tc.TABLE_NAME = '{1}'
        ) pk ON c.COLUMN_NAME = pk.COLUMN_NAME
        LEFT JOIN (
            SELECT ku.COLUMN_NAME
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
            JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
            WHERE tc.CONSTRAINT_TYPE = 'FOREIGN KEY'
                AND tc.TABLE_SCHEMA = '{0}'
                AND tc.TABLE_NAME = '{1}'
        ) fk ON c.COLUMN_NAME = fk.COLUMN_NAME
        WHERE c.TABLE_SCHEMA = '{0}'
            AND c.TABLE_NAME = '{1}'
        ORDER BY c.ORDINAL_POSITION
        """.format(schema_name.replace("'", "''"), table_name.replace("'", "''"))
        
        return q.QuerySql(columns_sql)
    
    def get_saved_queries(self):
        """SQL scripts in Special Content, split into the ones this page saved
        (keyword 'SQL Query Explorer') and everything else."""
        sql = """
        SELECT
            c.Id,
            c.[Name],
            c.Body,
            CASE WHEN EXISTS (SELECT 1 FROM ContentKeyWords k
                              WHERE k.Id = c.Id AND k.Word = '%s')
                 THEN 1 ELSE 0 END AS FromSqe
        FROM Content c
        WHERE c.TypeID = 4  -- SQL Script type
        ORDER BY c.[Name]
        """ % SQE_KEYWORD.replace("'", "''")
        try:
            return list(q.QuerySql(sql))
        except Exception:
            try:
                return list(q.QuerySql("SELECT c.Id, c.[Name], c.Body, 0 AS FromSqe FROM Content c WHERE c.TypeID = 4 ORDER BY c.[Name]"))
            except Exception:
                return []

    def load_my_queries(self):
        """This user's private saved queries, newest format: list of dicts."""
        if not self.user_id:
            return []
        try:
            raw = model.ExtraValueText(self.user_id, SAVED_FIELD)
        except Exception:
            raw = ''
        if not raw:
            return []
        try:
            data = json.loads(raw)
        except Exception:
            return []
        queries = data.get('queries', []) if isinstance(data, dict) else []
        return [x for x in queries if isinstance(x, dict) and x.get('name')]

    def store_my_queries(self, queries):
        text = safe_json({'v': 1, 'queries': queries})
        if len(text) > 1000000:
            raise Exception('Your saved queries would pass 1 MB. Delete a few first.')
        model.AddExtraValueText(self.user_id, SAVED_FIELD, text)

    def save_query(self, name, sql):
        """Save (or replace) a private query on the user's own record"""
        if not self.user_id:
            raise Exception('No signed-in user to save the query for')
        name = name.strip()[:200]
        queries = [x for x in self.load_my_queries() if x.get('name') != name]
        queries.append({
            'name': name,
            'sql': sql,
            'saved': datetime.now().strftime('%Y-%m-%d %H:%M')
        })
        queries.sort(key=lambda x: x.get('name', '').lower())
        self.store_my_queries(queries)
        return True

    def delete_query(self, name):
        if not self.user_id:
            raise Exception('No signed-in user')
        queries = self.load_my_queries()
        kept = [x for x in queries if x.get('name') != name]
        if len(kept) == len(queries):
            return False
        self.store_my_queries(kept)
        return True

    def get_common_queries(self):
        """Predefined useful queries from TouchPoint SQL Documentation
        Reference: https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html
        """
        return [
            {
                'name': 'Active Church Members',
                'category': 'People',
                'description': 'Find active church members with key details',
                'sql': '''SELECT TOP 1000 
    p.PeopleId,
    p.Name,
    p.EmailAddress,
    p.CellPhone,
    ms.Description as MemberStatus,
    DATEDIFF(year, p.BDate, GETDATE()) as Age,
    c.Description as Campus
FROM People p
JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
LEFT JOIN lookup.Campus c ON p.CampusId = c.Id
WHERE p.IsDeceased = 0 
    AND p.ArchivedFlag = 0
    AND p.MemberStatusId = 10  -- Members only
ORDER BY p.Name'''
            },
            {
                'name': 'Recent Attendance',
                'category': 'Attendance',
                'description': 'Track recent organization attendance (last week)',
                'sql': '''SELECT TOP 500
    p.Name,
    o.OrganizationName,
    m.MeetingDate,
    at.Description as AttendanceType
FROM Attend a
JOIN People p ON a.PeopleId = p.PeopleId
JOIN Organizations o ON a.OrganizationId = o.OrganizationId
JOIN Meetings m ON a.MeetingId = m.MeetingId
JOIN lookup.AttendType at ON a.AttendanceTypeId = at.Id
WHERE a.AttendanceFlag = 1  -- Actual attendance
    AND m.MeetingDate >= DATEADD(week, -1, GETDATE())
    AND m.DidNotMeet = 0
    AND p.IsDeceased = 0 
    AND p.ArchivedFlag = 0
ORDER BY m.MeetingDate DESC, o.OrganizationName, p.Name'''
            },
            {
                'name': 'Monthly Giving Summary',
                'category': 'Contributions',
                'description': 'Analyze giving by fund for monthly reporting',
                'sql': '''SELECT 
    YEAR(c.ContributionDate) AS Year,
    MONTH(c.ContributionDate) AS Month,
    DATENAME(month, c.ContributionDate) AS MonthName,
    f.FundName,
    COUNT(DISTINCT c.PeopleId) AS Givers,
    COUNT(*) AS GiftCount,
    SUM(c.ContributionAmount) AS TotalAmount,
    AVG(c.ContributionAmount) AS AverageGift
FROM Contribution c
JOIN ContributionFund f ON c.FundId = f.FundId
WHERE c.ContributionDate >= DATEADD(month, -12, GETDATE())
    AND c.ContributionStatusId = 0  -- Recorded contributions
    AND c.ContributionTypeId IN (1, 5, 8)  -- Online, Check, Cash
GROUP BY YEAR(c.ContributionDate), 
    MONTH(c.ContributionDate),
    DATENAME(month, c.ContributionDate),
    f.FundName
ORDER BY Year DESC, Month DESC, f.FundName'''
            },
            {
                'name': 'Organization Members with Leaders',
                'category': 'Organizations',
                'description': 'List active organizations with member counts and leaders',
                'sql': '''SELECT 
    o.OrganizationName,
    o.LeaderName,
    ot.Description as OrgType,
    COUNT(om.PeopleId) as MemberCount,
    AVG(CASE WHEN a.AttendanceFlag = 1 THEN 1.0 ELSE 0.0 END) * 100 as AttendancePct
FROM Organizations o
LEFT JOIN OrganizationMembers om ON o.OrganizationId = om.OrganizationId
LEFT JOIN lookup.OrganizationType ot ON o.OrganizationTypeId = ot.Id
LEFT JOIN (
    SELECT OrganizationId, PeopleId, AttendanceFlag
    FROM Attend
    WHERE MeetingDate >= DATEADD(week, -4, GETDATE())
) a ON a.OrganizationId = o.OrganizationId AND a.PeopleId = om.PeopleId
WHERE o.OrganizationStatusId = 30  -- Active
GROUP BY o.OrganizationName, o.LeaderName, ot.Description
HAVING COUNT(om.PeopleId) > 0
ORDER BY MemberCount DESC'''
            },
            {
                'name': 'Family Units Analysis',
                'category': 'Families',
                'description': 'Analyze family composition and giving patterns',
                'sql': '''SELECT TOP 500
    p.FamilyId,
    MIN(CASE WHEN p.PositionInFamilyId = 10 THEN p.Name2 ELSE NULL END) as FamilyName,
    MIN(p.HomePhone) as HomePhone,
    COUNT(p.PeopleId) as FamilySize,
    SUM(CASE WHEN p.PositionInFamilyId = 10 THEN 1 ELSE 0 END) as Adults,
    SUM(CASE WHEN p.PositionInFamilyId = 30 THEN 1 ELSE 0 END) as Children,
    MAX(CASE WHEN p.PositionInFamilyId = 10 THEN p.Name ELSE NULL END) as HeadOfHousehold,
    MAX(c.LastGift) as LastContribution
FROM People p
LEFT JOIN (
    SELECT PeopleId, MAX(ContributionDate) as LastGift
    FROM Contribution
    WHERE ContributionStatusId = 0
    GROUP BY PeopleId
) c ON p.PeopleId = c.PeopleId
WHERE p.IsDeceased = 0 
    AND p.ArchivedFlag = 0
    AND p.FamilyId IS NOT NULL
GROUP BY p.FamilyId
HAVING COUNT(p.PeopleId) > 1  -- Multi-person families
ORDER BY FamilySize DESC'''
            },
            {
                'name': 'Email Engagement Metrics',
                'category': 'Communications',
                'description': 'Track email open rates and engagement',
                'sql': '''SELECT TOP 100
    e.Subject,
    e.FromAddr,
    e.Sent,
    COUNT(DISTINCT et.PeopleId) as Recipients,
    SUM(CASE WHEN er.Type = 'o' THEN 1 ELSE 0 END) as Opens,
    CAST(SUM(CASE WHEN er.Type = 'o' THEN 1.0 ELSE 0.0 END) * 100.0 / 
         NULLIF(COUNT(DISTINCT et.PeopleId), 0) as DECIMAL(10,2)) as OpenRate
FROM EmailQueue e
JOIN EmailQueueTo et ON e.Id = et.Id
LEFT JOIN EmailResponses er ON e.Id = er.EmailQueueId
WHERE e.Sent >= DATEADD(day, -30, GETDATE())
    AND e.Transactional = 0  -- Mass emails only
    AND e.Error IS NULL
GROUP BY e.Id, e.Subject, e.FromAddr, e.Sent
HAVING COUNT(DISTINCT et.PeopleId) > 10  -- Mass emails
ORDER BY e.Sent DESC'''
            },
            {
                'name': 'Connect Group Attendance Patterns',
                'category': 'Connect Groups',
                'description': 'Analyze Connect group attendance consistency',
                'sql': '''WITH GroupAttendance AS (
    SELECT 
        o.OrganizationId,
        o.OrganizationName,
        a.PeopleId,
        COUNT(DISTINCT m.MeetingDate) as MeetingsAttended,
        COUNT(DISTINCT CASE WHEN a.AttendanceFlag = 1 THEN m.MeetingDate END) as TimesPresent
    FROM Organizations o
    JOIN Meetings m ON o.OrganizationId = m.OrganizationId
    LEFT JOIN Attend a ON m.MeetingId = a.MeetingId
    WHERE m.MeetingDate >= DATEADD(week, -12, GETDATE())
        AND o.OrganizationTypeId = 125  -- Connect Groups
        AND m.DidNotMeet = 0
    GROUP BY o.OrganizationId, o.OrganizationName, a.PeopleId
)
SELECT TOP 100
    OrganizationName,
    COUNT(DISTINCT PeopleId) as TotalMembers,
    AVG(CAST(TimesPresent as FLOAT) / NULLIF(MeetingsAttended, 0) * 100) as AvgAttendanceRate,
    SUM(CASE WHEN TimesPresent >= MeetingsAttended * 0.75 THEN 1 ELSE 0 END) as RegularAttenders,
    SUM(CASE WHEN TimesPresent < MeetingsAttended * 0.25 THEN 1 ELSE 0 END) as RareAttenders
FROM GroupAttendance
GROUP BY OrganizationId, OrganizationName
ORDER BY AvgAttendanceRate DESC'''
            },
            {
                'name': 'New Visitor Follow-up',
                'category': 'Visitors',
                'description': 'Find recent visitors needing follow-up',
                'sql': '''SELECT 
    p.PeopleId,
    p.Name,
    p.EmailAddress,
    p.CellPhone,
    p.CreatedDate as FirstVisit,
    DATEDIFF(day, p.CreatedDate, GETDATE()) as DaysSinceVisit,
    ms.Description as Status,
    MAX(a.MeetingDate) as LastAttendedDate,
    CASE 
        WHEN MAX(a.MeetingDate) IS NOT NULL 
        THEN DATEDIFF(day, MAX(a.MeetingDate), GETDATE())
        ELSE NULL 
    END as DaysSinceLastAttended
FROM People p
JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
LEFT JOIN Attend a ON p.PeopleId = a.PeopleId AND a.AttendanceFlag = 1
WHERE p.CreatedDate >= DATEADD(week, -4, GETDATE())
    AND p.MemberStatusId IN (50, 60)  -- Visitor/Guest status  
    AND p.IsDeceased = 0
    AND p.ArchivedFlag = 0
GROUP BY p.PeopleId, p.Name, p.EmailAddress, p.CellPhone, p.CreatedDate, ms.Description
ORDER BY p.CreatedDate DESC'''
            }
        ]
    
    def parse_sql_error(self, error_msg):
        """Parse SQL error messages and provide helpful suggestions"""
        error_str = str(error_msg)
        
        # Extract the key error information
        suggestions = []
        error_type = "SQL Error"
        
        # Invalid object name
        if "Invalid object name" in error_str:
            match = re.search(r"Invalid object name '([^']+)'", error_str)
            if match:
                table_name = match.group(1)
                error_type = "Invalid Table/View Name"
                
                # Check if it's a case sensitivity issue
                suggestions.append("Check if the table name is spelled correctly")
                if table_name.startswith('#'):
                    suggestions.append("Temp tables only last for one run. Run the statement that creates it in the same run")

                # Check for common table name mistakes
                if table_name.lower() == 'people1':
                    suggestions.append("Did you mean 'People' instead of 'people1'?")
                elif not '.' in table_name and not table_name.startswith('#'):
                    suggestions.append("If it is a lookup table, add its schema, for example lookup." + table_name)
                
                # Suggest using the schema browser
                suggestions.append("Use the Schema browser on the left to find the correct table name")
                
        # Invalid column name
        elif "Invalid column name" in error_str:
            match = re.search(r"Invalid column name '([^']+)'", error_str)
            if match:
                column_name = match.group(1)
                error_type = "Invalid Column Name"
                suggestions.append("Check if the column name is spelled correctly")
                suggestions.append("Make sure the column belongs to the table its alias points to")
                suggestions.append("Click on a table in the Schema browser to see available columns")
                
        # A variable that is neither declared nor given a value
        elif "Must declare the scalar variable" in error_str or "Must declare the table variable" in error_str:
            match = re.search(r'variable "@?([^"]+)"', error_str)
            name = match.group(1) if match else 'it'
            error_type = "Missing Parameter"
            if name.lower() == 'bluetoolbartagid':
                suggestions.append("Type a tag Id for @BlueToolbarTagId in the Parameters bar above the editor to test with")
            elif name in ('UserPeopleId', 'CurrentOrgId') or name.lower() not in ('userpeopleid', 'currentorgid'):
                suggestions.append("Type a value for @" + name + " in the Parameters bar above the editor, or DECLARE it in the query")
            else:
                suggestions.append("TouchPoint only fills this in when it is spelled exactly @UserPeopleId or @CurrentOrgId")
            suggestions.append("A variable DECLAREd in one run is gone by the next. Run the DECLARE and the query together")

        # Syntax errors
        elif "Incorrect syntax" in error_str:
            error_type = "SQL Syntax Error"
            suggestions.append("Check for missing commas between column names")
            suggestions.append("Ensure all quotes and parentheses are properly closed")
            suggestions.append("Verify JOIN conditions are properly specified")
            
        # Permission errors
        elif "permission" in error_str.lower() or "read-only" in error_str.lower() or "read only" in error_str.lower():
            error_type = "Permission Denied"
            suggestions.append("TouchPoint runs these queries with a read-only login, so changing data is not possible here")
            suggestions.append("Temp tables (#Name) and table variables (@Name) still work for building up results")
            
        # Ambiguous column
        elif "Ambiguous column name" in error_str:
            error_type = "Ambiguous Column"
            suggestions.append("When joining tables, prefix column names with table aliases")
            suggestions.append("Example: SELECT p.Name FROM People p")
            
        # Extract error number if available
        error_num_match = re.search(r"Error Number:(\d+)", error_str)
        if error_num_match:
            error_num = error_num_match.group(1)
            if error_num == "208":
                # Invalid object name
                pass  # Already handled above
            elif error_num == "207":
                # Invalid column name
                error_type = "Invalid Column Name"
                if not any("column name" in s for s in suggestions):
                    suggestions.append("The specified column does not exist in the table")
        
        return {
            'type': error_type,
            'suggestions': suggestions,
            'raw_error': error_str
        }
    
    def auto_alias_duplicate_columns(self, sql):
        """Automatically add aliases to duplicate column names in SELECT statement"""
        try:
            # Simple regex to find SELECT clause (handles basic cases)
            select_match = re.search(r'SELECT\s+(.*?)\s+FROM', sql, re.IGNORECASE | re.DOTALL)
            if not select_match:
                return sql
            
            select_clause = select_match.group(1)
            
            # If there are already aliases with AS keyword, don't modify
            # Check for functions and expressions
            if re.search(r'\w+\s*\([^)]*\)\s+AS\s+\w+', select_clause, re.IGNORECASE):
                # Already has proper aliases, return as-is
                return sql
            
            # Parse columns (simplified - handles basic cases)
            column_names = {}
            modified_columns = []
            changed = False
            
            # Split by comma but be careful with nested functions
            # Better splitting that handles nested parentheses
            parts = []
            current_part = ''
            paren_count = 0
            
            for char in select_clause:
                if char == '(':
                    paren_count += 1
                elif char == ')':
                    paren_count -= 1
                elif char == ',' and paren_count == 0:
                    parts.append(current_part)
                    current_part = ''
                    continue
                current_part += char
            
            if current_part:
                parts.append(current_part)
            
            for part in parts:
                part = part.strip()
                
                # Skip if already has an alias
                if re.search(r'\s+AS\s+\w+', part, re.IGNORECASE):
                    modified_columns.append(part)
                    continue
                
                # Extract simple column name (handle table.column format only)
                # Don't try to parse functions
                simple_col_match = re.match(r'^(\w+)\.(\w+)$', part)
                
                if simple_col_match:
                    table_alias = simple_col_match.group(1)
                    column_name = simple_col_match.group(2)
                    
                    # Check if column name already exists
                    if column_name in column_names:
                        # Add alias with table prefix or counter
                        column_names[column_name] += 1
                        new_alias = table_alias + '_' + column_name
                        modified_columns.append(part + ' AS ' + new_alias)
                        changed = True
                    else:
                        column_names[column_name] = 1
                        modified_columns.append(part)
                else:
                    # Not a simple column reference, keep as-is
                    modified_columns.append(part)
            
            # Leave the SQL exactly as typed unless an alias was added
            if not changed:
                return sql

            # Reconstruct SQL with aliased columns
            new_select_clause = ', '.join(modified_columns)
            new_sql = sql[:select_match.start(1)] + new_select_clause + sql[select_match.end(1):]
            
            return new_sql
            
        except:
            # If anything fails, return original SQL
            return sql
    
    def is_write_query(self, sql):
        """True when the SQL would change a permanent table.
        Temp tables (#Name) and table variables (@Name) are allowed."""
        s = re.sub(r'--[^\n]*', ' ', sql)
        s = re.sub(r'/\*.*?\*/', ' ', s, flags=re.DOTALL)
        s = re.sub(r"'(?:[^']|'')*'", "''", s)
        s = s.upper()
        patterns = [
            r'\bINSERT\s+(?:INTO\s+)?(?!INTO\b)(?![@#])[\w\[]',
            r'\bINTO\s+(?![@#])[\w\[]',
            r'\bUPDATE\s+(?:TOP\s*\([^)]*\)\s+)?(?!TOP\b)(?![@#])[\w\[]',
            r'\bDELETE\s+(?:TOP\s*\([^)]*\)\s+)?(?!TOP\b)(?:FROM\s+)?(?!FROM\b)(?![@#])[\w\[]',
            r'\bMERGE\s+(?:TOP\s*\([^)]*\)\s+)?(?!TOP\b)(?:INTO\s+)?(?!INTO\b)(?![@#])[\w\[]',
            r'\bTRUNCATE\s+TABLE\s+(?!#)',
            r'\bDROP\s+(?!TABLE\s+(?:IF\s+EXISTS\s+)?#)',
            r'\bCREATE\s+(?!TABLE\s+#)(?!(?:UNIQUE\s+)?(?:CLUSTERED\s+|NONCLUSTERED\s+)?INDEX\s+\w+\s+ON\s+#)',
            r'\bALTER\s+(?!TABLE\s+#)',
            r'\b(?:GRANT|REVOKE|DENY)\s',
            r'\bEXEC(?:UTE)?\s+(?:\w+\.)*(?:SP_|XP_)'
        ]
        for pattern in patterns:
            if re.search(pattern, s):
                return True
        return False

    def builds_temp_data(self, sql):
        """True when the SQL fills a temp table or table variable. SET ROWCOUNT
        would cut those short too, so the row limit is applied while reading."""
        s = re.sub(r'--[^\n]*', ' ', sql)
        s = re.sub(r'/\*.*?\*/', ' ', s, flags=re.DOTALL)
        return re.search(r'\b(INTO|INSERT|UPDATE|DELETE|MERGE)\b', s, re.IGNORECASE) is not None

    def guess_column_order(self, sql):
        """Fallback only: read column names from the first SELECT list."""
        column_order = []
        select_match = re.search(r'SELECT\s+(TOP\s+\d+\s+)?(.*?)\s+FROM', sql, re.IGNORECASE | re.DOTALL)
        if not select_match:
            return column_order
        select_clause = select_match.group(2).strip()
        if select_clause == '*':
            return column_order
        parts = []
        current = ''
        paren_depth = 0
        in_quotes = False
        quote_char = None
        for char in select_clause + ',':
            if char in ('"', "'", '[') and not in_quotes:
                in_quotes = True
                quote_char = ']' if char == '[' else char
                current += char
            elif in_quotes and char == quote_char:
                in_quotes = False
                current += char
            elif not in_quotes:
                if char == '(':
                    paren_depth += 1
                    current += char
                elif char == ')':
                    paren_depth -= 1
                    current += char
                elif char == ',' and paren_depth == 0:
                    parts.append(current.strip())
                    current = ''
                else:
                    current += char
            else:
                current += char
        for part in parts:
            if part:
                alias_match = re.search(r'\s+[Aa][Ss]\s+([^\s]+)$', part)
                if alias_match:
                    column_order.append(alias_match.group(1).strip('"\'[]'))
                else:
                    words = re.findall(r'[\w\[\]]+', part)
                    if words:
                        column_order.append(words[-1].strip('[]'))
        return column_order

    def row_pairs(self, row):
        """The row's (name, value) pairs in the order the query returned them,
        duplicate names included. None when the row cannot be read that way."""
        try:
            pairs = [(kv.Key, kv.Value) for kv in row]
        except Exception:
            return None
        if not pairs:
            return None
        for k, v in pairs:
            if not isinstance(k, basestring):
                return None
        return pairs

    def clean_value(self, value):
        """Make one value safe to send as JSON, keeping the full text."""
        if value is None:
            return None
        if isinstance(value, (bool, int, long, float)):
            return value
        tn = type(value).__name__
        try:
            if tn in ('DateTime', 'DateTimeOffset'):
                txt = value.ToString('yyyy-MM-dd HH:mm:ss')
                return txt[:-9] if txt.endswith(' 00:00:00') else txt
            if hasattr(value, 'strftime'):
                txt = value.strftime('%Y-%m-%d %H:%M:%S')
                return txt[:-9] if txt.endswith(' 00:00:00') else txt
            if tn.startswith('Array'):
                return '[binary, %d bytes]' % len(value)
        except Exception:
            pass
        try:
            txt = unicode(value)
        except Exception:
            try:
                txt = str(value).decode('utf-8', 'replace')
            except Exception:
                return '[unreadable value]'
        txt = txt.replace(u'\x00', u'')
        if len(txt) > MAX_VALUE_CHARS:
            txt = txt[:MAX_VALUE_CHARS] + u' [cut at 100,000 characters]'
        return txt

    def build_params(self, raw):
        """The page's parameter values as a dict for q.QuerySql, which sends
        them as real query parameters. Whole numbers go in as numbers and
        everything else as text; SQL Server converts text when it is compared
        with a date or decimal column. Empty means NULL."""
        if not raw:
            return {}
        try:
            data = json.loads(raw)
        except Exception:
            return {}
        if not isinstance(data, dict):
            return {}
        params = {}
        for k, v in data.items():
            k = unicode(k)
            if not re.match(r'^[A-Za-z_][A-Za-z0-9_]{0,127}$', k):
                continue
            # TouchPoint fills these itself and would overwrite ours anyway
            if k in ('UserPeopleId', 'CurrentOrgId'):
                continue
            v = u'' if v is None else unicode(v).strip()
            if v == u'':
                params[k] = None
            elif re.match(r'^-?[0-9]{1,9}$', v):
                params[k] = int(v)
            else:
                params[k] = v
        return params

    def execute_query(self, sql, max_rows=DEFAULT_MAX_ROWS, params=None):
        """Execute SQL and return columns plus rows (lists, in column order)"""
        if not self.has_permission:
            return {
                'success': False,
                'error': 'Access Denied: Admin or Developer role required'
            }

        if self.is_write_query(sql):
            return {
                'success': False,
                'error': 'This page only reads data. It will not change permanent tables.',
                'error_info': {
                    'type': 'Read Only',
                    'suggestions': [
                        'Temp tables (#Name) and table variables (@Name) work for building up results',
                        'Make data changes through TouchPoint itself or a Python script'
                    ],
                    'raw_error': 'Blocked before running: the SQL writes to a permanent table'
                }
            }

        try:
            max_rows = int(max_rows)
        except Exception:
            max_rows = DEFAULT_MAX_ROWS
        if max_rows <= 0 or max_rows > MAX_ROWS_ALL:
            max_rows = MAX_ROWS_ALL

        try:
            start_time = datetime.now()

            sql = self.auto_alias_duplicate_columns(sql)

            # Stop SQL Server itself after the limit (plus one row, to know there
            # were more). Kept on the first line so error line numbers still match.
            run_sql = sql
            if not self.builds_temp_data(sql):
                run_sql = 'SET ROWCOUNT %d; ' % (max_rows + 1) + sql

            try:
                if params:
                    results = q.QuerySql(run_sql, params)
                else:
                    results = q.QuerySql(run_sql)
            except Exception as query_error:
                try:
                    error_str = unicode(query_error)
                except Exception:
                    error_str = 'Database error (the message could not be read)'
                error_str = error_str.replace('SET ROWCOUNT %d; ' % (max_rows + 1), '')
                return {
                    'success': False,
                    'error': error_str,
                    'error_info': self.parse_sql_error(error_str)
                }

            columns = []
            rows = []
            use_pairs = None
            names = []
            truncated = False

            for row in results:
                if len(rows) >= max_rows:
                    truncated = True
                    break

                if use_pairs is None:
                    pairs = self.row_pairs(row)
                    available = [c for c in dir(row) if not c.startswith('_')]
                    if pairs and set(k for k, v in pairs) <= set(available):
                        use_pairs = True
                        seen = {}
                        for k, v in pairs:
                            if k in seen:
                                seen[k] += 1
                                columns.append('%s_%d' % (k, seen[k]))
                            else:
                                seen[k] = 1
                                columns.append(k)
                    else:
                        use_pairs = False
                        order = self.guess_column_order(sql)
                        remaining = list(available)
                        for col in order:
                            for avail in remaining:
                                if avail.lower() == col.lower():
                                    names.append(avail)
                                    remaining.remove(avail)
                                    break
                        names.extend(sorted(remaining))
                        columns = list(names)

                if use_pairs:
                    pairs = self.row_pairs(row) or []
                    rows.append([self.clean_value(v) for k, v in pairs])
                else:
                    values = []
                    for col in names:
                        try:
                            values.append(self.clean_value(getattr(row, col, None)))
                        except Exception:
                            values.append(None)
                    rows.append(values)

            exec_time = (datetime.now() - start_time).total_seconds() * 1000

            return {
                'success': True,
                'columns': columns,
                'rows': rows,
                'rowCount': len(rows),
                'truncated': truncated,
                'maxRows': max_rows,
                'orderKnown': bool(use_pairs),
                'executionTime': int(exec_time)
            }

        except Exception as e:
            try:
                error_msg = unicode(e)
            except Exception:
                error_msg = 'Database error occurred (the message could not be read)'
            return {
                'success': False,
                'error': error_msg,
                'error_info': self.parse_sql_error(error_msg)
            }

    def render_interface(self):
        """Render the complete interface with multi-tab support"""
        # Browser storage (tabs, history) is keyed by user so a shared
        # computer does not mix two people's work.
        print('<script>window.SQE_USER = "%s"; window.SQE_DEFAULT_ROWS = %d; window.SQE_MAX_ALL = %d;</script>'
              % (int(self.user_id or 0), DEFAULT_MAX_ROWS, MAX_ROWS_ALL))
        print("""
        <!-- CodeMirror CSS -->
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/theme/monokai.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/fold/foldgutter.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/show-hint.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/lint/lint.min.css">
        
        <!-- CodeMirror JS -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/mode/sql/sql.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/edit/matchbrackets.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/edit/closebrackets.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/show-hint.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/sql-hint.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/lint/lint.min.js"></script>
        
        <!-- CodeMirror Folding -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/fold/foldcode.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/fold/foldgutter.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/fold/brace-fold.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/fold/indent-fold.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/fold/comment-fold.min.js"></script>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/dialog/dialog.min.css">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/dialog/dialog.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/searchcursor.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/search.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/comment/comment.min.js"></script>
        
        <style>
            /* Layout */
            .query-explorer {
                display: flex;
                gap: 0;
                height: 80vh;
                position: relative;
                max-width: 1500px;
                margin: 0 auto;
            }
            
            .left-panel {
                width: 300px;
                min-width: 150px;
                overflow-y: auto;
                padding-right: 15px;
                position: relative;
                flex-shrink: 0;
                border-right: 1px solid #ddd;
            }
            
            .panel-resize-handle {
                position: absolute;
                right: -3px;
                top: 0;
                bottom: 0;
                width: 6px;
                background: transparent;
                cursor: col-resize;
                z-index: 100;
            }
            
            /* Panel toggle button */
            .panel-toggle {
                margin-left: auto;
                margin-right: 2px;
                flex-shrink: 0;
                align-self: center;
                width: 26px;
                height: 26px;
                background: transparent;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                cursor: pointer;
                padding: 0;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 18px;
                line-height: 1;
                color: #6b7280;
            }
            
            .panel-toggle:hover {
                background: #eef2ff;
                color: #2563eb;
                border-color: #2563eb;
            }
            
            /* Collapsed: a narrow strip that keeps only the toggle */
            .left-panel.collapsed {
                width: 40px !important;
                min-width: 40px !important;
                overflow: hidden;
                padding-right: 0;
            }
            
            .left-panel.collapsed .sqe-ltab,
            .left-panel.collapsed .sqe-pane {
                display: none !important;
            }
            
            .left-panel.collapsed .sqe-ltabs {
                border-bottom: none;
                justify-content: center;
                padding-top: 6px;
            }
            
            .left-panel.collapsed .panel-toggle {
                margin-left: 0;
            }
            
            .left-panel.collapsed .panel-resize-handle {
                display: none;
            }
            
            .panel-resize-handle:hover {
                background: #e0e0e0;
            }
            
            .panel-resize-handle.dragging {
                background: #2563eb;
            }
            
            .main-panel {
                flex: 1;
                display: flex;
                flex-direction: column;
                margin-left: 10px;
                overflow: hidden;
            }
            
            /* Editor Tabs */
            .editor-tabs {
                display: flex;
                border-bottom: 1px solid #e5e7eb;
                background: #f8f9fa;
                align-items: center;
                min-height: 36px;
            }
            
            .editor-tab {
                padding: 8px 16px;
                cursor: pointer;
                border: 1px solid transparent;
                border-bottom: none;
                background: #e5e7eb;
                margin-right: 2px;
                position: relative;
                display: flex;
                align-items: center;
                gap: 8px;
                max-width: 200px;
                user-select: none;
            }
            
            .editor-tab:hover {
                background: #d1d5db;
            }
            
            .editor-tab.active {
                background: white;
                border-color: #e5e7eb;
            }
            
            .editor-tab-name {
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
            }
            
            .editor-tab-close {
                font-size: 16px;
                line-height: 1;
                cursor: pointer;
                opacity: 0.6;
                padding: 0 4px;
            }
            
            .editor-tab-close:hover {
                opacity: 1;
                color: #dc2626;
            }
            
            .new-tab-btn {
                padding: 8px 12px;
                cursor: pointer;
                border: none;
                background: none;
                font-size: 18px;
                color: #6b7280;
            }
            
            .new-tab-btn:hover {
                color: #2563eb;
                background: #e5e7eb;
            }
            
            /* Schema Explorer */
            .schema-tree {
                font-size: 14px;
            }
            
            .schema-group {
                margin-bottom: 15px;
            }
            
            .schema-group-header {
                font-weight: bold;
                color: #374151;
                padding: 5px 0;
                cursor: pointer;
                user-select: none;
                display: flex;
                align-items: center;
                gap: 5px;
            }
            
            .schema-group-header:hover {
                color: #2563eb;
            }
            
            .schema-group-content {
                margin-left: 20px;
            }
            
            .expand-icon {
                font-size: 12px;
                transition: transform 0.2s;
            }
            
            .schema-group.collapsed .expand-icon {
                transform: rotate(-90deg);
            }
            
            .schema-group.collapsed .schema-group-content {
                display: none;
            }
            
            .table-item {
                cursor: pointer;
                padding: 5px;
                border-radius: 3px;
                position: relative;
            }
            
            .table-item:hover {
                background-color: #f0f0f0;
            }
            
            .table-actions {
                position: absolute;
                right: 5px;
                top: 50%;
                transform: translateY(-50%);
                display: none;
            }
            
            .table-item:hover .table-actions {
                display: inline-block;
            }
            
            .table-action-btn {
                background: #2563eb;
                color: white;
                border: none;
                padding: 2px 8px;
                border-radius: 3px;
                font-size: 11px;
                cursor: pointer;
            }
            
            .table-action-btn:hover {
                background: #1d4ed8;
            }
            
            .table-name {
                font-weight: 500;
                color: #1976d2;
                font-size: 14px;
            }
            
            .column-list {
                margin-left: 20px;
                display: none;
            }
            
            .column-item {
                padding: 3px 5px;
                font-size: 13px;
                color: #666;
            }
            
            .column-type {
                color: #999;
                font-size: 12px;
            }
            
            .key-badge {
                background: #fbbf24;
                color: #000;
                padding: 1px 4px;
                border-radius: 3px;
                font-size: 10px;
                margin-left: 5px;
            }
            
            /* SQL Editor */
            .sql-editor {
                display: none;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            
            .sql-editor.active {
                display: block;
            }
            
            /* CodeMirror Overrides */
            .CodeMirror {
                height: 300px;
                min-height: 150px;
                max-height: 70vh;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 14px;
                resize: vertical;
                overflow: auto;
            }
            
            /* Editor wrapper for resize handle */
            .editor-wrapper {
                position: relative;
                border: 1px solid #ddd;
                border-radius: 4px;
                overflow: hidden;
            }
            
            .editor-wrapper .resize-handle {
                position: absolute;
                bottom: 0;
                left: 0;
                right: 0;
                height: 5px;
                cursor: ns-resize;
                background: #f0f0f0;
                border-top: 1px solid #ddd;
            }
            
            .editor-wrapper .resize-handle:hover {
                background: #e0e0e0;
            }
            
            .CodeMirror-scroll {
                overflow-y: auto;
                overflow-x: auto;
            }
            .string { color: #008000; }
            .comment { color: #808080; font-style: italic; }
            .number { color: #FF0000; }
            .operator { color: #666666; }
            
            /* Results */
            .results-container {
                flex: 1;
                overflow: auto;
                margin-top: 20px;
                max-width: 100%;
            }
            
            /* Results wrapper for horizontal scroll */
            .results-wrapper {
                overflow: auto;
                max-height: 65vh;
                max-width: 100%;
                border: 1px solid #e5e7eb;
            }
            .results-table th { z-index: 2; white-space: nowrap; cursor: pointer; user-select: none; }
            .results-table th .sqe-arrow { color: #2563eb; margin-left: 4px; }
            .results-table tr.sqe-filter-row th { top: 37px; padding: 4px; cursor: default; background: #f9fafb; }
            .sqe-filter { width: 100%; min-width: 60px; font-size: 12px; padding: 2px 4px; border: 1px solid #d1d5db; border-radius: 3px; font-weight: normal; }
            .results-table td.sqe-rownum, .results-table th.sqe-rownum { color: #9ca3af; text-align: right; cursor: copy; font-size: 12px; width: 1%; }
            .results-table td.sqe-rownum:hover { color: #2563eb; }
            .results-table td { white-space: nowrap; max-width: 480px; overflow: hidden; text-overflow: ellipsis; }
            .results-table td.sqe-open { white-space: pre-wrap; max-width: 720px; overflow: visible; }
            .sqe-null { color: #9ca3af; font-style: italic; }
            .sqe-more { color: #2563eb; cursor: pointer; font-size: 12px; margin-left: 4px; }
            .sqe-flash { background: #dcfce7 !important; }
            .sqe-info { display: flex; flex-wrap: wrap; gap: 8px 16px; align-items: center; padding: 8px 10px; background: #f3f4f6; font-size: 13px; }
            .sqe-info .sqe-warn { color: #92400e; background: #fef3c7; padding: 2px 8px; border-radius: 4px; }
            .sqe-info .sqe-hint { color: #6b7280; font-size: 12px; }
            .sqe-info .sqe-actions { display: flex; gap: 6px; margin-left: auto; }
            .sqe-params { display: none; flex-wrap: wrap; gap: 8px 14px; align-items: center; padding: 8px 10px; margin-bottom: 8px; background: #eef2ff; border: 1px solid #c7d2fe; border-radius: 6px; font-size: 13px; }
            .sqe-params-title { font-weight: 600; color: #3730a3; }
            .sqe-param { display: inline-flex; align-items: center; gap: 6px; margin: 0; font-weight: normal; font-family: monospace; }
            .sqe-param input { font-family: inherit; font-size: 13px; padding: 3px 6px; border: 1px solid #a5b4fc; border-radius: 4px; width: 170px; }
            .sqe-params .sqe-hint { color: #4b5563; font-size: 12px; }
            .results-table tfoot td { position: sticky; bottom: 0; background: #f3f4f6; font-size: 12px; color: #374151; border-top: 2px solid #d1d5db; z-index: 1; white-space: nowrap; }
            .sqe-prof-btn { color: #9ca3af; margin-left: 6px; font-weight: normal; cursor: pointer; }
            .sqe-prof-btn:hover { color: #2563eb; }
            .sqe-panel { border: 1px solid #e5e7eb; border-radius: 6px; padding: 10px 12px; margin-bottom: 10px; background: #fff; }
            .sqe-panel-head { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 8px; }
            .sqe-panel-close { margin-left: auto; color: #6b7280; cursor: pointer; font-size: 12px; }
            .sqe-panel-close:hover { color: #2563eb; }
            .sqe-prof-stats { display: flex; flex-wrap: wrap; gap: 6px 18px; font-size: 13px; margin-bottom: 6px; }
            .sqe-prof-row { position: relative; display: flex; justify-content: space-between; gap: 10px; padding: 3px 6px; font-size: 13px; cursor: pointer; max-width: 640px; }
            .sqe-prof-row:hover { outline: 1px solid #2563eb; }
            .sqe-prof-bar { position: absolute; left: 0; top: 0; bottom: 0; background: #dbeafe; z-index: 0; }
            .sqe-prof-val, .sqe-prof-n { position: relative; z-index: 1; }
            .sqe-prof-val { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
            .sqe-chart-box { position: relative; height: 320px; max-width: 100%; }
            .sqe-colgroup { margin-bottom: 10px; }
            .sqe-colgroup-head { font-weight: 600; font-size: 13px; padding: 4px 0; color: #374151; }
            .sqe-colhit { font-size: 13px; padding: 3px 4px; border-bottom: 1px solid #f3f4f6; }
            .sqe-colhit-col { font-weight: 600; cursor: pointer; }
            .sqe-colhit-tbl { cursor: pointer; color: #2563eb; }
            .sqe-colhit-col:hover, .sqe-colhit-tbl:hover { text-decoration: underline; }
            .sqe-colhit-in { color: #9ca3af; }
            .sqe-mini-btn { font-size: 12px; padding: 3px 10px; border: 1px solid #d1d5db; background: #fff; border-radius: 4px; cursor: pointer; }
            .sqe-mini-btn:hover { background: #eef2ff; }
            .sqe-rows { font-size: 13px; padding: 6px; border: 1px solid #d1d5db; border-radius: 4px; background: #fff; }
            .sqe-banner { padding: 12px 14px; background: #fef3c7; border: 1px solid #f59e0b; border-radius: 6px; margin: 10px 0; }
            .sqe-banner button { margin: 8px 8px 0 0; }
            .sqe-note { font-size: 12px; color: #6b7280; padding: 4px 8px 8px; }
            .sqe-item-actions { float: right; font-size: 12px; }
            .sqe-item-actions a { color: #6b7280; cursor: pointer; margin-left: 6px; }
            .sqe-item-actions a:hover { color: #2563eb; }
            .sqe-hist { padding: 6px 8px; border-bottom: 1px solid #f3f4f6; cursor: pointer; font-size: 12px; }
            .sqe-hist:hover { background: #eef2ff; }
            .sqe-hist .sqe-hist-sql { font-family: monospace; color: #111827; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
            .sqe-hist .sqe-hist-meta { color: #6b7280; }
            .sqe-hist .sqe-hist-err { color: #b91c1c; }
            
            .results-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 14px;
            }
            
            .results-table th {
                background: #f3f4f6;
                padding: 8px;
                text-align: left;
                border: 1px solid #e5e7eb;
                position: sticky;
                top: 0;
            }
            
            .results-table td {
                padding: 6px 8px;
                border: 1px solid #e5e7eb;
            }
            
            .results-table tr:hover {
                background: #f9fafb;
            }
            
            /* Left panel tabs */
            .sqe-ltabs {
                display: flex;
                align-items: flex-end;
                min-height: 37px;
                border-bottom: 1px solid var(--sqe-border);
                margin-bottom: 10px;
            }
            .sqe-ltab {
                padding: 9px 8px 8px;
                font-size: 13px;
                color: #4b5563;
                white-space: nowrap;
                flex: 0 1 auto;
                min-width: 0;
                overflow: hidden;
                text-overflow: ellipsis;
                cursor: pointer;
                border-bottom: 2px solid transparent;
                margin-bottom: -1px;
            }
            .sqe-ltab:hover { color: var(--sqe-accent); }
            .sqe-ltab.active {
                color: var(--sqe-accent);
                border-bottom-color: var(--sqe-accent);
                font-weight: 600;
            }
            .sqe-ltabs .panel-toggle { align-self: center; }

            /* Toolbar and buttons (not .btn: TouchPoint's Bootstrap owns that name) */
            .sqe-toolbar {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 8px;
                margin: 10px 0;
            }
            .sqe-spacer { flex: 1; }
            .sqe-btn {
                display: inline-flex;
                align-items: center;
                gap: 6px;
                height: 32px;
                padding: 0 12px;
                font-size: 13px;
                line-height: 1;
                color: #374151;
                background: #fff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                cursor: pointer;
                white-space: nowrap;
            }
            .sqe-btn:hover { background: var(--sqe-soft); border-color: #9ca3af; }
            .sqe-btn-primary {
                background: var(--sqe-accent);
                border-color: var(--sqe-accent);
                color: #fff;
                font-weight: 600;
            }
            .sqe-btn-primary:hover { background: var(--sqe-accent-dark); border-color: var(--sqe-accent-dark); }
            .sqe-split { display: inline-flex; }
            .sqe-split .sqe-btn:first-child { border-radius: 6px 0 0 6px; }
            .sqe-split .sqe-btn + .sqe-btn {
                border-radius: 0 6px 6px 0;
                border-left: 1px solid rgba(255, 255, 255, 0.35);
                font-weight: 500;
            }
            .sqe-kbd { font-size: 11px; font-weight: 400; opacity: 0.75; }
            .sqe-btn-icon {
                width: 32px;
                padding: 0;
                justify-content: center;
                border-radius: 50%;
                color: var(--sqe-muted);
                font-weight: 600;
            }
            .query-explorer select.sqe-rows {
                height: 32px;
                width: auto;
                padding: 0 8px;
                font-size: 13px;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background: #fff;
            }

            /* Loading */
            .loading {
                text-align: center;
                padding: 20px;
                color: #6b7280;
            }
            
            /* Query History */
            .query-item {
                padding: 10px;
                border-bottom: 1px solid #e5e7eb;
                cursor: pointer;
                font-size: 13px;
            }
            
            .query-item:hover {
                background: #f9fafb;
            }
            
            .query-preview {
                color: #6b7280;
                font-size: 12px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                margin-top: 4px;
            }
            
            .keyboard-shortcuts {
                color: #6b7280;
                font-size: 12px;
                margin-left: 20px;
            }
            
            /* Alerts */
            .alert {
                padding: 10px 15px;
                border-radius: 4px;
                margin: 10px 0;
            }
            
            .alert-success {
                background-color: #d4edda;
                border: 1px solid #c3e6cb;
                color: #155724;
            }
            
            .alert-danger {
                background-color: #f8d7da;
                border: 1px solid #f5c6cb;
                color: #721c24;
            }
            
            .alert-warning {
                background-color: #fff3cd;
                border: 1px solid #ffeaa7;
                color: #856404;
            }
            
            /* SQL Error Display Styles */
            .error-container {
                background: #fff5f5;
                border: 1px solid #feb2b2;
                border-radius: 6px;
                padding: 16px;
                margin: 10px 0;
            }
            
            .error-header {
                display: flex;
                align-items: center;
                margin-bottom: 12px;
            }
            
            .error-type {
                background: #e53e3e;
                color: white;
                padding: 4px 8px;
                border-radius: 4px;
                font-size: 12px;
                font-weight: bold;
                margin-right: 10px;
            }
            
            .error-suggestions {
                background: #fef2e5;
                border: 1px solid #fcd9a6;
                border-radius: 4px;
                padding: 12px;
                margin: 10px 0;
            }
            
            .error-suggestions h4 {
                color: #c05621;
                margin: 0 0 8px 0;
                font-size: 14px;
            }
            
            .error-suggestions ul {
                margin: 0;
                padding-left: 20px;
            }
            
            .error-suggestions li {
                color: #7c3927;
                margin: 4px 0;
            }
            
            .error-details {
                background: #f7fafc;
                border: 1px solid #e2e8f0;
                border-radius: 4px;
                padding: 8px 12px;
                margin-top: 10px;
            }
            
            .error-details strong {
                color: #2d3748;
            }
            
            .error-raw {
                font-family: monospace;
                font-size: 12px;
                color: #718096;
                white-space: pre-wrap;
                word-wrap: break-word;
                background: #edf2f7;
                padding: 8px;
                border-radius: 4px;
                margin-top: 8px;
            }
            
            /* Enhanced Autocomplete Styles */
            .CodeMirror-hints {
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 12px;
                max-height: 300px;
                background: #fff;
                border: 1px solid #d0d0d0;
                box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            }
            
            .CodeMirror-hint {
                padding: 4px 8px;
                white-space: pre;
                cursor: pointer;
            }
            
            .CodeMirror-hint-active {
                background: #0066cc;
                color: white;
            }
            
            .sql-table-hint {
                color: #0066cc;
                font-weight: bold;
            }
            
            .sql-column-hint {
                color: #008800;
            }
            
            .sql-keyword-hint {
                color: #cc00cc;
                text-transform: uppercase;
            }
            
            /* Linting Styles */
            .CodeMirror-lint-tooltip {
                background: #ffffcc;
                border: 1px solid #888;
                border-radius: 4px;
                color: #333;
                font-size: 12px;
                padding: 4px 8px;
                white-space: pre-wrap;
                max-width: 400px;
                z-index: 10000 !important;
            }
            
            .CodeMirror-lint-mark-error {
                border-bottom: 2px solid #ff0000;
            }
            
            .CodeMirror-lint-mark-warning {
                border-bottom: 2px solid #ffa500;
            }
            
            .CodeMirror-lint-mark-info {
                border-bottom: 1px dotted #0066cc;
            }
            
            .CodeMirror-lint-marker-error {
                color: #ff0000;
            }
            
            .CodeMirror-lint-marker-warning {
                color: #ffa500;
            }
            
            .CodeMirror-lint-marker-info {
                color: #0066cc;
            }
            
            .CodeMirror-lint-markers {
                width: 16px;
            }

            /* ---- Look (2.1) ---- */
            .query-explorer {
                --sqe-accent: #2563eb;
                --sqe-accent-dark: #1d4ed8;
                --sqe-border: #e5e7eb;
                --sqe-muted: #6b7280;
                --sqe-soft: #f8fafc;
                font-size: 14px;
                min-height: 520px;
            }
            .left-panel { padding-right: 12px; border-right-color: var(--sqe-border); }
            .sqe-pane { padding: 0; border: none; }
            .sqe-search {
                display: block;
                width: 100%;
                height: 32px;
                padding: 0 10px;
                font-size: 13px;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                margin-bottom: 10px;
                box-sizing: border-box;
            }
            .sqe-search:focus, .sqe-param input:focus, .sqe-filter:focus {
                outline: none;
                border-color: var(--sqe-accent);
                box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.15);
            }
            .schema-group { margin-bottom: 8px; }
            .schema-group-header { font-size: 13px; padding: 4px 0; }
            .schema-group-content { margin-left: 10px; }
            .table-item { padding: 3px 6px; font-size: 13px; }
            .table-name { font-size: 13px; font-weight: 400; color: var(--sqe-accent); overflow-wrap: anywhere; }

            .editor-tabs {
                background: transparent;
                border-bottom: 1px solid var(--sqe-border);
                min-height: 37px;
                align-items: flex-end;
                gap: 2px;
            }
            .editor-tab {
                padding: 7px 8px 7px 14px;
                background: #f3f4f6;
                border: 1px solid var(--sqe-border);
                border-bottom: none;
                border-radius: 6px 6px 0 0;
                margin: 0 0 -1px 0;
                font-size: 13px;
                color: #4b5563;
            }
            .editor-tab:hover { background: #e5e7eb; }
            .editor-tab.active { background: #fff; color: #111827; box-shadow: inset 0 2px 0 var(--sqe-accent); }
            .new-tab-btn { padding: 4px 10px; border-radius: 6px; align-self: center; }

            .sql-editor { border: none; }
            .editor-wrapper { border: 1px solid #d1d5db; border-radius: 6px; }
            .CodeMirror { resize: none; font-size: 13px; }
            .editor-wrapper .resize-handle { height: 6px; background: #f3f4f6; border-top: 1px solid var(--sqe-border); }
            .editor-wrapper .resize-handle:hover { background: #dbeafe; }
            .sqe-params { border-radius: 6px; }

            /* Results fill the rest of the window; the grid scrolls inside */
            .results-container { margin-top: 12px; display: flex; flex-direction: column; min-height: 0; overflow: hidden; }
            #results { flex: 1; min-height: 0; display: flex; flex-direction: column; overflow: auto; }
            #results > .results-wrapper { flex: 1 1 auto; min-height: 120px; max-height: none; }
            #sqe-side { flex: 0 0 auto; max-height: 45%; overflow: auto; }
            #sqe-side:empty { display: none; }
            .sqe-info { border: 1px solid var(--sqe-border); border-radius: 6px; background: var(--sqe-soft); margin-bottom: 6px; }
            .results-wrapper { border-color: var(--sqe-border); border-radius: 6px; }
            .results-table { font-size: 13px; }
            .results-table th { background: #f3f4f6; color: #374151; font-weight: 600; padding: 7px 8px; border-color: var(--sqe-border); }
            .results-table td { padding: 5px 8px; border-color: #eef0f3; }
            .results-table tbody tr:nth-child(even) { background: #fafbfc; }
            .results-table tbody tr:hover { background: #eef2ff; }
            .sqe-mini-btn {
                height: 26px;
                padding: 0 10px;
                font-size: 12px;
                color: #374151;
                background: #fff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                cursor: pointer;
            }
            .sqe-mini-btn:hover { background: #eef2ff; border-color: var(--sqe-accent); color: var(--sqe-accent); }

            .error-container { background: #fff; border: 1px solid #fecaca; border-left: 4px solid #dc2626; border-radius: 6px; padding: 12px 14px; }
            .error-type { background: #fee2e2; color: #991b1b; }
            .error-suggestions { background: var(--sqe-soft); border: 1px solid var(--sqe-border); }
            .error-suggestions li { color: #374151; }
            .error-details { background: transparent; border: none; padding: 0; }
            .error-raw { background: #f3f4f6; color: #4b5563; }

            .sqe-modal { display: none; position: fixed; inset: 0; background: rgba(17, 24, 39, 0.45); z-index: 1000; }
            .sqe-modal-box {
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                width: min(520px, calc(100vw - 32px));
                max-height: 80vh;
                overflow-y: auto;
                background: #fff;
                padding: 18px 20px;
                border-radius: 10px;
                box-shadow: 0 20px 40px rgba(0, 0, 0, 0.2);
                font-size: 13px;
            }
            .sqe-modal-box h3 { margin: 0 0 10px; font-size: 16px; }
            .sqe-modal-box table { width: 100%; border-collapse: collapse; }
            .sqe-modal-box td { vertical-align: top; }
        </style>
        
        <div class="query-explorer">
            <!-- Left Panel -->
            <div class="left-panel" id="left-panel">
                <div class="sqe-ltabs">
                    <div class="sqe-ltab active" onclick="showTab('schema')">Schema</div>
                    <div class="sqe-ltab" onclick="showTab('saved')">Saved</div>
                    <div class="sqe-ltab" onclick="showTab('history')">History</div>
                    <div class="sqe-ltab" onclick="showTab('examples')">Examples</div>
                    <button id="panel-toggle" class="panel-toggle" onclick="togglePanel()" title="Hide this panel">
                        <span id="toggle-icon">&#8249;</span>
                    </button>
                </div>
                
                <div id="schema-tab" class="sqe-pane">
                    <input type="text" id="schema-search" placeholder="Search tables or columns..." 
                           class="sqe-search"
                           onkeyup="filterSchema()">
                    <div id="column-matches"></div>
                    <div id="schema-tree" class="schema-tree">
                        <div class="loading">Loading schema...</div>
                    </div>
                </div>
                
                <div id="saved-tab" class="sqe-pane" style="display:none;">
                    <input type="text" id="saved-search" placeholder="Search saved queries..." 
                           class="sqe-search"
                           onkeyup="filterSaved()">
                    <div id="saved-queries">
                        <div class="loading">Loading saved queries...</div>
                    </div>
                </div>
                
                <div id="history-tab" class="sqe-pane" style="display:none;">
                    <input type="text" id="history-search" placeholder="Search history..."
                           class="sqe-search"
                           onkeyup="renderHistory()">
                    <div class="sqe-note">The last 50 queries you ran, kept in this browser only. Click one to open it in a new tab.
                        <a style="cursor:pointer;color:#2563eb;" onclick="clearHistory(this)">Clear</a></div>
                    <div id="history-list"></div>
                </div>

                <div id="examples-tab" class="sqe-pane" style="display:none;">
                    <div id="example-queries"></div>
                </div>
                
                <!-- Resize Handle -->
                <div class="panel-resize-handle" id="panel-resize-handle"></div>
            </div>
            
            
            <!-- Main Panel -->
            <div class="main-panel">
                <div class="editor-tabs" id="editor-tabs">
                    <div class="editor-tab active" data-tab-id="1" onclick="switchTab(1)" ondblclick="renameTab(1)">
                        <span class="editor-tab-name">Query 1</span>
                        <span class="editor-tab-close" onclick="closeTab(1, event)" title="Close tab (Ctrl+W)">&times;</span>
                    </div>
                    <button class="new-tab-btn" onclick="newTab()" title="New Query Tab (Ctrl+T)">+</button>
                </div>
                
                <div class="sqe-toolbar">
                    <span class="sqe-split">
                        <button class="sqe-btn sqe-btn-primary" onclick="executeQuery()" title="Run the query, or just the selected text (Ctrl+Enter)">
                            &#9654; Run <span class="sqe-kbd">Ctrl+Enter</span>
                        </button>
                        <button class="sqe-btn sqe-btn-primary" onclick="executeSelected()" title="Run only the statement the cursor is in (Ctrl+Shift+Enter)">
                            Statement
                        </button>
                    </span>
                    <select id="sqe-rows" class="sqe-rows" title="Stop after this many rows">
                        <option value="1000">1,000 rows</option>
                        <option value="10000">10,000 rows</option>
                        <option value="0">All rows (up to 50,000)</option>
                    </select>
                    <span class="sqe-spacer"></span>
                    <button class="sqe-btn" onclick="formatSQL()" title="Format the SQL, or just the selected part">Format</button>
                    <button class="sqe-btn" onclick="saveQuery()" title="Save to My queries">Save</button>
                    <button class="sqe-btn sqe-btn-icon" onclick="showKeyboardShortcuts()" title="Keyboard shortcuts">?</button>
                </div>
                
                <div id="sqe-params" class="sqe-params"></div>

                <div id="editors-container">
                    <div id="sql-editor-1" class="sql-editor active" data-tab-id="1">
                        <div class="editor-wrapper">
                            <textarea id="sql-editor-textarea-1">SELECT TOP 10 
    p.Name2 as FullName,
    p.EmailAddress,
    p.JoinDate,
    ms.Description as MemberStatus
FROM People p
JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
WHERE p.MemberStatusId = 10
ORDER BY p.JoinDate DESC</textarea>
                            <div class="resize-handle" onmousedown="initResize(event)"></div>
                        </div>
                    </div>
                </div>
                
                <div class="results-container">
                    <div id="results-info" class="sqe-info" style="display: none;"></div>
                    <div id="results"></div>
                </div>
            </div>
        </div>
        
        <!-- Keyboard Shortcuts Modal -->
        <div id="shortcuts-modal" class="sqe-modal">
            <div class="sqe-modal-box">
                <h3>Keyboard Shortcuts</h3>
                <table>
                    <tr><td style="padding: 5px;"><strong>Query Execution:</strong></td><td></td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Enter or F5</td><td>Execute Query (or selected text if any)</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Shift+Enter</td><td>Run only the statement the cursor is in</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+F / Ctrl+H</td><td>Find / Find and replace (Cmd+F, Cmd+Alt+F on a Mac)</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+/</td><td>Comment selected lines on or off</td></tr>
                    <tr><td style="padding: 5px;"><strong>Tab Management:</strong></td><td></td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+T</td><td>New Tab</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+W</td><td>Close Current Tab</td></tr>
                    <tr><td style="padding: 5px 20px;">Alt+Tab</td><td>Next Tab</td></tr>
                    <tr><td style="padding: 5px 20px;">Alt+Shift+Tab</td><td>Previous Tab</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+1 to Ctrl+9</td><td>Go to Tab 1-9</td></tr>
                    <tr><td style="padding: 5px 20px;">Double-click Tab</td><td>Rename Tab</td></tr>
                    <tr><td style="padding: 5px;"><strong>Editor:</strong></td><td></td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Space</td><td>Autocomplete</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Z / Cmd+Z</td><td>Undo</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Y / Cmd+Y</td><td>Redo</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Shift+Z</td><td>Redo (alternative)</td></tr>
                    <tr><td style="padding: 5px 20px;">Tab</td><td>Indent / Indent Selection</td></tr>
                    <tr><td style="padding: 5px 20px;">Shift+Tab</td><td>Outdent / Outdent Selection</td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Q</td><td>Toggle Code Folding at Cursor</td></tr>
                    <tr><td style="padding: 5px 20px;">Click fold icon</td><td>Fold/Unfold Code Block</td></tr>
                </table>
                <button class="sqe-btn" onclick="document.getElementById('shortcuts-modal').style.display='none'" style="margin-top: 15px;">Close</button>
            </div>
        </div>
        
        <script>
            // Dynamic schema tracking for autocomplete
            let dynamicSchema = {
                tables: {},
                views: {},
                loaded: false,
                loading: false
            };
            
            // Load complete schema on initialization
            function loadFullSchema() {
                if (dynamicSchema.loading || dynamicSchema.loaded) return;
                
                dynamicSchema.loading = true;
                
                // Convert /PyScript/ to /PyScriptForm/ for AJAX requests
                const url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                console.log('Loading full schema from:', url);
                
                const formData = new FormData();
                formData.append('action', 'get_full_schema');
                
                fetch(url, {
                    method: 'POST',
                    body: formData
                })
                .then(response => {
                    // Check if response is OK
                    if (!response.ok) {
                        throw new Error(`HTTP error! status: ${response.status}`);
                    }
                    // TouchPoint might not set correct content-type, so let's try to parse as JSON first
                    return response.text();
                })
                .then(text => {
                    // Try to parse as JSON
                    try {
                        const data = JSON.parse(text);
                        return data;
                    } catch (e) {
                        // If it's not JSON, check if it looks like HTML
                        if (text.trim().startsWith('<') || text.includes('<!DOCTYPE')) {
                            console.error('HTML response from server:', text.substring(0, 200));
                            throw new Error("Server returned HTML instead of JSON.");
                        } else {
                            console.error('Invalid response from server:', text.substring(0, 200));
                            throw new Error("Invalid response format from server.");
                        }
                    }
                })
                .then(data => {
                    // Check if we have schema data (with or without success flag)
                    const schemaData = data.schema || data;
                    if (schemaData && (schemaData.tables || schemaData.views)) {
                        // Store the complete schema
                        dynamicSchema.tables = schemaData.tables || {};
                        dynamicSchema.views = schemaData.views || {};
                        dynamicSchema.loaded = true;
                        
                        console.log('Schema successfully parsed:', {
                            hasTables: !!schemaData.tables,
                            hasViews: !!schemaData.views,
                            tableCount: Object.keys(schemaData.tables || {}).length,
                            viewCount: Object.keys(schemaData.views || {}).length
                        });
                        
                        // Enhanced logging to debug the issue
                        const sampleTables = ['ZipCodes', 'lookup.MemberStatus', 'People'];
                        const sampleInfo = {};
                        sampleTables.forEach(name => {
                            // Check different variations
                            const found = dynamicSchema.tables[name] || 
                                         dynamicSchema.tables[name.toLowerCase()] ||
                                         Object.keys(dynamicSchema.tables).find(k => k.toLowerCase() === name.toLowerCase());
                            if (found) {
                                const table = dynamicSchema.tables[found] || dynamicSchema.tables[name];
                                sampleInfo[name] = {
                                    found: true,
                                    storedAs: found,
                                    columnCount: table.columns ? table.columns.length : 0,
                                    firstColumns: table.columns ? table.columns.slice(0, 3) : []
                                };
                            } else {
                                sampleInfo[name] = { found: false };
                            }
                        });
                        
                        console.log('Full schema loaded:', {
                            tableCount: Object.keys(dynamicSchema.tables).length,
                            viewCount: Object.keys(dynamicSchema.views).length,
                            lookupTables: Object.keys(dynamicSchema.tables).filter(t => t.startsWith('lookup.')),
                            sampleTableInfo: sampleInfo
                        });
                    } else {
                        console.warn('Schema data not in expected format:', data);
                    }
                })
                .catch(error => {
                    console.error('Failed to load full schema:', error);
                    console.error('URL attempted:', url);
                    console.error('Current location:', window.location.pathname);
                    // Fall back to loading schema piece by piece
                    dynamicSchema.loaded = false;
                })
                .finally(() => {
                    dynamicSchema.loading = false;
                });
            }
            
            // Enhanced TouchPoint SQL Schema for Autocomplete (initial defaults)
            const touchpointSchema = {
                tables: {
                    // Core People Tables
                    "People": {
                        columns: ["PeopleId", "TitleCode", "FirstName", "PreferredName", "MiddleName", "LastName", 
                                 "SuffixCode", "NickName", "Name", "Name2", "AltName", "MaidenName",
                                 "EmailAddress", "EmailAddress2", "SendEmailAddress1", "SendEmailAddress2",
                                 "CellPhone", "WorkPhone", "HomePhone", "PrimaryAddress", "PrimaryCity",
                                 "PrimaryState", "PrimaryZip", "PrimaryCountry", "BirthDate", "BirthDay",
                                 "BirthMonth", "BirthYear", "DeceasedDate", "Age", "GenderId", "MaritalStatusId",
                                 "FamilyId", "PositionInFamilyId", "MemberStatusId", "CampusId", "JoinDate",
                                 "DecisionDate", "BaptismDate", "WeddingDate", "EnvelopeOptionsId",
                                 "ContributionOptionsId", "ElectronicStatement", "IsDeceased", "ArchivedFlag",
                                 "CreatedDate", "ModifiedDate", "LastAttended", "PictureId", "Comments",
                                 "Grade", "CheckInNotes", "Bio", "IsBusiness", "CustodyIssue", "OkTransport",
                                 "ReceiveSMS", "DoNotPublishPhones", "DoNotCallFlag", "DoNotMailFlag",
                                 "DoNotVisitFlag", "BadAddressFlag", "BlockMobileAccess", "SpouseId"],
                        schema: "dbo"
                    },
                    "Families": {
                        columns: ["FamilyId", "FamilyName", "HeadOfHouseholdId", "HeadOfHouseholdSpouseId",
                                 "AddressLineOne", "AddressLineTwo", "CityName", "StateCode", "ZipCode",
                                 "CountryCode", "HomePhone", "CreatedDate", "ModifiedDate"],
                        schema: "dbo"
                    },
                    "Organizations": {
                        columns: ["OrganizationId", "OrganizationName", "LeaderId", "LeaderMemberTypeId",
                                 "DivisionId", "ProgramId", "CampusId", "OrganizationStatusId", "SecurityTypeId",
                                 "LimitToRole", "Location", "Description", "CreatedDate", "ModifiedDate",
                                 "RegistrationTypeId", "RegistrationClosed", "ClassFilled", "AppCategory"],
                        schema: "dbo"
                    },
                    "OrganizationMembers": {
                        columns: ["OrganizationId", "PeopleId", "MemberTypeId", "EnrollmentDate", 
                                 "InactiveDate", "Pending", "Request", "Grade", "ShirtSize", "Tickets",
                                 "RegisterEmail", "Score", "LastAttended", "UserData"],
                        schema: "dbo"
                    },
                    "Attend": {
                        columns: ["PeopleId", "MeetingId", "OrganizationId", "MeetingDate", "AttendanceFlag",
                                 "AttendanceTypeId", "OtherOrgId", "OtherAttends", "MemberTypeId",
                                 "EffAttendFlag", "SeqNo", "Pager"],
                        schema: "dbo"
                    },
                    "Meetings": {
                        columns: ["MeetingId", "OrganizationId", "MeetingDate", "MeetingName", "Location",
                                 "Description", "NumPresent", "NumMembers", "NumVstMembers", "NumRepeatVst",
                                 "NumNewVst", "HeadCount", "MaxCount", "DidNotMeet", "CreatedDate"],
                        schema: "dbo"
                    },
                    "Contribution": {
                        columns: ["ContributionId", "PeopleId", "ContributionDate", "ContributionAmount",
                                 "ContributionDesc", "ContributionStatusId", "ContributionTypeId", "FundId",
                                 "CheckNo", "BankAccount", "CreatedDate", "ModifiedDate", "PostingDate",
                                 "Origin", "Source", "Campus", "ExtraDataId"],
                        schema: "dbo"
                    },
                    "ContributionFund": {
                        columns: ["FundId", "FundName", "FundDescription", "FundStatusId", "FundTypeId",
                                 "FundPledgeFlag", "FundAccountCode", "OnlineSort", "NonTaxDeductible",
                                 "CanPledge", "ShowList"],
                        schema: "dbo"
                    },
                    "EmailQueue": {
                        columns: ["Id", "PeopleId", "ToAddr", "FromAddr", "FromName", "Subject", "Body",
                                 "QueuedBy", "Queued", "Sent", "SendWhen", "Transactional", "Error",
                                 "Redacted", "Retry", "CClist", "BCClist"],
                        schema: "dbo"
                    },
                    "Tag": {
                        columns: ["Id", "Name", "TypeId", "Owner", "Created", "Modified", "Active",
                                 "DisplayOrder", "Description"],
                        schema: "dbo"
                    },
                    "TagPerson": {
                        columns: ["Id", "PeopleId", "DateCreated", "LastModified"],
                        schema: "dbo"
                    },
                    "ActivityLog": {
                        columns: ["Id", "ActivityDate", "UserId", "Activity", "PageUrl", "Machine",
                                 "OrgId", "PeopleId", "DatumId", "ClientIp"],
                        schema: "dbo"
                    },
                    "lookup.MemberStatus": {
                        columns: ["Id", "Code", "Description", "Hardwired"],
                        schema: "lookup",
                        name: "MemberStatus"
                    },
                    "lookup.Campus": {
                        columns: ["Id", "Code", "Description", "Hardwired"],
                        schema: "lookup",
                        name: "Campus"
                    },
                    "lookup.MemberType": {
                        columns: ["Id", "Code", "Description", "AttendanceTypeId", "Hardwired"],
                        schema: "lookup",
                        name: "MemberType"
                    },
                    "lookup.AttendType": {
                        columns: ["Id", "Code", "Description", "Hardwired"],
                        schema: "lookup",
                        name: "AttendType"
                    }
                },
                keywords: [
                    "SELECT", "FROM", "WHERE", "JOIN", "LEFT", "RIGHT", "INNER", "OUTER", "FULL",
                    "ON", "AS", "AND", "OR", "NOT", "IN", "EXISTS", "BETWEEN", "LIKE", "IS", "NULL",
                    "GROUP", "BY", "HAVING", "ORDER", "ASC", "DESC", "TOP", "DISTINCT", "ALL",
                    "UNION", "EXCEPT", "INTERSECT", "WITH", "CASE", "WHEN", "THEN", "ELSE", "END",
                    "COUNT", "SUM", "AVG", "MIN", "MAX", "GETDATE", "DATEADD", "DATEDIFF", "DATEPART",
                    "YEAR", "MONTH", "DAY", "CAST", "CONVERT", "SUBSTRING", "CHARINDEX", "LEN",
                    "UPPER", "LOWER", "LTRIM", "RTRIM", "REPLACE", "COALESCE", "ISNULL"
                ]
            };
            
            // TouchPoint SQL Linter
            function touchpointSQLLinter(text) {
                const errors = [];
                const lines = text.split('\\n');
                
                for (let i = 0; i < lines.length; i++) {
                    const line = lines[i];
                    
                    // Skip comment lines
                    if (line.trim().startsWith('--') || line.trim().startsWith('/*')) continue;
                    
                    // Check for dangerous operations without WHERE
                    if (/\\b(DELETE|DROP|TRUNCATE)\\b/i.test(line) && !/\\bWHERE\\b/i.test(line)) {
                        errors.push({
                            from: CodeMirror.Pos(i, 0),
                            to: CodeMirror.Pos(i, line.length),
                            message: "⚠️ Dangerous operation without WHERE clause",
                            severity: "error"
                        });
                    }
                    
                    // Check for missing TOP in large table queries
                    const largeTablePattern = /\\bFROM\\s+(TagPerson|ActivityLog|EmailResponses|Attend|EngagementScore|Contribution)\\b/i;
                    if (largeTablePattern.test(line)) {
                        let hasTop = false;
                        for (let j = Math.max(0, i - 5); j <= i; j++) {
                            if (/\\bSELECT\\s+TOP\\b/i.test(lines[j])) {
                                hasTop = true;
                                break;
                            }
                        }
                        if (!hasTop) {
                            errors.push({
                                from: CodeMirror.Pos(i, line.search(largeTablePattern)),
                                to: CodeMirror.Pos(i, line.length),
                                message: "⚠️ Large table query without TOP - may timeout",
                                severity: "warning"
                            });
                        }
                    }
                    
                    // Check for missing schema prefix on lookup tables
                    const lookupTables = ['MemberStatus', 'Campus', 'MemberType', 'AttendType', 'Gender', 'MaritalStatus'];
                    for (let tableName of lookupTables) {
                        // Check if table is used in FROM or JOIN without lookup. prefix
                        const tableUsageRegex = new RegExp(`(FROM|JOIN)\\\\s+${tableName}\\\\b`, 'i');
                        if (tableUsageRegex.test(line)) {
                            // Make sure it doesn't already have the lookup. prefix
                            const withPrefixRegex = new RegExp(`(FROM|JOIN)\\\\s+lookup\\\\.${tableName}\\\\b`, 'i');
                            if (!withPrefixRegex.test(line)) {
                                errors.push({
                                    from: CodeMirror.Pos(i, 0),
                                    to: CodeMirror.Pos(i, line.length),
                                    message: `Table '${tableName}' should use 'lookup.' schema prefix`,
                                    severity: "info"
                                });
                            }
                        }
                    }
                    
                    // Check for reserved words that need brackets
                    const reservedWords = ['User', 'Order', 'Group', 'Public', 'Index', 'Key', 'Default', 
                                          'View', 'Table', 'Database', 'Schema', 'Procedure', 'Function',
                                          'Trigger', 'Check', 'Constraint', 'Primary', 'Foreign', 'References'];
                    
                    reservedWords.forEach(word => {
                        // Look for unbracketed reserved words as column names (after SELECT, comma, or in ORDER BY)
                        const unbracketed = new RegExp(`(select|,|order\\\\s+by)\\\\s+.*\\\\b${word}\\\\b(?!\\\\])`, 'i');
                        if (unbracketed.test(line) && !line.includes(`[${word}]`)) {
                            errors.push({
                                from: CodeMirror.Pos(i, 0),
                                to: CodeMirror.Pos(i, line.length),
                                message: `'${word}' is a reserved word - use [${word}] instead`,
                                severity: "error"
                            });
                        }
                    });
                    
                    // Check for common typos
                    const typos = {
                        'CheckInActivity': 'CheckInActivities',
                        'Organization[^s]': 'Organizations',
                        'Meeting[^s]': 'Meetings'
                    };
                    
                    for (let typo in typos) {
                        const regex = new RegExp(`\\\\b${typo}\\\\b`, 'i');
                        if (regex.test(line)) {
                            errors.push({
                                from: CodeMirror.Pos(i, 0),
                                to: CodeMirror.Pos(i, line.length),
                                message: `Did you mean '${typos[typo]}'?`,
                                severity: "warning"
                            });
                        }
                    }
                }
                
                return errors;
            }
            
            // Custom hint renderer with type-ahead search
            function customHintRenderer(element, hints, cur) {
                // Create search input at the top of hints dropdown
                const searchInput = document.createElement('input');
                searchInput.type = 'text';
                searchInput.placeholder = 'Type to filter...';
                searchInput.className = 'hint-search-input';
                searchInput.style.cssText = 'width: 100%; padding: 4px; border: none; border-bottom: 1px solid #ddd; outline: none; box-sizing: border-box;';
                
                // Create wrapper for hints list
                const hintsWrapper = document.createElement('div');
                hintsWrapper.style.cssText = 'max-height: 200px; overflow-y: auto;';
                
                // Store original hints
                const originalHints = hints.list.slice();
                
                // Function to filter hints
                function filterHints(searchText) {
                    const filtered = originalHints.filter(hint => {
                        const displayText = hint.displayText || hint.text;
                        return displayText.toLowerCase().includes(searchText.toLowerCase());
                    });
                    
                    // Clear current hints
                    while (hintsWrapper.firstChild) {
                        hintsWrapper.removeChild(hintsWrapper.firstChild);
                    }
                    
                    // Add filtered hints
                    filtered.forEach((hint, index) => {
                        const hintElement = document.createElement('div');
                        hintElement.className = 'CodeMirror-hint ' + (hint.className || '');
                        hintElement.textContent = hint.displayText || hint.text;
                        hintElement.style.cssText = 'padding: 2px 4px; cursor: pointer;';
                        
                        // Add hover effect
                        hintElement.onmouseenter = function() {
                            document.querySelectorAll('.CodeMirror-hint-active').forEach(el => {
                                el.classList.remove('CodeMirror-hint-active');
                            });
                            hintElement.classList.add('CodeMirror-hint-active');
                        };
                        
                        // Add click handler
                        hintElement.onclick = function() {
                            hints.pick(originalHints.indexOf(hint));
                        };
                        
                        // Set active for first item
                        if (index === 0) {
                            hintElement.classList.add('CodeMirror-hint-active');
                        }
                        
                        hintsWrapper.appendChild(hintElement);
                    });
                    
                    // Update hints list for keyboard navigation
                    hints.list = filtered;
                }
                
                // Initial render
                filterHints('');
                
                // Add search functionality
                searchInput.oninput = function() {
                    filterHints(searchInput.value);
                };
                
                // Focus search input
                setTimeout(() => searchInput.focus(), 0);
                
                // Handle keyboard events
                searchInput.onkeydown = function(e) {
                    if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter' || e.key === 'Escape') {
                        e.preventDefault();
                        // Pass keyboard events to CodeMirror hint widget
                        const activeHint = hintsWrapper.querySelector('.CodeMirror-hint-active');
                        if (e.key === 'Enter' && activeHint) {
                            activeHint.click();
                        } else if (e.key === 'Escape') {
                            hints.close();
                        } else if (e.key === 'ArrowDown') {
                            const next = activeHint ? activeHint.nextSibling : hintsWrapper.firstChild;
                            if (next) {
                                if (activeHint) activeHint.classList.remove('CodeMirror-hint-active');
                                next.classList.add('CodeMirror-hint-active');
                                next.scrollIntoView({ block: 'nearest' });
                            }
                        } else if (e.key === 'ArrowUp') {
                            const prev = activeHint ? activeHint.previousSibling : hintsWrapper.lastChild;
                            if (prev) {
                                if (activeHint) activeHint.classList.remove('CodeMirror-hint-active');
                                prev.classList.add('CodeMirror-hint-active');
                                prev.scrollIntoView({ block: 'nearest' });
                            }
                        }
                    }
                };
                
                // Build final element
                element.appendChild(searchInput);
                element.appendChild(hintsWrapper);
            }
            
            // Enhanced SQL Hint Function with Dynamic Schema
            function touchpointSQLHint(cm) {
                const cur = cm.getCursor();
                const token = cm.getTokenAt(cur);
                // Trim whitespace and handle empty tokens
                const rawString = token.string || '';
                const string = rawString.trim().toLowerCase();
                
                let result = [];
                let start = token.start;
                let end = token.end;
                
                // If token is just whitespace, adjust the position
                if (!string && rawString) {
                    start = cur.ch;
                    end = cur.ch;
                }
                
                const line = cm.getLine(cur.line);
                const beforeCursor = line.substring(0, cur.ch).toLowerCase();
                
                // Get full query text to parse aliases
                const fullText = cm.getValue().toLowerCase();
                
                // Get text immediately before cursor (looking back for alias)
                // This handles both single-line and multi-line cases
                const textBeforeCursor = beforeCursor.trim();
                
                // Prefer dynamic schema if loaded, fall back to hardcoded
                let allTables = {};
                
                if (dynamicSchema.loaded) {
                    // Use dynamic schema as primary source
                    allTables = Object.assign({}, dynamicSchema.tables, dynamicSchema.views);
                } else {
                    // Fall back to static schema
                    allTables = Object.assign({}, touchpointSchema.tables);
                    
                    // Merge any partially loaded dynamic schema
                    Object.keys(dynamicSchema.tables).forEach(tableName => {
                        const table = dynamicSchema.tables[tableName];
                        if (table.columns && table.columns.length > 0) {
                            allTables[tableName] = table;
                        } else if (!allTables[tableName]) {
                            allTables[tableName] = table;
                        }
                    });
                    
                    Object.keys(dynamicSchema.views).forEach(viewName => {
                        const view = dynamicSchema.views[viewName];
                        allTables[viewName] = view;
                    });
                }
                
                // Parse table aliases from the query
                const aliasMap = {};
                // Match patterns like "FROM TableName alias" or "JOIN TableName alias"
                const aliasPatterns = [
                    /\\bfrom\\s+([\\w\\.\\[\\]]+)\\s+(?:as\\s+)?(\\w+)\\b/gi,
                    /\\bjoin\\s+([\\w\\.\\[\\]]+)\\s+(?:as\\s+)?(\\w+)\\b/gi
                ];
                
                aliasPatterns.forEach(pattern => {
                    let match;
                    while ((match = pattern.exec(fullText)) !== null) {
                        const tableName = match[1].replace(/\\[|\\]/g, '');
                        const alias = match[2];
                        if (alias && !alias.match(/^(on|where|inner|left|right|full|cross|join)$/i)) {
                            // Store the full table name (including schema if present)
                            aliasMap[alias.toLowerCase()] = tableName;
                        }
                    }
                });
                
                // Check context - after FROM, JOIN, etc. or typing a temp table
                if (beforeCursor.match(/\\b(from|join|update|into|table)\\s+#?(\\w*)$/i)) {
                    const match = beforeCursor.match(/\\b(from|join|update|into|table)\\s+(#?)(\\w*)$/i);
                    const isTemp = match[2] === '#';
                    const searchStr = match[3] || '';
                    
                    if (isTemp) {
                        // Suggest common temp table patterns
                        const tempTables = [
                            '#TempResults',
                            '#TempPeople', 
                            '#TempOrg',
                            '#TempAttend',
                            '#TempContribution',
                            '#TempMembers',
                            '#TempData',
                            '#Temp'
                        ];
                        
                        tempTables.forEach(tempTable => {
                            if (tempTable.toLowerCase().indexOf('#' + searchStr.toLowerCase()) === 0) {
                                result.push({
                                    text: tempTable.substring(1), // Remove # since it's already typed
                                    displayText: `${tempTable} (temp table)`,
                                    className: 'sql-table-hint'
                                });
                            }
                        });
                    } else {
                        // Suggest regular tables
                        for (let tableName in allTables) {
                            const table = allTables[tableName];
                            // tableName already includes schema prefix for non-dbo tables (e.g., "lookup.MemberStatus")
                            // Only match if search string matches the table name
                            if (tableName.toLowerCase().indexOf(searchStr.toLowerCase()) === 0 || !searchStr) {
                                const colCount = table.columns ? table.columns.length : 0;
                                // Don't add schema prefix again - tableName already has it for non-dbo tables
                                result.push({
                                    text: tableName,
                                    displayText: `${tableName} ${colCount > 0 ? `(${colCount} cols)` : ''}`,
                                    className: 'sql-table-hint'
                                });
                            }
                        }
                    }
                } else if (token.string === '.' || textBeforeCursor.match(/(\w+)\.(\w*)$/) || beforeCursor.match(/\b(\w+)\.(\w*)$/)) {
                    // Table.column completion with partial matching
                    // Check both trimmed text (for multiline with indentation) and full text
                    let tableMatch = textBeforeCursor.match(/(\w+)\.(\w*)$/);
                    if (!tableMatch) {
                        tableMatch = beforeCursor.match(/\b(\w+)\.(\w*)$/);
                    }
                    if (tableMatch) {
                        let tableName = tableMatch[1];
                        const partialColumn = tableMatch[2] || '';  // Text after dot (could be empty)
                        
                        // Remove brackets if present
                        const cleanTableName = tableName.replace(/\\[|\\]/g, '');
                        
                        // Check if this is a schema name (like 'lookup', 'dbo', 'custom')
                        const knownSchemas = ['lookup', 'dbo', 'custom'];
                        if (knownSchemas.includes(cleanTableName.toLowerCase())) {
                            // User typed "schema." or "schema.partial" - show tables in that schema
                            const schemaPrefix = cleanTableName.toLowerCase() + '.';
                            const searchTerm = partialColumn.toLowerCase();
                            
                            for (let tableName in allTables) {
                                // Check if table belongs to this schema
                                if (cleanTableName.toLowerCase() === 'dbo') {
                                    // For dbo schema, tables might not have prefix
                                    if (!tableName.includes('.') || tableName.toLowerCase().startsWith('dbo.')) {
                                        const displayName = tableName.startsWith('dbo.') ? tableName : tableName;
                                        const tableNameOnly = displayName.startsWith('dbo.') ? displayName.substring(4) : displayName;
                                        
                                        // Check if it matches the search term
                                        if (!searchTerm || tableNameOnly.toLowerCase().indexOf(searchTerm) >= 0) {
                                            const table = allTables[tableName];
                                            const colCount = table.columns ? table.columns.length : 0;
                                            result.push({
                                                text: displayName,
                                                displayText: `${displayName} ${colCount > 0 ? `(${colCount} cols)` : ''}`,
                                                className: 'sql-table-hint'
                                            });
                                        }
                                    }
                                } else {
                                    // For other schemas (lookup, custom), check for prefix
                                    if (tableName.toLowerCase().startsWith(schemaPrefix)) {
                                        const table = allTables[tableName];
                                        const colCount = table.columns ? table.columns.length : 0;
                                        // Remove the schema prefix from the completion text since user already typed it
                                        const tableNameOnly = tableName.substring(schemaPrefix.length);
                                        
                                        // Check if it matches the search term
                                        if (!searchTerm || tableNameOnly.toLowerCase().indexOf(searchTerm) >= 0) {
                                            result.push({
                                                text: tableNameOnly,
                                                displayText: `${tableName} ${colCount > 0 ? `(${colCount} cols)` : ''}`,
                                                className: 'sql-table-hint'
                                            });
                                        }
                                    }
                                }
                            }
                            
                            // If we found schema tables, return them
                            if (result.length > 0) {
                                return {
                                    list: result,
                                    from: CodeMirror.Pos(cur.line, tableMatch.index + tableMatch[1].length + 1),
                                    to: CodeMirror.Pos(cur.line, end)
                                };
                            }
                        }
                        
                        // Check if it's an alias first
                        let actualTableName = aliasMap[cleanTableName.toLowerCase()];
                        if (!actualTableName) {
                            actualTableName = cleanTableName;
                        }
                        
                        // Debug logging for troubleshooting
                        const debugTables = ['ms', 'zipcodes', 'lookup.memberstatus'];
                        if (debugTables.includes(cleanTableName.toLowerCase())) {
                            console.log(`Autocomplete debug for ${cleanTableName}:`, {
                                cleanTableName: cleanTableName,
                                actualTableName: actualTableName,
                                aliasMap: aliasMap,
                                matchingTables: Object.keys(allTables).filter(t => 
                                    t.toLowerCase().includes(actualTableName.toLowerCase()) ||
                                    t.toLowerCase() === actualTableName.toLowerCase()
                                ),
                                allTablesKeys: Object.keys(allTables).slice(0, 10) // Show first 10 for reference
                            });
                        }
                        
                        // Find matching table in allTables
                        let matchedTable = null;
                        let matchedTableName = null;
                        
                        // Try to find the table - handle case sensitivity issues
                        for (let tName in allTables) {
                            // Case 1: Exact match (case-insensitive) including schema if present
                            if (tName.toLowerCase() === actualTableName.toLowerCase()) {
                                matchedTable = allTables[tName];
                                matchedTableName = tName;
                                break;
                            }
                            
                            // Case 2: actualTableName is just table name, tName includes schema
                            // e.g., looking for "MemberStatus" when stored as "lookup.MemberStatus"
                            const tableNameOnly = tName.split('.').pop();
                            if (tableNameOnly.toLowerCase() === actualTableName.toLowerCase()) {
                                matchedTable = allTables[tName];
                                matchedTableName = tName;
                                break;
                            }
                            
                            // Case 3: actualTableName has schema (e.g., "dbo.People"), but tName is just table name
                            // This happens when using TOP 100 which adds dbo. prefix
                            const actualTableNameOnly = actualTableName.split('.').pop();
                            if (tName.toLowerCase() === actualTableNameOnly.toLowerCase()) {
                                matchedTable = allTables[tName];
                                matchedTableName = tName;
                                break;
                            }
                            
                            // Case 4: Both have schemas but different ones (e.g., "dbo.People" vs "People")
                            // Compare just the table names
                            if (actualTableName.includes('.') && tableNameOnly.toLowerCase() === actualTableNameOnly.toLowerCase()) {
                                matchedTable = allTables[tName];
                                matchedTableName = tName;
                                break;
                            }
                        }
                        
                        // Debug: log match result
                        if (debugTables.includes(cleanTableName.toLowerCase())) {
                            if (matchedTable) {
                                console.log(`Found match for ${cleanTableName}:`, {
                                    matchedTableName: matchedTableName,
                                    columnsCount: matchedTable.columns ? matchedTable.columns.length : 0,
                                    firstFewColumns: matchedTable.columns ? matchedTable.columns.slice(0, 5) : []
                                });
                            } else {
                                console.log(`NO MATCH found for ${cleanTableName}:`, {
                                    actualTableName: actualTableName,
                                    searchedIn: Object.keys(allTables).length + ' tables',
                                    exampleTableNames: Object.keys(allTables).filter(t => 
                                        t.toLowerCase().includes('member') || 
                                        t.toLowerCase().includes('zip')
                                    )
                                });
                            }
                        }
                        
                        if (matchedTable && matchedTable.columns && matchedTable.columns.length > 0) {
                            const table = matchedTable;
                            // List of SQL reserved words that need brackets
                            const reservedWords = ['User', 'Order', 'Group', 'Public', 'Index', 'Key', 'Default', 
                                                  'View', 'Table', 'Database', 'Schema', 'Procedure', 'Function',
                                                  'Trigger', 'Check', 'Constraint', 'Primary', 'Foreign', 'References'];
                            
                            table.columns.forEach(col => {
                                // Filter columns based on substring match anywhere (case-insensitive)
                                if (col.toLowerCase().indexOf(partialColumn.toLowerCase()) >= 0) {
                                    // Check if column name is a reserved word
                                    const needsBrackets = reservedWords.some(word => 
                                        word.toLowerCase() === col.toLowerCase()
                                    );
                                    
                                    const columnText = needsBrackets ? `[${col}]` : col;
                                    const displayText = needsBrackets ? `[${col}] (reserved)` : col;
                                    
                                    // Sort priority: starts-with matches first, then contains matches
                                    const startsWithMatch = col.toLowerCase().indexOf(partialColumn.toLowerCase()) === 0;
                                    result.push({
                                        text: columnText,
                                        displayText: displayText,
                                        className: 'sql-column-hint',
                                        sortPriority: startsWithMatch ? 0 : 1
                                    });
                                }
                            });
                            
                            // Sort results to show starts-with matches first
                            result.sort((a, b) => {
                                if (a.sortPriority !== b.sortPriority) {
                                    return a.sortPriority - b.sortPriority;
                                }
                                return a.text.localeCompare(b.text);
                            });
                        }
                        
                        // Adjust start position based on whether we have partial text
                        if (partialColumn) {
                            // Replace the partial text
                            start = cur.ch - partialColumn.length;
                            end = cur.ch;
                        } else {
                            // Just after the dot
                            start = cur.ch;
                            end = cur.ch;
                        }
                    }
                } else {
                    // General context - show relevant keywords and suggestions
                    
                    // Check if we should show keywords based on context
                    const showKeywords = !beforeCursor.match(/\b\w+\.\w*$/) && // Not after a dot (fixed to match word boundary)
                                       !beforeCursor.match(/\\b\\w+\\s*\\(\\s*$/); // Not in function call
                    
                    if (showKeywords && touchpointSchema.keywords) {
                        // Determine which keywords are most relevant based on context
                        let relevantKeywords = [];
                        
                        // At start of line or after semicolon - show main statement keywords
                        if (beforeCursor.match(/(^|;|\\n)\\s*\\w*$/i) || beforeCursor.trim() === '') {
                            relevantKeywords = ['SELECT', 'INSERT', 'UPDATE', 'DELETE', 'WITH', 'CREATE', 'ALTER', 'DROP', 'DECLARE', 'SET'];
                        }
                        // After SELECT - show SELECT modifiers
                        else if (beforeCursor.match(/\\bselect\\s+\\w*$/i)) {
                            relevantKeywords = ['TOP', 'DISTINCT', 'ALL', '*'];
                        }
                        // After FROM or JOIN - already handled above, but as fallback
                        else if (beforeCursor.match(/\\b(from|join)\\s+\\w*$/i)) {
                            // This is handled in the table section above
                        }
                        // After WHERE, AND, OR - show logical operators and functions
                        else if (beforeCursor.match(/\\b(where|and|or|having)\\s+\\w*$/i)) {
                            relevantKeywords = ['NOT', 'EXISTS', 'IN', 'BETWEEN', 'LIKE', 'IS'];
                        }
                        // General context - show all keywords
                        else {
                            relevantKeywords = touchpointSchema.keywords;
                        }
                        
                        relevantKeywords.forEach(keyword => {
                            // Show keyword if string is empty or keyword starts with typed string
                            if (!string || keyword.toLowerCase().indexOf(string) === 0) {
                                result.push({
                                    text: keyword,
                                    displayText: keyword,
                                    className: 'sql-keyword-hint'
                                });
                            }
                        });
                        
                        // Limit results to avoid overwhelming the user
                        if (result.length > 20 && !string) {
                            result = result.slice(0, 20);
                        }
                    }
                }
                
                return {
                    list: result,
                    from: CodeMirror.Pos(cur.line, start),
                    to: CodeMirror.Pos(cur.line, end)
                };
            }
            
            // Tab management
            let currentTabId = 1;
            let tabCounter = 1;
            let tabs = {
                1: {
                    name: 'Query 1',
                    content: '',
                    editor: null,
                    results: null,
                    resultsInfo: null
                }
            };
            
            // CodeMirror editors map
            let editors = {};
            
            // Initialize
            document.addEventListener('DOMContentLoaded', function() {
                // Initialize first CodeMirror editor
                const textarea1 = document.getElementById('sql-editor-textarea-1');
                if (textarea1) {
                    editors[1] = CodeMirror.fromTextArea(textarea1, {
                        mode: 'text/x-sql',
                        theme: 'default',
                        lineNumbers: true,
                        matchBrackets: true,
                        autoCloseBrackets: true,
                        indentWithTabs: true,
                        smartIndent: true,
                        lineWrapping: false,
                        foldGutter: true,
                        gutters: ["CodeMirror-linenumbers", "CodeMirror-foldgutter", "CodeMirror-lint-markers"],
                        lint: {
                            getAnnotations: touchpointSQLLinter,
                            async: false
                        },
                        extraKeys: {
                            "Ctrl-Space": function(cm) { cm.showHint(); },
                            "Ctrl-Enter": function() { executeQuery(); },
                            "F5": function() { executeQuery(); },
                            "Cmd-Enter": function() { executeQuery(); },
                            "Ctrl-Shift-Enter": function() { executeSelected(); },
                            "Cmd-Shift-Enter": function() { executeSelected(); },
                            "Ctrl-/": function(cm) { cm.toggleComment({ indent: true }); },
                            "Cmd-/": function(cm) { cm.toggleComment({ indent: true }); },
                            "Ctrl-F": "findPersistent",
                            "Cmd-F": "findPersistent",
                            "Ctrl-H": "replace",
                            "Ctrl-Z": "undo",
                            "Cmd-Z": "undo",
                            "Ctrl-Y": "redo",
                            "Cmd-Y": "redo",
                            "Ctrl-Shift-Z": "redo",
                            "Cmd-Shift-Z": "redo",
                            "Ctrl-Q": function(cm) { cm.foldCode(cm.getCursor()); },
                            "Tab": function(cm) {
                                if (cm.somethingSelected()) {
                                    cm.indentSelection("add");
                                } else {
                                    cm.replaceSelection("    ", "end");
                                }
                            },
                            "Shift-Tab": function(cm) {
                                cm.indentSelection("subtract");
                            }
                        },
                        hintOptions: {
                            hint: function(cm) {
                                const result = touchpointSQLHint(cm);
                                if (result && result.list && result.list.length > 0) {
                                    result.render = customHintRenderer;
                                }
                                return result;
                            },
                            completeSingle: false,
                            closeOnUnfocus: true
                        }
                    });
                    
                    tabs[1].content = editors[1].getValue();
                    tabs[1].editor = editors[1];
                    tabs[1].isDirty = false;
                    
                    // Add auto-complete triggers
                    editors[1].on('inputRead', function(cm, change) {
                        if (change.text.length === 1) {
                            const char = change.text[0];
                            // Auto-show hints after space following keywords, or after dot
                            if (char === ' ') {
                                const cursor = cm.getCursor();
                                const line = cm.getLine(cursor.line);
                                const beforeCursor = line.substring(0, cursor.ch).toLowerCase();
                                if (beforeCursor.match(/\\b(from|join|update|into|table)\\s+$/i)) {
                                    cm.showHint({completeSingle: false});
                                }
                            } else if (char === '.') {
                                cm.showHint({completeSingle: false});
                            }
                        }
                    });
                    
                    // Add change listener to track unsaved changes
                    editors[1].on('change', function() {
                        schedulePersist();
                        if (!tabs[1].isDirty) {
                            tabs[1].isDirty = true;
                            const tabElement = document.querySelector(`.editor-tab[data-tab-id="1"] .editor-tab-name`);
                            if (tabElement && !tabElement.textContent.endsWith(' *')) {
                                tabElement.textContent = tabElement.textContent + ' *';
                            }
                        }
                    });
                }
                
                // Fill the window below TouchPoint's header
                fitExplorerHeight();
                let fitTimer = null;
                window.addEventListener('resize', () => {
                    clearTimeout(fitTimer);
                    fitTimer = setTimeout(fitExplorerHeight, 150);
                });

                // Bring back the tabs open at the last visit, and the history list
                restoreTabs();
                renderHistory();
                updateParamBar();
                window.addEventListener('beforeunload', persistTabs);

                // Load full schema first for autocomplete
                loadFullSchema();
                
                // Then load visual schema tree
                loadSchema();
                loadSavedQueries();
                loadExampleQueries();
                
                // Initialize panel resize
                initPanelResize();
                
                // Ensure panel is visible
                setTimeout(ensurePanelVisible, 100);
                
                // Add global keyboard shortcuts
                document.addEventListener('keydown', function(event) {
                    // Only handle global shortcuts when not in an editor
                    if (event.target.classList.contains('sql-editor')) {
                        return;
                    }
                    
                    // Ctrl+T for new tab
                    if (event.ctrlKey && event.key === 't') {
                        event.preventDefault();
                        newTab();
                    }
                    
                    // Alt+Tab to switch tabs (Ctrl+Tab is reserved by browser)
                    if (event.altKey && event.key === 'Tab') {
                        event.preventDefault();
                        const tabIds = Object.keys(tabs).map(id => parseInt(id));
                        const currentIndex = tabIds.indexOf(currentTabId);
                        
                        if (event.shiftKey) {
                            // Backward
                            const prevIndex = (currentIndex - 1 + tabIds.length) % tabIds.length;
                            switchTab(tabIds[prevIndex]);
                        } else {
                            // Forward
                            const nextIndex = (currentIndex + 1) % tabIds.length;
                            switchTab(tabIds[nextIndex]);
                        }
                    }
                    
                    // Also support Ctrl+1, Ctrl+2, etc. for direct tab access
                    if (event.ctrlKey && event.key >= '1' && event.key <= '9') {
                        event.preventDefault();
                        const tabIds = Object.keys(tabs).map(id => parseInt(id));
                        const targetIndex = parseInt(event.key) - 1;
                        if (targetIndex < tabIds.length) {
                            switchTab(tabIds[targetIndex]);
                        }
                    }
                });
            });
            
            // Panel resize functionality
            let isPanelResizing = false;
            let panelStartX = 0;
            let panelStartWidth = 0;
            
            function initPanelResize() {
                const resizeHandle = document.getElementById('panel-resize-handle');
                const leftPanel = document.getElementById('left-panel');
                const explorer = document.querySelector('.query-explorer');
                
                resizeHandle.addEventListener('mousedown', function(e) {
                    isPanelResizing = true;
                    panelStartX = e.clientX;
                    panelStartWidth = parseInt(document.defaultView.getComputedStyle(leftPanel).width, 10);
                    resizeHandle.classList.add('dragging');
                    document.body.style.cursor = 'col-resize';
                    document.body.style.userSelect = 'none';
                    e.preventDefault();
                });
                
                document.addEventListener('mousemove', function(e) {
                    if (!isPanelResizing) return;
                    
                    const width = panelStartWidth + e.clientX - panelStartX;
                    // Allow more flexible sizing: minimum 150px, maximum 50% of window
                    const maxWidth = window.innerWidth * 0.5;
                    if (width >= 150 && width <= maxWidth) {
                        leftPanel.style.width = width + 'px';
                        // Store the width preference
                        if (window.localStorage) {
                            localStorage.setItem('sqlExplorerPanelWidth', width);
                        }
                    }
                });
                
                document.addEventListener('mouseup', function() {
                    if (isPanelResizing) {
                        isPanelResizing = false;
                        resizeHandle.classList.remove('dragging');
                        document.body.style.cursor = '';
                        document.body.style.userSelect = '';
                    }
                });
                
                // Restore saved width
                if (window.localStorage) {
                    const savedWidth = localStorage.getItem('sqlExplorerPanelWidth');
                    if (savedWidth && !isNaN(savedWidth)) {
                        leftPanel.style.width = savedWidth + 'px';
                    }
                }
            }
            
            function ensurePanelVisible() {
                // This function should do nothing - just maintain current state
                return;
            }
            
            function togglePanel() {
                const leftPanel = document.getElementById('left-panel');
                const toggleIcon = document.getElementById('toggle-icon');
                const resizeHandle = document.getElementById('panel-resize-handle');
                
                if (leftPanel.classList.contains('collapsed')) {
                    // Show panel
                    leftPanel.classList.remove('collapsed');
                    toggleIcon.textContent = '\u2039';
                    document.getElementById('panel-toggle').title = 'Hide this panel';
                    
                    // Restore saved width or default
                    if (window.localStorage) {
                        const savedWidth = localStorage.getItem('sqlExplorerPanelWidth');
                        leftPanel.style.width = savedWidth ? savedWidth + 'px' : '300px';
                    } else {
                        leftPanel.style.width = '300px';
                    }
                    
                    // Re-enable resize handle
                    if (resizeHandle) {
                        resizeHandle.style.display = 'block';
                    }
                } else {
                    // Save current width before hiding
                    if (window.localStorage && leftPanel.style.width) {
                        const currentWidth = parseInt(leftPanel.style.width);
                        if (!isNaN(currentWidth) && currentWidth > 0) {
                            localStorage.setItem('sqlExplorerPanelWidth', currentWidth);
                        }
                    }
                    
                    // Hide panel
                    leftPanel.classList.add('collapsed');
                    toggleIcon.textContent = '\u203A';
                    document.getElementById('panel-toggle').title = 'Show tables, saved queries and history';
                }
            }
            
            function getCurrentEditor() {
                return editors[currentTabId];
            }
            
            function newTab() {
                // Save current tab content
                const currentEditor = getCurrentEditor();
                if (currentEditor) {
                    tabs[currentTabId].content = currentEditor.getValue();
                }
                
                // Create new tab
                tabCounter++;
                const newTabId = tabCounter;
                tabs[newTabId] = {
                    name: 'Query ' + newTabId,
                    content: 'SELECT TOP 10\\n    *\\nFROM People\\nWHERE MemberStatusId = 10',
                    results: null,
                    resultsInfo: null
                };
                
                // Add tab to UI
                const tabsContainer = document.getElementById('editor-tabs');
                const newTabBtn = tabsContainer.querySelector('.new-tab-btn');
                
                const newTab = document.createElement('div');
                newTab.className = 'editor-tab';
                newTab.setAttribute('data-tab-id', newTabId);
                newTab.onclick = function() { switchTab(newTabId); };
                newTab.ondblclick = function() { renameTab(newTabId); };
                newTab.innerHTML = `
                    <span class="editor-tab-name">Query ${newTabId}</span>
                    <span class="editor-tab-close" onclick="closeTab(${newTabId}, event)" title="Close tab (Ctrl+W)">&times;</span>
                `;
                
                tabsContainer.insertBefore(newTab, newTabBtn);
                
                // Create new editor container
                const editorsContainer = document.getElementById('editors-container');
                const newEditorDiv = document.createElement('div');
                newEditorDiv.id = 'sql-editor-' + newTabId;
                newEditorDiv.className = 'sql-editor';
                newEditorDiv.setAttribute('data-tab-id', newTabId);
                
                // Create editor wrapper with resize handle
                const wrapperDiv = document.createElement('div');
                wrapperDiv.className = 'editor-wrapper';
                
                // Create textarea for CodeMirror
                const newTextarea = document.createElement('textarea');
                newTextarea.id = 'sql-editor-textarea-' + newTabId;
                newTextarea.value = tabs[newTabId].content;
                
                // Create resize handle
                const resizeHandle = document.createElement('div');
                resizeHandle.className = 'resize-handle';
                resizeHandle.onmousedown = initResize;
                
                wrapperDiv.appendChild(newTextarea);
                wrapperDiv.appendChild(resizeHandle);
                newEditorDiv.appendChild(wrapperDiv);
                editorsContainer.appendChild(newEditorDiv);
                
                // Initialize CodeMirror on the new textarea
                editors[newTabId] = CodeMirror.fromTextArea(newTextarea, {
                    mode: 'text/x-sql',
                    theme: 'default',
                    lineNumbers: true,
                    matchBrackets: true,
                    autoCloseBrackets: true,
                    indentWithTabs: true,
                    smartIndent: true,
                    lineWrapping: false,
                    foldGutter: true,
                    gutters: ["CodeMirror-linenumbers", "CodeMirror-foldgutter", "CodeMirror-lint-markers"],
                    lint: {
                        getAnnotations: touchpointSQLLinter,
                        async: false
                    },
                    extraKeys: {
                        "Ctrl-Space": function(cm) { cm.showHint(); },
                        "Ctrl-Enter": function() { executeQuery(); },
                        "F5": function() { executeQuery(); },
                        "Cmd-Enter": function() { executeQuery(); },
                        "Ctrl-Shift-Enter": function() { executeSelected(); },
                        "Cmd-Shift-Enter": function() { executeSelected(); },
                        "Ctrl-/": function(cm) { cm.toggleComment({ indent: true }); },
                        "Cmd-/": function(cm) { cm.toggleComment({ indent: true }); },
                        "Ctrl-F": "findPersistent",
                        "Cmd-F": "findPersistent",
                        "Ctrl-H": "replace",
                        "Ctrl-Z": "undo",
                        "Cmd-Z": "undo",
                        "Ctrl-Y": "redo",
                        "Cmd-Y": "redo",
                        "Ctrl-Shift-Z": "redo",
                        "Cmd-Shift-Z": "redo",
                        "Ctrl-Q": function(cm) { cm.foldCode(cm.getCursor()); },
                        "Tab": function(cm) {
                            if (cm.somethingSelected()) {
                                cm.indentSelection("add");
                            } else {
                                cm.replaceSelection("    ", "end");
                            }
                        },
                        "Shift-Tab": function(cm) {
                            cm.indentSelection("subtract");
                        }
                    },
                    hintOptions: {
                        hint: function(cm) {
                            const result = touchpointSQLHint(cm);
                            if (result && result.list && result.list.length > 0) {
                                result.render = customHintRenderer;
                            }
                            return result;
                        },
                        completeSingle: false,
                        closeOnUnfocus: true
                    }
                });
                
                tabs[newTabId].editor = editors[newTabId];
                tabs[newTabId].isDirty = false;
                
                // Add auto-complete triggers
                editors[newTabId].on('inputRead', function(cm, change) {
                    if (change.text.length === 1) {
                        const char = change.text[0];
                        // Auto-show hints after space following keywords, or after dot
                        if (char === ' ') {
                            const cursor = cm.getCursor();
                            const line = cm.getLine(cursor.line);
                            const beforeCursor = line.substring(0, cursor.ch).toLowerCase();
                            if (beforeCursor.match(/\\b(from|join|update|into|table)\\s+$/i)) {
                                cm.showHint({completeSingle: false});
                            }
                        } else if (char === '.') {
                            cm.showHint({completeSingle: false});
                        }
                    }
                });
                
                // Add change listener to track unsaved changes
                editors[newTabId].on('change', function() {
                    schedulePersist();
                    if (!tabs[newTabId].isDirty) {
                        tabs[newTabId].isDirty = true;
                        const tabElement = document.querySelector(`.editor-tab[data-tab-id="${newTabId}"] .editor-tab-name`);
                        if (tabElement && !tabElement.textContent.endsWith(' *')) {
                            tabElement.textContent = tabElement.textContent + ' *';
                        }
                    }
                });
                
                // Switch to new tab
                switchTab(newTabId);
                
                // Ensure panel stays visible
                ensurePanelVisible();
            }
            
            function switchTab(tabId) {
                // Save current tab content and results
                const currentEditor = getCurrentEditor();
                if (currentEditor) {
                    tabs[currentTabId].content = currentEditor.getValue();
                    // Grids are redrawn from their data; anything else (errors,
                    // messages) is kept as it looks now
                    if (!tabs[currentTabId].grid) {
                        tabs[currentTabId].results = document.getElementById('results').innerHTML;
                    }
                }
                
                // Hide all editors
                document.querySelectorAll('.sql-editor').forEach(editor => {
                    editor.classList.remove('active');
                });
                
                // Remove active class from all tabs
                document.querySelectorAll('.editor-tab').forEach(tab => {
                    tab.classList.remove('active');
                });
                
                // Show selected editor
                const selectedEditor = document.getElementById('sql-editor-' + tabId);
                if (selectedEditor) {
                    selectedEditor.classList.add('active');
                    
                    // Mark tab as active
                    const selectedTab = document.querySelector(`.editor-tab[data-tab-id="${tabId}"]`);
                    if (selectedTab) {
                        selectedTab.classList.add('active');
                    }
                    
                    currentTabId = tabId;
                    
                    // Refresh CodeMirror to ensure proper display
                    if (editors[tabId]) {
                        setTimeout(() => {
                            editors[tabId].refresh();
                        }, 1);
                    }
                    
                    // Restore results for this tab
                    restoreResults(tabId);
                    schedulePersist();
                    updateParamBar();
                    
                    // Ensure panel stays visible when switching tabs
                    ensurePanelVisible();
                }
            }
            
            function closeTab(tabId, event) {
                event.stopPropagation();
                
                // Don't close if it's the only tab
                const tabCount = Object.keys(tabs).length;
                if (tabCount <= 1) {
                    alert('Cannot close the last tab');
                    return;
                }
                
                // Clean up CodeMirror instance
                if (editors[tabId]) {
                    editors[tabId].toTextArea();
                    delete editors[tabId];
                }
                
                // Remove tab from data
                delete tabs[tabId];
                
                // Remove tab from UI
                const tabElement = document.querySelector(`.editor-tab[data-tab-id="${tabId}"]`);
                if (tabElement) {
                    tabElement.remove();
                }
                
                // Remove editor
                const editorElement = document.getElementById('sql-editor-' + tabId);
                if (editorElement) {
                    editorElement.remove();
                }
                
                // If we closed the current tab, switch to another one
                if (currentTabId === tabId) {
                    const remainingTabIds = Object.keys(tabs);
                    if (remainingTabIds.length > 0) {
                        switchTab(parseInt(remainingTabIds[0]));
                    }
                }
                schedulePersist();
            }
            
            function renameTab(tabId) {
                const newName = prompt('Enter new tab name:', tabs[tabId].name);
                if (newName) {
                    tabs[tabId].name = newName;
                    schedulePersist();
                    const tabElement = document.querySelector(`.editor-tab[data-tab-id="${tabId}"] .editor-tab-name`);
                    if (tabElement) {
                        tabElement.textContent = newName;
                    }
                }
            }
            
            // CodeMirror handles syntax highlighting now
            
            function showKeyboardShortcuts() {
                document.getElementById('shortcuts-modal').style.display = 'block';
            }
            
            // Click outside modal to close
            document.getElementById('shortcuts-modal').addEventListener('click', function(e) {
                if (e.target === this) {
                    this.style.display = 'none';
                }
            });
            
            function handleKeyPress(event) {
                if (event.ctrlKey && event.key === 'Enter') {
                    event.preventDefault();
                    executeQuery();
                }
                
                // Ctrl+T for new tab
                if (event.ctrlKey && event.key === 't') {
                    event.preventDefault();
                    newTab();
                }
                
                // Alt+Tab to switch tabs (Ctrl+Tab is reserved by browser)
                if (event.altKey && event.key === 'Tab') {
                    event.preventDefault();
                    const tabIds = Object.keys(tabs).map(id => parseInt(id));
                    const currentIndex = tabIds.indexOf(currentTabId);
                    
                    if (event.shiftKey) {
                        // Backward
                        const prevIndex = (currentIndex - 1 + tabIds.length) % tabIds.length;
                        switchTab(tabIds[prevIndex]);
                    } else {
                        // Forward
                        const nextIndex = (currentIndex + 1) % tabIds.length;
                        switchTab(tabIds[nextIndex]);
                    }
                }
                
                // Ctrl+W to close current tab
                if (event.ctrlKey && event.key === 'w') {
                    event.preventDefault();
                    closeTab(currentTabId, event);
                }
            }
            
            function executeQuery(sql = null, force = false) {
                const editor = getCurrentEditor();
                if (!editor) return;
                
                // Check if there's selected text first
                if (!sql) {
                    const selectedText = editor.getSelection();
                    if (selectedText && selectedText.trim()) {
                        sql = selectedText;
                    } else {
                        sql = editor.getValue();
                    }
                }
                
                if (!sql || sql.trim() === '') {
                    document.getElementById('results').innerHTML = '<div class="alert alert-warning">No SQL to execute</div>';
                    return;
                }

                const maxRows = document.getElementById('sqe-rows').value;

                // Ask first before a slow query on one of the largest tables
                if (!force) {
                    const big = slowBigTableQuery(sql, maxRows);
                    if (big) {
                        window.sqePendingSql = sql;
                        tabs[currentTabId].grid = null;
                        renderInfo();
                        document.getElementById('results').innerHTML =
                            '<div class="sqe-banner"><b>' + escapeAttr(big) + '</b> is one of the largest tables in TouchPoint, ' +
                            'and this query has no WHERE. Sorting, grouping, joining or counting it can take minutes. ' +
                            'Add a WHERE (a date range or an Id) to make it fast.<br>' +
                            '<button class="sqe-mini-btn" onclick="runPending()">Run anyway</button>' +
                            '<button class="sqe-mini-btn" onclick="cancelPending()">Cancel</button></div>';
                        return;
                    }
                }

                const tabId = currentTabId;
                tabs[tabId].grid = null;
                renderInfo();
                document.getElementById('results').innerHTML = '<div class="loading">Executing query...</div>';
                
                const found = detectParams(sql);
                const vals = readParamValues();
                const params = {};
                found.names.forEach(nm => { params[nm] = String(vals[nm.toLowerCase()] || '').trim(); });

                postAction({ action: 'execute_query', sql_query: sql, max_rows: maxRows, sql_params: JSON.stringify(params) })
                .then(data => {
                    addHistory({
                        sql: sql,
                        at: Date.now(),
                        ok: !!data.success,
                        rows: data.rowCount || 0,
                        ms: data.executionTime || 0
                    });
                    showResults(data, tabId);
                    ensurePanelVisible();
                })
                .catch(error => {
                    console.error('Query error:', error);
                    tabs[tabId].results = '<div class="alert alert-danger">Error: ' + escapeAttr(String(error.message || error)) + '</div>';
                    if (tabId === currentTabId) {
                        document.getElementById('results').innerHTML = tabs[tabId].results;
                    }
                });
            }

            function runPending() {
                const sql = window.sqePendingSql;
                window.sqePendingSql = null;
                if (sql) executeQuery(sql, true);
            }

            function cancelPending() {
                window.sqePendingSql = null;
                document.getElementById('results').innerHTML = '';
            }

            // TouchPoint's biggest tables. A query on one of these with no WHERE
            // is only quick when it just reads the first rows.
            const SQE_BIG_TABLES = ['TagPerson', 'ActivityLog', 'EmailResponses', 'EngagementScore',
                'Attend', 'Contribution', 'SMSItems', 'SMSList', 'RegistrationData', 'TaskNote',
                'Meetings', 'EmailQueueTo', 'ChangeLog', 'PeopleExtra'];

            function slowBigTableQuery(sql, maxRows) {
                let s = sql.replace(/--[^\\n]*/g, ' ').replace(/\\/\\*[\\s\\S]*?\\*\\//g, ' ').replace(/'(?:[^']|'')*'/g, "''");
                if (/\\bWHERE\\b/i.test(s)) return null;
                const heavy = maxRows === '0' || /\\b(ORDER\\s+BY|GROUP\\s+BY|DISTINCT|JOIN|COUNT\\s*\\(|SUM\\s*\\(|AVG\\s*\\(|MAX\\s*\\(|MIN\\s*\\()/i.test(s);
                if (!heavy) return null;
                for (const t of SQE_BIG_TABLES) {
                    if (new RegExp('(^|[^\\\\w.#@])(dbo\\\\.)?\\\\[?' + t + '\\\\]?(?![\\\\w])', 'i').test(s)) return t;
                }
                return null;
            }
            
            function executeSelected() {
                const editor = getCurrentEditor();
                if (!editor) return;
                
                // Get selected text
                let selectedText = editor.getSelection();
                
                // If no selection, try to get the current statement at cursor position
                if (!selectedText) {
                    const cursor = editor.getCursor();
                    const content = editor.getValue();
                    const lines = content.split('\\n');
                    
                    // Find the SQL statement at the cursor position
                    // Look for statement boundaries (empty lines or semicolons)
                    let startLine = cursor.line;
                    let endLine = cursor.line;
                    
                    // Find start of statement
                    while (startLine > 0) {
                        const line = lines[startLine - 1].trim();
                        if (line === '' || line.endsWith(';')) {
                            break;
                        }
                        startLine--;
                    }
                    
                    // Find end of statement
                    while (endLine < lines.length - 1) {
                        const line = lines[endLine].trim();
                        if (line.endsWith(';')) {
                            break;
                        }
                        if (endLine < lines.length - 1 && lines[endLine + 1].trim() === '') {
                            break;
                        }
                        endLine++;
                    }
                    
                    // Extract the statement
                    const statement = lines.slice(startLine, endLine + 1).join('\\n').trim();
                    if (statement) {
                        selectedText = statement;
                        
                        // Highlight the detected statement
                        editor.setSelection(
                            {line: startLine, ch: 0},
                            {line: endLine, ch: lines[endLine].length}
                        );
                    }
                }
                
                if (selectedText) {
                    executeQuery(selectedText);
                } else {
                    document.getElementById('results').innerHTML = '<div class="alert alert-warning">No text selected. Select SQL text or place cursor within a statement.</div>';
                }
            }
            
            function escapeHtml(text) {
                const div = document.createElement('div');
                div.textContent = text;
                return div.innerHTML;
            }
            
            function showResults(result, tabId) {
                if (tabId === undefined) tabId = currentTabId;
                const tab = tabs[tabId];
                if (!tab) return;
                tab.grid = null;
                tab.results = null;

                if (!result.success) {
                    let errorHtml = '<div class="error-container">';
                    if (result.error_info) {
                        errorHtml += '<div class="error-header">';
                        errorHtml += '<span class="error-type">' + escapeAttr(result.error_info.type || 'SQL Error') + '</span>';
                        errorHtml += '</div>';
                        if (result.error_info.suggestions && result.error_info.suggestions.length > 0) {
                            errorHtml += '<div class="error-suggestions">';
                            errorHtml += '<strong>Suggestions to fix this error:</strong>';
                            errorHtml += '<ul>';
                            result.error_info.suggestions.forEach(function(suggestion) {
                                errorHtml += '<li>' + escapeAttr(suggestion) + '</li>';
                            });
                            errorHtml += '</ul>';
                            errorHtml += '</div>';
                        }
                        errorHtml += '<div class="error-details">';
                        errorHtml += '<details>';
                        errorHtml += '<summary>Show technical details</summary>';
                        errorHtml += '<pre class="error-raw">' + escapeHtml(result.error || '') + '</pre>';
                        errorHtml += '</details>';
                        errorHtml += '</div>';
                    } else {
                        errorHtml += '<div class="alert alert-danger">Error: ' + escapeHtml(result.error || 'Unknown error') + '</div>';
                    }
                    errorHtml += '</div>';
                    tab.results = errorHtml;
                } else if (!result.columns || result.columns.length === 0) {
                    tab.results = '<div style="padding: 20px; color: #6b7280;">The query ran and returned no rows.</div>';
                } else {
                    tab.grid = {
                        columns: result.columns,
                        rows: result.rows || [],
                        sortCol: -1,
                        sortDir: 1,
                        filters: {},
                        view: null,
                        meta: {
                            rowCount: result.rowCount || 0,
                            ms: result.executionTime || 0,
                            truncated: !!result.truncated,
                            maxRows: result.maxRows || 0,
                            orderKnown: !!result.orderKnown
                        }
                    };
                }

                if (tabId === currentTabId) {
                    restoreResults(tabId);
                }
            }

            function restoreResults(tabId) {
                const tab = tabs[tabId];
                const resultsDiv = document.getElementById('results');
                if (tab && tab.grid) {
                    renderGrid();
                } else {
                    resultsDiv.innerHTML = (tab && tab.results) || '';
                    renderInfo();
                }
            }

            // ---- Results grid ----
            const GRID_DOM_LIMIT = 2000;
            const CELL_SHORT = 200;
            const LINK_COLUMNS = {
                peopleid: '/Person2/', personid: '/Person2/', leaderid: '/Person2/',
                headofhouseholdid: '/Person2/', headofhouseholdspouseid: '/Person2/', spouseid: '/Person2/',
                organizationid: '/Org/', orgid: '/Org/',
                familyid: '/Family/',
                meetingid: '/Meeting/',
                bundleid: '/Bundle/', bundleheaderid: '/Bundle/'
            };

            function escapeAttr(text) {
                return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            }

            function fmtNum(n) {
                return Number(n || 0).toLocaleString();
            }

            function isBlank(v) {
                return v === null || v === undefined || v === '';
            }

            function compareValues(x, y) {
                const tx = String(x).trim(), ty = String(y).trim();
                const nx = Number(tx), ny = Number(ty);
                if (tx !== '' && ty !== '' && !isNaN(nx) && !isNaN(ny)) return nx - ny;
                return tx.localeCompare(ty, undefined, { numeric: true, sensitivity: 'base' });
            }

            function gridView(grid) {
                const active = Object.keys(grid.filters).filter(k => grid.filters[k]);
                const idx = [];
                for (let i = 0; i < grid.rows.length; i++) {
                    let keep = true;
                    for (const k of active) {
                        const v = grid.rows[i][k];
                        const t = (v === null || v === undefined) ? 'null' : String(v).toLowerCase();
                        if (t.indexOf(grid.filters[k].toLowerCase()) < 0) { keep = false; break; }
                    }
                    if (keep) idx.push(i);
                }
                if (grid.sortCol >= 0) {
                    const c = grid.sortCol, d = grid.sortDir;
                    idx.sort((a, b) => {
                        const x = grid.rows[a][c], y = grid.rows[b][c];
                        const xb = isBlank(x), yb = isBlank(y);
                        if (xb || yb) return xb === yb ? a - b : (xb ? 1 : -1);
                        return (compareValues(x, y) * d) || (a - b);
                    });
                }
                return idx;
            }

            function renderGrid() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                const resultsDiv = document.getElementById('results');
                if (!grid) return;
                let html = '<div class="results-wrapper"><table class="results-table"><thead><tr>' +
                    '<th class="sqe-rownum" title="Click a row number to copy that row">#</th>';
                grid.columns.forEach((col, i) => {
                    html += '<th onclick="sortGrid(' + i + ')" title="Click to sort">' + escapeAttr(col) +
                        '<span class="sqe-arrow" id="sqe-arrow-' + i + '"></span>' +
                        '<span class="sqe-prof-btn" onclick="profileColumn(' + i + ', event)" title="Column profile">&#9432;</span></th>';
                });
                html += '</tr><tr class="sqe-filter-row"><th class="sqe-rownum"></th>';
                grid.columns.forEach((col, i) => {
                    html += '<th><input class="sqe-filter" data-c="' + i + '" placeholder="filter" value="' +
                        escapeAttr(grid.filters[i] || '') + '" oninput="filterGrid(this)"></th>';
                });
                html += '</tr></thead><tbody id="sqe-tbody"></tbody><tfoot id="sqe-tfoot"></tfoot></table></div>';
                resultsDiv.innerHTML = '<div id="sqe-side"></div>' + html;
                const head = resultsDiv.querySelector('thead tr');
                const h = head ? head.getBoundingClientRect().height : 37;
                resultsDiv.querySelectorAll('tr.sqe-filter-row th').forEach(th => { th.style.top = h + 'px'; });
                fillGridBody();
            }

            function fillGridBody() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                const tbody = document.getElementById('sqe-tbody');
                if (!grid || !tbody) return;
                const view = gridView(grid);
                grid.view = view;
                const shown = Math.min(view.length, GRID_DOM_LIMIT);
                const links = grid.columns.map(c => LINK_COLUMNS[String(c).toLowerCase()] || null);
                const parts = [];
                for (let n = 0; n < shown; n++) {
                    const r = view[n];
                    const row = grid.rows[r];
                    parts.push('<tr><td class="sqe-rownum" data-r="' + r + '" onclick="copyRow(this)" title="Copy this row">' + (n + 1) + '</td>');
                    for (let c = 0; c < grid.columns.length; c++) {
                        parts.push(cellHtml(row[c], links[c], r, c));
                    }
                    parts.push('</tr>');
                }
                tbody.innerHTML = parts.join('');
                renderTotals(grid, view);
                renderSide();
                grid.columns.forEach((col, i) => {
                    const a = document.getElementById('sqe-arrow-' + i);
                    if (a) a.textContent = grid.sortCol === i ? (grid.sortDir > 0 ? '\\u25B2' : '\\u25BC') : '';
                });
                renderInfo();
            }

            function cellHtml(v, link, r, c) {
                const attrs = ' data-r="' + r + '" data-c="' + c + '" ondblclick="copyCell(this)"';
                if (v === null || v === undefined) {
                    return '<td' + attrs + '><span class="sqe-null">NULL</span></td>';
                }
                const t = String(v);
                if (link && /^[0-9]+$/.test(t) && t !== '0') {
                    return '<td' + attrs + '><a href="' + link + t + '" target="_blank" style="color: #2563eb; text-decoration: underline;">' + t + '</a></td>';
                }
                const flat = t.replace(/\\s+/g, ' ');
                if (flat.length > CELL_SHORT) {
                    return '<td' + attrs + '>' + escapeAttr(flat.substring(0, CELL_SHORT)) +
                        '\\u2026<span class="sqe-more" onclick="openCell(this, event)">more</span></td>';
                }
                const title = flat.length > 60 ? ' title="' + escapeAttr(flat) + '"' : '';
                return '<td' + attrs + title + '>' + escapeAttr(flat) + '</td>';
            }

            function openCell(el, event) {
                event.stopPropagation();
                const td = el.parentNode;
                const grid = tabs[currentTabId].grid;
                const v = grid.rows[td.dataset.r][td.dataset.c];
                td.classList.add('sqe-open');
                td.innerHTML = escapeAttr(String(v)) + '<span class="sqe-more" onclick="closeCell(this, event)">less</span>';
            }

            function closeCell(el, event) {
                event.stopPropagation();
                const td = el.parentNode;
                const grid = tabs[currentTabId].grid;
                const r = td.dataset.r, c = td.dataset.c;
                td.outerHTML = cellHtml(grid.rows[r][c], LINK_COLUMNS[String(grid.columns[c]).toLowerCase()] || null, r, c);
            }

            function sortGrid(i) {
                const grid = tabs[currentTabId].grid;
                if (!grid) return;
                if (grid.sortCol !== i) { grid.sortCol = i; grid.sortDir = 1; }
                else if (grid.sortDir === 1) { grid.sortDir = -1; }
                else { grid.sortCol = -1; grid.sortDir = 1; }
                fillGridBody();
            }

            let filterTimer = null;
            function filterGrid(input) {
                const grid = tabs[currentTabId].grid;
                if (!grid) return;
                grid.filters[input.dataset.c] = input.value;
                clearTimeout(filterTimer);
                filterTimer = setTimeout(fillGridBody, grid.rows.length > 5000 ? 250 : 80);
            }

            function renderInfo() {
                const info = document.getElementById('results-info');
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) {
                    info.style.display = 'none';
                    info.innerHTML = '';
                    return;
                }
                const m = grid.meta;
                const view = grid.view || [];
                let h = '<span><b>' + fmtNum(m.rowCount) + '</b> rows in ' + fmtNum(m.ms) + ' ms</span>';
                if (m.truncated) {
                    if (m.maxRows >= (window.SQE_MAX_ALL || 50000)) {
                        h += '<span class="sqe-warn">Stopped at ' + fmtNum(m.maxRows) + ' rows, the most this page reads. Add a WHERE to narrow it.</span>';
                    } else {
                        h += '<span class="sqe-warn">Stopped at ' + fmtNum(m.maxRows) + ' rows. There are more: pick a bigger limit next to Run, or add a WHERE.</span>';
                    }
                }
                if (view.length !== grid.rows.length) {
                    h += '<span>' + fmtNum(view.length) + ' match the filters</span>';
                }
                if (view.length > GRID_DOM_LIMIT) {
                    h += '<span class="sqe-warn">Showing the first ' + fmtNum(GRID_DOM_LIMIT) + ' here. Export CSV and Copy grid include all ' + fmtNum(view.length) + '.</span>';
                }
                if (!m.orderKnown && grid.columns.length > 1) {
                    h += '<span class="sqe-hint">Column order is a best guess for this query.</span>';
                }
                h += '<span class="sqe-hint">Click a header to sort, &#9432; for a column profile. Double-click a cell to copy it. Click a row number to copy the row.</span>';
                h += '<span class="sqe-actions">' +
                    '<button class="sqe-mini-btn" onclick="exportResults()">Export CSV</button>' +
                    '<button class="sqe-mini-btn" onclick="exportExcel()">Excel</button>' +
                    '<button class="sqe-mini-btn" onclick="copyGrid()">Copy grid</button>' +
                    '<button class="sqe-mini-btn" onclick="openChart()">Chart</button></span>';
                info.innerHTML = h;
                info.style.display = 'flex';
            }

            // ---- Copy ----
            function copyText(text, el) {
                const done = () => {
                    if (el) {
                        el.classList.add('sqe-flash');
                        setTimeout(() => el.classList.remove('sqe-flash'), 600);
                    }
                };
                if (navigator.clipboard && window.isSecureContext) {
                    navigator.clipboard.writeText(text).then(done, () => fallbackCopy(text, done));
                } else {
                    fallbackCopy(text, done);
                }
            }

            function fallbackCopy(text, done) {
                const ta = document.createElement('textarea');
                ta.value = text;
                ta.style.position = 'fixed';
                ta.style.opacity = '0';
                document.body.appendChild(ta);
                ta.select();
                try { document.execCommand('copy'); done(); } catch (e) { }
                document.body.removeChild(ta);
            }

            function tsvValue(v) {
                if (v === null || v === undefined) return '';
                const t = String(v);
                return /[\\t\\n\\r"]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
            }

            function copyCell(td) {
                const grid = tabs[currentTabId].grid;
                const v = grid.rows[td.dataset.r][td.dataset.c];
                copyText(v === null || v === undefined ? '' : String(v), td);
            }

            function copyRow(td) {
                const grid = tabs[currentTabId].grid;
                copyText(grid.rows[td.dataset.r].map(tsvValue).join('\\t'), td.parentNode);
            }

            function copyGrid() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) {
                    showMessage('No results to copy. Run a query first.', 'error');
                    return;
                }
                const view = grid.view || gridView(grid);
                const lines = [grid.columns.map(tsvValue).join('\\t')];
                view.forEach(r => lines.push(grid.rows[r].map(tsvValue).join('\\t')));
                copyText(lines.join('\\r\\n'), null);
                showMessage('Copied ' + fmtNum(view.length) + ' rows. Paste into Excel or Sheets.', 'success');
            }

            // ---- History (this browser only) ----
            function historyKey() {
                return 'sqe_history_' + (window.SQE_USER || '0');
            }

            function readHistory() {
                try {
                    const h = JSON.parse(localStorage.getItem(historyKey()) || '[]');
                    return Array.isArray(h) ? h : [];
                } catch (e) {
                    return [];
                }
            }

            function addHistory(entry) {
                try {
                    let h = readHistory().filter(x => x.sql !== entry.sql);
                    h.unshift(entry);
                    localStorage.setItem(historyKey(), JSON.stringify(h.slice(0, 50)));
                } catch (e) { }
                renderHistory();
            }

            function renderHistory() {
                const list = document.getElementById('history-list');
                if (!list) return;
                const box = document.getElementById('history-search');
                const term = (box && box.value || '').toLowerCase();
                const h = readHistory();
                window.sqeHistory = h;
                const parts = [];
                h.forEach((x, i) => {
                    if (!x || typeof x.sql !== 'string') return;
                    if (term && x.sql.toLowerCase().indexOf(term) < 0) return;
                    const when = new Date(x.at).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
                    const meta = x.ok ? fmtNum(x.rows) + ' rows, ' + fmtNum(x.ms) + ' ms' : '<span class="sqe-hist-err">failed</span>';
                    parts.push('<div class="sqe-hist" onclick="openHistory(' + i + ')" title="' + escapeAttr(x.sql.substring(0, 600)) + '">' +
                        '<div class="sqe-hist-sql">' + escapeAttr(x.sql.replace(/\\s+/g, ' ').substring(0, 140)) + '</div>' +
                        '<div class="sqe-hist-meta">' + escapeAttr(when) + ' &middot; ' + meta + '</div></div>');
                });
                list.innerHTML = parts.length ? parts.join('') :
                    '<div class="sqe-note">' + (h.length ? 'Nothing matches.' : 'Queries you run will show here.') + '</div>';
            }

            function openHistory(i) {
                const x = (window.sqeHistory || [])[i];
                if (x) openSqlInNewTab(x.sql, 'From history');
            }

            function clearHistory(el) {
                if (el.dataset.armed !== '1') {
                    el.dataset.armed = '1';
                    el.textContent = 'Click again to clear';
                    setTimeout(() => { el.dataset.armed = ''; el.textContent = 'Clear'; }, 3000);
                    return;
                }
                try { localStorage.removeItem(historyKey()); } catch (e) { }
                el.dataset.armed = '';
                el.textContent = 'Clear';
                renderHistory();
            }

            // ---- Tabs survive a reload (this browser only) ----
            function tabsKey() {
                return 'sqe_tabs_' + (window.SQE_USER || '0');
            }

            let persistTimer = null;
            let restoring = false;
            function schedulePersist() {
                scheduleParamBar();
                if (restoring) return;
                clearTimeout(persistTimer);
                persistTimer = setTimeout(persistTabs, 400);
            }

            function persistTabs() {
                if (restoring) return;
                try {
                    const ids = Object.keys(tabs).map(Number).sort((a, b) => a - b);
                    const list = ids.map(id => ({
                        name: tabs[id].name,
                        sql: editors[id] ? editors[id].getValue() : (tabs[id].content || '')
                    }));
                    localStorage.setItem(tabsKey(), JSON.stringify({ v: 1, current: ids.indexOf(currentTabId), tabs: list }));
                } catch (e) { }
            }

            function restoreTabs() {
                let saved = null;
                try { saved = JSON.parse(localStorage.getItem(tabsKey()) || 'null'); } catch (e) { }
                if (!saved || !Array.isArray(saved.tabs) || !saved.tabs.length) return;
                restoring = true;
                try {
                    saved.tabs.forEach((t, i) => {
                        if (i > 0) newTab();
                        const id = currentTabId;
                        if (editors[id]) editors[id].setValue(typeof t.sql === 'string' ? t.sql : '');
                        setTabName(id, t.name || ('Query ' + id));
                        markClean(id);
                    });
                    const ids = Object.keys(tabs).map(Number).sort((a, b) => a - b);
                    switchTab(ids[saved.current] || ids[0]);
                } finally {
                    restoring = false;
                }
                updateParamBar();
            }

            function setTabName(id, name) {
                tabs[id].name = name;
                const el = document.querySelector(`.editor-tab[data-tab-id="${id}"] .editor-tab-name`);
                if (el) el.textContent = name;
                schedulePersist();
            }

            function markClean(id) {
                tabs[id].isDirty = false;
                const el = document.querySelector(`.editor-tab[data-tab-id="${id}"] .editor-tab-name`);
                if (el && el.textContent.endsWith(' *')) el.textContent = el.textContent.slice(0, -2);
            }

            function openSqlInNewTab(sql, name) {
                newTab();
                const editor = getCurrentEditor();
                if (editor) {
                    editor.setValue(sql);
                    setTabName(currentTabId, name);
                    markClean(currentTabId);
                }
            }


            function fitExplorerHeight() {
                const ex = document.querySelector('.query-explorer');
                if (!ex) return;
                const top = ex.getBoundingClientRect().top + window.scrollY;
                ex.style.height = Math.max(520, window.innerHeight - top - 24) + 'px';
                Object.keys(editors).forEach(id => { if (editors[id]) editors[id].refresh(); });
            }

            // ---- 2.1: libraries loaded only when first used ----
            const SQE_LIBS = {
                formatter: 'https://cdnjs.cloudflare.com/ajax/libs/sql-formatter/15.4.2/sql-formatter.min.js',
                chart: 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
                xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
            };
            const sqeLibCache = {};
            function loadLib(url, globalName) {
                if (window[globalName]) return Promise.resolve(window[globalName]);
                if (!sqeLibCache[url]) {
                    sqeLibCache[url] = new Promise((resolve, reject) => {
                        const tag = document.createElement('script');
                        tag.src = url;
                        tag.onload = () => window[globalName] ? resolve(window[globalName]) : reject(new Error(globalName + ' did not load'));
                        tag.onerror = () => { delete sqeLibCache[url]; reject(new Error('Could not load ' + url)); };
                        document.head.appendChild(tag);
                    });
                }
                return sqeLibCache[url];
            }

            // ---- 2.1: parameters ----
            // TouchPoint fills these in itself, but only when spelled exactly so
            const SQE_TP_PARAMS = { UserPeopleId: 'you', CurrentOrgId: 'the involvement you last opened' };

            function stripSqlNoise(sql) {
                return sql.replace(/--[^\\n]*/g, ' ')
                    .replace(/\\/\\*[\\s\\S]*?\\*\\//g, ' ')
                    .replace(/N?'(?:[^']|'')*'/g, "''")
                    .replace(/\\[[^\\]]*\\]/g, 'x');
            }

            // @Names used in the SQL that it does not DECLARE itself
            function detectParams(sql) {
                const s = stripSqlNoise(sql || '');
                const declared = {};
                const decl = /\\bDECLARE\\b/gi;
                const stop = /^(SELECT|INSERT|UPDATE|DELETE|SET|WITH|IF|WHILE|BEGIN|DECLARE|EXEC|EXECUTE|CREATE|DROP|MERGE|RETURN|PRINT|TRUNCATE|OPEN|FETCH|CLOSE|DEALLOCATE)\\b/i;
                let m;
                while ((m = decl.exec(s)) !== null) {
                    let i = m.index + 7, depth = 0, expectName = true;
                    while (i < s.length) {
                        const ch = s[i];
                        if (ch === '(') depth++;
                        else if (ch === ')') depth--;
                        else if (depth === 0 && ch === ';') break;
                        else if (depth === 0 && ch === ',') expectName = true;
                        else if (depth === 0 && ch === '@' && expectName) {
                            const nm = /^@(\\w+)/.exec(s.substring(i, i + 130));
                            if (nm) {
                                declared[nm[1].toLowerCase()] = 1;
                                i += nm[0].length;
                                expectName = false;
                                continue;
                            }
                        } else if (depth === 0 && !expectName && /\\s/.test(ch) && stop.test(s.substring(i + 1, i + 12))) {
                            break;
                        }
                        i++;
                    }
                }
                const names = [], touchpoint = [], seen = {};
                const use = /(^|[^@\\w])@(\\w+)/g;
                while ((m = use.exec(s)) !== null) {
                    const nm = m[2], key = nm.toLowerCase();
                    if (declared[key] || seen[key]) continue;
                    seen[key] = 1;
                    if (SQE_TP_PARAMS[nm]) touchpoint.push(nm);
                    else names.push(nm);
                }
                return { names: names, touchpoint: touchpoint };
            }

            function paramsKey() {
                return 'sqe_params_' + (window.SQE_USER || '0');
            }

            function readParamValues() {
                try {
                    const v = JSON.parse(localStorage.getItem(paramsKey()) || '{}');
                    return v && typeof v === 'object' ? v : {};
                } catch (e) {
                    return {};
                }
            }

            function saveParamValue(input) {
                const vals = readParamValues();
                vals[input.dataset.p.toLowerCase()] = input.value;
                try { localStorage.setItem(paramsKey(), JSON.stringify(vals)); } catch (e) { }
            }

            function paramKeyDown(event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    executeQuery();
                }
            }

            let paramTimer = null;
            function scheduleParamBar() {
                clearTimeout(paramTimer);
                paramTimer = setTimeout(updateParamBar, 300);
            }

            function updateParamBar() {
                const bar = document.getElementById('sqe-params');
                const editor = getCurrentEditor();
                if (!bar || !editor) return;
                const found = detectParams(editor.getValue());
                const sig = found.names.join(',') + '|' + found.touchpoint.join(',');
                if (bar.dataset.sig === sig) return;  // unchanged: keep focus while typing
                bar.dataset.sig = sig;
                if (!found.names.length && !found.touchpoint.length) {
                    bar.style.display = 'none';
                    bar.innerHTML = '';
                    return;
                }
                const vals = readParamValues();
                let h = '<span class="sqe-params-title">Parameters</span>';
                found.names.forEach(nm => {
                    const hint = nm.toLowerCase() === 'bluetoolbartagid' ? 'a tag Id to test with' : 'empty means NULL';
                    h += '<label class="sqe-param">@' + escapeAttr(nm) + ' <input data-p="' + escapeAttr(nm) + '" value="' +
                        escapeAttr(vals[nm.toLowerCase()] || '') + '" placeholder="' + hint +
                        '" oninput="saveParamValue(this)" onkeydown="paramKeyDown(event)"></label>';
                });
                if (found.names.length) {
                    h += '<span class="sqe-hint">Whole numbers go in as numbers, anything else as text, so 2026-01-31 works for a date.</span>';
                }
                if (found.touchpoint.length) {
                    h += '<span class="sqe-hint">Filled in by TouchPoint: ' +
                        found.touchpoint.map(nm => '@' + nm + ' is ' + SQE_TP_PARAMS[nm]).join(', ') + '.</span>';
                }
                bar.innerHTML = h;
                bar.style.display = 'flex';
            }

            // ---- 2.1: column search ----
            function renderColumnMatches(term) {
                const box = document.getElementById('column-matches');
                if (!box) return;
                term = (term || '').toLowerCase();
                if (term.length < 2 || !dynamicSchema.loaded) {
                    box.innerHTML = '';
                    return;
                }
                const hits = [];
                [['tables', 0], ['views', 1]].forEach(([kind, isView]) => {
                    const src = dynamicSchema[kind] || {};
                    Object.keys(src).forEach(key => {
                        (src[key].columns || []).forEach(col => {
                            const lc = String(col).toLowerCase();
                            const pos = lc.indexOf(term);
                            if (pos < 0) return;
                            hits.push({ col: col, table: key, view: isView, rank: lc === term ? 0 : (pos === 0 ? 1 : 2) });
                        });
                    });
                });
                hits.sort((a, b) => a.rank - b.rank || a.view - b.view ||
                    String(a.col).localeCompare(String(b.col)) || a.table.localeCompare(b.table));
                const shown = hits.slice(0, 150);
                let h = '<div class="sqe-colgroup"><div class="sqe-colgroup-head">Columns matching "' + escapeAttr(term) + '" (' + fmtNum(hits.length) + ')</div>';
                if (!hits.length) h += '<div class="sqe-note">No column has that in its name.</div>';
                shown.forEach(x => {
                    h += '<div class="sqe-colhit"><span class="sqe-colhit-col" data-v="' + escapeAttr(x.col) +
                        '" onclick="insertText(this.dataset.v)" title="Insert the column name">' + escapeAttr(x.col) + '</span>' +
                        ' <span class="sqe-colhit-in">in</span> <span class="sqe-colhit-tbl" data-v="' + escapeAttr(x.table) +
                        '" onclick="insertText(this.dataset.v)" title="Insert the table name">' + escapeAttr(x.table) +
                        (x.view ? ' <i>(view)</i>' : '') + '</span></div>';
                });
                if (hits.length > shown.length) {
                    h += '<div class="sqe-note">Showing the first ' + shown.length + '. Type more to narrow it.</div>';
                }
                box.innerHTML = h + '</div>';
            }

            function insertText(text) {
                const editor = getCurrentEditor();
                if (!editor) return;
                editor.replaceSelection(text);
                editor.focus();
            }

            // ---- 2.1: totals, column profile, chart, Excel ----
            function fmtDec(n) {
                return Number(n).toLocaleString(undefined, { maximumFractionDigits: 2 });
            }

            function shortText(t, max) {
                max = max || 80;
                t = String(t).replace(/\\s+/g, ' ');
                return t.length > max ? t.substring(0, max) + '\\u2026' : t;
            }

            function columnStats(grid, c, view) {
                let n = 0, nulls = 0, blanks = 0, numeric = true, sum = 0, nums = 0;
                for (const r of view) {
                    const v = grid.rows[r][c];
                    if (v === null || v === undefined) { nulls++; continue; }
                    if (v === '') { blanks++; continue; }
                    n++;
                    const x = typeof v === 'number' ? v : (/^-?[0-9]+(\\.[0-9]+)?$/.test(String(v).trim()) ? Number(v) : NaN);
                    if (typeof v === 'boolean' || isNaN(x)) numeric = false;
                    else { sum += x; nums++; }
                }
                return { n: n, nulls: nulls, blanks: blanks, numeric: numeric && nums > 0, sum: sum, avg: nums ? sum / nums : 0 };
            }

            function renderTotals(grid, view) {
                const foot = document.getElementById('sqe-tfoot');
                if (!foot) return;
                let h = '<tr><td class="sqe-rownum" title="Totals for the rows matching the filters">&Sigma;</td>';
                grid.columns.forEach((col, c) => {
                    const st = columnStats(grid, c, view);
                    if (st.numeric && !/id$/i.test(String(col))) {
                        h += '<td title="Sum and average">sum ' + fmtDec(st.sum) + '<br>avg ' + fmtDec(st.avg) + '</td>';
                    } else {
                        h += '<td title="Values that are not NULL or empty">' + fmtNum(st.n) + ' values</td>';
                    }
                });
                foot.innerHTML = h + '</tr>';
            }

            function renderSide() {
                const host = document.getElementById('sqe-side');
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!host || !grid) return;
                if (!grid.side) {
                    if (grid.chart) { grid.chart.destroy(); grid.chart = null; }
                    host.innerHTML = '';
                } else if (grid.side.kind === 'profile') {
                    if (grid.chart) { grid.chart.destroy(); grid.chart = null; }
                    renderProfile(grid, grid.side.c, host);
                } else if (grid.side.kind === 'chart') {
                    renderChart(grid, grid.side, host);
                }
            }

            function closeSide() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) return;
                grid.side = null;
                renderSide();
            }

            function profileColumn(c, event) {
                event.stopPropagation();
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) return;
                grid.side = { kind: 'profile', c: c };
                renderSide();
            }

            function renderProfile(grid, c, host) {
                const view = grid.view || gridView(grid);
                const counts = new Map();
                view.forEach(r => {
                    const v = grid.rows[r][c];
                    if (v === null || v === undefined || v === '') return;
                    const k = String(v);
                    counts.set(k, (counts.get(k) || 0) + 1);
                });
                const st = columnStats(grid, c, view);
                let min = null, max = null;
                counts.forEach((n, k) => {
                    if (min === null || compareValues(k, min) < 0) min = k;
                    if (max === null || compareValues(k, max) > 0) max = k;
                });
                const top = [...counts.entries()].sort((a, b) => b[1] - a[1] || compareValues(a[0], b[0])).slice(0, 10);
                const most = top.length ? top[0][1] : 1;
                let h = '<div class="sqe-panel"><div class="sqe-panel-head"><b>' + escapeAttr(grid.columns[c]) + '</b>' +
                    '<span class="sqe-hint">rows matching the filters</span>' +
                    '<a class="sqe-panel-close" onclick="closeSide()">close</a></div><div class="sqe-prof-stats">';
                [['rows', view.length], ['distinct', counts.size], ['NULL', st.nulls], ['empty text', st.blanks]].forEach(x => {
                    h += '<span><b>' + fmtNum(x[1]) + '</b> ' + x[0] + '</span>';
                });
                if (min !== null) {
                    h += '<span>lowest <b>' + escapeAttr(shortText(min, 40)) + '</b></span><span>highest <b>' + escapeAttr(shortText(max, 40)) + '</b></span>';
                }
                if (st.numeric) {
                    h += '<span>sum <b>' + fmtDec(st.sum) + '</b></span><span>average <b>' + fmtDec(st.avg) + '</b></span>';
                }
                h += '</div>';
                if (top.length) {
                    h += '<div class="sqe-note">Most common values. Click one to filter to it.</div>';
                    top.forEach(([k, n]) => {
                        h += '<div class="sqe-prof-row" data-v="' + escapeAttr(k) + '" onclick="filterTo(' + c + ', this.dataset.v)">' +
                            '<span class="sqe-prof-bar" style="width:' + Math.max(2, Math.round(n / most * 100)) + '%"></span>' +
                            '<span class="sqe-prof-val">' + escapeAttr(shortText(k)) + '</span>' +
                            '<span class="sqe-prof-n">' + fmtNum(n) + '</span></div>';
                    });
                }
                host.innerHTML = h + '</div>';
            }

            function filterTo(c, value) {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) return;
                grid.filters[c] = value;
                const input = document.querySelector('.sqe-filter[data-c="' + c + '"]');
                if (input) input.value = value;
                fillGridBody();
            }

            function openChart() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) {
                    showMessage('No results to chart. Run a query first.', 'error');
                    return;
                }
                if (!grid.side || grid.side.kind !== 'chart') {
                    const view = grid.view || gridView(grid);
                    let label = -1, value = -1;
                    grid.columns.forEach((col, c) => {
                        const st = columnStats(grid, c, view);
                        if (st.numeric && !/id$/i.test(String(col))) {
                            if (value < 0) value = c;
                        } else if (label < 0 && !st.numeric) {
                            label = c;
                        }
                    });
                    grid.side = { kind: 'chart', label: label < 0 ? 0 : label, value: value };
                }
                renderSide();
            }

            function setChart(select) {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid || !grid.side) return;
                grid.side[select.dataset.k] = parseInt(select.value, 10);
                renderSide();
            }

            function renderChart(grid, side, host) {
                const view = grid.view || gridView(grid);
                const pick = (k, cur, allowCount) => '<select class="sqe-rows" data-k="' + k + '" onchange="setChart(this)">' +
                    (allowCount ? '<option value="-1"' + (cur < 0 ? ' selected' : '') + '>Count of rows</option>' : '') +
                    grid.columns.map((col, c) => '<option value="' + c + '"' + (c === cur ? ' selected' : '') + '>' + escapeAttr(col) + '</option>').join('') +
                    '</select>';
                const agg = new Map();
                view.forEach(r => {
                    const row = grid.rows[r];
                    const lv = row[side.label];
                    const lab = lv === null || lv === undefined ? '(NULL)' : String(lv);
                    let add = 1;
                    if (side.value >= 0) {
                        const v = row[side.value];
                        const x = typeof v === 'number' ? v : Number(v);
                        add = (v === null || v === undefined || v === '' || isNaN(x)) ? 0 : x;
                    }
                    agg.set(lab, (agg.get(lab) || 0) + add);
                });
                let entries = [...agg.entries()];
                const LIMIT = 40;
                let note = 'Rows with the same ' + grid.columns[side.label] + ' are added together.';
                if (entries.length > LIMIT) {
                    entries.sort((a, b) => b[1] - a[1]);
                    note = 'Showing the ' + LIMIT + ' largest of ' + fmtNum(entries.length) + ' bars. ' + note;
                    entries = entries.slice(0, LIMIT);
                }
                const horizontal = entries.length > 12;
                const height = horizontal ? Math.max(240, entries.length * 20 + 40) : 320;
                host.innerHTML = '<div class="sqe-panel"><div class="sqe-panel-head"><b>Chart</b>' +
                    '<span>Bars</span>' + pick('label', side.label, false) +
                    '<span>Size</span>' + pick('value', side.value, true) +
                    '<a class="sqe-panel-close" onclick="closeSide()">close</a></div>' +
                    '<div class="sqe-chart-box" style="height:' + height + 'px"><canvas id="sqe-chart"></canvas></div>' +
                    '<div class="sqe-note">' + escapeAttr(note) + '</div></div>';
                loadLib(SQE_LIBS.chart, 'Chart').then(ChartLib => {
                    if (grid.chart) { grid.chart.destroy(); grid.chart = null; }
                    const canvas = document.getElementById('sqe-chart');
                    if (!canvas || tabs[currentTabId].grid !== grid) return;
                    grid.chart = new ChartLib(canvas, {
                        type: 'bar',
                        data: {
                            labels: entries.map(e => shortText(e[0], 40)),
                            datasets: [{
                                label: side.value >= 0 ? grid.columns[side.value] : 'Rows',
                                data: entries.map(e => e[1]),
                                backgroundColor: '#2563eb'
                            }]
                        },
                        options: {
                            indexAxis: horizontal ? 'y' : 'x',
                            maintainAspectRatio: false,
                            animation: false,
                            plugins: { legend: { display: false } }
                        }
                    });
                }).catch(e => {
                    host.querySelector('.sqe-note').textContent = 'The chart library did not load: ' + e.message;
                });
            }

            function exportExcel() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) {
                    showMessage('No results to export. Run a query first.', 'error');
                    return;
                }
                loadLib(SQE_LIBS.xlsx, 'XLSX').then(XL => {
                    const view = grid.view || gridView(grid);
                    const dateRe = /^([0-9]{4})-([0-9]{2})-([0-9]{2})(?: ([0-9]{2}):([0-9]{2}):([0-9]{2}))?$/;
                    const numRe = /^-?(0|[1-9][0-9]{0,14})(\\.[0-9]+)?$/;
                    let cut = 0;
                    const aoa = [grid.columns.slice()];
                    view.forEach(r => {
                        aoa.push(grid.rows[r].map(v => {
                            if (v === null || v === undefined) return null;
                            if (typeof v === 'number' || typeof v === 'boolean') return v;
                            const t = String(v);
                            const d = dateRe.exec(t);
                            if (d) return new Date(+d[1], +d[2] - 1, +d[3], +(d[4] || 0), +(d[5] || 0), +(d[6] || 0));
                            if (numRe.test(t)) return Number(t);
                            if (t.length > 32767) {
                                cut++;
                                return t.substring(0, 32700) + ' [cut to fit an Excel cell]';
                            }
                            return t;
                        }));
                    });
                    const ws = XL.utils.aoa_to_sheet(aoa, { cellDates: true });
                    Object.keys(ws).forEach(addr => {
                        const cell = ws[addr];
                        if (addr[0] === '!' || !cell || cell.t !== 'd') return;
                        const dt = cell.v;
                        cell.z = (dt.getHours() || dt.getMinutes() || dt.getSeconds()) ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd';
                    });
                    ws['!cols'] = grid.columns.map(col => ({ wch: Math.min(50, Math.max(String(col).length + 2, 10)) }));
                    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
                    const wb = XL.utils.book_new();
                    XL.utils.book_append_sheet(wb, ws, 'Results');
                    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
                    const filename = 'query_results_' + timestamp + '.xlsx';
                    XL.writeFile(wb, filename);
                    showMessage('Exported ' + fmtNum(view.length) + ' rows to ' + filename +
                        (cut ? ' (' + cut + ' long values were cut to fit Excel)' : ''), 'success');
                }).catch(e => showMessage('Excel export failed: ' + e.message, 'error'));
            }

            function parseJsonText(text) {
                const a = text.indexOf('{');
                const b = text.lastIndexOf('}') + 1;
                return JSON.parse(a >= 0 && b > a ? text.substring(a, b) : text);
            }

            function postAction(params) {
                const fd = new FormData();
                Object.keys(params).forEach(k => fd.append(k, params[k]));
                const url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                return fetch(url, { method: 'POST', body: fd, credentials: 'same-origin' })
                    .then(r => {
                        if (!r.ok) throw new Error('HTTP ' + r.status);
                        return r.text();
                    })
                    .then(text => {
                        try {
                            return parseJsonText(text);
                        } catch (e) {
                            console.log('Raw response:', text);
                            throw new Error('Invalid response from server. Check the console for details.');
                        }
                    });
            }
            
            function formatSQL() {
                const editor = getCurrentEditor();
                if (!editor) return;
                const hasSel = editor.somethingSelected();
                const src = hasSel ? editor.getSelection() : editor.getValue();
                if (!src.trim()) return;
                loadLib(SQE_LIBS.formatter, 'sqlFormatter').then(lib => {
                    let out;
                    try {
                        out = lib.format(src, { language: 'transactsql', keywordCase: 'upper', tabWidth: 4, linesBetweenQueries: 1 });
                    } catch (e) {
                        showMessage('The formatter could not read this SQL, so the simple formatter was used. (' + e.message + ')', 'error');
                        if (!hasSel) formatSQLBasic();
                        return;
                    }
                    if (hasSel) editor.replaceSelection(out, 'around');
                    else editor.setValue(out);
                }).catch(() => {
                    if (!hasSel) formatSQLBasic();
                });
            }

            function formatSQLBasic() {
                const editor = getCurrentEditor();
                if (!editor) return;
                
                let sql = editor.getValue();
                
                // Enhanced SQL formatting with proper indentation
                // First, normalize whitespace and preserve strings
                const stringMatches = [];
                const stringPlaceholder = '___STRING_PLACEHOLDER___';
                
                // Temporarily replace strings to avoid formatting them
                sql = sql.replace(/'([^']*)'/g, function(match) {
                    stringMatches.push(match);
                    return stringPlaceholder + (stringMatches.length - 1) + stringPlaceholder;
                });
                
                // Remove extra whitespace
                sql = sql.replace(/\\s+/g, ' ').trim();
                
                // Add newlines before major SQL clauses
                sql = sql.replace(/\\b(SELECT|FROM|WHERE|JOIN|LEFT JOIN|RIGHT JOIN|INNER JOIN|CROSS JOIN|FULL OUTER JOIN|ORDER BY|GROUP BY|HAVING|UNION|UNION ALL|EXCEPT|INTERSECT)\\b/gi, function(match, p1, offset) {
                    // Don't add newline at the beginning
                    if (offset === 0) return match.toUpperCase();
                    return '\\n' + match.toUpperCase();
                });
                
                // Format SELECT columns with proper indentation
                sql = sql.replace(/SELECT\\s+/gi, 'SELECT\\n    ');
                sql = sql.replace(/,\\s*(?![^()]*\\))/g, ',\\n    ');
                
                // Format JOIN conditions
                sql = sql.replace(/(JOIN\\s+\\S+)\\s+ON/gi, '$1\\n    ON');
                
                // Format WHERE conditions
                sql = sql.replace(/\\b(AND|OR)\\b/gi, function(match, p1, offset) {
                    // Check if it's in a WHERE clause
                    const beforeText = sql.substring(0, offset);
                    const lastClause = beforeText.match(/\\b(WHERE|ON|HAVING)\\b/gi);
                    if (lastClause) {
                        return '\\n    ' + match.toUpperCase();
                    }
                    return match.toUpperCase();
                });
                
                // Format CASE statements
                sql = sql.replace(/\\bCASE\\b/gi, '\\n    CASE');
                sql = sql.replace(/\\bWHEN\\b/gi, '\\n        WHEN');
                sql = sql.replace(/\\bTHEN\\b/gi, ' THEN');
                sql = sql.replace(/\\bELSE\\b/gi, '\\n        ELSE');
                sql = sql.replace(/\\bEND\\b/gi, '\\n    END');
                
                // Clean up extra spaces and newlines
                sql = sql.replace(/\\n\\s*\\n/g, '\\n');
                sql = sql.replace(/  +/g, ' ');
                
                // Restore strings
                sql = sql.replace(new RegExp(stringPlaceholder + '(\\d+)' + stringPlaceholder, 'g'), function(match, index) {
                    return stringMatches[parseInt(index)];
                });
                
                editor.setValue(sql.trim());
            }
            
            // Rest of the functions remain the same...
            function loadSchema() {
                var url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                var formData = new FormData();
                formData.append('action', 'load_schema');
                
                fetch(url, {
                    method: 'POST',
                    body: formData
                })
                .then(response => response.text())
                .then(text => {
                    try {
                        const data = JSON.parse(text);
                        if (data.tables || data.views) {
                            renderSchema(data);
                        }
                    } catch (e) {
                        console.error('Failed to parse response:', text);
                        document.getElementById('schema-tree').innerHTML = '<div class="alert alert-danger">Error loading schema</div>';
                    }
                })
                .catch(error => {
                    document.getElementById('schema-tree').innerHTML = '<div class="alert alert-danger">Error loading schema: ' + error + '</div>';
                });
            }
            
            function renderSchema(data) {
                const tree = document.getElementById('schema-tree');
                let html = '';
                
                // Only update dynamic schema if full schema isn't already loaded
                if (!dynamicSchema.loaded) {
                    // Store table names in dynamic schema
                    if (data.tables && data.tables.length > 0) {
                        data.tables.forEach(table => {
                            const fullName = table.schema === 'dbo' ? table.name : table.schema + '.' + table.name;
                            // Only add if not already present (from full schema load)
                            if (!dynamicSchema.tables[fullName]) {
                                dynamicSchema.tables[fullName] = {
                                    schema: table.schema,
                                    name: table.name,
                                    columns: []  // Will be populated when columns are loaded
                                };
                            }
                        });
                    }
                    
                    // Store view names in dynamic schema
                    if (data.views && data.views.length > 0) {
                        data.views.forEach(view => {
                            const fullName = view.schema === 'dbo' ? view.name : view.schema + '.' + view.name;
                            // Only add if not already present (from full schema load)
                            if (!dynamicSchema.views[fullName]) {
                                dynamicSchema.views[fullName] = {
                                    schema: view.schema,
                                    name: view.name,
                                    columns: []  // Will be populated when columns are loaded
                                };
                            }
                        });
                    }
                }
                
                // Tables section
                if (data.tables && data.tables.length > 0) {
                    html += `
                        <div class="schema-group" id="tables-group">
                            <div class="schema-group-header" onclick="toggleGroup('tables-group')">
                                <span class="expand-icon">▼</span>
                                <span>Tables (${data.tables.length})</span>
                            </div>
                            <div class="schema-group-content">
                    `;
                    
                    // Group tables by schema
                    const tablesBySchema = {};
                    data.tables.forEach(table => {
                        if (!tablesBySchema[table.schema]) {
                            tablesBySchema[table.schema] = [];
                        }
                        tablesBySchema[table.schema].push(table);
                    });
                    
                    // Render tables grouped by schema
                    Object.keys(tablesBySchema).sort().forEach(schema => {
                        // Add schema header for non-dbo schemas
                        if (schema !== 'dbo') {
                            html += `<div style="font-weight: bold; margin: 10px 0 5px 0; color: #4b5563;">${schema}</div>`;
                        }
                        
                        tablesBySchema[schema].forEach(table => {
                            const tableId = table.schema + '.' + table.name;
                            const displayName = schema === 'dbo' ? table.name : table.name;
                            html += `
                                <div class="table-item" style="${schema !== 'dbo' ? 'margin-left: 15px;' : ''}">
                                    <span class="table-name" onclick="toggleColumns('${tableId}')">
                                        ${displayName} <span style="color: #999; font-size: 12px;">(${table.columns})</span>
                                    </span>
                                    <div class="table-actions">
                                        <button class="table-action-btn" onclick="selectTop100('${tableId}', event)" title="Select Top 100 Rows">Top 100</button>
                                    </div>
                                    <div id="columns-${tableId.replace(/\\./g, '_')}" class="column-list" style="display: none;">
                                        <div class="loading" style="padding: 10px;">Loading columns...</div>
                                    </div>
                                </div>
                            `;
                        });
                    });
                    
                    html += '</div></div>';
                }
                
                // Views section
                if (data.views && data.views.length > 0) {
                    html += `
                        <div class="schema-group" id="views-group">
                            <div class="schema-group-header" onclick="toggleGroup('views-group')">
                                <span class="expand-icon">▼</span>
                                <span>Views (${data.views.length})</span>
                            </div>
                            <div class="schema-group-content">
                    `;
                    
                    // Group views by schema
                    const viewsBySchema = {};
                    data.views.forEach(view => {
                        if (!viewsBySchema[view.schema]) {
                            viewsBySchema[view.schema] = [];
                        }
                        viewsBySchema[view.schema].push(view);
                    });
                    
                    // Render views grouped by schema
                    Object.keys(viewsBySchema).sort().forEach(schema => {
                        // Add schema header for non-dbo schemas
                        if (schema !== 'dbo') {
                            html += `<div style="font-weight: bold; margin: 10px 0 5px 0; color: #4b5563;">${schema}</div>`;
                        }
                        
                        viewsBySchema[schema].forEach(view => {
                            const viewId = view.schema + '.' + view.name;
                            const displayName = schema === 'dbo' ? view.name : view.name;
                            html += `
                                <div class="table-item" style="${schema !== 'dbo' ? 'margin-left: 15px;' : ''}">
                                    <span class="table-name" onclick="toggleColumns('${viewId}')">
                                        [VIEW] ${displayName} <span style="color: #999; font-size: 12px;">(${view.columns})</span>
                                    </span>
                                    <div class="table-actions">
                                        <button class="table-action-btn" onclick="selectTop100('${viewId}', event)" title="Select Top 100 Rows">Top 100</button>
                                    </div>
                                    <div id="columns-${viewId.replace(/\\./g, '_')}" class="column-list" style="display: none;">
                                        <div class="loading" style="padding: 10px;">Loading columns...</div>
                                    </div>
                                </div>
                            `;
                        });
                    });
                    
                    html += '</div></div>';
                }
                
                tree.innerHTML = html;
            }
            
            function toggleGroup(groupId) {
                const group = document.getElementById(groupId);
                if (group) {
                    group.classList.toggle('collapsed');
                }
            }
            
            function toggleColumns(tableId) {
                const [schema, table] = tableId.split('.');
                const columnsDiv = document.getElementById('columns-' + tableId.replace(/\\./g, '_'));
                
                if (columnsDiv.style.display === 'none') {
                    columnsDiv.style.display = 'block';
                    
                    // Check if columns are already loaded from full schema
                    const fullTableName = schema === 'dbo' ? table : tableId;
                    const schemaTable = dynamicSchema.tables[fullTableName] || dynamicSchema.views[fullTableName];
                    
                    if (schemaTable && schemaTable.columns && schemaTable.columns.length > 0) {
                        // Use already loaded columns
                        let html = '';
                        schemaTable.columns.forEach(colName => {
                            const fullName = tableId + '.' + colName;
                            html += `
                                <div class="column-item" onclick="insertColumn('${fullName}', event)">
                                    ${colName}
                                </div>
                            `;
                        });
                        columnsDiv.innerHTML = html;
                    } else if (columnsDiv.innerHTML.indexOf('loading') > -1) {
                        // Load columns if not already loaded
                        var url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                        var formData = new FormData();
                        formData.append('action', 'load_columns');
                        formData.append('schema', schema);
                        formData.append('table', table);
                        
                        fetch(url, {
                            method: 'POST',
                            body: formData
                        })
                        .then(response => response.text())
                        .then(text => {
                            try {
                                const data = JSON.parse(text);
                                if (data.columns) {
                                    // Store columns in dynamic schema
                                    const fullTableName = schema === 'dbo' ? table : tableId;
                                    if (dynamicSchema.tables[fullTableName]) {
                                        dynamicSchema.tables[fullTableName].columns = data.columns.map(col => col.name);
                                    }
                                    
                                    let html = '';
                                    data.columns.forEach(col => {
                                        const fullName = tableId + '.' + col.name;
                                        html += `
                                            <div class="column-item" onclick="insertColumn('${fullName}', event)">
                                                ${col.name} 
                                                <span class="column-type">${col.type}${col.length ? '(' + col.length + ')' : ''}</span>
                                                ${col.key ? `<span class="key-badge">${col.key}</span>` : ''}
                                            </div>
                                        `;
                                    });
                                    columnsDiv.innerHTML = html;
                                }
                            } catch (e) {
                                columnsDiv.innerHTML = '<div style="color: red; padding: 10px;">Error loading columns</div>';
                            }
                        })
                        .catch(error => {
                            columnsDiv.innerHTML = '<div style="color: red; padding: 10px;">Error: ' + error + '</div>';
                        });
                    }
                } else {
                    columnsDiv.style.display = 'none';
                }
            }
            
            function insertColumn(columnName, event) {
                event.stopPropagation();
                
                const editor = getCurrentEditor();
                if (!editor) return;
                
                // Insert at cursor position in CodeMirror
                const cursor = editor.getCursor();
                editor.replaceSelection(columnName);
                editor.focus();
            }
            
            function showTab(tabName) {
                // Don't hide tabs if we're in the middle of other operations
                const clickTarget = event ? event.target : null;
                
                // Only proceed if this is an actual tab click
                if (!clickTarget || !clickTarget.classList.contains('sqe-ltab')) {
                    return;
                }
                
                document.querySelectorAll('.sqe-pane').forEach(tab => {
                    tab.style.display = 'none';
                });
                
                document.querySelectorAll('.sqe-ltab').forEach(tab => {
                    tab.classList.remove('active');
                });
                
                document.getElementById(`${tabName}-tab`).style.display = 'block';
                clickTarget.classList.add('active');
            }
            
            function filterSchema() {
                const searchTerm = document.getElementById('schema-search').value.toLowerCase();
                const items = document.querySelectorAll('#schema-tab .table-item');
                const groups = document.querySelectorAll('#schema-tab .schema-group');
                
                items.forEach(item => {
                    const itemName = item.querySelector('.table-name').textContent.toLowerCase();
                    item.style.display = itemName.includes(searchTerm) ? 'block' : 'none';
                });
                
                // Update group visibility based on visible children
                groups.forEach(group => {
                    const visibleItems = group.querySelectorAll('.table-item:not([style*="display: none"])');
                    group.style.display = visibleItems.length > 0 ? 'block' : 'none';
                });

                renderColumnMatches(searchTerm.trim());
            }
            
            function filterSaved() {
                const searchTerm = document.getElementById('saved-search').value.toLowerCase();
                const items = document.querySelectorAll('#saved-tab .saved-query-item');
                const groups = document.querySelectorAll('#saved-tab .schema-group');
                if (!searchTerm) {
                    items.forEach(item => { item.style.display = 'block'; });
                    groups.forEach(group => { group.style.display = 'block'; });
                    return;
                }
                
                items.forEach(item => {
                    const itemName = item.querySelector('.table-name').textContent.toLowerCase();
                    item.style.display = itemName.includes(searchTerm) ? 'block' : 'none';
                });
                
                // Update group visibility based on visible children
                groups.forEach(group => {
                    const visibleItems = group.querySelectorAll('.saved-query-item:not([style*="display: none"])');
                    group.style.display = visibleItems.length > 0 ? 'block' : 'none';
                });
            }
            
            function saveQuery() {
                const editor = getCurrentEditor();
                if (!editor) return;
                const sql = editor.getValue();
                if (!sql.trim()) {
                    showMessage('Please enter a query to save', 'error');
                    return;
                }
                let name = prompt('Save to My queries (only you can see them). Name:', tabs[currentTabId].name);
                if (!name || !name.trim()) return;
                name = name.trim();
                const mine = (window.savedLists && window.savedLists.mine) || [];
                if (mine.some(x => x.name === name) && !confirm('You already have a query named "' + name + '". Replace it?')) return;
                postAction({ action: 'save_query', query_name: name, query_sql: sql })
                    .then(data => {
                        if (data.success) {
                            showMessage('Saved "' + name + '" to My queries', 'success');
                            loadSavedQueries();
                            setTabName(currentTabId, name);
                            markClean(currentTabId);
                        } else {
                            showMessage('Error saving query: ' + (data.error || 'Unknown error'), 'error');
                        }
                    })
                    .catch(error => showMessage('Error saving query: ' + error.message, 'error'));
            }
            
            function exportResults() {
                const grid = tabs[currentTabId] && tabs[currentTabId].grid;
                if (!grid) {
                    showMessage('No results to export. Run a query first.', 'error');
                    return;
                }
                const view = grid.view || gridView(grid);
                const csvCell = v => {
                    if (v === null || v === undefined) return '';
                    const t = String(v);
                    return /[",\\r\\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
                };
                const lines = [grid.columns.map(csvCell).join(',')];
                view.forEach(r => lines.push(grid.rows[r].map(csvCell).join(',')));

                // The BOM tells Excel the file is UTF-8, so accented names survive
                const blob = new Blob(['\\ufeff' + lines.join('\\r\\n')], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement('a');
                const url = URL.createObjectURL(blob);
                const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
                const filename = 'query_results_' + timestamp + '.csv';
                link.setAttribute('href', url);
                link.setAttribute('download', filename);
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                showMessage('Exported ' + fmtNum(view.length) + ' rows to ' + filename, 'success');
            }
            
            function showMessage(msg, type) {
                const resultsDiv = document.getElementById('results');
                const messageDiv = document.createElement('div');
                messageDiv.className = type === 'error' ? 'alert alert-danger' : 'alert alert-success';
                messageDiv.textContent = msg;
                messageDiv.style.marginTop = '10px';
                
                // Insert message at top of results
                resultsDiv.insertBefore(messageDiv, resultsDiv.firstChild);
                
                // Auto-hide success messages after 3 seconds
                if (type === 'success') {
                    setTimeout(() => {
                        messageDiv.remove();
                    }, 3000);
                }
            }
            
            function loadSavedQueries() {
                postAction({ action: 'get_saved' })
                    .then(data => {
                        if (data.success === false) throw new Error(data.error || 'Unknown error');
                        renderSaved(data);
                    })
                    .catch(error => {
                        console.error('Error loading saved queries:', error);
                        document.getElementById('saved-queries').innerHTML =
                            '<div style="color: red; padding: 10px;">Error loading saved queries</div>';
                    });
            }

            function renderSaved(data) {
                const L = window.savedLists = { mine: data.mine || [], sqe: data.sqe || [], other: data.other || [] };
                let html = savedGroup('saved-mine', 'My queries (' + L.mine.length + ')',
                    'Only you can see these. They are kept on your own TouchPoint record.', L.mine, 'mine', false);
                if (L.sqe.length) {
                    html += savedGroup('saved-sqe', 'Saved by earlier versions (' + L.sqe.length + ')',
                        'These are in Special Content, which anyone can read without signing in. Use "copy to mine", then delete them in Special Content.',
                        L.sqe, 'sqe', false);
                }
                if (L.other.length) {
                    html += savedGroup('saved-other', 'Other SQL in Special Content (' + L.other.length + ')',
                        'Reports and scripts saved elsewhere in TouchPoint. Opening one here does not change it.',
                        L.other, 'other', true);
                }
                document.getElementById('saved-queries').innerHTML = html;
            }

            function savedGroup(id, title, note, list, kind, collapsed) {
                let h = '<div class="schema-group' + (collapsed ? ' collapsed' : '') + '" id="' + id + '">' +
                    '<div class="schema-group-header" onclick="toggleGroup(this.parentNode.id)">' +
                    '<span class="expand-icon">&#9660;</span><span>' + escapeAttr(title) + '</span></div>' +
                    '<div class="schema-group-content"><div class="sqe-note">' + escapeAttr(note) + '</div>';
                if (!list.length) {
                    h += '<div class="sqe-note">Nothing saved yet. Use Save Query above the editor.</div>';
                }
                list.forEach((x, i) => {
                    h += '<div class="table-item saved-query-item"><span class="sqe-item-actions">';
                    if (kind === 'mine') h += '<a data-i="' + i + '" onclick="deleteSaved(this, event)">delete</a>';
                    if (kind === 'sqe') h += '<a data-i="' + i + '" onclick="copySavedToMine(this, event)">copy to mine</a>';
                    h += '</span><span class="table-name" data-kind="' + kind + '" data-i="' + i + '" onclick="openSaved(this)" title="' +
                        escapeAttr(x.name + (x.saved ? ' (saved ' + x.saved + ')' : '')) + '">' + escapeAttr(x.name) + '</span></div>';
                });
                return h + '</div></div>';
            }

            function openSaved(el) {
                const x = (window.savedLists[el.dataset.kind] || [])[+el.dataset.i];
                if (x) openSqlInNewTab(x.sql || '', x.name);
            }

            function deleteSaved(a, event) {
                event.stopPropagation();
                const x = window.savedLists.mine[+a.dataset.i];
                if (!x) return;
                if (a.dataset.armed !== '1') {
                    a.dataset.armed = '1';
                    a.textContent = 'click again to delete';
                    setTimeout(() => { a.dataset.armed = ''; a.textContent = 'delete'; }, 3000);
                    return;
                }
                postAction({ action: 'delete_query', query_name: x.name })
                    .then(data => {
                        if (data.success) {
                            showMessage('Deleted "' + x.name + '"', 'success');
                            loadSavedQueries();
                        } else {
                            showMessage('Could not delete: ' + (data.error || 'Unknown error'), 'error');
                        }
                    })
                    .catch(error => showMessage('Could not delete: ' + error.message, 'error'));
            }

            function copySavedToMine(a, event) {
                event.stopPropagation();
                const x = window.savedLists.sqe[+a.dataset.i];
                if (!x) return;
                postAction({ action: 'save_query', query_name: x.name, query_sql: x.sql || '' })
                    .then(data => {
                        if (data.success) {
                            showMessage('Copied "' + x.name + '" to My queries. You can now delete it in Special Content.', 'success');
                            loadSavedQueries();
                        } else {
                            showMessage('Could not copy: ' + (data.error || 'Unknown error'), 'error');
                        }
                    })
                    .catch(error => showMessage('Could not copy: ' + error.message, 'error'));
            }
            
            
            function loadExampleQueries() {
                var url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                var formData = new FormData();
                formData.append('action', 'get_examples');
                
                fetch(url, {
                    method: 'POST',
                    body: formData
                })
                .then(response => response.text())
                .then(text => {
                    try {
                        const data = JSON.parse(text);
                        if (data.examples) {
                            const examplesDiv = document.getElementById('example-queries');
                            
                            // Add documentation link at the top
                            let html = `
                                <div style="padding: 10px; background: #f0f9ff; border-left: 4px solid #3b82f6; margin-bottom: 10px;">
                                    <div style="font-weight: bold; margin-bottom: 5px;">
                                        📚 SQL Documentation Examples
                                    </div>
                                    <div style="font-size: 12px; color: #666;">
                                        From: <a href="https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html" 
                                                target="_blank" 
                                                style="color: #3b82f6;">
                                            TouchPoint SQL Documentation
                                        </a>
                                    </div>
                                </div>
                            `;
                            
                            // Group examples by category
                            const categories = {};
                            data.examples.forEach((example, index) => {
                                const category = example.category || 'Other';
                                if (!categories[category]) {
                                    categories[category] = [];
                                }
                                example.index = index;
                                categories[category].push(example);
                            });
                            
                            // Render examples by category
                            Object.keys(categories).sort().forEach(category => {
                                html += `<div style="font-weight: bold; margin: 10px 0 5px; color: #4a5568;">${category}</div>`;
                                categories[category].forEach(example => {
                                    html += `
                                        <div class="query-item" onclick="loadQuery(${example.index})" 
                                             title="${example.description || ''}">
                                            <div style="font-weight: 500;">${example.name}</div>
                                            ${example.description ? 
                                                `<div style="font-size: 11px; color: #666; margin: 2px 0;">${example.description}</div>` : 
                                                ''}
                                            <div class="query-preview">${example.sql.substring(0, 50)}...</div>
                                        </div>
                                    `;
                                });
                            });
                            
                            examplesDiv.innerHTML = html;
                            window.exampleQueries = data.examples;
                        }
                    } catch (e) {
                        console.error('Error parsing examples:', e);
                    }
                })
                .catch(error => {
                    console.error('Error loading examples:', error);
                });
            }
            
            function loadQuery(index) {
                if (window.exampleQueries && window.exampleQueries[index]) {
                    const editor = getCurrentEditor();
                    if (editor) {
                        editor.setValue(window.exampleQueries[index].sql);
                        tabs[currentTabId].name = window.exampleQueries[index].name;
                        const tabElement = document.querySelector(`.editor-tab[data-tab-id="${currentTabId}"] .editor-tab-name`);
                        if (tabElement) {
                            tabElement.textContent = tabs[currentTabId].name;
                        }
                    }
                }
            }
            
            function selectTop100(tableId, event) {
                event.stopPropagation();
                
                // First get the columns for this table
                const [schema, table] = tableId.split('.');
                const colsDiv = document.getElementById('cols-' + tableId.replace(/\./g, '_'));
                
                // Create a new tab
                newTab();
                
                // Update tab name to just the table name
                tabs[currentTabId].name = tableId;
                const tabElement = document.querySelector(`.editor-tab[data-tab-id="${currentTabId}"] .editor-tab-name`);
                if (tabElement) {
                    tabElement.textContent = tabs[currentTabId].name;
                }
                
                // Check if we already have the columns loaded
                if (colsDiv && columnsCache[tableId]) {
                    // Parse the cached columns to build the query
                    const tempDiv = document.createElement('div');
                    tempDiv.innerHTML = columnsCache[tableId];
                    const columnItems = tempDiv.querySelectorAll('.column-item');
                    
                    if (columnItems.length > 0) {
                        const columnNames = [];
                        columnItems.forEach(item => {
                            const onclick = item.getAttribute('onclick');
                            if (onclick) {
                                const match = onclick.match(/insertColumn\('([^.]+\.[^.]+\.([^']+))'/);
                                if (match && match[2]) {
                                    columnNames.push(match[2]);
                                }
                            }
                        });
                        
                        if (columnNames.length > 0) {
                            const sql = `SELECT TOP 100\n    ${columnNames.join(',\\n    ')}\nFROM ${tableId}\nORDER BY 1`;
                            const editor = getCurrentEditor();
                            if (editor) {
                                editor.setValue(sql);
                                executeQuery();
                            }
                            return;
                        }
                    }
                }
                
                // If columns not loaded, load them first then build query
                const url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                const formData = new FormData();
                formData.append('action', 'load_columns');
                formData.append('schema', schema);
                formData.append('table', table);
                
                fetch(url, {
                    method: 'POST',
                    body: formData
                })
                .then(response => response.text())
                .then(text => {
                    try {
                        const data = JSON.parse(text);
                        if (data.columns && data.columns.length > 0) {
                            const columnNames = data.columns.map(col => col.name);
                            const sql = `SELECT TOP 100\n    ${columnNames.join(',\\n    ')}\nFROM ${tableId}\nORDER BY 1`;
                            
                            const editor = getCurrentEditor();
                            if (editor) {
                                editor.setValue(sql);
                                executeQuery();
                            }
                            
                            // Also cache the columns for the schema display
                            if (colsDiv) {
                                let html = '';
                                data.columns.forEach(col => {
                                    const fullName = tableId + '.' + col.name;
                                    const badge = col.key ? `<span class="key-badge ${col.key.toLowerCase()}-badge">${col.key}</span>` : '';
                                    html += `
                                        <div class="column-item" onclick="insertColumn('${fullName}')">
                                            ${col.name} 
                                            <span class="column-type">${col.type}${col.length ? '(' + col.length + ')' : ''}</span>
                                            ${badge}
                                        </div>
                                    `;
                                });
                                columnsCache[tableId] = html;
                                colsDiv.innerHTML = html;
                            }
                        } else {
                            // Fallback to * if no columns found
                            const sql = `SELECT TOP 100\n    *\nFROM ${tableId}\nORDER BY 1`;
                            const editor = getCurrentEditor();
                            if (editor) {
                                editor.setValue(sql);
                                executeQuery();
                            }
                        }
                    } catch (e) {
                        // Fallback to * on error
                        const sql = `SELECT TOP 100\n    *\nFROM ${tableId}\nORDER BY 1`;
                        const editor = getCurrentEditor();
                        if (editor) {
                            editor.setValue(sql);
                            executeQuery();
                        }
                    }
                })
                .catch(error => {
                    // Fallback to * on error
                    const sql = `SELECT TOP 100\n    *\nFROM ${tableId}\nORDER BY 1`;
                    const editor = getCurrentEditor();
                    if (editor) {
                        editor.setValue(sql);
                        executeQuery();
                    }
                });
            }
            
            // Editor resizing functionality
            let isResizing = false;
            let startY = 0;
            let startHeight = 0;
            let currentEditor = null;
            
            function initResize(e) {
                isResizing = true;
                startY = e.clientY;
                const editorWrapper = e.target.parentElement;
                const cmElement = editorWrapper.querySelector('.CodeMirror');
                if (cmElement) {
                    startHeight = cmElement.offsetHeight;
                    currentEditor = editors[currentTabId];
                }
                
                document.addEventListener('mousemove', doResize);
                document.addEventListener('mouseup', stopResize);
                e.preventDefault();
            }
            
            function doResize(e) {
                if (!isResizing) return;
                
                const diff = e.clientY - startY;
                const newHeight = Math.max(150, startHeight + diff);  // Removed upper limit
                
                if (currentEditor) {
                    currentEditor.setSize(null, newHeight);
                }
            }
            
            function stopResize() {
                isResizing = false;
                document.removeEventListener('mousemove', doResize);
                document.removeEventListener('mouseup', stopResize);
            }
            
        </script>
        """)

# Main execution - instantiate the explorer
try:
    # Check if this is an AJAX request first
    if hasattr(model.Data, 'action') and model.Data.action:
        # Remove debug code that was causing JSON parsing issues
        # This was printing extra JSON before the actual response
        pass
    
    explorer = QueryExplorer()
    
    # Check if user has permission (but still handle AJAX requests)
    if explorer.has_permission:
        if hasattr(model.Data, 'action') and model.Data.action:
            # Set content type for JSON responses
            model.Header = ""
            action = model.Data.action
            
            try:
                if action == 'load_schema':
                    # Return tables and views organized by type
                    objects = explorer.get_database_objects()
                    tables_list = []
                    views_list = []
                    
                    for obj in objects:
                        item = {
                            'schema': obj.TABLE_SCHEMA,
                            'name': obj.TABLE_NAME,
                            'columns': int(obj.COLUMN_COUNT)
                        }
                        
                        if obj.TABLE_TYPE == 'BASE TABLE':
                            tables_list.append(item)
                        elif obj.TABLE_TYPE == 'VIEW':
                            views_list.append(item)
                    
                    print(safe_json({
                        'tables': tables_list,
                        'views': views_list
                    }))
            
                elif action == 'load_columns':
                    # Load columns for a specific table
                    schema_name = getattr(model.Data, 'schema', 'dbo')
                    table_name = getattr(model.Data, 'table', '')
                    
                    if table_name:
                        columns = explorer.get_table_columns(schema_name, table_name)
                        columns_list = []
                        
                        for col in columns:
                            columns_list.append({
                                'name': col.COLUMN_NAME,
                                'type': col.DATA_TYPE,
                                'length': getattr(col, 'CHARACTER_MAXIMUM_LENGTH', None),
                                'nullable': col.IS_NULLABLE,
                                'key': col.KEY_TYPE
                            })
                        
                        print(safe_json({'columns': columns_list}))
                    else:
                        print(safe_json({'error': 'No table name provided'}))
                
                elif action == 'execute_query':
                    sql = get_param('sql_query')
                    if sql.strip():
                        max_rows = get_param('max_rows', str(DEFAULT_MAX_ROWS))
                        params = explorer.build_params(get_param('sql_params'))
                        print(safe_json(explorer.execute_query(sql, max_rows, params)))
                    else:
                        print(safe_json({'success': False, 'error': 'No SQL query provided'}))

                elif action == 'get_full_schema':
                    # Get complete schema with all columns
                    try:
                        schema = explorer.get_full_schema()
                        print(safe_json({'success': True, 'schema': schema}))
                    except Exception as e:
                        print(safe_json({'success': False, 'error': str(e)}))
                
                elif action == 'get_examples':
                    # Return example queries
                    examples = explorer.get_common_queries()
                    print(safe_json({'examples': examples}))
                
                elif action == 'get_saved':
                    mine = []
                    for x in explorer.load_my_queries():
                        mine.append({'name': x.get('name'), 'sql': x.get('sql', ''), 'saved': x.get('saved', '')})
                    sqe = []
                    other = []
                    for row in explorer.get_saved_queries():
                        try:
                            item = {
                                'id': int(getattr(row, 'Id', 0) or 0),
                                'name': unicode(getattr(row, 'Name', '') or ''),
                                'sql': unicode(getattr(row, 'Body', '') or '')
                            }
                        except Exception:
                            continue
                        if not item['name'] or item['id'] <= 0:
                            continue
                        if getattr(row, 'FromSqe', 0):
                            sqe.append(item)
                        else:
                            other.append(item)
                    print(safe_json({'success': True, 'mine': mine, 'sqe': sqe, 'other': other}))

                elif action == 'save_query':
                    name = get_param('query_name').strip()
                    sql = get_param('query_sql')
                    if name and sql.strip():
                        try:
                            explorer.save_query(name, sql)
                            print(safe_json({'success': True, 'message': 'Query saved'}))
                        except Exception as save_error:
                            print(safe_json({'success': False, 'error': unicode(save_error)}))
                    else:
                        print(safe_json({'success': False, 'error': 'Name and SQL are required'}))

                elif action == 'delete_query':
                    name = get_param('query_name')
                    try:
                        found = explorer.delete_query(name)
                        print(safe_json({'success': found, 'error': '' if found else 'Query not found'}))
                    except Exception as del_error:
                        print(safe_json({'success': False, 'error': unicode(del_error)}))

                else:
                    # Unknown action
                    print(safe_json({'success': False, 'error': 'Unknown action: ' + action}))
                    
            except Exception as e:
                # Always return JSON for AJAX requests
                print(safe_json({'success': False, 'error': str(e)}))
        
        else:
            # Render the interface
            explorer.render_interface()
    else:
        # Check if this is an AJAX request
        if hasattr(model.Data, 'action') and model.Data.action:
            # For AJAX requests, still allow them to go through to get proper JSON errors
            model.Header = ""
            action = model.Data.action
            
            # All actions will return permission denied from their respective methods
            if action == 'execute_query':
                sql = model.Data.sql_query if hasattr(model.Data, 'sql_query') else ''
                if sql:
                    result = explorer.execute_query(sql)  # This will return permission denied
                    print(safe_json(result))
                else:
                    print(safe_json({'success': False, 'error': 'No SQL query provided'}))
            else:
                print(safe_json({'success': False, 'error': 'Access Denied: Admin or Developer role required'}))
        else:
            print('<div class="alert alert-danger">Access Denied: Admin or Developer role required</div>')
        
except Exception as e:
    # Check if this is an AJAX request
    if hasattr(model.Data, 'action') and model.Data.action:
        # Return JSON error for AJAX requests
        print(safe_json({'success': False, 'error': 'Initialization error: ' + str(e)}))
    else:
        # Return HTML error for page load
        print('<div class="alert alert-danger">Error initializing Query Explorer: {0}</div>'.format(str(e)))
