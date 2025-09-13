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
#####################################################################

# Written By: Ben Swaby
# Email: bswaby@fbchtn.org

import json
import re
from datetime import datetime

# Custom JSON encoder for Python 2.7 Unicode handling
class SafeJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        # Handle datetime objects
        if hasattr(obj, 'isoformat'):
            return obj.isoformat()
        elif hasattr(obj, 'strftime'):
            return obj.strftime('%Y-%m-%d %H:%M:%S')
        # Let the base class handle other types
        return super(SafeJSONEncoder, self).default(obj)

def safe_json_dumps(obj):
    """Safely serialize object to JSON, handling encoding issues"""
    try:
        return json.dumps(obj, cls=SafeJSONEncoder, ensure_ascii=True)
    except Exception as e:
        # If that fails, try cleaning the data
        try:
            def clean_for_json(item):
                """Recursively clean data for JSON serialization"""
                if isinstance(item, dict):
                    return {k: clean_for_json(v) for k, v in item.items()}
                elif isinstance(item, list):
                    return [clean_for_json(i) for i in item]
                elif 'unicode' in dir(__builtins__) and isinstance(item, unicode):
                    # Python 2.7 unicode string
                    # Replace non-ASCII characters with XML character references
                    return item.encode('ascii', 'xmlcharrefreplace').decode('ascii')
                elif isinstance(item, str):
                    # Python 2.7 byte string or Python 3 string
                    try:
                        # Try to decode as UTF-8 first
                        if 'unicode' in dir(__builtins__):
                            # Python 2.7: decode then encode
                            return item.decode('utf-8', 'replace').encode('ascii', 'xmlcharrefreplace').decode('ascii')
                        else:
                            # Python 3: just encode
                            return item.encode('ascii', 'xmlcharrefreplace').decode('ascii')
                    except:
                        # If that fails, try latin-1
                        try:
                            if 'unicode' in dir(__builtins__):
                                return item.decode('latin-1', 'replace').encode('ascii', 'xmlcharrefreplace').decode('ascii')
                            else:
                                return item.encode('ascii', 'replace').decode('ascii')
                        except:
                            # Last resort - return placeholder
                            return "[Unable to decode text]"
                else:
                    # Return as-is for other types (numbers, booleans, None, etc.)
                    return item
            
            cleaned = clean_for_json(obj)
            return json.dumps(cleaned, ensure_ascii=True)
        except Exception as fallback_error:
            # Last resort - return error as JSON
            return json.dumps({'success': False, 'error': 'Unable to encode response: ' + str(fallback_error)})

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
        """Get user's saved queries from Content table"""
        # Simple query that we know works - Name is a reserved word so needs brackets
        sql = """
        SELECT 
            c.Id,
            c.[Name],
            c.Body
        FROM Content c
        WHERE c.TypeID = 4  -- SQL Script type
        ORDER BY c.[Name]
        """
        
        try:
            results = q.QuerySql(sql)
            # Force evaluation of the query results
            results_list = []
            if results:
                for r in results:
                    results_list.append(r)
            return results_list
        except Exception as e:
            # Return empty list if error
            return []
    
    def save_query(self, name, sql):
        """Save a query to TouchPoint Content table"""
        try:
            # Try to use WriteContentSql first - it might work
            model.WriteContentSql(name, sql, "SQL Query Explorer")
            return True
        except Exception as e:
            # If WriteContentSql fails, try alternative approach
            # Return a message that saving requires different approach
            raise Exception("Failed to save query: " + str(e))
    
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
                suggestions.append("Table names are case-sensitive - try: " + table_name.lower() + " or " + table_name.upper())
                
                # Check for common table name mistakes
                if table_name.lower() == 'people1':
                    suggestions.append("Did you mean 'People' instead of 'people1'?")
                elif not '.' in table_name:
                    suggestions.append("Try adding the schema prefix (e.g., dbo." + table_name + ")")
                
                # Suggest using the schema browser
                suggestions.append("Use the Schema browser on the left to find the correct table name")
                
        # Invalid column name
        elif "Invalid column name" in error_str:
            match = re.search(r"Invalid column name '([^']+)'", error_str)
            if match:
                column_name = match.group(1)
                error_type = "Invalid Column Name"
                suggestions.append("Check if the column name is spelled correctly")
                suggestions.append("Column names are case-sensitive")
                suggestions.append("Click on a table in the Schema browser to see available columns")
                
        # Syntax errors
        elif "Incorrect syntax" in error_str:
            error_type = "SQL Syntax Error"
            suggestions.append("Check for missing commas between column names")
            suggestions.append("Ensure all quotes and parentheses are properly closed")
            suggestions.append("Verify JOIN conditions are properly specified")
            
        # Permission errors
        elif "permission" in error_str.lower():
            error_type = "Permission Denied"
            suggestions.append("You may not have access to this table or operation")
            suggestions.append("Contact your database administrator for access")
            
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
                    else:
                        column_names[column_name] = 1
                        modified_columns.append(part)
                else:
                    # Not a simple column reference, keep as-is
                    modified_columns.append(part)
            
            # Reconstruct SQL with aliased columns
            new_select_clause = ', '.join(modified_columns)
            new_sql = sql[:select_match.start(1)] + new_select_clause + sql[select_match.end(1):]
            
            return new_sql
            
        except:
            # If anything fails, return original SQL
            return sql
    
    def execute_query(self, sql):
        """Execute SQL query and return results"""
        # Check permission first
        if not self.has_permission:
            return {
                'success': False,
                'error': 'Access Denied: Admin or Developer role required'
            }
        
        try:
            start_time = datetime.now()
            
            # Auto-alias duplicate columns to prevent data loss
            sql = self.auto_alias_duplicate_columns(sql)
            
            # Handle Unicode in SQL query for Python 2.7
            try:
                # Python 2.7 compatibility - check if unicode exists
                if 'unicode' in dir(__builtins__):
                    if isinstance(sql, unicode):
                        # If it's already unicode, encode to utf-8 for processing
                        sql = sql.encode('utf-8', 'ignore')
                # Handle regular strings
                if isinstance(sql, str):
                    # Try to ensure it's properly encoded
                    try:
                        sql = sql.decode('utf-8', 'ignore').encode('utf-8', 'ignore')
                    except:
                        # If decode fails, just use as-is
                        pass
            except:
                # If any Unicode handling fails, continue with original SQL
                pass
            
            # Remove or replace emoji and other problematic Unicode characters
            # This is necessary because SQL Server may not handle all Unicode characters
            if isinstance(sql, str):
                # Remove emoji and other non-ASCII characters from comments
                # First, let's preserve the structure by only cleaning comments
                lines = sql.split('\n')
                cleaned_lines = []
                for line in lines:
                    # Check if line contains a comment
                    comment_pos = line.find('--')
                    if comment_pos >= 0:
                        # Clean only the comment part
                        before_comment = line[:comment_pos]
                        comment = line[comment_pos:]
                        # Replace non-ASCII characters in comments with spaces
                        cleaned_comment = ''.join(c if ord(c) < 128 else ' ' for c in comment)
                        cleaned_lines.append(before_comment + cleaned_comment)
                    else:
                        cleaned_lines.append(line)
                sql = '\n'.join(cleaned_lines)
            
            # Add safety check for destructive queries
            sql_upper = sql.upper()
            # Use word boundaries to match whole words only
            destructive_patterns = [r'\bDELETE\b', r'\bUPDATE\b', r'\bINSERT\b', r'\bDROP\b', r'\bCREATE\b', r'\bALTER\b', r'\bTRUNCATE\b']
            if any(re.search(pattern, sql_upper) for pattern in destructive_patterns):
                if not self.is_developer:
                    return {
                        'error': 'Only developers can execute write operations',
                        'success': False
                    }
            
            # Try to parse column names from the SELECT statement to preserve order
            column_order = []
            select_match = re.search(r'SELECT\s+(TOP\s+\d+\s+)?(.*?)\s+FROM', sql, re.IGNORECASE | re.DOTALL)
            if select_match:
                select_clause = select_match.group(2).strip()
                if select_clause != '*':
                    # Parse column names/aliases from SELECT clause
                    # This is a simplified parser - won't handle all SQL perfectly
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
                    
                    # Extract column names/aliases
                    for part in parts:
                        if part:
                            # Look for alias after AS or just last word
                            alias_match = re.search(r'\s+[Aa][Ss]\s+([^\s]+)$', part)
                            if alias_match:
                                column_order.append(alias_match.group(1).strip('"\'[]'))
                            else:
                                # Get the last identifier (handles table.column)
                                words = re.findall(r'[\w\[\]]+', part)
                                if words:
                                    column_order.append(words[-1].strip('[]'))
            
            # Execute query with encoding error handling
            try:
                results = q.QuerySql(sql)
            except Exception as query_error:
                # Check if it's an encoding error
                try:
                    error_str = str(query_error)
                except:
                    # The error message itself has encoding issues
                    return {
                        'success': False,
                        'error': 'Data encoding error: The query results contain characters that cannot be displayed.',
                        'error_info': {
                            'type': 'Encoding Error',
                            'suggestions': [
                                'The data contains special characters that cannot be processed',
                                'Try selecting specific columns to isolate the problematic data',
                                'Use CAST(column AS VARCHAR) to convert problematic columns',
                                'Add WHERE conditions to exclude rows with special characters',
                                'Check for non-English characters or corrupted data in text fields'
                            ],
                            'raw_error': 'Unable to display error details due to encoding issues'
                        }
                    }
                
                if 'codec' in error_str or 'decode' in error_str or 'Unicode' in error_str:
                    # This is likely an encoding issue with the data
                    # Try to extract meaningful info from the error
                    return {
                        'success': False,
                        'error': 'Data encoding error: The query results contain characters that cannot be displayed. This often happens with special characters or corrupted data.',
                        'error_info': {
                            'type': 'Encoding Error',
                            'suggestions': [
                                'Try selecting specific columns instead of all columns',
                                'Check if any text columns contain special characters',
                                'Consider using CAST or CONVERT to handle the data',
                                'Add WHERE conditions to filter out problematic rows'
                            ],
                            'raw_error': error_str
                        }
                    }
                else:
                    # Re-raise other types of errors
                    raise
            
            # Convert to list of dictionaries
            data = []
            columns = []
            
            if results:
                # Get column names from first row
                first_row = True
                for row in results:
                    if first_row:
                        # Get all available columns from the row
                        available_cols = [col for col in dir(row) if not col.startswith('_')]
                        
                        # If we parsed column order from SQL, use that order
                        if column_order:
                            # Use parsed order, but only for columns that exist
                            ordered_cols = []
                            remaining_cols = set(available_cols)
                            
                            # First add columns in the order specified in SELECT
                            for col in column_order:
                                # Try exact match first
                                if col in available_cols:
                                    ordered_cols.append(col)
                                    remaining_cols.discard(col)
                                else:
                                    # Try case-insensitive match
                                    for avail_col in available_cols:
                                        if col.lower() == avail_col.lower():
                                            ordered_cols.append(avail_col)
                                            remaining_cols.discard(avail_col)
                                            break
                            
                            # Add any remaining columns that weren't in the SELECT
                            # (this handles SELECT * or computed columns)
                            ordered_cols.extend(sorted(remaining_cols))
                            columns = ordered_cols
                        else:
                            # No parsed order, use alphabetical
                            columns = sorted(available_cols)
                        
                        first_row = False
                    
                    row_dict = {}
                    for col in columns:
                        try:
                            value = getattr(row, col, None)
                            
                            # Skip binary/blob columns
                            if col.lower() in ['body', 'content'] and value and isinstance(value, (bytes, bytearray)):
                                value = "[Binary Data - {} bytes]".format(len(value))
                            # Convert datetime objects to strings
                            elif hasattr(value, 'strftime'):
                                value = value.strftime('%Y-%m-%d %H:%M:%S')
                            elif hasattr(value, '__class__') and 'datetime' in str(value.__class__).lower():
                                # Handle TouchPoint's datetime objects
                                value = str(value)
                            elif value is not None and not isinstance(value, (str, int, float, bool)):
                                # Convert any other non-JSON-serializable objects to string
                                value = str(value)
                            
                            # Handle encoding issues - clean up non-UTF8 characters
                            if isinstance(value, str):
                                try:
                                    # First try to handle as UTF-8
                                    value = value.encode('utf-8', 'ignore').decode('utf-8', 'ignore')
                                except:
                                    try:
                                        # Try Latin-1 encoding (common for legacy data)
                                        value = value.encode('latin-1', 'ignore').decode('latin-1', 'ignore')
                                    except:
                                        # Last resort - strip non-ASCII characters
                                        value = ''.join(char if ord(char) < 128 else ' ' for char in str(value))
                                
                                # Replace common problematic characters
                                value = value.replace('\xa0', ' ')  # Non-breaking space
                                value = value.replace('\x00', '')   # Null character
                                value = value.replace('\r\n', ' ')  # Windows line endings
                                value = value.replace('\n', ' ')    # Unix line endings
                                value = value.replace('\t', ' ')    # Tabs
                                
                                # Truncate very long values (likely HTML content)
                                if len(value) > 1000:
                                    value = value[:1000] + '...'
                            row_dict[col] = value
                        except Exception as e:
                            # If we can't get the value, set it to None
                            row_dict[col] = None
                    
                    data.append(row_dict)
            
            exec_time = (datetime.now() - start_time).total_seconds() * 1000
            
            return {
                'success': True,
                'data': data,
                'columns': columns,
                'rowCount': len(data),
                'executionTime': int(exec_time)
            }
            
        except Exception as e:
            # Handle encoding errors in the exception message itself
            try:
                error_msg = str(e)
            except:
                try:
                    # Try to get the error message with encoding handling
                    if 'unicode' in dir(__builtins__):
                        error_msg = unicode(e).encode('utf-8', 'ignore')
                    else:
                        error_msg = repr(e)
                except:
                    # If all else fails, provide a generic message
                    error_msg = "Database error occurred (encoding issue)"
            
            # Parse the error for better user experience
            error_info = self.parse_sql_error(error_msg)
            return {
                'success': False,
                'error': error_msg,
                'error_info': error_info
            }
    
    def render_interface(self):
        """Render the complete interface with multi-tab support"""
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
                position: absolute;
                top: 5px;
                right: 5px;
                width: 30px;
                height: 30px;
                background: #0066cc;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                padding: 0;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 16px;
                color: white;
                transition: all 0.2s ease;
                z-index: 1000;
            }
            
            .panel-toggle:hover {
                background: #0052a3;
            }
            
            .left-panel.collapsed {
                width: 0 !important;
                min-width: 0 !important;
                overflow: hidden;
                padding: 0;
                border: none;
            }
            
            .left-panel.collapsed .panel-toggle {
                position: fixed;
                left: 5px;
                top: 80px;
                right: auto;
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
                overflow-x: auto;
                overflow-y: visible;
                max-width: 100%;
            }
            
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
            
            /* Tabs */
            .tabs {
                display: flex;
                border-bottom: 2px solid #e5e7eb;
                margin-bottom: 15px;
            }
            
            .tab {
                padding: 10px 20px;
                cursor: pointer;
                border-bottom: 2px solid transparent;
                transition: all 0.2s;
            }
            
            .tab:hover {
                background: #f3f4f6;
            }
            
            .tab.active {
                border-bottom-color: #2563eb;
                color: #2563eb;
                font-weight: 500;
            }
            
            /* Buttons */
            .btn-group {
                margin: 10px 0;
            }
            
            .btn-run {
                background: #10b981;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
            }
            
            .btn-run:hover {
                background: #059669;
            }
            
            .btn {
                padding: 8px 16px;
                border: 1px solid #ddd;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                margin-right: 10px;
                background: white;
            }
            
            .btn:hover {
                background: #f3f4f6;
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
        </style>
        
        <div class="query-explorer">
            <!-- Left Panel -->
            <div class="left-panel" id="left-panel">
                <button id="panel-toggle" class="panel-toggle" onclick="togglePanel()" title="Toggle panel">
                    <span id="toggle-icon">◀</span>
                </button>
                <div class="tabs">
                    <div class="tab active" onclick="showTab('schema')">Schema</div>
                    <div class="tab" onclick="showTab('saved')">Saved</div>
                    <div class="tab" onclick="showTab('examples')">Examples</div>
                </div>
                
                <div id="schema-tab" class="tab-content">
                    <input type="text" id="schema-search" placeholder="Search tables..." 
                           class="form-control" style="margin-bottom: 10px;"
                           onkeyup="filterSchema()">
                    <div id="schema-tree" class="schema-tree">
                        <div class="loading">Loading schema...</div>
                    </div>
                </div>
                
                <div id="saved-tab" class="tab-content" style="display:none;">
                    <input type="text" id="saved-search" placeholder="Search saved queries..." 
                           class="form-control" style="margin-bottom: 10px;"
                           onkeyup="filterSaved()">
                    <div id="saved-queries">
                        <div class="loading">Loading saved queries...</div>
                    </div>
                </div>
                
                <div id="examples-tab" class="tab-content" style="display:none;">
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
                
                <div class="btn-group">
                    <button class="btn-run" onclick="executeQuery()">
                        ▶ Run Query (Ctrl+Enter)
                    </button>
                    <button class="btn btn-secondary" onclick="formatSQL()">
                        Format SQL
                    </button>
                    <button class="btn btn-secondary" onclick="saveQuery()">
                        Save Query
                    </button>
                    <button class="btn btn-secondary" onclick="exportResults()">
                        Export CSV
                    </button>
                    <span class="keyboard-shortcuts" onclick="showKeyboardShortcuts()" style="cursor: help; text-decoration: underline;">
                        Keyboard Shortcuts (click for more)
                    </span>
                </div>
                
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
                    <div id="results-info" style="padding: 10px; background: #f3f4f6; display: none;">
                        <span id="row-count"></span> rows returned in <span id="exec-time"></span>ms
                    </div>
                    <div id="results"></div>
                </div>
            </div>
        </div>
        
        <!-- Keyboard Shortcuts Modal -->
        <div id="shortcuts-modal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000;">
            <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 20px; border-radius: 8px; max-width: 500px; max-height: 80vh; overflow-y: auto;">
                <h3 style="margin-top: 0;">Keyboard Shortcuts</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 5px;"><strong>Query Execution:</strong></td><td></td></tr>
                    <tr><td style="padding: 5px 20px;">Ctrl+Enter or F5</td><td>Execute Query (or selected text if any)</td></tr>
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
                <button onclick="document.getElementById('shortcuts-modal').style.display='none'" style="margin-top: 15px; padding: 8px 16px;">Close</button>
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
                        if (!tabs[1].isDirty) {
                            tabs[1].isDirty = true;
                            const tabElement = document.querySelector(`.editor-tab[data-tab-id="1"] .editor-tab-name`);
                            if (tabElement && !tabElement.textContent.endsWith(' *')) {
                                tabElement.textContent = tabElement.textContent + ' *';
                            }
                        }
                    });
                }
                
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
                    toggleIcon.textContent = '◀';
                    
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
                    toggleIcon.textContent = '▶';
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
                    // Save current results
                    const resultsDiv = document.getElementById('results');
                    const resultsInfo = document.getElementById('results-info');
                    tabs[currentTabId].results = resultsDiv.innerHTML;
                    tabs[currentTabId].resultsInfo = {
                        display: resultsInfo.style.display,
                        rowCount: document.getElementById('row-count').textContent,
                        execTime: document.getElementById('exec-time').textContent
                    };
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
                    const resultsDiv = document.getElementById('results');
                    const resultsInfo = document.getElementById('results-info');
                    if (tabs[tabId].results !== null) {
                        resultsDiv.innerHTML = tabs[tabId].results;
                        if (tabs[tabId].resultsInfo) {
                            resultsInfo.style.display = tabs[tabId].resultsInfo.display;
                            document.getElementById('row-count').textContent = tabs[tabId].resultsInfo.rowCount;
                            document.getElementById('exec-time').textContent = tabs[tabId].resultsInfo.execTime;
                        }
                    } else {
                        resultsDiv.innerHTML = '';
                        resultsInfo.style.display = 'none';
                    }
                    
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
            }
            
            function renameTab(tabId) {
                const newName = prompt('Enter new tab name:', tabs[tabId].name);
                if (newName) {
                    tabs[tabId].name = newName;
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
            
            function executeQuery(sql = null) {
                const editor = getCurrentEditor();
                if (!editor) return;
                
                // Check if there's selected text first
                if (!sql) {
                    const selectedText = editor.getSelection();
                    if (selectedText && selectedText.trim()) {
                        // Use selected text
                        sql = selectedText;
                    } else {
                        // Use full editor content
                        sql = editor.getValue();
                    }
                }
                
                if (!sql || sql.trim() === '') {
                    document.getElementById('results').innerHTML = '<div class="alert alert-warning">No SQL to execute</div>';
                    return;
                }
                
                document.getElementById('results').innerHTML = '<div class="loading">Executing query...</div>';
                
                // Get PyScriptForm URL
                var url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                
                // Create form data
                var formData = new FormData();
                formData.append('action', 'execute_query');
                formData.append('sql_query', sql);
                
                // Make AJAX request
                fetch(url, {
                    method: 'POST',
                    body: formData
                })
                .then(response => {
                    if (!response.ok) {
                        throw new Error('Network response was not ok');
                    }
                    return response.text();
                })
                .then(text => {
                    try {
                        // Try to find JSON in the response
                        // TouchPoint sometimes logs the response to console
                        const jsonStart = text.indexOf('{');
                        const jsonEnd = text.lastIndexOf('}') + 1;
                        
                        if (jsonStart >= 0 && jsonEnd > jsonStart) {
                            const jsonStr = text.substring(jsonStart, jsonEnd);
                            const data = JSON.parse(jsonStr);
                            showResults(data);
                            
                            // Ensure panel stays visible after query execution
                            ensurePanelVisible();
                        } else {
                            throw new Error('No JSON found in response');
                        }
                    } catch (e) {
                        // Only show error if we really can't parse the response
                        console.error('Response parsing error:', e);
                        console.log('Raw response:', text);
                        document.getElementById('results').innerHTML = '<div class="alert alert-danger">Error: Invalid response from server. Check console for details.</div>';
                    }
                })
                .catch(error => {
                    document.getElementById('results').innerHTML = '<div class="alert alert-danger">Error: ' + error + '</div>';
                });
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
            
            function showResults(result) {
                const resultsDiv = document.getElementById('results');
                const resultsInfo = document.getElementById('results-info');
                
                if (!result.success) {
                    let errorHtml = '<div class="error-container">';
                    
                    // Check if we have parsed error info
                    if (result.error_info) {
                        errorHtml += '<div class="error-header">';
                        errorHtml += '<span class="error-type">' + result.error_info.type + '</span>';
                        errorHtml += '</div>';
                        
                        if (result.error_info.suggestions && result.error_info.suggestions.length > 0) {
                            errorHtml += '<div class="error-suggestions">';
                            errorHtml += '<strong>Suggestions to fix this error:</strong>';
                            errorHtml += '<ul>';
                            result.error_info.suggestions.forEach(function(suggestion) {
                                errorHtml += '<li>' + suggestion + '</li>';
                            });
                            errorHtml += '</ul>';
                            errorHtml += '</div>';
                        }
                        
                        errorHtml += '<div class="error-details">';
                        errorHtml += '<details>';
                        errorHtml += '<summary>Show technical details</summary>';
                        errorHtml += '<pre class="error-raw">' + escapeHtml(result.error) + '</pre>';
                        errorHtml += '</details>';
                        errorHtml += '</div>';
                    } else {
                        // Fallback to raw error
                        errorHtml += '<div class="alert alert-danger">Error: ' + escapeHtml(result.error) + '</div>';
                    }
                    
                    errorHtml += '</div>';
                    
                    resultsDiv.innerHTML = errorHtml;
                    resultsInfo.style.display = 'none';
                    
                    // Save results to current tab
                    tabs[currentTabId].results = resultsDiv.innerHTML;
                    tabs[currentTabId].resultsInfo = {
                        display: 'none',
                        rowCount: '0',
                        execTime: '0'
                    };
                    return;
                }
                
                if (!result.data || result.data.length === 0) {
                    resultsDiv.innerHTML = '<div style="padding: 20px; color: #6b7280;">No results returned</div>';
                    resultsInfo.style.display = 'none';
                    // Save results to current tab
                    tabs[currentTabId].results = resultsDiv.innerHTML;
                    tabs[currentTabId].resultsInfo = {
                        display: 'none',
                        rowCount: '0',
                        execTime: '0'
                    };
                    return;
                }
                
                // Build table with wrapper for scrolling
                let html = '<div class="results-wrapper"><table class="results-table"><thead><tr>';
                
                result.columns.forEach(col => {
                    html += `<th>${col}</th>`;
                });
                html += '</tr></thead><tbody>';
                
                result.data.forEach(row => {
                    html += '<tr>';
                    result.columns.forEach(col => {
                        let value = row[col];
                        // Escape HTML by default to prevent rendering
                        let displayValue = value !== null && value !== undefined 
                            ? String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;')
                            : '';
                        
                        // Create links for ID columns
                        if (value && !isNaN(value) && value > 0) {
                            const colLower = col.toLowerCase();
                            const intValue = parseInt(value);
                            
                            // Check for PeopleId variations
                            if (colLower === 'peopleid' || 
                                colLower === 'personid' ||
                                colLower === 'leaderid' ||
                                colLower === 'headofhouseholdid' ||
                                colLower === 'createdby' ||
                                colLower === 'modifiedby' ||
                                (colLower === 'id' && result.columns.some(c => c.toLowerCase().includes('name') || c.toLowerCase().includes('email')))) {
                                // Keep the original escaped value for display in the link
                                displayValue = `<a href="/Person2/${intValue}" target="_blank" style="color: #2563eb; text-decoration: underline;">${displayValue}</a>`;
                            }
                            // Check for Organization IDs
                            else if (colLower === 'organizationid' || 
                                     colLower === 'orgid' ||
                                     colLower === 'divisionid') {
                                displayValue = `<a href="/Org/${intValue}" target="_blank" style="color: #2563eb; text-decoration: underline;">${displayValue}</a>`;
                            }
                            // Check for Family ID
                            else if (colLower === 'familyid') {
                                displayValue = `<a href="/Family/${intValue}" target="_blank" style="color: #2563eb; text-decoration: underline;">${displayValue}</a>`;
                            }
                            // Check for Meeting ID
                            else if (colLower === 'meetingid') {
                                displayValue = `<a href="/Meeting/${intValue}" target="_blank" style="color: #2563eb; text-decoration: underline;">${displayValue}</a>`;
                            }
                            // Check for Contribution ID
                            else if (colLower === 'contributionid' || colLower === 'bundleid') {
                                displayValue = `<a href="/Finance/Bundle/${intValue}" target="_blank" style="color: #2563eb; text-decoration: underline;">${displayValue}</a>`;
                            }
                            // Check for Fund ID
                            else if (colLower === 'fundid') {
                                displayValue = `<a href="/Funds" target="_blank" style="color: #2563eb; text-decoration: underline;">${displayValue}</a>`;
                            }
                        }
                        
                        html += `<td>${displayValue != null ? displayValue : ''}</td>`;
                    });
                    html += '</tr>';
                });
                
                html += '</tbody></table></div>';
                resultsDiv.innerHTML = html;
                
                // Show info
                document.getElementById('row-count').textContent = result.rowCount;
                document.getElementById('exec-time').textContent = result.executionTime;
                resultsInfo.style.display = 'block';
                
                // Save results to current tab
                tabs[currentTabId].results = resultsDiv.innerHTML;
                tabs[currentTabId].resultsInfo = {
                    display: 'block',
                    rowCount: result.rowCount,
                    execTime: result.executionTime
                };
                
                // Always ensure panel visibility after showing results
                ensurePanelVisible();
            }
            
            function formatSQL() {
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
                if (!clickTarget || !clickTarget.classList.contains('tab')) {
                    return;
                }
                
                document.querySelectorAll('.tab-content').forEach(tab => {
                    tab.style.display = 'none';
                });
                
                document.querySelectorAll('.tab').forEach(tab => {
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
            }
            
            function filterSaved() {
                const searchTerm = document.getElementById('saved-search').value.toLowerCase();
                const items = document.querySelectorAll('#saved-tab .saved-query-item');
                const groups = document.querySelectorAll('#saved-tab .schema-group');
                
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
                
                const name = prompt('Query name:', tabs[currentTabId].name);
                if (!name) return;
                
                // Show saving indicator
                showMessage('Saving query...', 'info');
                
                // Use AJAX to save the query
                const formData = new FormData();
                formData.append('action', 'save_query');
                formData.append('query_name', name.trim());
                formData.append('query_sql', sql);
                
                // TouchPoint requires /PyScriptForm/ for POST requests
                const actionUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                
                fetch(actionUrl, {
                    method: 'POST',
                    body: formData,
                    credentials: 'same-origin'
                })
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`HTTP error! status: ${response.status}`);
                    }
                    return response.text();
                })
                .then(text => {
                    try {
                        const data = JSON.parse(text);
                        if (data.success) {
                            showMessage(`Query "${name}" saved successfully!`, 'success');
                            // Reload saved queries list
                            loadSavedQueries();
                            // Update tab name if it was renamed
                            if (tabs[currentTabId].name !== name) {
                                tabs[currentTabId].name = name;
                                const tabElement = document.querySelector(`.editor-tab[data-tab-id="${currentTabId}"] .editor-tab-name`);
                                if (tabElement) {
                                    tabElement.textContent = name;
                                }
                            }
                            // Mark as saved (remove dirty indicator)
                            tabs[currentTabId].isDirty = false;
                            const tabElement = document.querySelector(`.editor-tab[data-tab-id="${currentTabId}"] .editor-tab-name`);
                            if (tabElement && tabElement.textContent.endsWith(' *')) {
                                tabElement.textContent = tabElement.textContent.replace(' *', '');
                            }
                        } else {
                            showMessage('Error saving query: ' + (data.error || 'Unknown error'), 'error');
                        }
                    } catch (e) {
                        // If response isn't JSON, it might be an HTML error page
                        if (text.includes('<!DOCTYPE') || text.includes('<html')) {
                            showMessage('Error: Server returned HTML instead of JSON. Check permissions.', 'error');
                        } else {
                            showMessage('Error parsing response: ' + e.message, 'error');
                        }
                    }
                })
                .catch(error => {
                    console.error('Save error:', error);
                    showMessage('Error saving query: ' + error.message, 'error');
                });
            }
            
            function exportResults() {
                // Check if we have results to export
                const resultsTable = document.querySelector('#results .results-table');
                if (!resultsTable) {
                    alert('No results to export. Please run a query first.');
                    return;
                }
                
                // Get all rows including headers
                const rows = resultsTable.querySelectorAll('tr');
                const csvData = [];
                
                // Process each row
                rows.forEach((row, index) => {
                    const cols = row.querySelectorAll(index === 0 ? 'th' : 'td');
                    const rowData = [];
                    
                    cols.forEach(col => {
                        // Get text content and handle special characters
                        let text = col.textContent.trim();
                        
                        // Remove any HTML (like links) and get just text
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = col.innerHTML;
                        text = tempDiv.textContent || tempDiv.innerText || '';
                        
                        // Escape quotes and wrap in quotes if contains comma, newline, or quotes
                        if (text.includes('"')) {
                            text = text.replace(/"/g, '""');
                        }
                        if (text.includes(',') || text.includes('\\n') || text.includes('"')) {
                            text = '"' + text + '"';
                        }
                        
                        rowData.push(text);
                    });
                    
                    csvData.push(rowData.join(','));
                });
                
                // Create CSV content
                const csvContent = csvData.join('\\n');
                
                // Create blob and download
                const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement('a');
                const url = URL.createObjectURL(blob);
                
                // Generate filename with timestamp
                const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
                const filename = 'query_results_' + timestamp + '.csv';
                
                link.setAttribute('href', url);
                link.setAttribute('download', filename);
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                
                // Show success message
                const rowCount = rows.length - 1; // Exclude header row
                showMessage('Exported ' + rowCount + ' rows to ' + filename, 'success');
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
                var url = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
                var formData = new FormData();
                formData.append('action', 'get_saved');
                
                fetch(url, {
                    method: 'POST',
                    body: formData
                })
                .then(response => response.text())
                .then(text => {
                    try {
                        // Try to find JSON in the response
                        const jsonStart = text.indexOf('{');
                        const jsonEnd = text.lastIndexOf('}') + 1;
                        
                        if (jsonStart >= 0 && jsonEnd > jsonStart) {
                            text = text.substring(jsonStart, jsonEnd);
                        }
                        
                        const data = JSON.parse(text);
                        console.log('Saved queries loaded:', data);
                        if (data.debug) {
                            console.log('Debug info:', data.debug);
                        }
                        if (data.error) {
                            console.error('Error from backend:', data.error);
                        }
                        if (data.queries && data.queries.length > 0) {
                            const savedDiv = document.getElementById('saved-queries');
                            let html = '';
                            
                            
                            // Count valid queries
                            const validCount = data.queries.filter(q => q.name && q.name !== 'Unnamed' && q.id > 0).length;
                            
                            // Simple approach - show all queries in one list
                            html += `
                                <div class="schema-group" id="all-queries-group">
                                    <div class="schema-group-header" onclick="toggleGroup('all-queries-group')">
                                        <span class="expand-icon">▼</span>
                                        <span>SQL Scripts (${validCount})</span>
                                    </div>
                                    <div class="schema-group-content">
                            `;
                            
                            // Sort queries by name
                            data.queries.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
                            
                            data.queries.forEach((query) => {
                                // Skip invalid entries
                                if (query.name && query.name !== 'Unnamed' && query.id > 0) {
                                    html += `
                                        <div class="table-item saved-query-item" data-query-id="${query.id}">
                                            <span class="table-name" onclick="openSavedQuery(${query.id})" title="${query.name}">
                                                ${query.name}
                                            </span>
                                        </div>
                                    `;
                                }
                            });
                            
                            html += '</div></div>';
                            
                            savedDiv.innerHTML = html;
                            window.savedQueries = data.queries;
                        } else {
                            document.getElementById('saved-queries').innerHTML = 
                                '<div style="padding: 10px; color: #666;">No saved SQL scripts found</div>';
                        }
                    } catch (e) {
                        console.error('Error parsing saved queries:', e);
                        document.getElementById('saved-queries').innerHTML = 
                            '<div style="color: red; padding: 10px;">Error loading saved queries</div>';
                    }
                })
                .catch(error => {
                    console.error('Error loading saved queries:', error);
                    document.getElementById('saved-queries').innerHTML = 
                        '<div style="color: red; padding: 10px;">Error loading saved queries</div>';
                });
            }
            
            function openSavedQuery(queryId) {
                // Find the query by ID
                const query = window.savedQueries.find(q => q.id === queryId);
                if (!query) return;
                
                // Create a new tab
                newTab();
                
                // Load the SQL into the new tab
                const editor = getCurrentEditor();
                if (editor) {
                    editor.setValue(query.sql);
                    
                    // Update tab name
                    tabs[currentTabId].name = query.name;
                    const tabElement = document.querySelector(`.editor-tab[data-tab-id="${currentTabId}"] .editor-tab-name`);
                    if (tabElement) {
                        tabElement.textContent = tabs[currentTabId].name;
                    }
                }
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
        # Check for non-AJAX save action first (page reload) - DEPRECATED, using AJAX now
        if False and hasattr(model.Data, 'save_action') and model.Data.save_action == 'save_query':
            # Handle save query during page load where model.WriteContentSql works
            query_name = getattr(model.Data, 'query_name', '')
            query_sql = getattr(model.Data, 'query_sql', '')
            
            if query_name and query_sql:
                try:
                    model.WriteContentSql(query_name, query_sql, "SQL Query Explorer")
                    # Show only a nice confirmation page
                    success_html = """<html><head><title>Query Saved - TouchPoint</title><style>
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f5f5f5; }
.success-container { background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; max-width: 500px; }
.success-icon { font-size: 60px; color: #28a745; margin-bottom: 20px; }
h2 { color: #333; margin-bottom: 10px; }
.query-name { color: #0066cc; font-weight: bold; }
.message { color: #666; margin: 20px 0; }
.close-button { background: #0066cc; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 16px; }
.close-button:hover { background: #0052a3; }
</style></head><body>
<div class="success-container">
<div class="success-icon">✓</div>
<h2>Query Saved Successfully!</h2>
<p>The query <span class="query-name">"{0}"</span> has been saved to Special Content.</p>
<p class="message">You can now close this tab and return to the Query Explorer.</p>
<button class="close-button" onclick="window.close()">Close This Tab</button>
</div></body></html>""".format(query_name.replace('"', '&quot;'))
                    print(success_html)
                except Exception as e:
                    # Show error page
                    error_html = """<html><head><title>Error Saving Query - TouchPoint</title><style>
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f5f5f5; }
.error-container { background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; max-width: 500px; }
.error-icon { font-size: 60px; color: #dc3545; margin-bottom: 20px; }
h2 { color: #333; margin-bottom: 10px; }
.error-message { color: #dc3545; margin: 20px 0; padding: 10px; background: #f8d7da; border-radius: 4px; }
.close-button { background: #6c757d; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 16px; }
.close-button:hover { background: #5a6268; }
</style></head><body>
<div class="error-container">
<div class="error-icon">✗</div>
<h2>Error Saving Query</h2>
<p class="error-message">{0}</p>
<button class="close-button" onclick="window.close()">Close This Tab</button>
</div></body></html>""".format(str(e).replace('"', '&quot;'))
                    print(error_html)
            else:
                # Show validation error page
                warning_html = """<html><head><title>Invalid Request - TouchPoint</title><style>
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f5f5f5; }
.warning-container { background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; max-width: 500px; }
.warning-icon { font-size: 60px; color: #ffc107; margin-bottom: 20px; }
h2 { color: #333; margin-bottom: 10px; }
.message { color: #666; margin: 20px 0; }
.close-button { background: #6c757d; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 16px; }
.close-button:hover { background: #5a6268; }
</style></head><body>
<div class="warning-container">
<div class="warning-icon">⚠</div>
<h2>Invalid Request</h2>
<p class="message">Query name and SQL content are required to save a query.</p>
<button class="close-button" onclick="window.close()">Close This Tab</button>
</div></body></html>"""
                print(warning_html)
        # Check if this is an AJAX request
        elif hasattr(model.Data, 'action') and model.Data.action:
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
                    
                    print(json.dumps({
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
                        
                        print(json.dumps({'columns': columns_list}))
                    else:
                        print(json.dumps({'error': 'No table name provided'}))
                
                elif action == 'execute_query':
                    # Execute the query
                    sql = model.Data.sql_query if hasattr(model.Data, 'sql_query') else ''
                    if sql:
                        # Check if explorer has the execute_query method
                        if hasattr(explorer, 'execute_query'):
                            result = explorer.execute_query(sql)
                            # Try to serialize to JSON, handling encoding issues
                            try:
                                print(safe_json_dumps(result))
                            except (UnicodeDecodeError, UnicodeError, Exception) as json_error:
                                # JSON serialization failed due to encoding
                                # Check if it's specifically an encoding error
                                error_msg = str(json_error) if hasattr(json_error, '__str__') else 'Unknown error'
                                if 'codec' in error_msg or 'decode' in error_msg or 'encode' in error_msg:
                                    print(json.dumps({
                                    'success': False,
                                    'error': 'The query results contain special characters that cannot be displayed.',
                                    'error_info': {
                                        'type': 'Display Error',
                                        'suggestions': [
                                            'Try selecting fewer columns to isolate the problematic data',
                                            'Use CONVERT(VARCHAR, column) for text fields with special characters',
                                            'Add TOP 10 to limit results and identify problematic rows',
                                            'Check for names with accented characters (é, ñ, ó, etc.)',
                                            'Consider using ASCII() function to identify special characters'
                                        ],
                                        'raw_error': 'JSON encoding error - special characters in results'
                                    }
                                }))
                                else:
                                    # Other type of error
                                    print(json.dumps({
                                        'success': False,
                                        'error': 'Error processing query results: ' + error_msg
                                    }))
                        else:
                            print(json.dumps({'success': False, 'error': 'Explorer not properly initialized'}))
                    else:
                        print(json.dumps({'success': False, 'error': 'No SQL query provided'}))
                
                elif action == 'get_full_schema':
                    # Get complete schema with all columns
                    try:
                        schema = explorer.get_full_schema()
                        print(json.dumps({'success': True, 'schema': schema}))
                    except Exception as e:
                        print(json.dumps({'success': False, 'error': str(e)}))
                
                elif action == 'get_examples':
                    # Return example queries
                    examples = explorer.get_common_queries()
                    print(json.dumps({'examples': examples}))
                
                elif action == 'get_saved':
                    # Get saved queries
                    try:
                        queries = explorer.get_saved_queries()
                        queries_list = []
                        
                        # Debug: log query count
                        query_count = 0
                        if queries:
                            for query in queries:
                                query_count += 1
                                try:
                                    # Handle Body field that might have encoding issues
                                    body_text = getattr(query, 'Body', '')
                                    if body_text and isinstance(body_text, str):
                                        # Clean up any encoding issues
                                        body_text = body_text.encode('utf-8', 'ignore').decode('utf-8', 'ignore')
                                    elif body_text:
                                        body_text = str(body_text)
                                    
                                    queries_list.append({
                                        'id': getattr(query, 'Id', 0),
                                        'name': getattr(query, 'Name', 'Unnamed'),
                                        'sql': body_text
                                    })
                                except Exception as e:
                                    # Skip problematic records instead of adding error entries
                                    continue
                        
                        # Always return something to help debug
                        print(json.dumps({
                            'queries': queries_list,
                            'debug': {
                                'total_rows': query_count,
                                'processed': len(queries_list)
                            }
                        }))
                    except Exception as e:
                        print(json.dumps({'queries': [], 'error': str(e), 'debug': 'Exception in get_saved'}))
                
                elif action == 'save_query':
                    # Save a new query
                    try:
                        # Check both possible field names for compatibility
                        name = getattr(model.Data, 'query_name', '') or getattr(model.Data, 'name', '')
                        sql = getattr(model.Data, 'query_sql', '') or getattr(model.Data, 'sql', '')
                        
                        if name and sql:
                            try:
                                success = explorer.save_query(name, sql)
                                if success:
                                    print(json.dumps({'success': True, 'message': 'Query saved successfully'}))
                                else:
                                    print(json.dumps({'success': False, 'error': 'Failed to save query'}))
                            except Exception as save_error:
                                print(json.dumps({'success': False, 'error': str(save_error)}))
                        else:
                            print(json.dumps({'success': False, 'error': 'Name and SQL are required'}))
                    except Exception as e:
                        print(json.dumps({'success': False, 'error': 'Action handler error: ' + str(e)}))
                
                else:
                    # Unknown action
                    print(json.dumps({'success': False, 'error': 'Unknown action: ' + action}))
                    
            except Exception as e:
                # Always return JSON for AJAX requests
                print(json.dumps({'success': False, 'error': str(e)}))
        
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
                    print(json.dumps(result))
                else:
                    print(json.dumps({'success': False, 'error': 'No SQL query provided'}))
            else:
                print(json.dumps({'success': False, 'error': 'Access Denied: Admin or Developer role required'}))
        else:
            print('<div class="alert alert-danger">Access Denied: Admin or Developer role required</div>')
        
except Exception as e:
    # Check if this is an AJAX request
    if hasattr(model.Data, 'action') and model.Data.action:
        # Return JSON error for AJAX requests
        print(json.dumps({'success': False, 'error': 'Initialization error: ' + str(e)}))
    else:
        # Return HTML error for page load
        print('<div class="alert alert-danger">Error initializing Query Explorer: {0}</div>'.format(str(e)))
