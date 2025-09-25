#roles=Edit

#####################################################################
# TPxi Membership Analysis Report - 10 Year Trends
#####################################################################
# Purpose: Comprehensive analysis of church membership trends over 10 years
# for leadership to understand growth patterns, demographics shifts,
# attendance behaviors, and member engagement pathways
#
# Features:
# 1. 10-year membership growth trends with yearly breakdowns
# 2. Age demographics evolution over time
# 3. Connection pathways analysis (how people found the church)
# 4. Attendance patterns before/after membership
# 5. Family unit analysis and trends
# 6. Retention and attrition metrics
# 7. Campus-specific breakdowns (if multi-campus)
# 8. Interactive charts and visualizations
#
# Upload Instructions:
# 1. Admin > Advanced > Special Content > Python
# 2. New Python Script File
# 3. Name: TPxi_MembershipAnalysisReport
# 4. Run manually or schedule quarterly/annually
#
# Configuration: Update all settings below for your church
#
# Note: This report uses the current Age field instead of BirthDate
# for age calculations, which provides an approximation of age groups
#####################################################################

#written by: Ben Swaby
#email:bswaby@fbchtn.org

# ===== CONFIGURATION SECTION =====
# Place all configuration at the top for easy customization
class Config:
    # Report settings
    YEARS_TO_ANALYZE = 5  # Number of years to look back for main analysis (default: 5)
    COHORT_YEARS_TO_ANALYZE = 10  # Number of years for cohort retention analysis (default: 10)
    REPORT_TITLE = "Church Membership Analysis Report"
    
    # Fiscal Year Settings
    FISCAL_YEAR_START_MONTH = 10  # October = 10, January = 1, etc.
    FISCAL_YEAR_START_DAY = 1     # Day of month fiscal year starts
    USE_FISCAL_YEAR = True        # Set to False to use calendar year
    
    # Member status configuration
    MEMBER_STATUS_ID = 10  # Typically 10 = Member
    PREVIOUS_MEMBER_STATUS_ID = 40  # Typically 40 = Previous Member
    PROSPECT_STATUS_ID = 20  # Typically 20 = Prospect
    GUEST_STATUS_ID = 30  # Typically 30 = Guest
    
    # Age group definitions
    # Note: Uses current age, not age at joining (BirthDate field not available)
    AGE_GROUPS = [
        ("Children", 0, 12),
        ("Teens", 13, 17),
        ("Young Adults", 18, 29),
        ("Adults", 30, 49),
        ("Mature Adults", 50, 64),
        ("Seniors", 65, 150)
    ]
    
    # Attendance tracking
    # Note: For worship, we use MaxCount (total attendance) from Meetings table
    # MaxCount typically holds the total count, HeadCount is used as fallback
    # For Connect Groups, we track individual attendance records (AttendanceFlag = 1)
    PRIMARY_WORSHIP_PROGRAM_ID = 1124  # Main worship service program ID (updated)
    PRIMARY_WORSHIP_PROGRAM_NAME = "Worship"
    SMALL_GROUP_PROGRAM_ID = 1128  # Small groups/Connect groups program ID
    SMALL_GROUP_PROGRAM_NAME = "Connect Groups"
    
    # Campus settings (0 = all campuses)
    CAMPUS_ID = 0  # Set to specific campus ID to filter, or 0 for all
    
    # Display settings
    SHOW_CAMPUS_BREAKDOWN = False  # Show campus-specific analysis (set to False if no Campus table)
    SHOW_ORIGIN_ANALYSIS = False  # Show how people found the church (set to False if no Origin table)
    SHOW_RETENTION_METRICS = True  # Show retention/attrition analysis
    SHOW_FAMILY_ANALYSIS = True  # Show family unit trends
    SHOW_ATTENDANCE_IMPACT = True  # Show member engagement analysis (individual attendance patterns)
    
    # Chart colors (for consistency)
    CHART_COLORS = [
        "#667eea",  # Purple
        "#48bb78",  # Green
        "#ed8936",  # Orange
        "#e53e3e",  # Red
        "#38b2ac",  # Teal
        "#805ad5",  # Violet
        "#d69e2e",  # Yellow
        "#3182ce",  # Blue
    ]
    
    # Export settings
    ENABLE_EXPORT = True  # Allow data export to CSV
    
    # Performance settings
    USE_CACHING = True  # Cache expensive queries
    CACHE_DURATION = 60  # Cache duration in minutes

# Set page header
model.Header = Config.REPORT_TITLE

def main():
    """Main execution function"""
    try:
        # Check permissions
        if not check_permissions():
            return
        
        # Test basic database connectivity first
        try:
            test_count = q.QuerySqlScalar("""
                SELECT COUNT(*) FROM People WHERE MemberStatusId = {} AND IsDeceased = 0
            """.format(Config.MEMBER_STATUS_ID))
            print('<script>console.log("Found {} active members");</script>'.format(test_count))
            
            # Test if JoinDate column exists
            join_date_test = q.QuerySqlScalar("""
                SELECT COUNT(*) FROM People 
                WHERE MemberStatusId = {} 
                  AND IsDeceased = 0 
                  AND JoinDate IS NOT NULL
                  AND JoinDate >= DATEADD(year, -10, GETDATE())
            """.format(Config.MEMBER_STATUS_ID))
            print('<script>console.log("Found {} members with join dates in last 10 years");</script>'.format(join_date_test))
            
        except Exception as e:
            print('<div class="alert alert-danger">Database connectivity test failed: {}</div>'.format(str(e)))
            print('<div class="alert alert-info">This report requires access to People table with JoinDate field.</div>')
            return
        
        # Calculate date range
        end_date = model.DateTime
        start_date_str = q.QuerySqlScalar("""
            SELECT CONVERT(varchar, DATEADD(year, -{}, GETDATE()), 120)
        """.format(Config.YEARS_TO_ANALYZE))
        
        # For cohort analysis, use longer timeframe
        cohort_start_date_str = q.QuerySqlScalar("""
            SELECT CONVERT(varchar, DATEADD(year, -{}, GETDATE()), 120)
        """.format(Config.COHORT_YEARS_TO_ANALYZE))
        
        print('<script>console.log("Date range: {} to {}");</script>'.format(start_date_str, end_date.ToString("yyyy-MM-dd")))
        
        # Show loading message
        print('<div id="loadingMessage" class="alert alert-info"><i class="fa fa-spinner fa-spin"></i> Generating {} Year Analysis Report (with {} Year Cohort Analysis)... This may take a moment.</div>'.format(Config.YEARS_TO_ANALYZE, Config.COHORT_YEARS_TO_ANALYZE))
        
        # Gather all analytics data (pass cohort start date as well)
        analytics = gather_analytics_data(start_date_str, end_date.ToString("yyyy-MM-dd"), cohort_start_date_str)
        
        # Generate report
        generate_report(analytics, start_date_str, end_date.ToString("yyyy-MM-dd"))
        
        # Hide loading message
        print('<script>document.getElementById("loadingMessage").style.display = "none";</script>')
        
    except Exception as e:
        print('<div class="alert alert-danger">Error generating report: {}</div>'.format(str(e)))
        if model.UserIsInRole("Developer"):
            import traceback
            print('<pre>{}</pre>'.format(traceback.format_exc()))

def check_permissions():
    """Check if user has required permissions"""
    if not model.UserIsInRole("Edit") and not model.UserIsInRole("Admin"):
        print('<div class="alert alert-danger"><i class="fa fa-lock"></i> You need Edit or Admin role to access this report.</div>')
        return False
    return True

def gather_analytics_data(start_date, end_date, cohort_start_date_str=None):
    """Gather all analytics data for the report"""
    analytics = {}
    
    # Use cohort_start_date_str for retention metrics if provided
    if not cohort_start_date_str:
        cohort_start_date_str = start_date
    
    try:
        # 1. Overall membership trends by year
        print('<script>console.log("Loading yearly trends...");</script>')
        analytics['yearly_trends'] = get_yearly_membership_trends(start_date, end_date)
        
        # 2. Age demographics over time
        print('<script>console.log("Loading age demographics...");</script>')
        analytics['age_demographics'] = get_age_demographics_trends(start_date, end_date)
        
        # 3. Connection pathways
        if Config.SHOW_ORIGIN_ANALYSIS:
            print('<script>console.log("Loading origin trends...");</script>')
            analytics['origins'] = get_origin_trends(start_date, end_date)
        
        # 4. Attendance impact analysis
        if Config.SHOW_ATTENDANCE_IMPACT:
            print('<script>console.log("Loading attendance impact...");</script>')
            analytics['attendance_impact'] = get_attendance_impact_analysis(start_date, end_date)
        
        # 5. Family analysis
        if Config.SHOW_FAMILY_ANALYSIS:
            print('<script>console.log("Loading family trends...");</script>')
            analytics['family_trends'] = get_family_trends(start_date, end_date)
        
        # 6. Retention metrics (use cohort years)
        if Config.SHOW_RETENTION_METRICS:
            print('<script>console.log("Loading retention metrics...");</script>')
            analytics['retention'] = get_retention_metrics(cohort_start_date_str, end_date)
        
        # 7. Campus breakdown
        if Config.SHOW_CAMPUS_BREAKDOWN and Config.CAMPUS_ID == 0:
            print('<script>console.log("Loading campus breakdown...");</script>')
            analytics['campus_breakdown'] = get_campus_breakdown(start_date, end_date)
        
        # 8. Current totals
        print('<script>console.log("Loading current totals...");</script>')
        analytics['current_totals'] = get_current_totals()
        
    except Exception as e:
        print('<div class="alert alert-warning">Error loading analytics data: {}</div>'.format(str(e)))
        raise
    
    return analytics

def get_yearly_membership_trends(start_date, end_date):
    """Get membership trends by year"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    sql = """
    SELECT 
        {} AS JoinYear,
        COUNT(*) AS NewMembers,
        COUNT(DISTINCT p.FamilyId) AS NewFamilies,
        -- Gender breakdown
        SUM(CASE WHEN p.GenderId = 1 THEN 1 ELSE 0 END) AS Males,
        SUM(CASE WHEN p.GenderId = 2 THEN 1 ELSE 0 END) AS Females,
        -- Age at joining (using current age as approximation)
        AVG(p.Age) AS AvgAgeAtJoining,
        -- Marital status
        SUM(CASE WHEN p.MaritalStatusId = 20 THEN 1 ELSE 0 END) AS Married,
        SUM(CASE WHEN p.MaritalStatusId = 10 THEN 1 ELSE 0 END) AS Single
    FROM People p
    WHERE p.MemberStatusId = {}
      AND p.JoinDate >= '{}'
      AND p.JoinDate <= '{}'
      AND p.IsDeceased = 0
    GROUP BY {}
    ORDER BY {} DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        fiscal_year_sql,
        fiscal_year_sql
    )
    
    return q.QuerySql(sql)

def get_age_demographics_trends(start_date, end_date):
    """Get age demographics trends over time"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    sql = """
    SELECT 
        {} AS JoinYear,
    """.format(fiscal_year_sql)
    
    # Add age group columns dynamically
    for group_name, min_age, max_age in Config.AGE_GROUPS:
        sql += """
        SUM(CASE WHEN p.Age BETWEEN {} AND {} THEN 1 ELSE 0 END) AS {},
        """.format(min_age, max_age, group_name.replace(" ", ""))
    
    sql += """
        SUM(CASE WHEN p.Age IS NULL THEN 1 ELSE 0 END) AS UnknownAge
    FROM People p
    WHERE p.MemberStatusId = {}
      AND p.JoinDate >= '{}'
      AND p.JoinDate <= '{}'
      AND p.IsDeceased = 0
    GROUP BY {}
    ORDER BY {} DESC
    """.format(
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        fiscal_year_sql,
        fiscal_year_sql
    )
    
    return q.QuerySql(sql)

def get_origin_trends(start_date, end_date):
    """Get trends in how people found the church"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    # Note: Origin table may not exist in all TouchPoint instances
    # Fallback to simple count by year if Origin data is not available
    try:
        sql = """
        SELECT 
            {} AS JoinYear,
            ISNULL(o.Description, 'Not Specified') AS Origin,
            COUNT(*) AS Count
        FROM People p
        LEFT JOIN Origin o ON p.OriginId = o.Id
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        GROUP BY {}, o.Description
        ORDER BY {} DESC, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)
    except:
        # Fallback if Origin table doesn't exist
        sql = """
        SELECT 
            {} AS JoinYear,
            'Not Available' AS Origin,
            COUNT(*) AS Count
        FROM People p
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        GROUP BY {}
        ORDER BY {}, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)

def get_attendance_impact_analysis(start_date, end_date):
    """Analyze attendance patterns before and after membership
    
    This combines:
    1. Worship headcount data (from Meetings table) - using MaxCount for total attendance
    2. Connect Group individual attendance (from Attend table)
    
    It compares attendance patterns 6 months before and after membership.
    """
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    # First check if we can get worship headcount data
    # Check both HeadCount and MaxCount columns
    try:
        headcount_test = q.QuerySqlScalar("""
            SELECT COUNT(*) 
            FROM Meetings m
            WHERE (m.HeadCount > 0 OR m.MaxCount > 0)
              AND m.MeetingDate >= DATEADD(month, -1, GETDATE())
        """)
        use_headcount = headcount_test > 0
        
        # Debug: Check worship meetings specifically
        worship_meeting_test = q.QuerySqlTop1("""
            SELECT 
                COUNT(*) as TotalMeetings,
                SUM(CASE WHEN m.MaxCount > 0 THEN 1 ELSE 0 END) as WithMaxCount,
                SUM(CASE WHEN m.HeadCount > 0 THEN 1 ELSE 0 END) as WithHeadCount,
                AVG(CASE WHEN m.MaxCount > 0 THEN m.MaxCount ELSE m.HeadCount END) as AvgCount
            FROM Meetings m
            INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
            WHERE os.ProgId = {}
              AND m.MeetingDate >= DATEADD(month, -3, GETDATE())
        """.format(Config.PRIMARY_WORSHIP_PROGRAM_ID))
        
        if worship_meeting_test:
            print('<script>console.log("Worship meetings last 3 months: Total={}, WithMaxCount={}, WithHeadCount={}, AvgCount={}");</script>'.format(
                safe_get_value(worship_meeting_test, 'TotalMeetings', 0),
                safe_get_value(worship_meeting_test, 'WithMaxCount', 0),
                safe_get_value(worship_meeting_test, 'WithHeadCount', 0),
                int(safe_get_value(worship_meeting_test, 'AvgCount', 0))
            ))
            
        # Also check just the meetings table structure
        sample_meeting = q.QuerySqlTop1("""
            SELECT TOP 1 m.*, o.OrganizationName
            FROM Meetings m
            INNER JOIN Organizations o ON m.OrganizationId = o.OrganizationId
            INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
            WHERE os.ProgId = {}
              AND m.MeetingDate >= DATEADD(month, -1, GETDATE())
              AND (m.MaxCount > 0 OR m.HeadCount > 0)
            ORDER BY m.MeetingDate DESC
        """.format(Config.PRIMARY_WORSHIP_PROGRAM_ID))
        
        if sample_meeting:
            print('<script>console.log("Sample worship meeting - Org: {}, MaxCount: {}, HeadCount: {}");</script>'.format(
                safe_get_value(sample_meeting, 'OrganizationName', 'Unknown'),
                safe_get_value(sample_meeting, 'MaxCount', 0),
                safe_get_value(sample_meeting, 'HeadCount', 0)
            ))
    except Exception as e:
        use_headcount = False
        print('<script>console.log("Error checking worship meetings: {}");</script>'.format(str(e).replace('"', '\"')))
    
    # Get worship data separately to avoid aggregate function errors
    worship_data = {}
    if use_headcount:
        # Get fiscal years first
        year_sql = """
        SELECT DISTINCT {} AS Year
        FROM People p
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        ORDER BY Year
        """.format(
            get_fiscal_year_sql().format("p.JoinDate"),
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date
        )
        
        years = q.QuerySql(year_sql)
        
        # For each year, get average worship headcount
        for year_row in years:
            year = safe_get_value(year_row, 'Year', None)
            if not year:
                continue
                
            # Convert year to integer to avoid "invalid integer number literal" errors
            try:
                year_int = int(year)
            except (ValueError, TypeError):
                print('<script>console.log("Invalid year value - skipping");</script>')
                continue
            
            # Calculate fiscal year date range
            # Python 2.7.3 doesn't support {:02d} format
            if Config.USE_FISCAL_YEAR:
                if Config.FISCAL_YEAR_START_MONTH >= 10:  # Oct, Nov, Dec
                    month_str = str(Config.FISCAL_YEAR_START_MONTH).zfill(2)
                    day_str = str(Config.FISCAL_YEAR_START_DAY).zfill(2)
                    end_month_str = str(Config.FISCAL_YEAR_START_MONTH - 1).zfill(2)
                    year_start = "{}-{}-{}".format(year_int - 1, month_str, day_str)
                    year_end = "{}-{}-30".format(year_int, end_month_str)
                else:
                    month_str = str(Config.FISCAL_YEAR_START_MONTH).zfill(2)
                    day_str = str(Config.FISCAL_YEAR_START_DAY).zfill(2)
                    end_month_str = str(Config.FISCAL_YEAR_START_MONTH - 1).zfill(2)
                    year_start = "{}-{}-{}".format(year_int, month_str, day_str)
                    year_end = "{}-{}-28".format(year_int + 1, end_month_str)
            else:
                year_start = "{}-01-01".format(year_int)
                year_end = "{}-12-31".format(year_int)
            
            # Debug: Let's see what organizations are under worship
            debug_sql = """
            SELECT COUNT(*) as OrgCount, COUNT(DISTINCT os.OrgId) as UniqueOrgs
            FROM OrganizationStructure os
            WHERE os.ProgId = {}
            """.format(Config.PRIMARY_WORSHIP_PROGRAM_ID)
            
            try:
                org_count = q.QuerySqlTop1(debug_sql)
                org_count_val = safe_get_value(org_count, 'OrgCount', 0)
                print('<script>console.log("Worship orgs for year " + "{}" + ": " + "{}" + " orgs");</script>'.format(year_int, int(org_count_val) if org_count_val else 0))
            except:
                pass
            
            # Updated worship query - aggregate by week to get typical Sunday attendance
            # Since worship is typically once per week, we sum by week and then average those weekly totals
            worship_sql = """
            WITH WeeklyTotals AS (
                SELECT 
                    DATEPART(year, m.MeetingDate) as MeetingYear,
                    DATEPART(week, m.MeetingDate) as MeetingWeek,
                    SUM(CASE 
                        WHEN m.MaxCount > 0 THEN m.MaxCount 
                        WHEN m.HeadCount > 0 THEN m.HeadCount
                        ELSE 0 
                    END) AS WeeklyTotal
                FROM Meetings m
                INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                WHERE m.MeetingDate >= '{}'
                  AND m.MeetingDate <= '{}'
                  AND os.ProgId = {}
                GROUP BY DATEPART(year, m.MeetingDate), DATEPART(week, m.MeetingDate)
            )
            SELECT AVG(CAST(WeeklyTotal AS FLOAT)) AS AvgHeadcount
            FROM WeeklyTotals
            WHERE WeeklyTotal > 0
            """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID)
            
            try:
                avg_headcount = q.QuerySqlScalar(worship_sql)
                if avg_headcount:
                    worship_data[year_int] = float(avg_headcount)
                    avg_int = int(avg_headcount) if avg_headcount else 0
                    print('<script>console.log("Year " + "{}" + " worship avg: " + "{}");</script>'.format(year_int, avg_int))
                    
                    # Also get meeting count for context
                    # Get count of unique weeks with worship meetings
                    week_count_sql = """
                    SELECT COUNT(DISTINCT DATEPART(year, m.MeetingDate) * 100 + DATEPART(week, m.MeetingDate)) as UniqueWeeks
                    FROM Meetings m
                    INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                    WHERE m.MeetingDate >= '{}'
                      AND m.MeetingDate <= '{}'
                      AND os.ProgId = {}
                    """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID)
                    week_count = q.QuerySqlScalar(week_count_sql)
                    week_count_val = int(week_count) if week_count else 0
                    print('<script>console.log("Year " + "{}" + " had " + "{}" + " worship weeks");</script>'.format(year_int, week_count_val))
                else:
                    # Try a simpler query to see if there's any data
                    simple_test = q.QuerySqlScalar("""
                        SELECT COUNT(*) 
                        FROM Meetings m
                        INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                        WHERE m.MeetingDate >= '{}'
                          AND m.MeetingDate <= '{}'
                          AND os.ProgId = {}
                    """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID))
                    simple_test_val = int(simple_test) if simple_test else 0
                    print('<script>console.log("Year " + "{}" + " has " + "{}" + " worship meetings but no headcount data");</script>'.format(year_int, simple_test_val))
                    
                    # Let's check a sample of what MaxCount/HeadCount values look like
                    sample_sql = """
                    SELECT TOP 5 m.MeetingDate, o.OrganizationName, m.MaxCount, m.HeadCount
                    FROM Meetings m
                    INNER JOIN Organizations o ON m.OrganizationId = o.OrganizationId
                    INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                    WHERE m.MeetingDate >= '{}'
                      AND m.MeetingDate <= '{}'
                      AND os.ProgId = {}
                    ORDER BY m.MeetingDate DESC
                    """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID)
                    
                    samples = q.QuerySql(sample_sql)
                    if samples:
                        for s in samples[:2]:  # Just show first 2
                            meeting_date = safe_get_value(s, 'MeetingDate', 'Unknown')
                            org_name = safe_get_value(s, 'OrganizationName', 'Unknown')
                            max_count = int(safe_get_value(s, 'MaxCount', 0))
                            head_count = int(safe_get_value(s, 'HeadCount', 0))
                            print('<script>console.log("Sample: {} - {} - MaxCount: {}, HeadCount: {}");</script>'.format(
                                str(meeting_date), str(org_name), max_count, head_count))
            except Exception as e:
                error_msg = str(e).replace('"', '').replace("'", '')
                print('<script>console.log("Error getting worship data for year " + "{}" + ": " + "{}");</script>'.format(year_int, error_msg))
                pass
    
    # Now get connect group and overall ministry attendance data
    sql = """
    WITH MembershipData AS (
        SELECT 
            p.PeopleId,
            p.JoinDate,
            {} AS JoinYear
        FROM People p
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
    ),
    EngagementData AS (
        SELECT 
            md.PeopleId,
            md.JoinYear,
            -- Connect Group attendance 6 months before/after joining
            SUM(CASE WHEN a.MeetingDate >= DATEADD(month, -6, md.JoinDate) 
                     AND a.MeetingDate < md.JoinDate
                     AND os.ProgId = {}
                THEN 1 ELSE 0 END) AS CGAttendanceBefore,
            SUM(CASE WHEN a.MeetingDate >= md.JoinDate 
                     AND a.MeetingDate < DATEADD(month, 6, md.JoinDate)
                     AND os.ProgId = {}
                THEN 1 ELSE 0 END) AS CGAttendanceAfter,
            -- Overall ministry engagement (all programs except worship) 6 months before/after
            SUM(CASE WHEN a.MeetingDate >= DATEADD(month, -6, md.JoinDate) 
                     AND a.MeetingDate < md.JoinDate
                     AND os.ProgId != {}
                THEN 1 ELSE 0 END) AS AllMinistryBefore,
            SUM(CASE WHEN a.MeetingDate >= md.JoinDate 
                     AND a.MeetingDate < DATEADD(month, 6, md.JoinDate)
                     AND os.ProgId != {}
                THEN 1 ELSE 0 END) AS AllMinistryAfter
        FROM MembershipData md
        LEFT JOIN Attend a ON md.PeopleId = a.PeopleId AND a.AttendanceFlag = 1
        LEFT JOIN Meetings m ON a.MeetingId = m.MeetingId
        LEFT JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
        GROUP BY md.PeopleId, md.JoinYear
    )
    SELECT 
        JoinYear,
        COUNT(DISTINCT PeopleId) AS TotalMembers,
        -- Connect Group metrics
        AVG(CAST(CGAttendanceBefore AS FLOAT)) AS AvgSmallGroupBefore,
        AVG(CAST(CGAttendanceAfter AS FLOAT)) AS AvgSmallGroupAfter,
        COUNT(CASE WHEN CGAttendanceAfter > CGAttendanceBefore THEN 1 END) AS ImprovedSmallGroupCount,
        -- Overall ministry metrics (excluding worship)
        AVG(CAST(AllMinistryBefore AS FLOAT)) AS AvgMinistryBefore,
        AVG(CAST(AllMinistryAfter AS FLOAT)) AS AvgMinistryAfter,
        COUNT(CASE WHEN AllMinistryAfter > AllMinistryBefore THEN 1 END) AS ImprovedMinistryCount
    FROM EngagementData
    GROUP BY JoinYear
    ORDER BY JoinYear DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        Config.SMALL_GROUP_PROGRAM_ID,
        Config.SMALL_GROUP_PROGRAM_ID,
        Config.PRIMARY_WORSHIP_PROGRAM_ID,
        Config.PRIMARY_WORSHIP_PROGRAM_ID
    )
    
    results = q.QuerySql(sql)
    
    # Merge worship data with connect group data
    final_results = []
    for row in results:
        # Create a new object with all the data
        join_year = safe_get_value(row, 'JoinYear', None)
        result = type('obj', (object,), {
            'JoinYear': join_year,
            'TotalMembers': safe_get_value(row, 'TotalMembers', 0),
            'AvgWorshipHeadcount': worship_data.get(join_year, 0),
            'AvgSmallGroupBefore': safe_get_value(row, 'AvgSmallGroupBefore', 0),
            'AvgSmallGroupAfter': safe_get_value(row, 'AvgSmallGroupAfter', 0),
            'ImprovedSmallGroupCount': safe_get_value(row, 'ImprovedSmallGroupCount', 0),
            'AvgMinistryBefore': safe_get_value(row, 'AvgMinistryBefore', 0),
            'AvgMinistryAfter': safe_get_value(row, 'AvgMinistryAfter', 0),
            'ImprovedMinistryCount': safe_get_value(row, 'ImprovedMinistryCount', 0)
        })()
        final_results.append(result)
    
    return final_results

def get_family_trends(start_date, end_date):
    """Analyze family unit trends"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    sql = """
    WITH FamilyData AS (
        SELECT 
            FamilyId,
            COUNT(*) AS FamilySize,
            MAX(CASE WHEN Age < 18 THEN 1 ELSE 0 END) AS HasChildren
        FROM People
        WHERE IsDeceased = 0
        GROUP BY FamilyId
    )
    SELECT 
        {} AS JoinYear,
        COUNT(DISTINCT p.FamilyId) AS UniqueFamilies,
        -- Family size categories
        COUNT(DISTINCT CASE WHEN f.FamilySize = 1 THEN p.FamilyId END) AS SinglePersonFamilies,
        COUNT(DISTINCT CASE WHEN f.FamilySize = 2 THEN p.FamilyId END) AS CouplesFamilies,
        COUNT(DISTINCT CASE WHEN f.FamilySize BETWEEN 3 AND 4 THEN p.FamilyId END) AS SmallFamilies,
        COUNT(DISTINCT CASE WHEN f.FamilySize >= 5 THEN p.FamilyId END) AS LargeFamilies,
        -- Family composition
        COUNT(DISTINCT CASE WHEN f.HasChildren = 1 THEN p.FamilyId END) AS FamiliesWithChildren,
        AVG(CAST(f.FamilySize AS FLOAT)) AS AvgFamilySize
    FROM People p
    INNER JOIN FamilyData f ON p.FamilyId = f.FamilyId
    WHERE p.MemberStatusId = {}
      AND p.JoinDate >= '{}'
      AND p.JoinDate <= '{}'
      AND p.IsDeceased = 0
    GROUP BY {}
    ORDER BY {} DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        fiscal_year_sql,
        fiscal_year_sql
    )
    
    return q.QuerySql(sql)

def get_retention_metrics(start_date, end_date):
    """Calculate retention and attrition metrics"""
    fiscal_year_sql = get_fiscal_year_sql().format("JoinDate")
    
    sql = """
    WITH MemberCohorts AS (
        SELECT 
            {} AS CohortYear,
            PeopleId,
            JoinDate
        FROM People
        WHERE MemberStatusId = {}
          AND JoinDate >= '{}'
          AND JoinDate <= DATEADD(year, -1, GETDATE())  -- Exclude current year for retention calc
          AND IsDeceased = 0
    ),
    RetentionData AS (
        SELECT 
            mc.CohortYear,
            COUNT(DISTINCT mc.PeopleId) AS CohortSize,
            -- 1 year retention
            COUNT(DISTINCT CASE 
                WHEN p.MemberStatusId = {} 
                AND DATEADD(year, 1, mc.JoinDate) <= GETDATE()
                THEN mc.PeopleId END) AS RetainedYear1,
            -- 3 year retention
            COUNT(DISTINCT CASE 
                WHEN p.MemberStatusId = {} 
                AND DATEADD(year, 3, mc.JoinDate) <= GETDATE()
                THEN mc.PeopleId END) AS RetainedYear3,
            -- 5 year retention
            COUNT(DISTINCT CASE 
                WHEN p.MemberStatusId = {} 
                AND DATEADD(year, 5, mc.JoinDate) <= GETDATE()
                THEN mc.PeopleId END) AS RetainedYear5
        FROM MemberCohorts mc
        INNER JOIN People p ON mc.PeopleId = p.PeopleId
        GROUP BY mc.CohortYear
    )
    SELECT 
        CohortYear,
        CohortSize,
        RetainedYear1,
        CASE WHEN CohortSize > 0 THEN (RetainedYear1 * 100.0 / CohortSize) ELSE 0 END AS RetentionRate1Year,
        RetainedYear3,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 3, CAST(CAST(CohortYear AS varchar) + '-01-01' AS datetime)) <= GETDATE() 
             THEN (RetainedYear3 * 100.0 / CohortSize) ELSE NULL END AS RetentionRate3Year,
        RetainedYear5,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 5, CAST(CAST(CohortYear AS varchar) + '-01-01' AS datetime)) <= GETDATE()
             THEN (RetainedYear5 * 100.0 / CohortSize) ELSE NULL END AS RetentionRate5Year
    FROM RetentionData
    ORDER BY CohortYear DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        Config.MEMBER_STATUS_ID,
        Config.MEMBER_STATUS_ID,
        Config.MEMBER_STATUS_ID
    )
    
    return q.QuerySql(sql)

def get_campus_breakdown(start_date, end_date):
    """Get membership breakdown by campus"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    # Note: Campus table may not exist in all TouchPoint instances
    try:
        sql = """
        SELECT 
            {} AS JoinYear,
            ISNULL(c.Description, 'No Campus') AS Campus,
            COUNT(*) AS NewMembers,
            COUNT(DISTINCT p.FamilyId) AS NewFamilies
        FROM People p
        LEFT JOIN Campus c ON p.CampusId = c.Id
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        GROUP BY {}, c.Description
        ORDER BY {} DESC, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)
    except:
        # Fallback if Campus table doesn't exist - just group by CampusId
        sql = """
        SELECT 
            {} AS JoinYear,
            CASE WHEN p.CampusId IS NULL THEN 'No Campus' 
                 ELSE 'Campus ' + CAST(p.CampusId AS varchar) END AS Campus,
            COUNT(*) AS NewMembers,
            COUNT(DISTINCT p.FamilyId) AS NewFamilies
        FROM People p
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        GROUP BY {}, p.CampusId
        ORDER BY {} DESC, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)

def get_current_totals():
    """Get current membership totals"""
    sql = """
    SELECT 
        COUNT(*) AS TotalMembers,
        COUNT(DISTINCT FamilyId) AS TotalFamilies,
        COUNT(DISTINCT CampusId) AS TotalCampuses,
        AVG(Age) AS AverageAge,
        SUM(CASE WHEN GenderId = 1 THEN 1 ELSE 0 END) AS Males,
        SUM(CASE WHEN GenderId = 2 THEN 1 ELSE 0 END) AS Females
    FROM People
    WHERE MemberStatusId = {}
      AND IsDeceased = 0
    """.format(Config.MEMBER_STATUS_ID)
    
    if Config.CAMPUS_ID > 0:
        sql += " AND CampusId = {}".format(Config.CAMPUS_ID)
    
    return q.QuerySqlTop1(sql)

def generate_report(analytics, start_date, end_date):
    """Generate the HTML report with all visualizations"""
    
    # Report header
    print("""
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">
                <div class="page-header">
                    <h1>{} <small>{} Year Main Analysis | {} Year Cohort Analysis</small></h1>
                    <p class="text-muted">Report Period: {} to {}</p>
                </div>
            </div>
        </div>
    """.format(
        Config.REPORT_TITLE,
        Config.YEARS_TO_ANALYZE,
        Config.COHORT_YEARS_TO_ANALYZE,
        format_date_display(start_date),
        format_date_display(end_date)
    ))
    
    try:
        # Executive Summary
        print('<script>console.log("Rendering executive summary...");</script>')
        render_executive_summary(analytics)
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering executive summary: {}</div>'.format(str(e)))
    
    try:
        # Yearly Trends Chart
        if 'yearly_trends' in analytics:
            print('<script>console.log("Rendering yearly trends...");</script>')
            render_yearly_trends_chart(analytics['yearly_trends'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering yearly trends: {}</div>'.format(str(e)))
    
    try:
        # Age Demographics Evolution
        if 'age_demographics' in analytics:
            print('<script>console.log("Rendering age demographics...");</script>')
            render_age_demographics_chart(analytics['age_demographics'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering age demographics: {}</div>'.format(str(e)))
    
    try:
        # Connection Pathways
        if Config.SHOW_ORIGIN_ANALYSIS and 'origins' in analytics:
            print('<script>console.log("Rendering origin analysis...");</script>')
            render_origin_analysis(analytics['origins'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering origin analysis: {}</div>'.format(str(e)))
    
    try:
        # Attendance Impact
        if Config.SHOW_ATTENDANCE_IMPACT and 'attendance_impact' in analytics:
            print('<script>console.log("Rendering attendance impact...");</script>')
            render_attendance_impact(analytics['attendance_impact'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering attendance impact: {}</div>'.format(str(e)))
    
    try:
        # Family Analysis
        if Config.SHOW_FAMILY_ANALYSIS and 'family_trends' in analytics:
            print('<script>console.log("Rendering family analysis...");</script>')
            render_family_analysis(analytics['family_trends'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering family analysis: {}</div>'.format(str(e)))
    
    try:
        # Retention Metrics
        if Config.SHOW_RETENTION_METRICS and 'retention' in analytics:
            print('<script>console.log("Rendering retention metrics...");</script>')
            render_retention_metrics(analytics['retention'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering retention metrics: {}</div>'.format(str(e)))
    
    try:
        # Campus Breakdown
        if Config.SHOW_CAMPUS_BREAKDOWN and 'campus_breakdown' in analytics:
            print('<script>console.log("Rendering campus breakdown...");</script>')
            render_campus_breakdown(analytics['campus_breakdown'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering campus breakdown: {}</div>'.format(str(e)))
    
    # Export button
    if Config.ENABLE_EXPORT:
        render_export_button()
    
    print("</div>")  # Close container

def render_executive_summary(analytics):
    """Render executive summary section"""
    current = analytics.get('current_totals')
    yearly = analytics.get('yearly_trends', [])
    
    # Debug logging
    print('<script>console.log("Current totals type:", "{}");</script>'.format(type(current)))
    print('<script>console.log("Yearly trends count:", {});</script>'.format(len(yearly)))
    if current:
        try:
            # Try to see what attributes are available
            attrs = []
            for attr in dir(current):
                if not attr.startswith('_'):
                    attrs.append(attr)
            print('<script>console.log("Current attributes:", {});</script>'.format(str(attrs[:10])))
        except Exception as e:
            print('<script>console.log("Error getting attributes:", "{}");</script>'.format(str(e)))
    
    # Handle case where queries failed
    if not current:
        current = type('obj', (object,), {
            'TotalMembers': 0,
            'TotalFamilies': 0,
            'TotalCampuses': 0,
            'AverageAge': 0,
            'Males': 0,
            'Females': 0
        })()
    
    # Safe access to properties
    total_members = int(safe_get_value(current, 'TotalMembers', 0))
    
    # Calculate key metrics safely
    total_joined_years = 0
    try:
        if yearly:
            # Convert to list if it's not already
            yearly_list = list(yearly) if not isinstance(yearly, list) else yearly
            for y in yearly_list:
                total_joined_years += int(safe_get_value(y, 'NewMembers', 0))
    except Exception as e:
        print('<script>console.log("Error calculating total joined:", "{}");</script>'.format(str(e)))
        total_joined_years = 0
        
    avg_per_year = total_joined_years / len(yearly) if len(yearly) > 0 else 0
    
    # Calculate members joined this year and last year through same date
    current_date = model.DateTime
    current_fy_start_month = Config.FISCAL_YEAR_START_MONTH
    current_fy_start_day = Config.FISCAL_YEAR_START_DAY
    
    # Determine current fiscal year dates
    if Config.USE_FISCAL_YEAR:
        if current_date.Month >= current_fy_start_month:
            current_fy = current_date.Year + 1
            fy_start_date = "{}-{:02d}-{:02d}".format(current_date.Year, current_fy_start_month, current_fy_start_day)
        else:
            current_fy = current_date.Year
            fy_start_date = "{}-{:02d}-{:02d}".format(current_date.Year - 1, current_fy_start_month, current_fy_start_day)
    else:
        current_fy = current_date.Year
        fy_start_date = "{}-01-01".format(current_date.Year)
    
    # Get actual YTD numbers from database for accuracy
    joined_this_year_sql = """
        SELECT COUNT(*) FROM People 
        WHERE MemberStatusId = {} 
          AND JoinDate >= '{}'
          AND JoinDate <= GETDATE()
          AND IsDeceased = 0
    """.format(Config.MEMBER_STATUS_ID, fy_start_date)
    
    # Last year through same date
    if Config.USE_FISCAL_YEAR:
        last_fy_start = "{}-{:02d}-{:02d}".format(
            current_date.Year - 1 if current_date.Month >= current_fy_start_month else current_date.Year - 2,
            current_fy_start_month, 
            current_fy_start_day
        )
    else:
        last_fy_start = "{}-01-01".format(current_date.Year - 1)
    
    last_year_same_date_sql = """
        SELECT COUNT(*) FROM People 
        WHERE MemberStatusId = {} 
          AND JoinDate >= '{}'
          AND JoinDate <= DATEADD(year, -1, GETDATE())
          AND IsDeceased = 0
    """.format(Config.MEMBER_STATUS_ID, last_fy_start)
    
    joined_this_year = 0
    joined_last_year_ytd = 0
    
    try:
        joined_this_year = int(q.QuerySqlScalar(joined_this_year_sql) or 0)
        joined_last_year_ytd = int(q.QuerySqlScalar(last_year_same_date_sql) or 0)
        print('<script>console.log("Joined this FY: {}, Last FY same date: {}");</script>'.format(
            joined_this_year, joined_last_year_ytd))
    except Exception as e:
        print('<script>console.log("Error getting YTD numbers: {}");</script>'.format(str(e).replace('"', '')))
    
    # Calculate growth percentage
    if joined_last_year_ytd > 0:
        growth_pct = ((float(joined_this_year) / float(joined_last_year_ytd)) - 1) * 100
    else:
        growth_pct = 100 if joined_this_year > 0 else 0
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Executive Summary</h2>
            <div class="row">
                <div class="col-md-3">
                    <div class="panel panel-primary">
                        <div class="panel-body text-center">
                            <h2>{}</h2>
                            <p>Current Total Members</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="panel panel-success">
                        <div class="panel-body text-center">
                            <h2>{}</h2>
                            <p>Joined This {} YTD</p>
                            <small style="font-size: 10px; color: #666;">Through today's date</small>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="panel panel-info">
                        <div class="panel-body text-center">
                            <h2>{}</h2>
                            <p>Last {} Same Period</p>
                            <small style="font-size: 10px; color: #666;">Through same date last year</small>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="panel panel-{}">
                        <div class="panel-body text-center">
                            <h2>{}{}%</h2>
                            <p>YTD Growth</p>
                            <small style="font-size: 10px; color: #666;">Year-over-year comparison</small>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    """.format(
        format(total_members, ',d'),
        format(joined_this_year, ',d'),
        "FY" if Config.USE_FISCAL_YEAR else "Year",
        format(joined_last_year_ytd, ',d'),
        "FY" if Config.USE_FISCAL_YEAR else "Year",
        "success" if growth_pct >= 0 else "danger",
        "+" if growth_pct >= 0 else "",
        int(round(growth_pct))
    ))

def render_yearly_trends_chart(yearly_trends):
    """Render yearly membership trends chart"""
    # Reverse the data for charts (keep tables in DESC order but charts in ASC for chronological view)
    try:
        yearly_trends_chart = list(yearly_trends) if yearly_trends else []
        yearly_trends_chart.reverse()
    except:
        yearly_trends_chart = yearly_trends
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Membership Growth Trends</h2>
            <div class="panel panel-default">
                <div class="panel-body">
                    <canvas id="yearlyTrendsChart" style="max-height: 400px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
    var ctx = document.getElementById('yearlyTrendsChart').getContext('2d');
    var yearlyTrendsChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in yearly_trends_chart]) + """],
            datasets: [{
                label: 'New Members',
                data: [""" + ",".join([str(safe_get_value(y, 'NewMembers', 0)) for y in yearly_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                backgroundColor: '""" + Config.CHART_COLORS[0] + """20',
                tension: 0.1,
                fill: true
            }, {
                label: 'New Families',
                data: [""" + ",".join([str(safe_get_value(y, 'NewFamilies', 0)) for y in yearly_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                backgroundColor: '""" + Config.CHART_COLORS[1] + """20',
                tension: 0.1,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            elements: {
                line: {
                    borderWidth: 3
                },
                point: {
                    radius: 5,
                    hoverRadius: 7,
                    backgroundColor: 'white',
                    borderWidth: 2
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'New Members and Families by Year'
                },
                tooltip: {
                    mode: 'index',
                    intersect: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Gender breakdown table
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Gender and Marital Status Breakdown</h3>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>New Members</th>
                            <th>New Families</th>
                            <th>Male</th>
                            <th>Female</th>
                            <th>% Male</th>
                            <th>Married</th>
                            <th>Single</th>
                            <th>Avg Age</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    for year in yearly_trends:
        males = safe_get_value(year, 'Males', 0)
        new_members = safe_get_value(year, 'NewMembers', 0)
        male_pct = (float(males) / float(new_members) * 100) if new_members > 0 else 0
        avg_age = safe_get_value(year, 'AvgAgeAtJoining', None)
        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                        </tr>
        """.format(
            get_year_label() + str(safe_get_value(year, 'JoinYear', '')),
            new_members,
            safe_get_value(year, 'NewFamilies', 0),
            males,
            safe_get_value(year, 'Females', 0),
            int(round(male_pct)),
            safe_get_value(year, 'Married', 0),
            safe_get_value(year, 'Single', 0),
            int(avg_age) if avg_age else "N/A"
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_age_demographics_chart(age_demographics):
    """Render age demographics evolution chart"""
    # Debug logging
    print('<script>console.log("Age demographics count:", {});</script>'.format(len(age_demographics) if age_demographics else 0))
    
    # Reverse for chronological chart display
    try:
        age_demographics_chart = list(age_demographics) if age_demographics else []
        age_demographics_chart.reverse()
    except:
        age_demographics_chart = age_demographics
    
    # Prepare data for stacked bar chart
    age_groups = []
    datasets = []
    
    # Get all age group names from Config
    # This is more reliable than trying to extract from the query results
    for group_name, min_age, max_age in Config.AGE_GROUPS:
        age_groups.append(group_name.replace(" ", ""))
    age_groups.append("UnknownAge")
    
    print('<script>console.log("Age groups:", {});</script>'.format(str(age_groups)))
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Age Demographics Evolution</h2>
            <div class="panel panel-default">
                <div class="panel-body">
                    <canvas id="ageDemographicsChart" style="max-height: 500px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    var ctx2 = document.getElementById('ageDemographicsChart').getContext('2d');
    var ageDemographicsChart = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in age_demographics_chart]) + """],
            datasets: [
    """)
    
    # Create datasets for each age group
    for i, group in enumerate(age_groups):
        if i < len(Config.CHART_COLORS):
            color = Config.CHART_COLORS[i]
        else:
            color = "#999999"
        
        data_values = []
        for year in age_demographics_chart:
            value = safe_get_value(year, group, 0)
            data_values.append(str(value))
        
        # Format the group name for display
        display_name = group.replace("_", " ")
        if group == "UnknownAge":
            display_name = "Unknown Age"
        
        print("""
            {
                label: '""" + display_name + """',
                data: [""" + ",".join(data_values) + """],
                backgroundColor: '""" + color + """',
                borderColor: '""" + color + """',
                borderWidth: 1
            },
        """)
    
    print("""
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Age Distribution of New Members by Year',
                    font: {
                        size: 16
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false
                },
                legend: {
                    position: 'top'
                }
            },
            scales: {
                x: {
                    stacked: true,
                    title: {
                        display: true,
                        text: 'Fiscal Year'
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    },
                    title: {
                        display: true,
                        text: 'Number of New Members'
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Add a percentage breakdown table for clarity
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Age Distribution Percentages</h3>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Total</th>
                            <th>Children (0-12)</th>
                            <th>Teens (13-17)</th>
                            <th>Young Adults (18-29)</th>
                            <th>Adults (30-49)</th>
                            <th>Mature Adults (50-64)</th>
                            <th>Seniors (65+)</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    for year in age_demographics:
        total = 0
        values = {}
        for group_name, min_age, max_age in Config.AGE_GROUPS:
            key = group_name.replace(" ", "")
            values[key] = safe_get_value(year, key, 0)
            total += values[key]
        
        if total > 0:
            print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                        </tr>
            """.format(
                get_year_label() + str(safe_get_value(year, 'JoinYear', '')),
                total,
                int(round(values.get('Children', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('Teens', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('YoungAdults', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('Adults', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('MatureAdults', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('Seniors', 0) * 100.0 / total)) if total > 0 else 0
            ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_origin_analysis(origins):
    """Render analysis of how people found the church"""
    # Aggregate origins by type
    origin_totals = {}
    origin_by_year = {}
    
    for row in origins:
        origin = safe_get_value(row, 'Origin', 'Unknown')
        year = safe_get_value(row, 'JoinYear', '')
        count = safe_get_value(row, 'Count', 0)
        
        if origin not in origin_totals:
            origin_totals[origin] = 0
            origin_by_year[origin] = {}
        
        origin_totals[origin] += count
        origin_by_year[origin][year] = count
    
    # Sort origins by total count
    sorted_origins = sorted(origin_totals.items(), key=lambda x: x[1], reverse=True)
    top_origins = sorted_origins[:10]  # Top 10 origins
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Connection Pathways - How People Found Us</h2>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Top 10 Connection Pathways (10 Year Total)</h3>
                </div>
                <div class="panel-body">
                    <canvas id="originPieChart"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Connection Pathway Trends</h3>
                </div>
                <div class="panel-body">
                    <canvas id="originTrendsChart" height="200"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Pie chart for top origins
    var ctx3 = document.getElementById('originPieChart').getContext('2d');
    var originPieChart = new Chart(ctx3, {
        type: 'doughnut',
        data: {
            labels: [""" + ",".join(["'{}'".format(o[0]) for o in top_origins]) + """],
            datasets: [{
                data: [""" + ",".join([str(o[1]) for o in top_origins]) + """],
                backgroundColor: [""" + ",".join(["'{}'".format(Config.CHART_COLORS[i % len(Config.CHART_COLORS)]) for i in range(len(top_origins))]) + """]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: 'How New Members Found Us'
                },
                legend: {
                    position: 'right'
                }
            }
        }
    });
    
    // Line chart for trends
    var ctx4 = document.getElementById('originTrendsChart').getContext('2d');
    var originTrendsChart = new Chart(ctx4, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}'".format(y) for y in sorted(set([r.JoinYear for r in origins]))]) + """],
            datasets: [
    """)
    
    # Create datasets for top 5 origins
    for i, (origin, total) in enumerate(top_origins[:5]):
        years = sorted(set([r.JoinYear for r in origins]))
        data_points = []
        for year in years:
            value = origin_by_year[origin].get(year, 0)
            data_points.append(str(value))
        
        print("""
            {
                label: '""" + origin + """',
                data: [""" + ",".join(data_points) + """],
                borderColor: '""" + Config.CHART_COLORS[i % len(Config.CHART_COLORS)] + """',
                backgroundColor: '""" + Config.CHART_COLORS[i % len(Config.CHART_COLORS)] + """20',
                tension: 0.1
            },
        """)
    
    print("""
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Top 5 Connection Pathways Over Time'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    </script>
    """)

def render_attendance_impact(attendance_impact):
    """Render attendance impact analysis"""
    # Reverse for chronological display
    try:
        attendance_impact_chart = list(attendance_impact) if attendance_impact else []
        attendance_impact_chart.reverse()
    except:
        attendance_impact_chart = attendance_impact
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Member Engagement Analysis</h2>
            <p class="text-muted">Individual attendance tracking 6 months before and after membership. Shows engagement patterns and integration success.</p>
            <div class="alert alert-info">
                <strong>6-Month Analysis Periods:</strong>
                <ul class="mb-0">
                    <li><strong>Before:</strong> 6 months prior to joining date - shows pre-membership connection level</li>
                    <li><strong>After:</strong> 6 months following joining date - shows post-membership integration success</li>
                </ul>
            </div>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">{} Engagement</h3>
                </div>
                <div class="panel-body">
                    <canvas id="worshipAttendanceChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Overall Ministry Engagement</h3>
                </div>
                <div class="panel-body">
                    <canvas id="smallGroupAttendanceChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    """.format(Config.PRIMARY_WORSHIP_PROGRAM_NAME, Config.SMALL_GROUP_PROGRAM_NAME))
    
    # Create comparison charts
    print("""
    <script>
    // Worship headcount trend
    var ctx5 = document.getElementById('worshipAttendanceChart').getContext('2d');
    var worshipChart = new Chart(ctx5, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in attendance_impact_chart]) + """],
            datasets: [{
                label: 'Average Worship Headcount',
                data: [""" + ",".join([str(int(safe_get_value(y, 'AvgWorshipHeadcount', 0))) for y in attendance_impact_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                backgroundColor: '""" + Config.CHART_COLORS[0] + """20',
                borderWidth: 3,
                tension: 0.1,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            elements: {
                point: {
                    radius: 5,
                    hoverRadius: 7
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'Worship Service Average Headcount by Year'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
    
    // Small group attendance comparison
    var ctx6 = document.getElementById('smallGroupAttendanceChart').getContext('2d');
    var smallGroupChart = new Chart(ctx6, {
        type: 'bar',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in attendance_impact_chart]) + """],
            datasets: [{
                label: 'Before Membership',
                data: [""" + ",".join([str(int(safe_get_value(y, 'AvgSmallGroupBefore', 0))) for y in attendance_impact_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[3] + """'
            }, {
                label: 'After Membership',
                data: [""" + ",".join([str(int(safe_get_value(y, 'AvgSmallGroupAfter', 0))) for y in attendance_impact_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[1] + """'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Average Connect Group Attendance (6 month periods)'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
    </script>
    """)
    
    # Summary statistics
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Member Integration Success Metrics</h3>
            <div class="alert alert-info">
                <p><strong>Understanding the Metrics:</strong></p>
                <ul class="mb-0">
                    <li><strong>CG (Connect Groups):</strong> Specifically tracks attendance at Connect Group/Small Group meetings (Program ID: {})</li>
                    <li><strong>Ministry:</strong> Tracks attendance at ALL church programs and activities EXCEPT worship services - includes Connect Groups, classes, serve teams, events, committees, etc.</li>
                    <li><strong>Before/After:</strong> Average individual attendance counts in the 6-month periods before and after membership</li>
                    <li><strong>Improved:</strong> Number of people who had higher attendance after joining than before</li>
                    <li><strong>% Improved:</strong> Percentage of new members who increased their engagement level</li>
                </ul>
            </div>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>New Members</th>
                            <th>Avg CG Before</th>
                            <th>Avg CG After</th>
                            <th>CG Improved</th>
                            <th>% CG Improved</th>
                            <th>Avg Ministry Before</th>
                            <th>Avg Ministry After</th>
                            <th>Ministry Improved</th>
                            <th>% Ministry Improved</th>
                        </tr>
                    </thead>
                    <tbody>
    """.format(Config.SMALL_GROUP_PROGRAM_ID))
    
    for year in attendance_impact:
        total_members = safe_get_value(year, 'TotalMembers', 0)
        cg_before = safe_get_value(year, 'AvgSmallGroupBefore', 0)
        cg_after = safe_get_value(year, 'AvgSmallGroupAfter', 0)
        cg_improved = safe_get_value(year, 'ImprovedSmallGroupCount', 0)
        
        # Get ministry data from SQL results (now available in query)
        ministry_before = safe_get_value(year, 'AvgMinistryBefore', 0)
        ministry_after = safe_get_value(year, 'AvgMinistryAfter', 0)
        ministry_improved = safe_get_value(year, 'ImprovedMinistryCount', 0)
        
        cg_improved_pct = (float(cg_improved) / float(total_members) * 100) if total_members > 0 else 0
        ministry_improved_pct = (float(ministry_improved) / float(total_members) * 100) if total_members > 0 else 0
        
        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                        </tr>
        """.format(
            get_year_label() + str(safe_get_value(year, 'JoinYear', '')),
            total_members,
            round(cg_before * 10) / 10.0,
            round(cg_after * 10) / 10.0,
            cg_improved,
            int(round(cg_improved_pct)),
            round(ministry_before * 10) / 10.0,
            round(ministry_after * 10) / 10.0,
            ministry_improved,
            int(round(ministry_improved_pct))
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_family_analysis(family_trends):
    """Render family composition analysis"""
    # Reverse for chronological display
    try:
        family_trends_chart = list(family_trends) if family_trends else []
        family_trends_chart.reverse()
    except:
        family_trends_chart = family_trends
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Family Unit Analysis</h2>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Family Size Trends</h3>
                </div>
                <div class="panel-body">
                    <canvas id="familySizeChart" style="width: 100%; height: 300px;"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Families with/without Children</h3>
                </div>
                <div class="panel-body">
                    <canvas id="familyCompositionChart" style="width: 100%; height: 300px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Family size trends
    var ctx7 = document.getElementById('familySizeChart').getContext('2d');
    var familySizeChart = new Chart(ctx7, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}'".format(safe_get_value(y, 'JoinYear', '')) for y in family_trends_chart]) + """],
            datasets: [{
                label: 'Single Person',
                data: [""" + ",".join([str(safe_get_value(y, 'SinglePersonFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                tension: 0.1
            }, {
                label: 'Couples',
                data: [""" + ",".join([str(safe_get_value(y, 'CouplesFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                tension: 0.1
            }, {
                label: 'Small Families (3-4)',
                data: [""" + ",".join([str(safe_get_value(y, 'SmallFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[2] + """',
                tension: 0.1
            }, {
                label: 'Large Families (5+)',
                data: [""" + ",".join([str(safe_get_value(y, 'LargeFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[3] + """',
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Family Size Trends'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    
    // Family composition
    var ctx8 = document.getElementById('familyCompositionChart').getContext('2d');
    var familyCompChart = new Chart(ctx8, {
        type: 'bar',
        data: {
            labels: [""" + ",".join(["'{}'".format(safe_get_value(y, 'JoinYear', '')) for y in family_trends_chart]) + """],
            datasets: [{
                label: 'Families with Children',
                data: [""" + ",".join([str(safe_get_value(y, 'FamiliesWithChildren', 0)) for y in family_trends_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[1] + """',
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                borderWidth: 1
            }, {
                label: 'Families without Children',
                data: [""" + ",".join([str(safe_get_value(y, 'UniqueFamilies', 0) - safe_get_value(y, 'FamiliesWithChildren', 0)) for y in family_trends_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[4] + """',
                borderColor: '""" + Config.CHART_COLORS[4] + """',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Families with/without Children'
                }
            },
            scales: {
                x: {
                    stacked: true
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Average family size trend
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Average Family Size Trend</h3>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Unique Families</th>
                            <th>Average Family Size</th>
                            <th>% with Children</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    for year in family_trends:
        unique_families = safe_get_value(year, 'UniqueFamilies', 0)
        families_with_children = safe_get_value(year, 'FamiliesWithChildren', 0)
        children_pct = (float(families_with_children) / float(unique_families) * 100) if unique_families > 0 else 0
        
        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                        </tr>
        """.format(
            safe_get_value(year, 'JoinYear', ''),
            unique_families,
            int(round(float(safe_get_value(year, 'AvgFamilySize', 0)) * 10)) / 10.0,
            int(round(children_pct))
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_retention_metrics(retention_data):
    """Render retention metrics visualization"""
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Member Retention Analysis</h2>
            <p class="text-muted">Percentage of members from each year's cohort who remain active members</p>
        </div>
    </div>
    <div class="row">
        <div class="col-md-12">
            <div class="panel panel-default">
                <div class="panel-body">
                    <canvas id="retentionChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    var ctx9 = document.getElementById('retentionChart').getContext('2d');
    var retentionChart = new Chart(ctx9, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}'".format(safe_get_value(y, 'CohortYear', '')) for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
            datasets: [{
                label: '1 Year Retention',
                data: [""" + ",".join([str(int(safe_get_value(y, 'RetentionRate1Year', 0))) for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                backgroundColor: '""" + Config.CHART_COLORS[0] + """20',
                tension: 0.1
            }, {
                label: '3 Year Retention',
                data: [""" + ",".join([str(int(safe_get_value(y, 'RetentionRate3Year', 0))) if safe_get_value(y, 'RetentionRate3Year', None) is not None else "null" for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                backgroundColor: '""" + Config.CHART_COLORS[1] + """20',
                tension: 0.1,
                spanGaps: true
            }, {
                label: '5 Year Retention',
                data: [""" + ",".join([str(int(safe_get_value(y, 'RetentionRate5Year', 0))) if safe_get_value(y, 'RetentionRate5Year', None) is not None else "null" for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
                borderColor: '""" + Config.CHART_COLORS[2] + """',
                backgroundColor: '""" + Config.CHART_COLORS[2] + """20',
                tension: 0.1,
                spanGaps: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            elements: {
                line: {
                    borderWidth: 3
                },
                point: {
                    radius: 5,
                    hoverRadius: 7,
                    backgroundColor: 'white',
                    borderWidth: 2
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'Member Retention Rates by Cohort Year'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + (context.parsed.y || 0) + '%';
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Detailed retention table
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Retention Details by Cohort</h3>
            <div class="alert alert-info">
                <p><strong>Understanding Retention Metrics:</strong></p>
                <ul>
                    <li><strong>Cohort Year</strong> - The fiscal year when members joined</li>
                    <li><strong>Initial Size</strong> - Number of new members who joined that year</li>
                    <li><strong>1/3/5 Year columns</strong> - How many from that cohort are still active members after that many years</li>
                    <li><strong>Percentage columns</strong> - What percentage of the original cohort remains active</li>
                    <li>Dashes (-) indicate not enough time has passed to calculate that metric</li>
                </ul>
                <p><em>Example: If 100 people joined in FY2020 and 85 are still members in FY2021, the 1-year retention is 85%</em></p>
            </div>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Cohort Year</th>
                            <th>Initial Size</th>
                            <th>1 Year</th>
                            <th>1 Year %</th>
                            <th>3 Years</th>
                            <th>3 Years %</th>
                            <th>5 Years</th>
                            <th>5 Years %</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    for cohort in retention_data:
        rate1 = safe_get_value(cohort, 'RetentionRate1Year', None)
        rate3 = safe_get_value(cohort, 'RetentionRate3Year', None)
        rate5 = safe_get_value(cohort, 'RetentionRate5Year', None)
        
        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                        </tr>
        """.format(
            safe_get_value(cohort, 'CohortYear', ''),
            safe_get_value(cohort, 'CohortSize', 0),
            safe_get_value(cohort, 'RetainedYear1', 0),
            int(round(rate1)) if rate1 is not None else "-",
            safe_get_value(cohort, 'RetainedYear3', '-'),
            "{}%".format(int(round(rate3))) if rate3 is not None else "-",
            safe_get_value(cohort, 'RetainedYear5', '-'),
            "{}%".format(int(round(rate5))) if rate5 is not None else "-"
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_campus_breakdown(campus_data):
    """Render campus-specific membership analysis"""
    # Aggregate by campus
    campus_totals = {}
    for row in campus_data:
        campus = safe_get_value(row, 'Campus', 'Unknown')
        if campus not in campus_totals:
            campus_totals[campus] = 0
        campus_totals[campus] += safe_get_value(row, 'NewMembers', 0)
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Multi-Campus Analysis</h2>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">10-Year Total by Campus</h3>
                </div>
                <div class="panel-body">
                    <canvas id="campusPieChart"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Campus Growth Trends</h3>
                </div>
                <div class="panel-body">
                    <canvas id="campusTrendsChart" height="200"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Campus pie chart
    var ctx10 = document.getElementById('campusPieChart').getContext('2d');
    var campusPieChart = new Chart(ctx10, {
        type: 'pie',
        data: {
            labels: [""" + ",".join(["'{}'".format(c) for c in campus_totals.keys()]) + """],
            datasets: [{
                data: [""" + ",".join([str(v) for v in campus_totals.values()]) + """],
                backgroundColor: [""" + ",".join(["'{}'".format(Config.CHART_COLORS[i % len(Config.CHART_COLORS)]) for i in range(len(campus_totals))]) + """]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: 'Total New Members by Campus (10 Years)'
                }
            }
        }
    });
    </script>
    """)

def render_export_button():
    """Render export functionality"""
    print("""
    <div class="row">
        <div class="col-md-12">
            <hr>
            <div class="text-center">
                <button class="btn btn-primary" onclick="exportData()">
                    <i class="fa fa-download"></i> Export Full Data to CSV
                </button>
            </div>
        </div>
    </div>
    
    <script>
    function exportData() {
        alert('Export functionality would download detailed CSV data. Implement based on your needs.');
        // In production, this would trigger a download of the full dataset
    }
    </script>
    """)

def format_date_display(date_str):
    """Format date for display"""
    try:
        # Parse the date string and format nicely
        date_obj = model.ParseDate(date_str[:10])
        return date_obj.ToString("MMMM d, yyyy")
    except:
        return date_str[:10]

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

def get_fiscal_year_sql():
    """Get SQL expression to calculate fiscal year from a date"""
    if not Config.USE_FISCAL_YEAR:
        return "YEAR({0})"  # Calendar year
    
    # For fiscal year starting Oct 1: dates from Oct-Dec are in next fiscal year
    # Example: Oct 1, 2023 is in fiscal year 2024
    if Config.FISCAL_YEAR_START_MONTH >= 10:  # Oct, Nov, Dec
        return """
        CASE 
            WHEN MONTH({0}) >= {1} THEN YEAR({0}) + 1
            ELSE YEAR({0})
        END""".format("{0}", Config.FISCAL_YEAR_START_MONTH)
    else:
        # For fiscal years starting Jan-Sep
        return """
        CASE 
            WHEN MONTH({0}) >= {1} THEN YEAR({0})
            ELSE YEAR({0}) - 1
        END""".format("{0}", Config.FISCAL_YEAR_START_MONTH)

def get_year_label():
    """Get label for year display"""
    if Config.USE_FISCAL_YEAR:
        return "FY"
    return ""

# Add print styles
print("""
<style>
@media print {
    .btn { display: none; }
    .page-break { page-break-before: always; }
}
.panel {
    border: 1px solid #ddd;
    border-radius: 4px;
    margin-bottom: 20px;
}
.panel-body {
    padding: 15px;
}
.panel-heading {
    padding: 10px 15px;
    background-color: #f5f5f5;
    border-bottom: 1px solid #ddd;
    border-radius: 3px 3px 0 0;
}
.panel-primary .panel-body {
    background-color: #f0f4ff;
}
.panel-success .panel-body {
    background-color: #f0fff4;
}
.panel-info .panel-body {
    background-color: #f0faff;
}
.panel-danger .panel-body {
    background-color: #fff0f0;
}
h2 {
    margin-top: 30px;
    margin-bottom: 20px;
    border-bottom: 2px solid #eee;
    padding-bottom: 10px;
}
h3 {
    margin-top: 20px;
    margin-bottom: 15px;
}
</style>
""")

# Execute main function
main()

# End of report
