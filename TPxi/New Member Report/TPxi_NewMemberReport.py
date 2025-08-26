#roles=Edit

#####################################################################
# TPxi New Member Report - Weekly Email Edition
#####################################################################
# Purpose: Comprehensive weekly report on new members with statistics,
# demographics, engagement metrics, and actionable insights
#
# Features:
# 1. New member listings with photos and contact info
# 2. Demographic breakdown and trends
# 3. Connection pathways (how they found us)
# 4. Ministry involvement and next steps
# 5. YTD statistics and comparisons
# 6. Follow-up status tracking
# 7. Beautiful HTML email format
#
# Upload Instructions:
# 1. Admin > Advanced > Special Content > Python
# 2. New Python Script File
# 3. Name: TPxi_NewMemberReport
# 4. Schedule as recurring email or run manually
#
# Configuration: Update email recipients and date ranges below

# written by: Ben Swaby
# email: bswaby@fbchtn.org
#####################################################################

# ===== CONFIGURATION =====
class Config:
    # Email settings
    EMAIL_SAVED_QUERY = "NewMemberTeam"  # Name of saved query with recipients
    EMAIL_FROM_ADDRESS = "someemail@somedomain.com"
    EMAIL_FROM_NAME = "New Member Guru"
    EMAIL_FROM_ID = 3134  # PeopleId of sender
    EMAIL_SUBJECT = "New Members - Last 7 Days"
    
    # Report settings
    DAYS_BACK = 7  # Look back this many days for new members
    INCLUDE_PHOTOS = True
    PHOTO_SIZE = 100  # pixels
    
    # Member status IDs (adjust for your church)
    NEW_MEMBER_STATUS_ID = 10  # Typically 10 = Member
    
    # Campus filtering (0 = all campuses)
    CAMPUS_ID = 0
    
    # Ministry categories to track
    TRACK_MINISTRIES = ["Life Groups", "Serve Teams", "Classes"]
    
    # Fiscal year settings (October 1 - September 30)
    FISCAL_YEAR_START_MONTH = 10  # October
    FISCAL_YEAR_START_DAY = 1
    
    # Feature toggles
    SHOW_CONNECTION_PATHWAYS = False  # Show/hide "How They Found Us" section
    SHOW_ENGAGEMENT_PROGRESSION = True  # Show engagement journey tracking
    
    # Attendance tracking
    PRIMARY_PROGRAM_NAME = "Connect Group"  # Name of primary program to track
    PRIMARY_PROGRAM_ID = 1128  # Program ID for attendance metrics
    
    # Baptism hour tracking
    SHOW_BAPTISM_HOUR = True  # Show baptism service/hour if available
    BAPTISM_HOUR_EXTRA_VALUE = "BaptismHour"  # Name of extra value field

# Set page header
model.Header = "New Member Report"

def main():
    """Main execution function"""
    try:
        # Calculate date range
        end_date = model.DateTime
        # Use SQL to calculate start date
        start_date_str = q.QuerySqlScalar("""
            SELECT CONVERT(varchar, DATEADD(day, -{}, GETDATE()), 120)
        """.format(Config.DAYS_BACK))
        
        # For display purposes, we'll use the dates as strings
        
        # Get new members using date strings
        new_members = get_new_members(start_date_str, model.DateTime.ToString("yyyy-MM-dd"))
        
        if not new_members:
            print('<div class="alert alert-info">No new members found in the past {} days.</div>'.format(Config.DAYS_BACK))
            return
        
        # Generate report sections
        html_content = generate_html_report(new_members, start_date_str, model.DateTime.ToString("yyyy-MM-dd"))
        
        # Display report
        print(html_content)
        
        # Optionally send email
        if getattr(model.Data, 'send_email', '') == "true":
            send_report_email(html_content, start_date_str, model.DateTime.ToString("yyyy-MM-dd"))
            print('<div class="alert alert-success">Report emailed to saved query: {}</div>'.format(Config.EMAIL_SAVED_QUERY))
        
    except Exception as e:
        print('<div class="alert alert-danger">Error generating report: {}</div>'.format(str(e)))
        if model.UserIsInRole("Developer"):
            import traceback
            print('<pre>{}</pre>'.format(traceback.format_exc()))

def get_new_members(start_date, end_date):
    """Get all new members in date range with comprehensive data - includes both JoinDate and BaptismDate criteria"""
    sql = """
    SELECT 
        p.PeopleId,
        p.Name2 AS FullName,
        p.FirstName,
        p.LastName,
        p.NickName,
        p.EmailAddress,
        p.CellPhone,
        p.HomePhone,
        p.PrimaryAddress,
        p.PrimaryCity,
        p.PrimaryState,
        p.PrimaryZip,
        p.Age,
        p.GenderId,
        p.MaritalStatusId,
        ms.Description AS MaritalStatus,
        p.FamilyId,
        pic.MediumUrl AS PhotoUrl,
        p.CampusId,
        c.Description AS Campus,
        p.JoinDate,
        p.DecisionTypeId,
        dt.Description AS DecisionType,
        p.NewMemberClassStatusId,
        p.BaptismStatusId,
        p.BaptismDate,
        p.WeddingDate,
        p.CreatedDate,
        -- Calculate days since the qualifying event
        CASE 
            WHEN p.JoinDate >= '{1}' AND p.JoinDate < DATEADD(day, 1, '{2}') AND p.BaptismDate >= '{1}' AND p.BaptismDate < DATEADD(day, 1, '{2}') 
                THEN DATEDIFF(day, CASE WHEN p.JoinDate <= p.BaptismDate THEN p.JoinDate ELSE p.BaptismDate END, GETDATE())
            WHEN p.JoinDate >= '{1}' AND p.JoinDate < DATEADD(day, 1, '{2}') THEN DATEDIFF(day, p.JoinDate, GETDATE())
            WHEN p.BaptismDate >= '{1}' AND p.BaptismDate < DATEADD(day, 1, '{2}') THEN DATEDIFF(day, p.BaptismDate, GETDATE())
            ELSE 0
        END AS DaysSinceJoining,
        fp.Description AS FamilyPosition,
        ISNULL(o.Description, 'Not Specified') AS Origin,
        -- Track which events qualify
        CASE 
            WHEN p.JoinDate >= '{1}' AND p.JoinDate < DATEADD(day, 1, '{2}') THEN 1 ELSE 0 
        END AS JoinDateInRange,
        CASE 
            WHEN p.BaptismDate >= '{1}' AND p.BaptismDate < DATEADD(day, 1, '{2}') THEN 1 ELSE 0 
        END AS BaptismDateInRange
    FROM People p
    LEFT JOIN lookup.MaritalStatus ms ON p.MaritalStatusId = ms.Id
    LEFT JOIN lookup.Campus c ON p.CampusId = c.Id
    LEFT JOIN lookup.DecisionType dt ON p.DecisionTypeId = dt.Id
    LEFT JOIN lookup.FamilyPosition fp ON p.PositionInFamilyId = fp.Id
    LEFT JOIN lookup.Origin o ON p.OriginId = o.Id
    LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
    WHERE p.MemberStatusId = {0}
      AND (
          (p.JoinDate >= '{1}' AND p.JoinDate < DATEADD(day, 1, '{2}'))
          OR 
          (p.BaptismDate >= '{1}' AND p.BaptismDate < DATEADD(day, 1, '{2}'))
      )
      AND p.IsDeceased = 0
    """

    # Add campus filter if specified
    if Config.CAMPUS_ID > 0:
        sql += " AND p.CampusId = {}".format(Config.CAMPUS_ID)
    
    sql += " ORDER BY p.LastName, p.FirstName"
    
    return q.QuerySql(sql.format(
        Config.NEW_MEMBER_STATUS_ID,
        start_date[:10],  # Extract date portion if datetime string
        end_date[:10]     # Extract date portion if datetime string
    ))

def get_member_involvements(people_id):
    """Get ministry involvements for a member"""
    sql = """
    SELECT 
        o.OrganizationName,
        o.Location,
        om.EnrollmentDate,
        CASE WHEN om.Pending = 1 THEN 'Pending' ELSE 'Active' END AS Status
    FROM OrganizationMembers om
    JOIN Organizations o ON om.OrganizationId = o.OrganizationId
    WHERE om.PeopleId = {}
      AND o.OrganizationStatusId = 30  -- Active orgs only
    ORDER BY om.EnrollmentDate DESC
    """.format(people_id)
    
    return q.QuerySql(sql)

def get_family_members(family_id, exclude_person_id):
    """Get other family members"""
    sql = """
    SELECT 
        p.Name2 AS Name,
        p.Age,
        fp.Description AS Relationship,
        ms.Description AS MemberStatus
    FROM People p
    LEFT JOIN lookup.FamilyPosition fp ON p.PositionInFamilyId = fp.Id
    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
    WHERE p.FamilyId = {}
      AND p.PeopleId != {}
      AND p.IsDeceased = 0
    ORDER BY p.PositionInFamilyId, p.Age DESC
    """.format(family_id, exclude_person_id)
    
    return q.QuerySql(sql)

def get_baptism_hour(people_id):
    """Get baptism hour extra value for a person"""
    if not Config.SHOW_BAPTISM_HOUR:
        return None
    
    # Get the extra value
    baptism_hour = model.ExtraValueText(people_id, Config.BAPTISM_HOUR_EXTRA_VALUE)
    return baptism_hour if baptism_hour else None

def get_statistics(start_date, end_date):
    """Get comprehensive statistics about new members - counts each qualifying date separately"""
    stats = {}
    
    # Total events (not people) - each qualifying date counts separately
    sql_total = """
    SELECT 
        SUM(
            CASE WHEN JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}') THEN 1 ELSE 0 END +
            CASE WHEN BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}') THEN 1 ELSE 0 END
        ) AS Total
    FROM People
    WHERE MemberStatusId = {}
      AND (
          (JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}'))
          OR 
          (BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}'))
      )
      AND IsDeceased = 0
    """.format(
        start_date[:10], end_date[:10], start_date[:10], end_date[:10], 
        Config.NEW_MEMBER_STATUS_ID,
        start_date[:10], end_date[:10], start_date[:10], end_date[:10]
    )
    
    stats['total'] = q.QuerySqlScalar(sql_total) or 0
    
    # Total active members count (unchanged)
    sql_total_members = """
    SELECT COUNT(*) AS TotalMembers
    FROM People
    WHERE MemberStatusId = {}
      AND IsDeceased = 0
    """.format(Config.NEW_MEMBER_STATUS_ID)
    
    stats['total_members'] = q.QuerySqlScalar(sql_total_members)
    
    # Demographics breakdown - based on people who had qualifying events
    sql_demographics = """
    SELECT 
        -- Age groups
        SUM(CASE WHEN Age < 18 THEN 1 ELSE 0 END) AS Children,
        SUM(CASE WHEN Age BETWEEN 18 AND 29 THEN 1 ELSE 0 END) AS YoungAdults,
        SUM(CASE WHEN Age BETWEEN 30 AND 49 THEN 1 ELSE 0 END) AS Adults,
        SUM(CASE WHEN Age BETWEEN 50 AND 64 THEN 1 ELSE 0 END) AS OlderAdults,
        SUM(CASE WHEN Age >= 65 THEN 1 ELSE 0 END) AS Seniors,
        SUM(CASE WHEN Age IS NULL OR Age = 0 THEN 1 ELSE 0 END) AS UnknownAge,
        -- Gender
        SUM(CASE WHEN GenderId = 1 THEN 1 ELSE 0 END) AS Males,
        SUM(CASE WHEN GenderId = 2 THEN 1 ELSE 0 END) AS Females,
        -- Marital Status
        SUM(CASE WHEN MaritalStatusId = 20 THEN 1 ELSE 0 END) AS Married,
        SUM(CASE WHEN MaritalStatusId = 10 THEN 1 ELSE 0 END) AS Single,
        -- Family composition
        COUNT(DISTINCT FamilyId) AS UniqueFamilies,
        -- Connection methods
        COUNT(DISTINCT CASE WHEN OriginId IS NOT NULL THEN PeopleId END) AS HasOrigin,
        COUNT(DISTINCT CASE WHEN DecisionTypeId IS NOT NULL THEN PeopleId END) AS HasDecision,
        -- Contact completeness
        SUM(CASE WHEN EmailAddress IS NOT NULL AND EmailAddress != '' THEN 1 ELSE 0 END) AS HasEmail,
        SUM(CASE WHEN CellPhone IS NOT NULL THEN 1 ELSE 0 END) AS HasCellPhone,
        SUM(CASE WHEN PrimaryAddress IS NOT NULL THEN 1 ELSE 0 END) AS HasAddress,
        -- Event type breakdown
        SUM(CASE WHEN JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}') THEN 1 ELSE 0 END) AS JoinEvents,
        SUM(CASE WHEN BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}') THEN 1 ELSE 0 END) AS BaptismEvents,
        SUM(CASE 
            WHEN JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}')
                 AND BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}')
            THEN 1 ELSE 0 END) AS BothEvents
    FROM People
    WHERE MemberStatusId = {}
      AND (
          (JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}'))
          OR 
          (BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}'))
      )
      AND IsDeceased = 0
    """.format(
        start_date[:10], end_date[:10], start_date[:10], end_date[:10],
        start_date[:10], end_date[:10], start_date[:10], end_date[:10],
        Config.NEW_MEMBER_STATUS_ID,
        start_date[:10], end_date[:10], start_date[:10], end_date[:10]
    )
    
    demo_result = q.QuerySqlTop1(sql_demographics)
    if demo_result:
        stats['demographics'] = demo_result
    
    # Get fiscal year dates
    current_date = model.DateTime
    current_year = current_date.Year
    current_month = current_date.Month
    
    # Determine fiscal year
    if current_month >= Config.FISCAL_YEAR_START_MONTH:
        fy_start = "%d-%02d-%02d" % (current_year, Config.FISCAL_YEAR_START_MONTH, Config.FISCAL_YEAR_START_DAY)
        fy_end = "%d-%02d-%02d" % (current_year + 1, Config.FISCAL_YEAR_START_MONTH - 1, 30)
        prev_fy_start = "%d-%02d-%02d" % (current_year - 1, Config.FISCAL_YEAR_START_MONTH, Config.FISCAL_YEAR_START_DAY)
        prev_fy_end = "%d-%02d-%02d" % (current_year, Config.FISCAL_YEAR_START_MONTH - 1, 30)
    else:
        fy_start = "%d-%02d-%02d" % (current_year - 1, Config.FISCAL_YEAR_START_MONTH, Config.FISCAL_YEAR_START_DAY)
        fy_end = "%d-%02d-%02d" % (current_year, Config.FISCAL_YEAR_START_MONTH - 1, 30)
        prev_fy_start = "%d-%02d-%02d" % (current_year - 2, Config.FISCAL_YEAR_START_MONTH, Config.FISCAL_YEAR_START_DAY)
        prev_fy_end = "%d-%02d-%02d" % (current_year - 1, Config.FISCAL_YEAR_START_MONTH - 1, 30)
    
    # Fiscal YTD comparison - count each qualifying date separately
    sql_fytd = """
    SELECT 
        SUM(
            CASE WHEN JoinDate >= '{}' AND JoinDate <= GETDATE() THEN 1 ELSE 0 END +
            CASE WHEN BaptismDate >= '{}' AND BaptismDate <= GETDATE() THEN 1 ELSE 0 END
        ) AS FYTDTotal,
        (SELECT COUNT(DISTINCT DATEPART(week, JoinDate)) FROM People WHERE MemberStatusId = {} AND JoinDate >= '{}' AND JoinDate <= GETDATE() AND IsDeceased = 0) +
        (SELECT COUNT(DISTINCT DATEPART(week, BaptismDate)) FROM People WHERE MemberStatusId = {} AND BaptismDate >= '{}' AND BaptismDate <= GETDATE() AND IsDeceased = 0) AS WeeksWithNewMembers
    FROM People
    WHERE MemberStatusId = {}
      AND (
          (JoinDate >= '{}' AND JoinDate <= GETDATE())
          OR
          (BaptismDate >= '{}' AND BaptismDate <= GETDATE())
      )
      AND IsDeceased = 0
    """.format(
        fy_start, fy_start, 
        Config.NEW_MEMBER_STATUS_ID, fy_start,
        Config.NEW_MEMBER_STATUS_ID, fy_start,
        Config.NEW_MEMBER_STATUS_ID, fy_start, fy_start
    )
    
    fytd_result = q.QuerySqlTop1(sql_fytd)
    if fytd_result:
        stats['fytd'] = fytd_result
    
    # Previous fiscal year stats - count each qualifying date separately
    days_into_fy = q.QuerySqlScalar("""
        SELECT DATEDIFF(day, '{}', GETDATE())
    """.format(fy_start))
    
    sql_prev_fy = """
    SELECT 
        SUM(
            CASE WHEN JoinDate >= '{}' AND JoinDate < '{}' THEN 1 ELSE 0 END +
            CASE WHEN BaptismDate >= '{}' AND BaptismDate < '{}' THEN 1 ELSE 0 END
        ) AS PrevFYTotal,
        SUM(
            CASE WHEN JoinDate >= '{}' AND JoinDate <= DATEADD(day, {}, '{}') THEN 1 ELSE 0 END +
            CASE WHEN BaptismDate >= '{}' AND BaptismDate <= DATEADD(day, {}, '{}') THEN 1 ELSE 0 END
        ) AS PrevFYTDTotal
    FROM People
    WHERE MemberStatusId = {}
      AND (
          (JoinDate >= '{}' AND JoinDate < '{}')
          OR
          (BaptismDate >= '{}' AND BaptismDate < '{}')
      )
      AND IsDeceased = 0
    """.format(
        prev_fy_start, fy_start, prev_fy_start, fy_start,  # PrevFYTotal
        prev_fy_start, days_into_fy, prev_fy_start, prev_fy_start, days_into_fy, prev_fy_start,  # PrevFYTDTotal
        Config.NEW_MEMBER_STATUS_ID,
        prev_fy_start, fy_start, prev_fy_start, fy_start  # WHERE clause
    )
    
    prev_fy_result = q.QuerySqlTop1(sql_prev_fy)
    if prev_fy_result:
        stats['prev_fy'] = prev_fy_result
    
    # Connection pathways (only if enabled)
    if Config.SHOW_CONNECTION_PATHWAYS:
        sql_origins = """
        SELECT 
            ISNULL(o.Description, 'Not Specified') AS Origin,
            COUNT(*) AS Count
        FROM People p
        LEFT JOIN lookup.Origin o ON p.OriginId = o.Id
        WHERE p.MemberStatusId = {}
          AND (
              (p.JoinDate >= '{}' AND p.JoinDate < DATEADD(day, 1, '{}'))
              OR 
              (p.BaptismDate >= '{}' AND p.BaptismDate < DATEADD(day, 1, '{}'))
          )
          AND p.IsDeceased = 0
        GROUP BY o.Description
        ORDER BY COUNT(*) DESC
        """.format(
            Config.NEW_MEMBER_STATUS_ID, 
            start_date[:10], end_date[:10], start_date[:10], end_date[:10]
        )
        
        stats['origins'] = q.QuerySql(sql_origins)
    
    # Fiscal Year Demographics - based on people who had qualifying events this fiscal year
    sql_fy_demographics = """
    SELECT 
        -- Age groups
        SUM(CASE WHEN Age < 18 THEN 1 ELSE 0 END) AS Children,
        SUM(CASE WHEN Age BETWEEN 18 AND 29 THEN 1 ELSE 0 END) AS YoungAdults,
        SUM(CASE WHEN Age BETWEEN 30 AND 49 THEN 1 ELSE 0 END) AS Adults,
        SUM(CASE WHEN Age BETWEEN 50 AND 64 THEN 1 ELSE 0 END) AS OlderAdults,
        SUM(CASE WHEN Age >= 65 THEN 1 ELSE 0 END) AS Seniors,
        SUM(CASE WHEN Age IS NULL OR Age = 0 THEN 1 ELSE 0 END) AS UnknownAge,
        -- Gender
        SUM(CASE WHEN GenderId = 1 THEN 1 ELSE 0 END) AS Males,
        SUM(CASE WHEN GenderId = 2 THEN 1 ELSE 0 END) AS Females,
        -- Marital Status
        SUM(CASE WHEN MaritalStatusId = 20 THEN 1 ELSE 0 END) AS Married,
        SUM(CASE WHEN MaritalStatusId = 10 THEN 1 ELSE 0 END) AS Single,
        -- Contact completeness
        SUM(CASE WHEN EmailAddress IS NOT NULL AND EmailAddress != '' THEN 1 ELSE 0 END) AS HasEmail,
        SUM(CASE WHEN CellPhone IS NOT NULL THEN 1 ELSE 0 END) AS HasCellPhone,
        SUM(CASE WHEN PrimaryAddress IS NOT NULL THEN 1 ELSE 0 END) AS HasAddress,
        -- Total count (unique people, not events)
        COUNT(DISTINCT PeopleId) AS Total
    FROM People
    WHERE MemberStatusId = {}
      AND (
          (JoinDate >= '{}' AND JoinDate <= GETDATE())
          OR 
          (BaptismDate >= '{}' AND BaptismDate <= GETDATE())
      )
      AND IsDeceased = 0
    """.format(Config.NEW_MEMBER_STATUS_ID, fy_start, fy_start)
    
    fy_demo_result = q.QuerySqlTop1(sql_fy_demographics)
    if fy_demo_result:
        stats['fy_demographics'] = fy_demo_result
    
    # Attendance metrics for fiscal year members
    if Config.PRIMARY_PROGRAM_ID:
        sql_attendance = """
        SELECT 
            COUNT(DISTINCT a.PeopleId) AS AttendingProgram,
            AVG(AttendCount) AS AvgAttendance
        FROM (
            SELECT 
                p.PeopleId,
                COUNT(DISTINCT a.MeetingDate) AS AttendCount
            FROM People p
            LEFT JOIN Attend a ON p.PeopleId = a.PeopleId
            LEFT JOIN Organizations o ON a.OrganizationId = o.OrganizationId
            LEFT JOIN Division d ON o.DivisionId = d.Id
            WHERE p.MemberStatusId = {}
              AND (
                  (p.JoinDate >= '{}' AND p.JoinDate <= GETDATE())
                  OR 
                  (p.BaptismDate >= '{}' AND p.BaptismDate <= GETDATE())
              )
              AND p.IsDeceased = 0
              AND a.AttendanceFlag = 1
              AND d.ProgId = {}
              AND a.MeetingDate >= COALESCE(p.JoinDate, p.BaptismDate)
            GROUP BY p.PeopleId
        ) a
        """.format(Config.NEW_MEMBER_STATUS_ID, fy_start, fy_start, Config.PRIMARY_PROGRAM_ID)
        
        attendance_result = q.QuerySqlTop1(sql_attendance)
        if attendance_result:
            stats['attendance'] = attendance_result
            # Get total FY members for percentage calculation
            stats['attendance'].TotalFYMembers = stats['fy_demographics'].Total if stats.get('fy_demographics') else 0
    
    return stats

def get_engagement_progression(people_id):
    """Get engagement timeline for a person"""
    sql = """
    SELECT 
        'First Attendance' AS Event,
        MIN(MeetingDate) AS EventDate
    FROM Attend
    WHERE PeopleId = {}
      AND AttendanceFlag = 1
    
    UNION ALL
    
    SELECT 
        'Decision Type: ' + dt.Description AS Event,
        p.DecisionDate AS EventDate
    FROM People p
    JOIN lookup.DecisionType dt ON p.DecisionTypeId = dt.Id
    WHERE p.PeopleId = {}
      AND p.DecisionDate IS NOT NULL
    
    UNION ALL
    
    SELECT 
        'New Member Class' AS Event,
        p.NewMemberClassDate AS EventDate
    FROM People p
    WHERE p.PeopleId = {}
      AND p.NewMemberClassDate IS NOT NULL
    
    UNION ALL
    
    SELECT 
        'Baptism' AS Event,
        p.BaptismDate AS EventDate
    FROM People p
    WHERE p.PeopleId = {}
      AND p.BaptismDate IS NOT NULL
    
    UNION ALL
    
    SELECT 
        'Joined: ' + o.OrganizationName AS Event,
        om.EnrollmentDate AS EventDate
    FROM OrganizationMembers om
    JOIN Organizations o ON om.OrganizationId = o.OrganizationId
    WHERE om.PeopleId = {}
      AND o.OrganizationStatusId = 30
    
    ORDER BY EventDate
    """.format(people_id, people_id, people_id, people_id, people_id)
    
    return q.QuerySql(sql)

def generate_html_report(new_members, start_date, end_date):
    """Generate beautiful HTML report"""
    stats = get_statistics(start_date, end_date)
    
    # Format dates for display using SQL
    date_info = q.QuerySqlTop1("""
        SELECT 
            DATENAME(month, CAST('{}' AS datetime)) + ' ' + CAST(DAY(CAST('{}' AS datetime)) AS varchar) AS StartDisplay,
            DATENAME(month, CAST('{}' AS datetime)) + ' ' + CAST(DAY(CAST('{}' AS datetime)) AS varchar) + ', ' + CAST(YEAR(CAST('{}' AS datetime)) AS varchar) AS EndDisplay
    """.format(start_date[:10], start_date[:10], end_date[:10], end_date[:10], end_date[:10]))
    
    html = """
    <table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#f7fafc">
        <tr>
            <td align="center" style="padding: 40px 20px;">
                <!-- Main Container -->
                <table border="0" cellpadding="0" cellspacing="0" width="600" bgcolor="#ffffff" style="border-radius: 8px;">
                    <!-- Header -->
                    <tr>
                        <td style="padding: 0;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td bgcolor="#667eea" style="padding: 40px 30px; text-align: center; border-radius: 8px 8px 0 0;">
                                        <h1 style="margin: 0 0 10px 0; color: #ffffff; font-size: 32px; font-weight: 700; line-height: 1.2; font-family: Arial, sans-serif;">New Member Report</h1>
                                        <p style="margin: 0; color: #e2e8f0; font-size: 18px; font-family: Arial, sans-serif;">{} - {}</p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
    """.format(
        date_info.StartDisplay if date_info else "Start Date",
        date_info.EndDisplay if date_info else "End Date"
    )
    
    # Key Metrics Section with enhanced stats
    html += """
                        <!-- Key Metrics -->
                        <tr>
                            <td style="padding: 30px;">
                                <h2 style="margin: 0 0 20px 0; color: #2d3748; font-size: 24px; font-weight: 600;">Key Metrics</h2>
                                
                                <!-- Stats Cards Table -->
                                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                        <td style="padding: 0 10px 20px 0;" width="50%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 25px; text-align: center;">
                                                        <div style="font-size: 42px; font-weight: 700; color: #667eea; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 14px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px;">New Events This Week</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td style="padding: 0 0 20px 10px;" width="50%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 25px; text-align: center;">
                                                        <div style="font-size: 42px; font-weight: 700; color: #48bb78; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 14px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px;">Fiscal YTD Events</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
    """.format(stats.get('total', 0), stats['fytd'].FYTDTotal if stats.get('fytd') else 0)
    
    # Event breakdown section if we have demographics
    if stats.get('demographics'):
        demo = stats['demographics']
        join_events = demo.JoinEvents or 0
        baptism_events = demo.BaptismEvents or 0
        both_events = demo.BothEvents or 0
        unique_people = join_events + baptism_events - both_events
        
        html += """
                                    <!-- Event Type Breakdown -->
                                    <tr>
                                        <td colspan="2" style="padding: 0 0 20px 0;">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f0f9ff; border: 1px solid #0ea5e9; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 20px;">
                                                        <h3 style="margin: 0 0 15px 0; color: #0369a1; font-size: 16px; font-weight: 600;">This Week's Event Breakdown</h3>
                                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                            <tr>
                                                                <td width="25%" style="text-align: center; padding: 10px;">
                                                                    <div style="font-size: 28px; font-weight: 700; color: #059669;">{}</div>
                                                                    <div style="font-size: 12px; color: #6b7280; margin-top: 5px;">Join Events</div>
                                                                </td>
                                                                <td width="25%" style="text-align: center; padding: 10px;">
                                                                    <div style="font-size: 28px; font-weight: 700; color: #2563eb;">{}</div>
                                                                    <div style="font-size: 12px; color: #6b7280; margin-top: 5px;">Baptisms</div>
                                                                </td>
                                                                <td width="25%" style="text-align: center; padding: 10px;">
                                                                    <div style="font-size: 28px; font-weight: 700; color: #dc2626;">{}</div>
                                                                    <div style="font-size: 12px; color: #6b7280; margin-top: 5px;">Both Events</div>
                                                                </td>
                                                                <td width="25%" style="text-align: center; padding: 10px;">
                                                                    <div style="font-size: 28px; font-weight: 700; color: #7c3aed;">{}</div>
                                                                    <div style="font-size: 12px; color: #6b7280; margin-top: 5px;">Unique People</div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
        """.format(join_events, baptism_events, both_events, unique_people)
    
    # Additional metrics row
    if stats.get('fytd') or stats.get('demographics'):
        weeks = max(stats['fytd'].WeeksWithNewMembers or 1, 1) if stats.get('fytd') else 1
        avg_per_week = float(stats['fytd'].FYTDTotal or 0) / weeks if stats.get('fytd') else 0
        
        html += """
                                    <tr>
                                        <td style="padding: 0 10px 20px 0;" width="50%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 25px; text-align: center;">
                                                        <div style="font-size: 42px; font-weight: 700; color: #ed8936; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 14px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px;">Avg Events/Week</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td style="padding: 0 0 20px 10px;" width="50%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 25px; text-align: center;">
                                                        <div style="font-size: 42px; font-weight: 700; color: #9f7aea; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 14px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px;">New Families</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
        """.format(int(round(avg_per_week)), stats['demographics'].UniqueFamilies if stats.get('demographics') else 0)
    
    # Add total member count row
    if stats.get('total_members'):
        html += """
                                    <tr>
                                        <td colspan="2" style="padding: 0 0 20px 0;">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #2d3748; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 30px; text-align: center;">
                                                        <div style="font-size: 48px; font-weight: 700; color: #ffffff; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 16px; color: #e2e8f0; text-transform: uppercase; letter-spacing: 1px;">Total Active Members</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
        """.format(format(int(stats['total_members']), ',d'))
    
    # Fiscal Year Comparison
    if stats.get('fytd') and stats.get('prev_fy'):
        fytd_total = stats['fytd'].FYTDTotal or 0
        prev_fytd_total = stats['prev_fy'].PrevFYTDTotal or 0
        prev_fy_total = stats['prev_fy'].PrevFYTotal or 0
        
        # Calculate growth percentage
        growth_pct = ((float(fytd_total) / float(prev_fytd_total)) - 1) * 100 if prev_fytd_total > 0 else 0
        growth_color = "#48bb78" if growth_pct >= 0 else "#e53e3e"
        growth_symbol = "+" if growth_pct >= 0 else ""
        
        html += """
                                </table>
                                
                                <!-- Fiscal Year Comparison -->
                                <h3 style="margin: 20px 0 15px 0; color: #2d3748; font-size: 20px; font-weight: 600;">Fiscal Year Comparison</h3>
                                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                        <!-- Previous FY Total -->
                                        <td style="padding: 0 10px 20px 0;" width="33%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f0f4f8; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 20px; text-align: center;">
                                                        <div style="font-size: 36px; font-weight: 700; color: #718096; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 12px; color: #718096; text-transform: uppercase;">Previous FY Total</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <!-- Previous FYTD -->
                                        <td style="padding: 0 5px 20px 5px;" width="33%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f0f4f8; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 20px; text-align: center;">
                                                        <div style="font-size: 36px; font-weight: 700; color: #718096; line-height: 1; margin-bottom: 8px;">{}</div>
                                                        <div style="font-size: 12px; color: #718096; text-transform: uppercase;">Previous FYTD</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <!-- Growth Percentage -->
                                        <td style="padding: 0 0 20px 10px;" width="33%">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f0f4f8; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 20px; text-align: center;">
                                                        <div style="font-size: 36px; font-weight: 700; color: {}; line-height: 1; margin-bottom: 8px;">{}{}%</div>
                                                        <div style="font-size: 12px; color: #718096; text-transform: uppercase;">YoY Growth</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
        """.format(
            prev_fy_total,
            prev_fytd_total,
            growth_color,
            growth_symbol,
            int(round(growth_pct))
        )
    else:
        html += "</table>"
    
    html += """
                            </td>
                        </tr>
    """
    
    # Demographics Section - Using Fiscal Year Data
    if stats.get('fy_demographics'):
        demo = stats['fy_demographics']
        total = float(demo.Total) if demo.Total > 0 else 1
        
        html += """
                        <!-- Demographics Section -->
                        <tr>
                            <td style="padding: 0 30px 30px 30px;">
                                <h2 style="margin: 0 0 20px 0; color: #2d3748; font-size: 24px; font-weight: 600;">Fiscal Year Demographics</h2>
                                
                                <!-- Age Distribution -->
                                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px; margin-bottom: 20px;">
                                    <tr>
                                        <td style="padding: 20px;">
                                            <h3 style="margin: 0 0 15px 0; color: #4a5568; font-size: 18px; font-weight: 600;">Age Distribution</h3>
        """
        
        # Age groups with colors
        age_groups = [
            ('Children (0-17)', demo.Children or 0, '#3182ce'),
            ('Young Adults (18-29)', demo.YoungAdults or 0, '#805ad5'),
            ('Adults (30-49)', demo.Adults or 0, '#38a169'),
            ('Older Adults (50-64)', demo.OlderAdults or 0, '#dd6b20'),
            ('Seniors (65+)', demo.Seniors or 0, '#e53e3e'),
            ('Unknown Age', demo.UnknownAge or 0, '#718096')
        ]
        
        for label, count, color in age_groups:
            if count > 0:
                percentage = (count / total) * 100
                html += """
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 10px;">
                                                <tr>
                                                    <td width="40%" style="padding: 5px 0; color: #4a5568; font-size: 14px;">{}</td>
                                                    <td width="60%">
                                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                            <tr>
                                                                <td style="background-color: #e2e8f0; border-radius: 4px; padding: 0; height: 20px;">
                                                                    <div style="background-color: {}; height: 20px; border-radius: 4px; width: {}%;"></div>
                                                                </td>
                                                                <td width="60" style="padding-left: 10px; color: #4a5568; font-size: 14px; text-align: right;">{} ({}%)</td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                """.format(label, color, int(percentage), count, int(percentage))
        
        html += """
                                        </td>
                                    </tr>
                                </table>
                                
                                <!-- Gender & Marital Status -->
                                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                        <td width="48%" style="padding-right: 2%;">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 20px;">
                                                        <h3 style="margin: 0 0 15px 0; color: #4a5568; font-size: 18px; font-weight: 600;">Gender</h3>
        """
        
        # Gender distribution
        male_pct = ((demo.Males or 0) / total) * 100 if total > 0 else 0
        female_pct = ((demo.Females or 0) / total) * 100 if total > 0 else 0
        
        html += """
                                                        <div style="text-align: center;">
                                                            <div style="display: inline-block; margin: 0 20px;">
                                                                <div style="font-size: 36px; font-weight: 700; color: #4299e1; margin-bottom: 5px;">{}%</div>
                                                                <div style="font-size: 14px; color: #718096;">Male ({})</div>
                                                            </div>
                                                            <div style="display: inline-block; margin: 0 20px;">
                                                                <div style="font-size: 36px; font-weight: 700; color: #ed64a6; margin-bottom: 5px;">{}%</div>
                                                                <div style="font-size: 14px; color: #718096;">Female ({})</div>
                                                            </div>
                                                        </div>
        """.format(int(male_pct), demo.Males or 0, int(female_pct), demo.Females or 0)
        
        html += """
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td width="48%" style="padding-left: 2%;">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                                <tr>
                                                    <td style="padding: 20px;">
                                                        <h3 style="margin: 0 0 15px 0; color: #4a5568; font-size: 18px; font-weight: 600;">Marital Status</h3>
        """
        
        # Marital status
        married_pct = ((demo.Married or 0) / total) * 100 if total > 0 else 0
        single_pct = ((demo.Single or 0) / total) * 100 if total > 0 else 0
        
        html += """
                                                        <div style="text-align: center;">
                                                            <div style="display: inline-block; margin: 0 20px;">
                                                                <div style="font-size: 36px; font-weight: 700; color: #48bb78; margin-bottom: 5px;">{}%</div>
                                                                <div style="font-size: 14px; color: #718096;">Married ({})</div>
                                                            </div>
                                                            <div style="display: inline-block; margin: 0 20px;">
                                                                <div style="font-size: 36px; font-weight: 700; color: #667eea; margin-bottom: 5px;">{}%</div>
                                                                <div style="font-size: 14px; color: #718096;">Single ({})</div>
                                                            </div>
                                                        </div>
        """.format(int(married_pct), demo.Married or 0, int(single_pct), demo.Single or 0)
        
        html += """
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                
                                <!-- Contact Completeness Section -->
                                <h3 style="margin: 20px 0 15px 0; color: #2d3748; font-size: 20px; font-weight: 600;">Contact & Engagement Metrics</h3>
                                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                    <tr>
                                        <td style="padding: 20px;">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
        """
        
        # Contact completeness metrics
        if demo:
            email_pct = (float(demo.HasEmail or 0) / total) * 100 if total > 0 else 0
            phone_pct = (float(demo.HasCellPhone or 0) / total) * 100 if total > 0 else 0
            address_pct = (float(demo.HasAddress or 0) / total) * 100 if total > 0 else 0
            
            html += """
                                                <tr>
                                                    <td width="33%" style="text-align: center; padding: 10px;">
                                                        <div style="font-size: 32px; font-weight: 700; color: #667eea;">{}%</div>
                                                        <div style="font-size: 14px; color: #718096; margin-top: 5px;">Have Email</div>
                                                    </td>
                                                    <td width="33%" style="text-align: center; padding: 10px;">
                                                        <div style="font-size: 32px; font-weight: 700; color: #48bb78;">{}%</div>
                                                        <div style="font-size: 14px; color: #718096; margin-top: 5px;">Have Phone</div>
                                                    </td>
                                                    <td width="33%" style="text-align: center; padding: 10px;">
                                                        <div style="font-size: 32px; font-weight: 700; color: #ed8936;">{}%</div>
                                                        <div style="font-size: 14px; color: #718096; margin-top: 5px;">Have Address</div>
                                                    </td>
                                                </tr>
            """.format(int(email_pct), int(phone_pct), int(address_pct))
        
        html += """
                                            </table>
                                        </td>
                                    </tr>
                                </table>
        """
        
        # Attendance metrics if available
        if stats.get('attendance') and Config.PRIMARY_PROGRAM_ID:
            attending_pct = 0
            if stats['attendance'].TotalFYMembers > 0:
                attending_pct = (float(stats['attendance'].AttendingProgram or 0) / float(stats['attendance'].TotalFYMembers)) * 100
            
            html += """
                                <!-- Attendance Metrics -->
                                <h3 style="margin: 20px 0 15px 0; color: #2d3748; font-size: 20px; font-weight: 600;">{} Attendance</h3>
                                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                    <tr>
                                        <td style="padding: 20px;">
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                <tr>
                                                    <td width="50%" style="text-align: center; padding: 10px;">
                                                        <div style="font-size: 32px; font-weight: 700; color: #38a169;">{}%</div>
                                                        <div style="font-size: 14px; color: #718096; margin-top: 5px;">Attending {}</div>
                                                        <div style="font-size: 12px; color: #a0aec0; margin-top: 5px;">({} of {} FY members)</div>
                                                    </td>
                                                    <td width="50%" style="text-align: center; padding: 10px;">
                                                        <div style="font-size: 32px; font-weight: 700; color: #dd6b20;">{}</div>
                                                        <div style="font-size: 14px; color: #718096; margin-top: 5px;">Avg Attendance</div>
                                                        <div style="font-size: 12px; color: #a0aec0; margin-top: 5px;">Since joining</div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
            """.format(
                Config.PRIMARY_PROGRAM_NAME,
                int(attending_pct),
                Config.PRIMARY_PROGRAM_NAME,
                stats['attendance'].AttendingProgram or 0,
                stats['attendance'].TotalFYMembers,
                int(round(stats['attendance'].AvgAttendance or 0))
            )
        
        # Connection pathways if available
        if stats.get('origins') and len(stats['origins']) > 0:
            html += """
                                <!-- Connection Pathways -->
                                <h3 style="margin: 20px 0 15px 0; color: #2d3748; font-size: 20px; font-weight: 600;">How They Found Us</h3>
                                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px;">
                                    <tr>
                                        <td style="padding: 20px;">
            """
            
            for i, origin in enumerate(stats['origins']):
                if i < 5:  # Show top 5 origins
                    origin_pct = (float(origin.Count) / total) * 100 if total > 0 else 0
                    colors = ['#667eea', '#48bb78', '#ed8936', '#e53e3e', '#38b2ac']
                    color = colors[i % len(colors)]
                    
                    html += """
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 10px;">
                                                <tr>
                                                    <td width="30%" style="padding: 5px 0; color: #4a5568; font-size: 14px;">{}</td>
                                                    <td width="70%">
                                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                            <tr>
                                                                <td style="background-color: #e2e8f0; border-radius: 4px; padding: 0; height: 20px;">
                                                                    <div style="background-color: {}; height: 20px; border-radius: 4px; width: {}%;"></div>
                                                                </td>
                                                                <td width="80" style="padding-left: 10px; color: #4a5568; font-size: 14px; text-align: right;">{} ({}%)</td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                    """.format(origin.Origin, color, int(origin_pct), origin.Count, int(origin_pct))
            
            html += """
                                        </td>
                                    </tr>
                                </table>
            """
        
        html += """
                            </td>
                        </tr>
        """
    
    # Member Cards - Group by families
    html += """
                        <!-- New Members Section -->
                        <tr>
                            <td style="padding: 0 30px 30px 30px;">
                                <h2 style="margin: 0 0 20px 0; color: #2d3748; font-size: 24px; font-weight: 600;">New Members This Week</h2>
    """
    
    # Group members by family
    processed_family_ids = set()
    
    for member in new_members:
        # Skip if we've already processed this family
        if member.FamilyId and member.FamilyId in processed_family_ids:
            continue
            
        # Get all family members who are new this week
        family_new_members = []
        if member.FamilyId:
            for m in new_members:
                if m.FamilyId == member.FamilyId:
                    family_new_members.append(m)
            processed_family_ids.add(member.FamilyId)
        else:
            family_new_members = [member]
        
        # If multiple family members, show family card
        if len(family_new_members) > 1:
            html += """
                                <!-- Family Card -->
                                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border: 2px solid #667eea; border-radius: 8px; margin-bottom: 20px;">
                                    <tr>
                                        <td style="padding: 20px;">
                                            <h3 style="margin: 0 0 15px 0; color: #667eea; font-size: 18px; font-weight: 600;">
                                                <span style="background-color: #667eea; color: white; padding: 2px 8px; border-radius: 4px; font-size: 14px; margin-right: 10px;">FAMILY</span>
                                                {} Family
                                            </h3>
            """.format(family_new_members[0].LastName)
            
            # Show each family member
            for fm in family_new_members:
                photo_html = get_member_photo_html(fm)
                html += render_member_info(fm, photo_html, include_family=False)
            
            html += """
                                        </td>
                                    </tr>
                                </table>
            """
        else:
            # Single member card
            member = family_new_members[0]
            photo_html = get_member_photo_html(member)
            
            html += """
                                <!-- Member Card -->
                                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8fafc; border-radius: 8px; margin-bottom: 20px;">
                                    <tr>
                                        <td style="padding: 20px;">
            """
            
            html += render_member_info(member, photo_html, include_family=True)
            
            html += """
                                        </td>
                                    </tr>
                                </table>
            """
    
    html += """
                            </td>
                        </tr>
                        
                        <!-- Footer -->
                        <tr>
                            <td bgcolor="#f8fafc" style="padding: 30px; text-align: center; border-radius: 0 0 8px 8px;">
                                <p style="margin: 0 0 10px 0; color: #718096; font-size: 14px; font-family: Arial, sans-serif;">
                                    Generated on {} by {} Church Management System
                                </p>
                                <p style="margin: 0; color: #a0aec0; font-size: 12px; font-family: Arial, sans-serif;">
                                    This report contains confidential member information. Please handle with care.
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    """.format(
        model.DateTime.ToString("MMMM d, yyyy 'at' h:mm tt"),
        model.Setting("NameOfChurch", "TouchPoint")
    )
    
    return html

def get_member_photo_html(member):
    """Generate photo HTML for a member"""
    if Config.INCLUDE_PHOTOS and hasattr(member, 'PhotoUrl') and member.PhotoUrl:
        # Make the photo clickable to view high resolution version
        # Replace MediumUrl with LargeUrl for high res version
        high_res_url = member.PhotoUrl.replace('MediumUrl', 'LargeUrl') if 'MediumUrl' in member.PhotoUrl else member.PhotoUrl
        return '<a href="{}" target="_blank" style="text-decoration: none;"><img src="{}" alt="{}" style="width: 80px; height: 80px; border-radius: 50%; object-fit: cover; cursor: pointer;" title="Click for high resolution"></a>'.format(high_res_url, member.PhotoUrl, member.FullName)
    else:
        # Use initials
        initials = (member.FirstName[:1] if member.FirstName else "") + (member.LastName[:1] if member.LastName else "")
        # Use table-based centering for email compatibility
        return '''<table width="80" height="80" cellpadding="0" cellspacing="0" style="background-color: #667eea; border-radius: 50%;">
            <tr>
                <td align="center" valign="middle" style="color: white; font-size: 28px; font-weight: 600; font-family: Arial, sans-serif;">
                    {}
                </td>
            </tr>
        </table>'''.format(initials)

def render_member_info(member, photo_html, include_family=True):
    """Render member information HTML"""
    html = """
                                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                <tr>
                                                    <td width="100" valign="top">
                                                        {}
                                                    </td>
                                                    <td style="padding-left: 20px;" valign="top">
                                                        <h3 style="margin: 0 0 10px 0; color: #2d3748; font-size: 20px; font-weight: 600;">
                                                            {}
                                                            <a href="/Person2/{}" target="_blank" style="text-decoration: none; margin-left: 10px;" title="View Profile">
                                                                <span style="font-size: 16px; color: #667eea;">&#128100;</span>
                                                            </a>
                                                        </h3>
    """.format(photo_html, member.FullName, member.PeopleId)
    
    # Basic info with email-safe styling
    if member.Age and member.Age > 0:
        age_info = "{} years old".format(member.Age)
        if member.MaritalStatus:
            age_info += ", " + member.MaritalStatus
        html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• {}</p>'.format(age_info)
    elif member.MaritalStatus:
        # Show marital status even if no age
        html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• {}</p>'.format(member.MaritalStatus)
    
    if member.EmailAddress:
        html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• Email: <a href="mailto:{}" style="color: #667eea; text-decoration: none;">{}</a></p>'.format(member.EmailAddress, member.EmailAddress)
    
    if member.CellPhone:
        formatted_phone = model.FmtPhone(member.CellPhone)
        # Create tel: link with just digits
        phone_digits = ''.join(c for c in member.CellPhone if c.isdigit())
        html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• Phone: <a href="tel:{}" style="color: #667eea; text-decoration: none;">{}</a></p>'.format(phone_digits, formatted_phone)
    
    if member.PrimaryAddress:
        address = "{}, {} {}".format(
            member.PrimaryAddress,
            member.PrimaryCity or "",
            member.PrimaryZip or ""
        )
        html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• Address: {}</p>'.format(address)
    
    # Show qualifying events - Join info and/or Baptism info
    if hasattr(member, 'JoinDateInRange') and hasattr(member, 'BaptismDateInRange'):
        if member.JoinDateInRange and member.JoinDate:
            html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• <strong>Joined:</strong> {}</p>'.format(
                member.JoinDate.ToString("MMM d, yyyy")
            )
        
        if member.BaptismDateInRange and member.BaptismDate:
            # Get baptism hour if configured
            baptism_text = '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• <strong>Baptized:</strong> {}'.format(
                member.BaptismDate.ToString("MMM d, yyyy")
            )
            
            if Config.SHOW_BAPTISM_HOUR:
                baptism_hour = get_baptism_hour(member.PeopleId)
                if baptism_hour:
                    baptism_text += ' <span style="color: #667eea; font-weight: 600;">({} Service)</span>'.format(baptism_hour)
            
            baptism_text += '</p>'
            html += baptism_text
    else:
        # Fallback for older logic - show join date
        if member.JoinDate:
            html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• Joined: <strong>{}</strong></p>'.format(
                member.JoinDate.ToString("MMM d, yyyy")
            )
    
    # Campus
    if member.Campus:
        html += '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• Campus: {}</p>'.format(member.Campus)
    
    # Show baptism information if not already shown above and they are baptized
    if hasattr(member, 'BaptismDateInRange') and not member.BaptismDateInRange:
        if hasattr(member, 'BaptismDate') and member.BaptismDate and member.BaptismStatusId == 30:  # 30 = Baptized
            baptism_info = '<p style="margin: 0 0 5px 0; color: #4a5568; font-size: 14px;">• Previously Baptized: <strong>{}</strong>'.format(
                member.BaptismDate.ToString("MMM d, yyyy")
            )
            
            # Get baptism hour if configured
            if Config.SHOW_BAPTISM_HOUR:
                baptism_hour = get_baptism_hour(member.PeopleId)
                if baptism_hour:
                    baptism_info += ' <span style="color: #667eea; font-weight: 600;">({} Service)</span>'.format(baptism_hour)
            
            baptism_info += '</p>'
            html += baptism_info
    
    # Family members (only if requested and available)
    if include_family and member.FamilyId:
        family_members = get_family_members(member.FamilyId, member.PeopleId)
        if family_members:
            family_list = ", ".join([fm.Name for fm in family_members])
            html += '<p style="margin: 0 0 10px 0; color: #4a5568; font-size: 14px;">• Family: {}</p>'.format(family_list)
    
    # Badges with inline styles - enhanced to show qualifying events
    badges_html = ""
    
    # Show which events qualified them for this report
    if hasattr(member, 'JoinDateInRange') and hasattr(member, 'BaptismDateInRange'):
        if member.JoinDateInRange and member.BaptismDateInRange:
            badges_html += '<span style="display: inline-block; background-color: #d69e2e; color: #ffffff; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 8px;">New Member & Baptized</span>'
        elif member.JoinDateInRange:
            badges_html += '<span style="display: inline-block; background-color: #c6f6d5; color: #22543d; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 8px;">New Member</span>'
        elif member.BaptismDateInRange:
            badges_html += '<span style="display: inline-block; background-color: #bee3f8; color: #2a4e7c; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 8px;">Recently Baptized</span>'
    else:
        # Fallback for older logic
        if member.DaysSinceJoining <= 7:
            badges_html += '<span style="display: inline-block; background-color: #c6f6d5; color: #22543d; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 8px;">New This Week</span>'
    
    if Config.SHOW_CONNECTION_PATHWAYS and member.Origin and member.Origin != "Not Specified":
        badges_html += '<span style="display: inline-block; background-color: #bee3f8; color: #2a4e7c; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 8px;">{}</span>'.format(member.Origin)
    
    if member.DecisionType:
        badges_html += '<span style="display: inline-block; background-color: #fefcbf; color: #744210; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 8px;">{}</span>'.format(member.DecisionType)
    
    if badges_html:
        html += '<div style="margin-top: 10px;">{}</div>'.format(badges_html)
    
    # Ministry involvements
    involvements = get_member_involvements(member.PeopleId)
    if involvements:
        html += '<div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #e2e8f0;">'
        html += '<p style="margin: 0 0 8px 0; color: #2d3748; font-size: 14px; font-weight: 600;">Ministry Involvements:</p>'
        for inv in involvements:
            html += '<span style="display: inline-block; background-color: #e2e8f0; color: #4a5568; padding: 4px 12px; border-radius: 15px; font-size: 12px; margin: 0 5px 5px 0;">{}</span>'.format(inv.OrganizationName)
        html += '</div>'
    
    # Engagement progression (if enabled)
    if Config.SHOW_ENGAGEMENT_PROGRESSION:
        progression = get_engagement_progression(member.PeopleId)
        if progression and len(list(progression)) > 0:
            html += '<div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #e2e8f0;">'
            html += '<p style="margin: 0 0 8px 0; color: #2d3748; font-size: 14px; font-weight: 600;">Engagement Journey:</p>'
            html += '<div style="padding-left: 20px; border-left: 2px solid #667eea;">'
            for i, event in enumerate(progression):
                if event.EventDate:
                    html += '<div style="margin-bottom: 8px; position: relative;">'
                    html += '<div style="position: absolute; left: -26px; top: 2px; width: 10px; height: 10px; background-color: #667eea; border-radius: 50%;"></div>'
                    html += '<span style="color: #718096; font-size: 12px;">{} - </span>'.format(event.EventDate.ToString("MMM d, yyyy"))
                    html += '<span style="color: #4a5568; font-size: 12px;">{}</span>'.format(event.Event)
                    html += '</div>'
            html += '</div></div>'
    
    html += """
                                                    </td>
                                                </tr>
                                            </table>
    """
    
    return html

def send_report_email(html_content, start_date, end_date):
    """Send the report via email"""
    # Format dates for email display
    date_range = q.QuerySqlTop1("""
        SELECT 
            DATENAME(month, CAST('{}' AS datetime)) + ' ' + CAST(DAY(CAST('{}' AS datetime)) AS varchar) AS StartDisplay,
            DATENAME(month, CAST('{}' AS datetime)) + ' ' + CAST(DAY(CAST('{}' AS datetime)) AS varchar) + ', ' + CAST(YEAR(CAST('{}' AS datetime)) AS varchar) AS EndDisplay
    """.format(start_date[:10], start_date[:10], end_date[:10], end_date[:10], end_date[:10]))
    
    # Add email header/footer
    email_html = """
    <div style="max-width: 900px; margin: 0 auto;">
        <p>Hello Team,</p>
        <p>Here's your weekly new member report for {} through {}. We welcomed {} new events this week!</p>
        {}
        <p style="margin-top: 30px;">Best regards,<br>{}</p>
    </div>
    """.format(
        date_range.StartDisplay if date_range else "this period",
        date_range.EndDisplay if date_range else "this period",
        q.QuerySqlScalar("""SELECT 
            SUM(
                CASE WHEN JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}') THEN 1 ELSE 0 END +
                CASE WHEN BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}') THEN 1 ELSE 0 END
            ) AS Total
        FROM People
        WHERE MemberStatusId = {}
          AND (
              (JoinDate >= '{}' AND JoinDate < DATEADD(day, 1, '{}'))
              OR 
              (BaptismDate >= '{}' AND BaptismDate < DATEADD(day, 1, '{}'))
          )
          AND IsDeceased = 0""".format(
            start_date[:10], end_date[:10], start_date[:10], end_date[:10],
            Config.NEW_MEMBER_STATUS_ID,
            start_date[:10], end_date[:10], start_date[:10], end_date[:10]
        )),
        html_content,
        Config.EMAIL_FROM_NAME
    )
    
    # Send email using saved query for recipients
    model.Email(
        Config.EMAIL_SAVED_QUERY,      # Saved query name for recipients
        Config.EMAIL_FROM_ID,           # PeopleId of sender
        Config.EMAIL_FROM_ADDRESS,      # From email address
        Config.EMAIL_FROM_NAME,         # From name
        Config.EMAIL_SUBJECT,           # Subject line
        email_html                      # HTML content
    )

# Add controls for running/emailing report
print("""
<div style="margin-bottom: 20px;">
    <button onclick="window.location.href='?send_email=true'" class="btn btn-primary">
        Email Report to Staff
    </button>
    <button onclick="window.print()" class="btn btn-secondary">
        Print Report
    </button>
</div>
""")

# Execute main function
main()
