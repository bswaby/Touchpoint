# ---------------------------------------------------------------
# TaskNote Activity Dashboard
# ---------------------------------------------------------------
# This script creates a dashboard showing task and note activities
# across the database for administrative purposes.
#
# Features:
# - Overview of recent activity statistics
# - Breakdown by task type, status and ownership
# - Detailed activity lists with filtering options
# - Simple visual representations of metrics
# - Tracking of overdue tasks and who owns them
# - Keyword trend analysis
# ---------------------------------------------------------------

# --Upload Instructions Start--
# To upload code to Touchpoint, use the following steps.
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python and paste all this code
# 4. Test and optionally add to menu
# --Upload Instructions End--

# Set script header and roles
model.Header = "TaskNote Activity Dashboard"
# Uncomment and modify the line below to restrict access based on role
# #roles=Admin,Finance

# ---------------- HELPER FUNCTIONS ----------------

def format_date(date):
    """Format date in a human-readable format"""
    if date is None:
        return "N/A"
    try:
        # TouchPoint dates are usually already in datetime format
        return date.strftime("%m/%d/%Y")
    except:
        # Handle string conversions if needed
        return str(date)

def format_time_ago(date):
    """Format date as time ago (e.g., "2 days ago")"""
    if date is None:
        return "N/A"
    
    from datetime import datetime
    try:
        # Convert string to datetime if needed
        if isinstance(date, str):
            from datetime import datetime
            date = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        
        # Calculate difference
        now = datetime.now()
        diff = now - date
        
        # Format based on time difference
        days = diff.days
        seconds = diff.seconds
        
        if days > 365:
            years = days // 365
            return "{0} year{1} ago".format(years, "s" if years != 1 else "")
        if days > 30:
            months = days // 30
            return "{0} month{1} ago".format(months, "s" if months != 1 else "")
        if days > 0:
            return "{0} day{1} ago".format(days, "s" if days != 1 else "")
        if seconds > 3600:
            hours = seconds // 3600
            return "{0} hour{1} ago".format(hours, "s" if hours != 1 else "")
        if seconds > 60:
            minutes = seconds // 60
            return "{0} minute{1} ago".format(minutes, "s" if minutes != 1 else "")
        
        return "Just now"
    except:
        return str(date)

def get_status_description(status_id):
    """Convert status ID to description"""
    statuses = {
        1: "Complete",
        2: "Pending",
        3: "Accepted",
        4: "Declined",
        5: "Note", 
        6: "Archived"
    }
    return statuses.get(status_id, "Unknown")

def get_person_link(people_id, name=None):
    """Create link to person profile"""
    if people_id is None or people_id == 0:
        return "N/A"
    
    if name is None:
        try:
            sql = "SELECT Name FROM People WHERE PeopleId = {0}".format(people_id)
            name = q.QuerySqlScalar(sql)
        except:
            name = "Person #{0}".format(people_id)
            
    return '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(people_id, name)

def truncate_text(text, max_length=100):
    """Truncate text with ellipsis if too long"""
    if text is None:
        return ""
    if len(text) <= max_length:
        return text
    return text[:max_length] + "..."

# Create a simple HTML bar for visualizations
def html_bar(value, max_value, color="#3498db", width=100, height=20, label=None):
    """Generate HTML for a simple horizontal bar chart"""
    if max_value <= 0:
        max_value = 1  # Prevent division by zero
    
    percent = min(100, max(0, (float(value) / max_value * 100)))
    
    html = """
    <div style="display: flex; align-items: center; margin-bottom: 5px;">
        <div style="width: {5}px; margin-right: 10px;">
            <div style="background-color: #eee; width: 100%; height: {3}px; border-radius: 3px;">
                <div style="background-color: {2}; width: {0}%; height: {3}px; border-radius: 3px;"></div>
            </div>
        </div>
        <div style="min-width: 70px;">
            {1} {4}
        </div>
    </div>
    """.format(
        percent,
        value, 
        color, 
        height, 
        "" if label is None else "({0})".format(label),
        width
    )
    
    return html

# ---------------- SQL QUERIES ----------------

def get_recent_activity_sql(days=30, limit=100, offset=0, type_filter=None, status_filter=None):
    """Get SQL for recent TaskNote activities with pagination and filters"""
    # Create WHERE clauses for filters
    where_clauses = ["tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())".format(days)]
    
    if type_filter:
        if type_filter == "Note":
            where_clauses.append("tn.IsNote = 1")
        elif type_filter == "Task":
            where_clauses.append("tn.IsNote = 0")
    
    if status_filter:
        where_clauses.append("tn.StatusId = {0}".format(status_filter))
    
    # Combine all WHERE clauses
    where_clause = " AND ".join(where_clauses)
    
    return """
        SELECT
            tn.TaskNoteId,
            tn.CreatedDate,
            CASE WHEN tn.IsNote = 1 THEN 'Note' ELSE 'Task' END AS ActivityType,
            tn.StatusId,
            tn.OwnerId,
            owner.Name AS OwnerName,
            tn.AboutPersonId,
            about.Name AS AboutName,
            tn.AssigneeId,
            assignee.Name AS AssigneeName,
            tn.Instructions,
            tn.Notes,
            tn.CompletedDate,
            tn.DueDate,
            (
                SELECT STRING_AGG(k.Description, ', ')
                FROM TaskNoteKeyword tnk
                JOIN Keyword k ON tnk.KeywordId = k.KeywordId
                WHERE tnk.TaskNoteId = tn.TaskNoteId
            ) AS Keywords
        FROM TaskNote tn
        LEFT JOIN People owner ON tn.OwnerId = owner.PeopleId
        LEFT JOIN People about ON tn.AboutPersonId = about.PeopleId
        LEFT JOIN People assignee ON tn.AssigneeId = assignee.PeopleId
        WHERE {3}
        ORDER BY tn.CreatedDate DESC
        OFFSET {2} ROWS
        FETCH NEXT {1} ROWS ONLY
    """.format(days, limit, offset, where_clause)

def get_activity_count_sql(days=30, type_filter=None, status_filter=None):
    """Get total count of activities for pagination with filters"""
    # Create WHERE clauses for filters
    where_clauses = ["tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())".format(days)]
    
    if type_filter:
        if type_filter == "Note":
            where_clauses.append("tn.IsNote = 1")
        elif type_filter == "Task":
            where_clauses.append("tn.IsNote = 0")
    
    if status_filter:
        where_clauses.append("tn.StatusId = {0}".format(status_filter))
    
    # Combine all WHERE clauses
    where_clause = " AND ".join(where_clauses)
    
    return """
        SELECT COUNT(*) AS TotalCount
        FROM TaskNote tn
        WHERE {0}
    """.format(where_clause)

def get_task_summary_sql(days=None):
    """Get SQL for task summary statistics with optional time filter"""
    time_filter = ""
    if days:
        time_filter = "WHERE CreatedDate >= DATEADD(day, -{0}, GETDATE())".format(days)
    
    return """
        SELECT
            COUNT(*) AS TotalTasks,
            SUM(CASE WHEN IsNote = 1 THEN 1 ELSE 0 END) AS TotalNotes,
            SUM(CASE WHEN IsNote = 0 THEN 1 ELSE 0 END) AS TotalActionTasks,
            SUM(CASE WHEN StatusId = 2 THEN 1 ELSE 0 END) AS PendingTasks,
            SUM(CASE WHEN StatusId = 3 THEN 1 ELSE 0 END) AS AcceptedTasks,
            SUM(CASE WHEN StatusId = 1 THEN 1 ELSE 0 END) AS CompletedTasks,
            SUM(CASE WHEN StatusId = 4 THEN 1 ELSE 0 END) AS DeclinedTasks,
            SUM(CASE WHEN StatusId = 6 THEN 1 ELSE 0 END) AS ArchivedTasks,
            SUM(CASE WHEN DueDate IS NOT NULL AND DueDate < GETDATE() AND StatusId NOT IN (1, 6) THEN 1 ELSE 0 END) AS OverdueTasks,
            SUM(CASE WHEN CreatedDate >= DATEADD(day, -7, GETDATE()) THEN 1 ELSE 0 END) AS CreatedLast7Days,
            SUM(CASE WHEN CompletedDate >= DATEADD(day, -7, GETDATE()) THEN 1 ELSE 0 END) AS CompletedLast7Days
        FROM TaskNote
        {0}
    """.format(time_filter)

def get_keyword_breakdown_sql(days=None, limit=10):
    """Get SQL for keyword usage breakdown with optional time filter"""
    time_filter = ""
    if days:
        time_filter = "AND tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())".format(days)
    
    return """
        SELECT TOP {0}
            k.Description AS KeywordName,
            COUNT(*) AS TaskCount
        FROM TaskNoteKeyword tnk
        JOIN Keyword k ON tnk.KeywordId = k.KeywordId
        JOIN TaskNote tn ON tnk.TaskNoteId = tn.TaskNoteId
        WHERE 1=1 {1}
        GROUP BY k.Description
        ORDER BY COUNT(*) DESC
    """.format(limit, time_filter)

def get_keyword_trend_sql(days=90):
    """Get SQL for keyword trends over time"""
    # Calculate the period sizes in Python instead of SQL
    period_size = days // 3
    period1_start = days
    period1_end = days - period_size
    period2_start = period1_end
    period2_end = period2_start - period_size
    # period3 goes from period2_end to 0 (today)
    
    return """
        WITH DateRanges AS (
            SELECT 
                DATEADD(day, -{0}, GETDATE()) AS StartDate,
                DATEADD(day, -{1}, GETDATE()) AS Period1End,
                DATEADD(day, -{2}, GETDATE()) AS Period2End,
                GETDATE() AS EndDate
        ),
        Period1Keywords AS (
            SELECT 
                k.KeywordId,
                k.Description AS KeywordName,
                COUNT(*) AS Period1Count
            FROM TaskNoteKeyword tnk
            JOIN Keyword k ON tnk.KeywordId = k.KeywordId
            JOIN TaskNote tn ON tnk.TaskNoteId = tn.TaskNoteId
            CROSS JOIN DateRanges dr
            WHERE tn.CreatedDate BETWEEN dr.StartDate AND dr.Period1End
            GROUP BY k.KeywordId, k.Description
        ),
        Period2Keywords AS (
            SELECT 
                k.KeywordId,
                k.Description AS KeywordName,
                COUNT(*) AS Period2Count
            FROM TaskNoteKeyword tnk
            JOIN Keyword k ON tnk.KeywordId = k.KeywordId
            JOIN TaskNote tn ON tnk.TaskNoteId = tn.TaskNoteId
            CROSS JOIN DateRanges dr
            WHERE tn.CreatedDate BETWEEN dr.Period1End AND dr.Period2End
            GROUP BY k.KeywordId, k.Description
        ),
        Period3Keywords AS (
            SELECT 
                k.KeywordId,
                k.Description AS KeywordName,
                COUNT(*) AS Period3Count
            FROM TaskNoteKeyword tnk
            JOIN Keyword k ON tnk.KeywordId = k.KeywordId
            JOIN TaskNote tn ON tnk.TaskNoteId = tn.TaskNoteId
            CROSS JOIN DateRanges dr
            WHERE tn.CreatedDate BETWEEN dr.Period2End AND dr.EndDate
            GROUP BY k.KeywordId, k.Description
        )
        SELECT 
            COALESCE(p1.KeywordName, p2.KeywordName, p3.KeywordName) AS KeywordName,
            ISNULL(p1.Period1Count, 0) AS Period1Count,
            ISNULL(p2.Period2Count, 0) AS Period2Count,
            ISNULL(p3.Period3Count, 0) AS Period3Count,
            ISNULL(p3.Period3Count, 0) - ISNULL(p1.Period1Count, 0) AS Trend
        FROM Period1Keywords p1
        FULL OUTER JOIN Period2Keywords p2 ON p1.KeywordId = p2.KeywordId
        FULL OUTER JOIN Period3Keywords p3 ON COALESCE(p1.KeywordId, p2.KeywordId) = p3.KeywordId
        WHERE 
            ISNULL(p1.Period1Count, 0) > 0 OR 
            ISNULL(p2.Period2Count, 0) > 0 OR 
            ISNULL(p3.Period3Count, 0) > 0
        ORDER BY ABS(ISNULL(p3.Period3Count, 0) - ISNULL(p1.Period1Count, 0)) DESC
    """.format(period1_start, period1_end, period2_end)

def get_user_activity_sql(days=None, limit=10):
    """Get SQL for user activity stats with optional time filter"""
    time_filter = ""
    if days:
        time_filter = "AND tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())".format(days)
    
    return """
        SELECT TOP {1}
            p.PeopleId,
            p.Name,
            COUNT(DISTINCT CASE WHEN tn.OwnerId = p.PeopleId THEN tn.TaskNoteId ELSE NULL END) AS TasksCreated,
            COUNT(DISTINCT CASE WHEN tn.AssigneeId = p.PeopleId THEN tn.TaskNoteId ELSE NULL END) AS TasksAssigned,
            COUNT(DISTINCT CASE WHEN tn.CompletedBy = p.PeopleId THEN tn.TaskNoteId ELSE NULL END) AS TasksCompleted
        FROM People p
        JOIN TaskNote tn ON p.PeopleId IN (tn.OwnerId, tn.AssigneeId, tn.CompletedBy)
        WHERE 1=1 {0}
        GROUP BY p.PeopleId, p.Name
        ORDER BY (
            COUNT(DISTINCT CASE WHEN tn.OwnerId = p.PeopleId THEN tn.TaskNoteId ELSE NULL END) +
            COUNT(DISTINCT CASE WHEN tn.AssigneeId = p.PeopleId THEN tn.TaskNoteId ELSE NULL END) +
            COUNT(DISTINCT CASE WHEN tn.CompletedBy = p.PeopleId THEN tn.TaskNoteId ELSE NULL END)
        ) DESC
    """.format(time_filter, limit)

def get_overdue_tasks_by_assignee_sql(days=None):
    """Get SQL for overdue tasks grouped by assignee with optional time filter"""
    time_filter = ""
    if days:
        time_filter = "AND tn.DueDate >= DATEADD(day, -{0}, GETDATE())".format(days)
    
    return """
        SELECT 
            p.PeopleId,
            p.Name,
            COUNT(*) AS OverdueTasks,
            MIN(tn.DueDate) AS OldestDueDate
        FROM TaskNote tn
        JOIN People p ON p.PeopleId = tn.AssigneeId
        WHERE 
            tn.DueDate IS NOT NULL  -- Only include tasks with due dates
            AND tn.DueDate < GETDATE() 
            AND tn.StatusId NOT IN (1, 6) -- Not completed or archived
            AND tn.IsNote = 0 -- Only tasks, not notes
            {0}
        GROUP BY p.PeopleId, p.Name
        ORDER BY COUNT(*) DESC
    """.format(time_filter)

def get_overdue_tasks_for_assignee_sql(assignee_id):
    """Get SQL for all overdue tasks for a specific assignee"""
    return """
        SELECT 
            tn.TaskNoteId,
            tn.DueDate,
            DATEDIFF(day, tn.DueDate, GETDATE()) AS DaysOverdue,
            tn.Instructions,
            tn.Notes,
            owner.Name AS OwnerName,
            about.Name AS AboutName,
            tn.StatusId,
            (
                SELECT STRING_AGG(k.Description, ', ')
                FROM TaskNoteKeyword tnk
                JOIN Keyword k ON tnk.KeywordId = k.KeywordId
                WHERE tnk.TaskNoteId = tn.TaskNoteId
            ) AS Keywords
        FROM TaskNote tn
        LEFT JOIN People owner ON owner.PeopleId = tn.OwnerId
        LEFT JOIN People about ON about.PeopleId = tn.AboutPersonId
        WHERE 
            tn.AssigneeId = {0}
            AND tn.DueDate IS NOT NULL  -- Only include tasks with due dates
            AND tn.DueDate < GETDATE() 
            AND tn.StatusId NOT IN (1, 6) -- Not completed or archived
            AND tn.IsNote = 0 -- Only tasks, not notes
        ORDER BY tn.DueDate ASC
    """.format(assignee_id)

def get_monthly_activity_sql(months=12):
    """Get SQL for activity data grouped by month"""
    return """
        WITH MonthDates AS (
            SELECT 
                DATEFROMPARTS(YEAR(DATEADD(MONTH, -n, GETDATE())), MONTH(DATEADD(MONTH, -n, GETDATE())), 1) AS MonthStart,
                DATEADD(DAY, -1, DATEFROMPARTS(YEAR(DATEADD(MONTH, -n+1, GETDATE())), MONTH(DATEADD(MONTH, -n+1, GETDATE())), 1)) AS MonthEnd
            FROM (
                SELECT TOP ({0}) 
                    ROW_NUMBER() OVER (ORDER BY object_id) - 1 AS n
                FROM sys.objects
            ) AS nums
        )
        SELECT 
            FORMAT(md.MonthStart, 'MMMM yyyy') AS MonthName,
            COUNT(CASE WHEN tn.IsNote = 1 THEN 1 END) AS NoteCount,
            COUNT(CASE WHEN tn.IsNote = 0 THEN 1 END) AS TaskCount,
            COUNT(*) AS TotalCount
        FROM MonthDates md
        LEFT JOIN TaskNote tn ON tn.CreatedDate >= md.MonthStart AND tn.CreatedDate <= md.MonthEnd
        GROUP BY md.MonthStart, FORMAT(md.MonthStart, 'MMMM yyyy')
        ORDER BY md.MonthStart DESC
    """.format(months)
    
def get_completion_kpi_sql(days=30):
    """Get SQL for task completion KPIs with optional time filter"""
    # Ensure days is treated as an integer (in case it comes from form data)
    try:
        days = int(days)
    except (ValueError, TypeError):
        days = 30  # Default if days is not a valid integer
        
    return """
        WITH CompletedTasks AS (
            SELECT
                tn.TaskNoteId,
                tn.CreatedDate,
                tn.CompletedDate,
                tn.DueDate,
                tn.AssigneeId,
                p.Name AS AssigneeName,
                DATEDIFF(day, tn.CreatedDate, tn.CompletedDate) AS DaysToComplete,
                CASE 
                    WHEN tn.DueDate IS NULL THEN 1  -- If no due date, consider it on time
                    WHEN tn.CompletedDate <= tn.DueDate THEN 1 
                    ELSE 0 
                END AS CompletedOnTime
            FROM TaskNote tn
            LEFT JOIN People p ON tn.AssigneeId = p.PeopleId
            WHERE 
                tn.IsNote = 0
                AND tn.StatusId = 1  -- Completed tasks
                AND tn.CompletedDate IS NOT NULL
                AND tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())
        )
        SELECT
            COUNT(*) AS TotalCompleted,
            ISNULL(SUM(CompletedOnTime), 0) AS CompletedOnTime,
            ISNULL(AVG(CAST(DaysToComplete AS FLOAT)), 0) AS AvgCompletionDays,
            ISNULL(MIN(DaysToComplete), 0) AS MinCompletionDays,
            ISNULL(MAX(DaysToComplete), 0) AS MaxCompletionDays,
            CASE 
                WHEN COUNT(*) > 0 THEN CAST(ISNULL(SUM(CompletedOnTime), 0) AS FLOAT) / COUNT(*) * 100
                ELSE 0
            END AS OnTimePercentage
        FROM CompletedTasks
    """.format(days)

def get_completion_by_assignee_sql(days=30, limit=10):
    """Get SQL for completion efficiency by assignee"""
    return """
        WITH AssigneeStats AS (
            SELECT
                tn.AssigneeId,
                p.Name AS AssigneeName,
                COUNT(CASE WHEN tn.StatusId = 1 THEN 1 END) AS CompletedTasks,
                COUNT(CASE WHEN tn.StatusId = 1 AND (tn.DueDate IS NULL OR tn.CompletedDate <= tn.DueDate) THEN 1 END) AS OnTimeTasks,
                COUNT(CASE WHEN tn.StatusId = 1 AND tn.DueDate IS NOT NULL AND tn.CompletedDate > tn.DueDate THEN 1 END) AS LateTasks,
                COUNT(CASE WHEN tn.StatusId = 2 THEN 1 END) AS PendingTasks,
                COUNT(CASE WHEN tn.StatusId = 3 THEN 1 END) AS AcceptedTasks,
                COUNT(CASE WHEN tn.DueDate IS NOT NULL AND tn.DueDate < GETDATE() AND tn.StatusId NOT IN (1, 6) THEN 1 END) AS OverdueTasks,
                AVG(CASE WHEN tn.StatusId = 1 THEN DATEDIFF(day, tn.CreatedDate, tn.CompletedDate) END) AS AvgCompletionDays
            FROM TaskNote tn
            JOIN People p ON tn.AssigneeId = p.PeopleId
            WHERE 
                tn.IsNote = 0
                AND tn.AssigneeId IS NOT NULL
                AND tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())
            GROUP BY tn.AssigneeId, p.Name
        )
        SELECT TOP {1}
            AssigneeId,
            AssigneeName,
            CompletedTasks,
            OnTimeTasks,
            LateTasks,
            PendingTasks,
            AcceptedTasks,
            OverdueTasks,
            AvgCompletionDays,
            CASE 
                WHEN CompletedTasks > 0 THEN CAST(OnTimeTasks AS FLOAT) / CompletedTasks * 100 
                ELSE 0 
            END AS OnTimePercentage,
            CASE
                WHEN (CompletedTasks + PendingTasks + AcceptedTasks) > 0 
                THEN CAST(CompletedTasks AS FLOAT) / (CompletedTasks + PendingTasks + AcceptedTasks) * 100
                ELSE 0
            END AS CompletionRate
        FROM AssigneeStats
        WHERE CompletedTasks > 0 OR PendingTasks > 0 OR AcceptedTasks > 0 OR OverdueTasks > 0
        ORDER BY CompletedTasks DESC, OnTimePercentage DESC
    """.format(days, limit)

def get_completion_trend_sql(weeks=12):
    """Get SQL for weekly completion trends over time"""
    return """
        WITH WeekDates AS (
            SELECT 
                DATEADD(WEEK, -n, DATEADD(DAY, -(DATEPART(WEEKDAY, GETDATE()) - 1), CAST(GETDATE() AS DATE))) AS WeekStart,
                DATEADD(DAY, 6, DATEADD(WEEK, -n, DATEADD(DAY, -(DATEPART(WEEKDAY, GETDATE()) - 1), CAST(GETDATE() AS DATE)))) AS WeekEnd
            FROM (
                SELECT TOP ({0}) 
                    ROW_NUMBER() OVER (ORDER BY object_id) - 1 AS n
                FROM sys.objects
            ) AS nums
        ),
        WeeklyStats AS (
            SELECT 
                wd.WeekStart,
                COUNT(CASE WHEN tn.IsNote = 0 AND tn.CreatedDate BETWEEN wd.WeekStart AND wd.WeekEnd THEN tn.TaskNoteId END) AS TasksCreated,
                COUNT(CASE WHEN tn.StatusId = 1 AND tn.CompletedDate BETWEEN wd.WeekStart AND wd.WeekEnd THEN tn.TaskNoteId END) AS TasksCompleted,
                COUNT(CASE WHEN tn.StatusId = 1 AND tn.CompletedDate BETWEEN wd.WeekStart AND wd.WeekEnd 
                          AND (tn.DueDate IS NULL OR tn.CompletedDate <= tn.DueDate) THEN tn.TaskNoteId END) AS TasksCompletedOnTime
            FROM WeekDates wd
            LEFT JOIN TaskNote tn ON 
                (tn.CreatedDate BETWEEN wd.WeekStart AND wd.WeekEnd) OR 
                (tn.StatusId = 1 AND tn.CompletedDate BETWEEN wd.WeekStart AND wd.WeekEnd)
            GROUP BY wd.WeekStart
        )
        SELECT 
            FORMAT(WeekStart, 'MMM d') + ' - ' + FORMAT(DATEADD(DAY, 6, WeekStart), 'MMM d, yyyy') AS WeekLabel,
            ISNULL(TasksCreated, 0) AS TasksCreated,
            ISNULL(TasksCompleted, 0) AS TasksCompleted,
            ISNULL(TasksCompletedOnTime, 0) AS TasksCompletedOnTime,
            CASE WHEN ISNULL(TasksCompleted, 0) > 0 
                 THEN CAST(ISNULL(TasksCompletedOnTime, 0) AS FLOAT) / ISNULL(TasksCompleted, 0) * 100 
                 ELSE 0 
            END AS OnTimePercentage,
            CASE WHEN ISNULL(TasksCreated, 0) > 0 
                 THEN CAST(ISNULL(TasksCompleted, 0) AS FLOAT) / ISNULL(TasksCreated, 0) * 100 
                 ELSE 0 
            END AS CompletionRate
        FROM WeeklyStats
        ORDER BY WeekStart DESC
    """.format(weeks)

# ---------------- DASHBOARD RENDERING ----------------

def render_dashboard():
    """Render the full dashboard"""
    days_filter = 90
    try:
        if hasattr(model.Data, 'days') and model.Data.days is not None:
            days_filter = int(model.Data.days)
    except:
        days_filter = 90
    
    assignee_id = None
    try:
        if hasattr(model.Data, 'assignee') and model.Data.assignee is not None:
            assignee_id = int(model.Data.assignee)
    except:
        assignee_id = None
        
    # Get current page for activity tab
    page = 1
    try:
        if hasattr(model.Data, 'page') and model.Data.page is not None:
            page = int(model.Data.page)
    except:
        page = 1
        
    # Get current active tab
    active_tab = "overview"
    try:
        if hasattr(model.Data, 'tab') and model.Data.tab is not None and model.Data.tab.strip():
            active_tab = model.Data.tab
    except:
        active_tab = "overview"
    
    # Get type and status filters for activity tab
    type_filter = None
    status_filter = None
    try:
        if hasattr(model.Data, 'type_filter') and model.Data.type_filter is not None and model.Data.type_filter.strip():
            type_filter = model.Data.type_filter
        if hasattr(model.Data, 'status_filter') and model.Data.status_filter is not None and model.Data.status_filter.strip():
            status_filter = model.Data.status_filter
    except:
        pass
    
    # CSS for dashboard styling
    print """
    <style>
        /* General styles */
        body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
        .dashboard-container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .card { background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
                margin-bottom: 20px; padding: 20px; overflow: hidden; }
        .card-title { font-size: 18px; font-weight: bold; margin-top: 0; margin-bottom: 15px; 
                     border-bottom: 1px solid #eee; padding-bottom: 10px; }
        
        /* Flex layouts */
        .flex-row { display: flex; flex-wrap: wrap; margin: 0 -10px; }
        .flex-col { flex: 1; padding: 0 10px; min-width: 250px; }
        
        /* Stat box styles */
        .stat-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 15px; }
        .stat-box { padding: 15px; border-radius: 5px; color: white; display: flex; flex-direction: column; text-align: center; }
        .stat-box .number { font-size: 24px; font-weight: bold; margin-bottom: 5px; text-align: center; }
        .stat-box .label { font-size: 14px; opacity: 0.9; text-align: center; }
        
        /* Color scheme */
        .bg-blue { background-color: #3498db; }
        .bg-green { background-color: #2ecc71; }
        .bg-orange { background-color: #e67e22; }
        .bg-red { background-color: #e74c3c; }
        .bg-purple { background-color: #9b59b6; }
        .bg-gray { background-color: #7f8c8d; }
        
        /* Table styles */
        .data-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        .data-table th { background: #f5f5f5; padding: 10px; text-align: left; font-weight: bold; }
        .data-table td { padding: 10px; border-top: 1px solid #eee; vertical-align: top; }
        .data-table tr:hover { background-color: #f9f9f9; }
        
        /* Filter controls */
        .filter-controls { margin-bottom: 20px; padding: 15px; background: #f5f5f5; border-radius: 5px; }
        .filter-controls select, .filter-controls input { padding: 8px; margin-right: 10px; }
        .filter-controls button { padding: 8px 15px; background: #3498db; color: white; 
                                 border: none; border-radius: 4px; cursor: pointer; }
        .filter-controls button:hover { background: #2980b9; }
        
        /* Loading indicator */
        .loading { text-align: center; padding: 20px; }
        .loading-spinner { border: 5px solid #f3f3f3; border-top: 5px solid #3498db; 
                         border-radius: 50%; width: 40px; height: 40px; margin: 0 auto;
                         animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        
        /* Responsive design */
        @media (max-width: 768px) {
            .flex-col { min-width: 100%; }
            .stat-grid { grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); }
        }
        
        /* Task status colors */
        .status-pending { color: #e67e22; }
        .status-accepted { color: #3498db; }
        .status-complete { color: #2ecc71; }
        .status-declined { color: #e74c3c; }
        .status-archived { color: #7f8c8d; }
        .status-note { color: #9b59b6; }
        
        /* Tooltip styles */
        .tooltip { position: relative; display: inline-block; }
        .tooltip .tooltip-text { visibility: hidden; width: 200px; background-color: #333; 
                              color: #fff; text-align: center; border-radius: 6px; padding: 5px;
                              position: absolute; z-index: 1; bottom: 125%; left: 50%;
                              margin-left: -100px; opacity: 0; transition: opacity 0.3s; }
        .tooltip:hover .tooltip-text { visibility: visible; opacity: 1; }
        
        /* Tab navigation */
        .tab-nav { display: flex; border-bottom: 1px solid #ddd; margin-bottom: 20px; overflow-x: auto; }
        .tab-nav button { padding: 10px 15px; background: none; border: none; cursor: pointer;
                        font-size: 16px; border-bottom: 3px solid transparent; white-space: nowrap; }
        .tab-nav button.active { border-bottom-color: #3498db; color: #3498db; font-weight: bold; }
        .tab-panel { display: none; }
        .tab-panel.active { display: block; }
        
        /* Simple visualization styles */
        .viz-container {
            padding: 10px; 
            background-color: #f9f9f9;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .viz-title {
            font-size: 16px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .viz-legend {
            display: flex;
            margin-bottom: 10px;
        }
        .viz-legend-item {
            display: flex;
            align-items: center;
            margin-right: 15px;
            font-size: 12px;
        }
        .viz-color-box {
            width: 12px;
            height: 12px;
            margin-right: 5px;
            border-radius: 2px;
        }
        
        /* Simple bar chart */
        .simple-bar-chart {
            width: 100%;
            padding: 15px;
            background: #f9f9f9;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        
        /* Simple pie chart */
        .simple-pie-container {
            display: flex;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }
        .simple-pie-segment {
            display: flex;
            align-items: center;
            margin: 5px 10px 5px 0;
        }
        .simple-pie-color {
            width: 15px;
            height: 15px;
            border-radius: 50%;
            margin-right: 5px;
        }
        .simple-pie-label {
            font-size: 14px;
        }
        .simple-pie-value {
            margin-left: 5px;
            font-weight: bold;
        }
        
        /* Trending indicators */
        .trend-up {
            color: #2ecc71;
        }
        .trend-down {
            color: #e74c3c;
        }
        .trend-neutral {
            color: #7f8c8d;
        }
        .trend-arrow {
            font-size: 16px;
            margin-left: 5px;
        }
        
        /* Overdue tasks list */
        .overdue-list {
            margin-top: 15px;
        }
        .overdue-task {
            padding: 15px;
            margin-bottom: 10px;
            border-left: 3px solid #e74c3c;
            background-color: #fff;
            border-radius: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .overdue-task-header {
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
        }
        .overdue-task-title {
            font-weight: bold;
        }
        .overdue-task-days {
            color: #e74c3c;
            font-weight: bold;
        }
        .overdue-task-details {
            margin-top: 5px;
            font-size: 14px;
            color: #666;
        }
        .overdue-task-keywords {
            margin-top: 5px;
            font-size: 12px;
        }
        .overdue-task-keyword {
            display: inline-block;
            background-color: #f5f5f5;
            padding: 2px 6px;
            border-radius: 3px;
            margin-right: 5px;
            margin-bottom: 5px;
        }
        
        /* Back button */
        .back-button {
            display: inline-block;
            margin-bottom: 15px;
            padding: 5px 10px;
            background-color: #f5f5f5;
            border: 1px solid #ddd;
            border-radius: 4px;
            text-decoration: none;
            color: #333;
        }
        .back-button:hover {
            background-color: #e5e5e5;
        }
        
        /* Pagination controls */
        .pagination {
            display: flex;
            justify-content: center;
            margin-top: 20px;
        }
        .pagination a, .pagination span {
            margin: 0 5px;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            text-decoration: none;
            color: #333;
        }
        .pagination a:hover {
            background-color: #eee;
        }
        .pagination .current {
            background-color: #3498db;
            color: white;
            border-color: #3498db;
        }
        .pagination .disabled {
            color: #ccc;
            cursor: not-allowed;
        }
        
        /* Loading overlay */
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.8);
            z-index: 1000;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-direction: column;
        }
        .loading-message {
            margin-top: 20px;
            font-size: 18px;
            color: #333;
        }
        #assignee-efficiency {
            max-height: 1800px;
            overflow-y: auto;
            scrollbar-width: thin;
            scrollbar-color: #ccc #f5f5f5;
            padding-right: 10px;
        }
        
        #assignee-efficiency::-webkit-scrollbar {
            width: 8px;
        }
        
        #assignee-efficiency::-webkit-scrollbar-track {
            background: #f5f5f5;
            border-radius: 4px;
        }
        
        #assignee-efficiency::-webkit-scrollbar-thumb {
            background-color: #ccc;
            border-radius: 4px;
        }
    </style>
    """
    
    # JavaScript for tab navigation and loading indicators
    print """
    <script>
        // Show loading overlay
        function showLoading() {{
            var overlay = document.createElement('div');
            overlay.className = 'loading-overlay';
            
            var spinner = document.createElement('div');
            spinner.className = 'loading-spinner';
            
            var message = document.createElement('div');
            message.className = 'loading-message';
            message.textContent = 'Loading data...';
            
            overlay.appendChild(spinner);
            overlay.appendChild(message);
            document.body.appendChild(overlay);
        }}
        
        // Hide loading overlay
        function hideLoading() {{
            var overlay = document.querySelector('.loading-overlay');
            if (overlay) {{
                document.body.removeChild(overlay);
            }}
        }}
        
        function switchTab(tabName) {{
            // Instead of just switching the display, navigate to the URL with the tab parameter
            // This will cause a full page reload with the correct tab data
            var currentUrl = new URL(window.location.href);
            
            // Preserve the days filter parameter if it exists
            var daysFilter = currentUrl.searchParams.get('days');
            
            // Create new URL with the tab parameter
            var newUrl = window.location.pathname + '?tab=' + tabName;
            if (daysFilter) {{
                newUrl += '&days=' + daysFilter;
            }}
            
            // Show loading indicator
            showLoading();
            
            // Navigate to the new URL
            window.location.href = newUrl;
            
            return false;
        }}
        
        // Initialize on page load
        window.onload = function() {{
            var initialTab = '{0}';
            if (!initialTab) {{
                initialTab = 'overview';
            }}
            
            // Just update the visual state without causing navigation
            var tabPanels = document.getElementsByClassName('tab-panel');
            for (var i = 0; i < tabPanels.length; i++) {{
                tabPanels[i].style.display = 'none';
            }}
            
            var targetPanel = document.getElementById(initialTab + '-panel');
            if (targetPanel) {{
                targetPanel.style.display = 'block';
            }}
            
            var tabButtons = document.getElementsByClassName('tab-button');
            for (var i = 0; i < tabButtons.length; i++) {{
                tabButtons[i].className = tabButtons[i].className.replace(' active', '');
            }}
            
            var targetButton = document.getElementById(initialTab + '-button');
            if (targetButton) {{
                targetButton.className += ' active';
            }}
            
            hideLoading();
        }};
        
        // Show loading overlay during form submission
        document.addEventListener('DOMContentLoaded', function() {{
            var forms = document.querySelectorAll('form');
            for (var i = 0; i < forms.length; i++) {{
                forms[i].addEventListener('submit', function() {{
                    showLoading();
                }});
            }}
        }});
    </script>
    """.format(active_tab)
    
    # Render overdue task details if assignee is specified
    if assignee_id is not None:
        render_overdue_tasks_detail(assignee_id)
        return
    
    # Dashboard header with filter controls
    print """
    <div class="dashboard-container">
        <h1>TaskNote Activity Dashboard
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
          </svg>        
        </h1>
        
        <div class="filter-controls">
            <form method="get" action="">
                <label for="days-filter">Time period:</label>
                <select id="days-filter" name="days" onchange="this.form.submit()">
                    <option value="30" {0}>Last 30 days</option>
                    <option value="90" {1}>Last 90 days</option>
                    <option value="180" {2}>Last 6 months</option>
                    <option value="365" {3}>Last year</option>
                </select>
                <input type="hidden" name="tab" value="{4}">
                <button type="submit">Apply Filters</button>
            </form>
        </div>
        
        <!-- Tab Navigation -->
        <div class="tab-nav">
            <button id="overview-button" class="tab-button" onclick="return switchTab('overview')">Overview</button>
            <button id="activity-button" class="tab-button" onclick="return switchTab('activity')">Recent Activity</button>
            <button id="trends-button" class="tab-button" onclick="return switchTab('trends')">Keyword Trends</button>
            <button id="overdue-button" class="tab-button" onclick="return switchTab('overdue')">Overdue Tasks</button>
            <button id="completion-button" class="tab-button" onclick="return switchTab('completion')">Completion KPIs</button>
            <button id="stats-button" class="tab-button" onclick="return switchTab('stats')">Analytics</button>
        </div>
    """.format(
        "selected" if days_filter == 30 else "",
        "selected" if days_filter == 90 else "",
        "selected" if days_filter == 180 else "",
        "selected" if days_filter == 365 else "",
        active_tab
    )
    
    # Add a loading spinner
    print """
    <div id="loading" class="loading">
        <div class="loading-spinner"></div>
        <p>Loading data...</p>
    </div>
    """
    
    # Render all tab containers but load data only for the active tab
    # Overview tab
    print """<div id="overview-panel" class="tab-panel">"""
    if active_tab == "overview":
        render_overview_tab(days_filter)
    else:
        print """<div class="loading"><div class="loading-spinner"></div><p>Loading overview data...</p></div>"""
    print """</div>"""
    
    # Activity tab
    print """<div id="activity-panel" class="tab-panel">"""
    if active_tab == "activity":
        render_activity_tab(days_filter, page)
    else:
        print """<div class="loading"><div class="loading-spinner"></div><p>Loading activity data...</p></div>"""
    print """</div>"""
    
    # Trends tab
    print """<div id="trends-panel" class="tab-panel">"""
    if active_tab == "trends":
        render_trends_tab(days_filter)
    else:
        print """<div class="loading"><div class="loading-spinner"></div><p>Loading trend data...</p></div>"""
    print """</div>"""
    
    # Overdue tasks tab
    print """<div id="overdue-panel" class="tab-panel">"""
    if active_tab == "overdue":
        # Pass the days_filter parameter
        if assignee_id is not None:
            render_overdue_tasks_detail(assignee_id, days_filter)
        else:
            render_overdue_tab(days_filter)
    else:
        print """<div class="loading"><div class="loading-spinner"></div><p>Loading overdue data...</p></div>"""
    print """</div>"""
    
    # Completion KPIs tab
    print """<div id="completion-panel" class="tab-panel">"""
    if active_tab == "completion":
        render_completion_kpi_tab(days_filter)
    else:
        print """<div class="loading"><div class="loading-spinner"></div><p>Loading completion KPI data...</p></div>"""
    print """</div>"""
    
    # Stats tab
    print """<div id="stats-panel" class="tab-panel">"""
    if active_tab == "stats":
        render_stats_tab(days_filter)
    else:
        print """<div class="loading"><div class="loading-spinner"></div><p>Loading stats data...</p></div>"""
    print """</div>"""
    
    # Close dashboard container
    print """
    </div>
    
    <!-- Script to hide the loading indicator -->
    <script>
        // Hide the main loading indicator once everything is loaded
        document.getElementById('loading').style.display = 'none';
    </script>
    """

def render_overview_tab(days_filter):
    """Render the overview tab with summary stats"""
    print """
        <div class="flex-row">
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Activity Summary (Last {0} Days)</h2>
                    <div class="stat-grid">
    """.format(days_filter)
    
    # Fetch summary data for the selected period
    try:
        summary_data = q.QuerySqlTop1(get_task_summary_sql(days_filter))
        
        # Display summary stat boxes
        stat_boxes = [
            ("Total Notes", summary_data.TotalNotes, "bg-purple"),
            ("Total Tasks", summary_data.TotalActionTasks, "bg-blue"),
            ("Pending Tasks", summary_data.PendingTasks, "bg-orange"),
            ("Completed Tasks", summary_data.CompletedTasks, "bg-green"),
            ("Overdue Tasks", summary_data.OverdueTasks, "bg-red"),
            ("Created (7 days)", summary_data.CreatedLast7Days, "bg-gray")
        ]
        
        for label, value, color_class in stat_boxes:
            print """
                <div class="stat-box {2}">
                    <div class="number">{1}</div>
                    <div class="label">{0}</div>
                </div>
            """.format(label, value, color_class)
            
        print """
                    </div>
        """
        
        # Add a divider after the stat boxes
        print """
            <div style="border-top: 1px solid #eee; margin: 15px 0; padding-top: 15px;">
                <h3 style="margin-top: 0;">Task Completion Summary</h3>
        """
        # Fetch KPI data for completion summary
        try:
            kpi_data = q.QuerySqlTop1(get_completion_kpi_sql(days_filter))
            
            # Simple metrics display
            if kpi_data.TotalCompleted > 0:
                on_time_pct = kpi_data.OnTimePercentage or 0
                on_time_color = "green" if on_time_pct >= 75 else "orange" if on_time_pct >= 50 else "red"
                
                print """
                    <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                        <div>
                            <strong>Task Completion Rate:</strong>
                        </div>
                        <div>
                            <span style="color: {0};">{1:.1f}%</span>
                        </div>
                    </div>
                    
                    <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                        <div>
                            <strong>On-Time Completion:</strong>
                        </div>
                        <div>
                            <span style="color: {2};">{3:.1f}%</span>
                        </div>
                    </div>
                    
                    <div style="display: flex; justify-content: space-between;">
                        <div>
                            <strong>Average Completion Time:</strong>
                        </div>
                        <div>
                            {4:.1f} days
                        </div>
                    </div>
                    
                    <div style="margin-top: 10px;">
                        <a href="?tab=completion&days={5}" style="color: #3498db;">View detailed completion analytics →</a>
                    </div>
                """.format(
                    "green" if (kpi_data.TotalCompleted / float(summary_data.TotalActionTasks) * 100 if summary_data.TotalActionTasks > 0 else 0) >= 75 else "orange",
                    kpi_data.TotalCompleted / float(summary_data.TotalActionTasks) * 100 if summary_data.TotalActionTasks > 0 else 0,
                    on_time_color,
                    on_time_pct,
                    kpi_data.AvgCompletionDays or 0,
                    days_filter
                )
            else:
                print "<p>No completed tasks in this time period.</p>"
        except Exception as e:
            print "<p>Error loading completion summary: {0}</p>".format(str(e))
        # Close the divider section
        print """</div>"""
        
        print """
                </div>
            </div>
        </div>
        
        <div class="flex-row">
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Top Active Users (Last {0} Days)</h2>
                    <div id="user-stats">
        """.format(days_filter)
        
        # User activity stats filtered by time period
        try:
            user_data = q.QuerySql(get_user_activity_sql(days_filter))
            
            # Find max value for scaling bars
            max_user_count = 0
            for user in user_data:
                max_user_count = max(max_user_count, 
                                    user.TasksCreated, 
                                    user.TasksAssigned, 
                                    user.TasksCompleted)
            
            # Create simple visual representation for user activity
            for user in user_data:
                print '<div style="margin-bottom: 15px;">'
                print '<div style="font-weight: bold; margin-bottom: 5px;">{0}</div>'.format(
                    get_person_link(user.PeopleId, user.Name)
                )
                # Created tasks
                print html_bar(user.TasksCreated, max_user_count, color="#3498db", width=300, label="Created")
                
                # Assigned tasks
                print html_bar(user.TasksAssigned, max_user_count, color="#e67e22", width=300, label="Assigned")
                
                # Completed tasks
                print html_bar(user.TasksCompleted, max_user_count, color="#2ecc71", width=300, label="Completed")
                
                print '</div>'
            
            # Also show data table
            print """<table class="data-table">
                <tr>
                    <th>User</th>
                    <th>Tasks Created</th>
                    <th>Tasks Assigned</th>
                    <th>Tasks Completed</th>
                </tr>
            """
            
            for user in user_data:
                print """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                </tr>
                """.format(
                    get_person_link(user.PeopleId, user.Name),
                    user.TasksCreated,
                    user.TasksAssigned,
                    user.TasksCompleted
                )
                
            print "</table>"
            
        except Exception as e:
            print "<p>Error loading user data: {0}</p>".format(str(e))
        
        print """
                    </div>
                </div>
            </div>
            
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Monthly Activity Trends</h2>
                    <div class="simple-bar-chart">
        """
        
        # Use monthly data instead of weekly for activity trends
        monthly_data = q.QuerySql(get_monthly_activity_sql(12))  # Last 12 months
        
        # Find max values for scaling
        max_count = 0
        for row in monthly_data:
            max_count = max(max_count, row.TotalCount)
        
        # Create simple bar visualization
        print """
            <div class="viz-legend">
                <div class="viz-legend-item">
                    <div class="viz-color-box" style="background-color: #3498db;"></div>
                    <span>Tasks</span>
                </div>
                <div class="viz-legend-item">
                    <div class="viz-color-box" style="background-color: #9b59b6;"></div>
                    <span>Notes</span>
                </div>
            </div>
        """
        
        # Generate HTML bars for each month
        for row in monthly_data:
            # Print month name
            print '<div style="margin-bottom: 15px;">'
            print '<div style="font-weight: bold; margin-bottom: 5px;">{0}</div>'.format(row.MonthName)
            
            # Generate task bar
            print html_bar(row.TaskCount, max_count, color="#3498db", width=300, label="Tasks")
            
            # Generate note bar
            print html_bar(row.NoteCount, max_count, color="#9b59b6", width=300, label="Notes")
            
            print '</div>'
        
        print """
                    </div>
                    
                    <!-- Data Table for Reference -->
                    <table class="data-table">
                        <tr>
                            <th>Month</th>
                            <th>Notes</th>
                            <th>Tasks</th>
                            <th>Total</th>
                        </tr>
        """
        
        # Reset the data query for the table
        monthly_data = q.QuerySql(get_monthly_activity_sql(12))  # Last 12 months
        for row in monthly_data:
            print """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                </tr>
            """.format(
                row.MonthName,
                row.NoteCount,
                row.TaskCount,
                row.TotalCount
            )
        
        print """
                    </table>
                </div>
            </div>
        </div>
        """
        
    except Exception as e:
        # Handle errors gracefully
        print "<p>Error loading summary data: {0}</p>".format(str(e))

def render_activity_tab(days_filter, page=1):
    """Render the recent activity tab with pagination and filters"""
    items_per_page = 20
    page = int(page)  # Ensure page is an integer
    offset = (page - 1) * items_per_page
    
    # Get filter values from request if they exist
    type_filter = None
    status_filter = None
    try:
        if hasattr(model.Data, 'type_filter') and model.Data.type_filter is not None and model.Data.type_filter.strip():
            type_filter = model.Data.type_filter
        if hasattr(model.Data, 'status_filter') and model.Data.status_filter is not None and model.Data.status_filter.strip():
            status_filter = model.Data.status_filter
    except:
        pass
    
    print """
        <div class="card">
            <h2 class="card-title">Recent Activity (Last {0} Days)</h2>
            
            <!-- Add filter form -->
            <form method="get" action="" class="filter-controls">
                <input type="hidden" name="tab" value="activity">
                <input type="hidden" name="days" value="{1}">
                <input type="hidden" name="page" value="1">
                
                <label for="type-filter">Type:</label>
                <select id="type-filter" name="type_filter">
                    <option value="" {2}>All Types</option>
                    <option value="Task" {3}>Tasks</option>
                    <option value="Note" {4}>Notes</option>
                </select>
                
                <label for="status-filter">Status:</label>
                <select id="status-filter" name="status_filter">
                    <option value="" {5}>All Statuses</option>
                    <option value="1" {6}>Complete</option>
                    <option value="2" {7}>Pending</option>
                    <option value="3" {8}>Accepted</option>
                    <option value="4" {9}>Declined</option>
                    <option value="5" {10}>Note</option>
                    <option value="6" {11}>Archived</option>
                </select>
                
                <button type="submit">Apply Filters</button>
            </form>
            
            <table class="data-table">
                <tr>
                    <th>Date</th>
                    <th>Type</th>
                    <th>Status</th>
                    <th>Owner</th>
                    <th>About</th>
                    <th>Assignee</th>
                    <th>Content</th>
                    <th>Keywords</th>
                </tr>
    """.format(
        days_filter,
        days_filter,
        "selected" if not type_filter else "",
        "selected" if type_filter == "Task" else "",
        "selected" if type_filter == "Note" else "",
        "selected" if not status_filter else "",
        "selected" if status_filter == "1" else "",
        "selected" if status_filter == "2" else "",
        "selected" if status_filter == "3" else "",
        "selected" if status_filter == "4" else "",
        "selected" if status_filter == "5" else "",
        "selected" if status_filter == "6" else "",
    )
    
    try:
        # Create WHERE clauses for filters
        where_clauses = ["tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())".format(days_filter)]
        
        if type_filter:
            if type_filter == "Note":
                where_clauses.append("tn.IsNote = 1")
            elif type_filter == "Task":
                where_clauses.append("tn.IsNote = 0")
        
        if status_filter:
            where_clauses.append("tn.StatusId = {0}".format(status_filter))
        
        # Combine all WHERE clauses
        where_clause = " AND ".join(where_clauses)
        
        # Modify the SQL query to include filters
        activity_sql = """
            SELECT
                tn.TaskNoteId,
                tn.CreatedDate,
                CASE WHEN tn.IsNote = 1 THEN 'Note' ELSE 'Task' END AS ActivityType,
                tn.StatusId,
                tn.OwnerId,
                owner.Name AS OwnerName,
                tn.AboutPersonId,
                about.Name AS AboutName,
                tn.AssigneeId,
                assignee.Name AS AssigneeName,
                tn.Instructions,
                tn.Notes,
                tn.CompletedDate,
                tn.DueDate,
                (
                    SELECT STRING_AGG(k.Description, ', ')
                    FROM TaskNoteKeyword tnk
                    JOIN Keyword k ON tnk.KeywordId = k.KeywordId
                    WHERE tnk.TaskNoteId = tn.TaskNoteId
                ) AS Keywords
            FROM TaskNote tn
            LEFT JOIN People owner ON tn.OwnerId = owner.PeopleId
            LEFT JOIN People about ON tn.AboutPersonId = about.PeopleId
            LEFT JOIN People assignee ON tn.AssigneeId = assignee.PeopleId
            WHERE {0}
            ORDER BY tn.CreatedDate DESC
            OFFSET {1} ROWS
            FETCH NEXT {2} ROWS ONLY
        """.format(where_clause, offset, items_per_page)
        
        # Count query with filters
        count_sql = """
            SELECT COUNT(*) AS TotalCount
            FROM TaskNote tn
            WHERE {0}
        """.format(where_clause)
        
        # Get total count for pagination
        total_count_result = q.QuerySqlTop1(count_sql)
        total_count = total_count_result.TotalCount if hasattr(total_count_result, 'TotalCount') else 0
        total_pages = (total_count + items_per_page - 1) // items_per_page if total_count > 0 else 1
        
        # Get paginated activity data
        activity_data = q.QuerySql(activity_sql)
        
        for activity in activity_data:
            # Determine status class
            status_class = "status-note" if activity.ActivityType == "Note" else {
                1: "status-complete",    # Was "status-pending"
                2: "status-pending",     # Was "status-accepted"
                3: "status-accepted",    # Was "status-completed"
                4: "status-declined",    # This one was correct
                5: "status-note",        # Was "status-archived"
                6: "status-archived"     # Was "status-note"
            }.get(activity.StatusId, "")
            
            # Format content for display
            content = activity.Notes if activity.Notes else activity.Instructions
            content = truncate_text(content, 150)
            
            print """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td class="{8}">{2}</td>
                <td>{3}</td>
                <td>{4}</td>
                <td>{5}</td>
                <td>{6}</td>
                <td>{7}</td>
            </tr>
            """.format(
                format_time_ago(activity.CreatedDate),
                activity.ActivityType,
                get_status_description(activity.StatusId),
                get_person_link(activity.OwnerId, activity.OwnerName),
                get_person_link(activity.AboutPersonId, activity.AboutName),
                get_person_link(activity.AssigneeId, activity.AssigneeName),
                content,
                activity.Keywords if activity.Keywords else "",
                status_class
            )
        
        # Render pagination controls
        if total_pages > 1:
            print """
            <tr>
                <td colspan="8">
                    <div class="pagination">
            """
            
            # Previous page link
            if page > 1:
                print """<a href="?days={0}&page={1}&tab=activity&type_filter={2}&status_filter={3}">&laquo; Previous</a>""".format(
                    days_filter, page - 1, type_filter or "", status_filter or "")
            else:
                print """<span class="disabled">&laquo; Previous</span>"""
            
            # Page numbers
            page_window = 5  # Show 5 page numbers at a time
            start_page = max(1, page - page_window // 2)
            end_page = min(total_pages, start_page + page_window - 1)
            
            # Adjust start_page if we're near the end
            start_page = max(1, end_page - page_window + 1)
            
            for p in range(start_page, end_page + 1):
                if p == page:
                    print """<span class="current">{0}</span>""".format(p)
                else:
                    print """<a href="?days={0}&page={1}&tab=activity&type_filter={2}&status_filter={3}">{1}</a>""".format(
                        days_filter, p, type_filter or "", status_filter or "")
            
            # Next page link
            if page < total_pages:
                print """<a href="?days={0}&page={1}&tab=activity&type_filter={2}&status_filter={3}">Next &raquo;</a>""".format(
                    days_filter, page + 1, type_filter or "", status_filter or "")
            else:
                print """<span class="disabled">Next &raquo;</span>"""
            
            print """
                    </div>
                </td>
            </tr>
            """
        
    except Exception as e:
        print "<tr><td colspan='8'>Error loading activity data: {0}</td></tr>".format(str(e))
        import traceback
        print "<tr><td colspan='8'><pre>{0}</pre></td></tr>".format(traceback.format_exc())
    
    print """
            </table>
        </div>
    """

def render_trends_tab(days_filter):
    """Render the keyword trends tab"""
    print """
        <div class="card">
            <h2 class="card-title">Keyword Trends (Last {0} Days)</h2>
            <p>This shows how keyword usage has changed over time, divided into three periods.</p>
            
            <table class="data-table">
                <tr>
                    <th>Keyword</th>
                    <th>First Period</th>
                    <th>Middle Period</th>
                    <th>Recent Period</th>
                    <th>Trend</th>
                </tr>
    """.format(days_filter)
    
    try:
        keyword_trend_data = q.QuerySql(get_keyword_trend_sql(days_filter))
        
        for trend in keyword_trend_data:
            # Determine trend direction
            trend_class = ""
            trend_arrow = ""
            
            if trend.Trend > 0:
                trend_class = "trend-up"
                trend_arrow = "&#x25B2;"  # Up arrow
            elif trend.Trend < 0:
                trend_class = "trend-down"
                trend_arrow = "&#x25BC;"  # Down arrow
            else:
                trend_class = "trend-neutral"
                trend_arrow = "&#x25AC;"  # Horizontal line
            
            print """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
                <td>{3}</td>
                <td class="{4}">{5} <span class="trend-arrow">{6}</span></td>
            </tr>
            """.format(
                trend.KeywordName,
                trend.Period1Count,
                trend.Period2Count,
                trend.Period3Count,
                trend_class,
                trend.Trend,
                trend_arrow
            )
        
    except Exception as e:
        print "<tr><td colspan='5'>Error loading trend data: {0}</td></tr>".format(str(e))
    
    print """
            </table>
        </div>
    """

def render_overdue_tab(days_filter=None):
    """Render the overdue tasks tab"""
    print """
        <div class="card">
            <h2 class="card-title">Overdue Tasks by Assignee</h2>
            
            <table class="data-table">
                <tr>
                    <th>Assignee</th>
                    <th>Overdue Tasks</th>
                    <th>Oldest Due Date</th>
                    <th>Action</th>
                </tr>
    """
    
    try:
        overdue_data = q.QuerySql(get_overdue_tasks_by_assignee_sql(days_filter))
        
        for assignee in overdue_data:
            print """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
                <td><a href="?assignee={3}&tab=overdue&days={4}" class="btn">View Details</a></td>
            </tr>
            """.format(
                get_person_link(assignee.PeopleId, assignee.Name),
                assignee.OverdueTasks,
                format_date(assignee.OldestDueDate),
                assignee.PeopleId,
                days_filter or ''  # Pass the current day filter to maintain state
            )
        
    except Exception as e:
        print "<tr><td colspan='4'>Error loading overdue data: {0}</td></tr>".format(str(e))
    
    print """
            </table>
        </div>
    """

def render_overdue_tasks_detail(assignee_id):
    """Render detailed view of overdue tasks for a specific assignee"""
    try:
        # Get the days filter from the request
        days_filter = None
        try:
            if model.Data.days is not None:
                days_filter = int(model.Data.days)
        except:
            days_filter = None
            
        # Get assignee name
        assignee_name = q.QuerySqlScalar("SELECT Name FROM People WHERE PeopleId = {0}".format(assignee_id))
        
        # Create time filter condition
        time_filter = ""
        if days_filter:
            time_filter = "AND tn.DueDate >= DATEADD(day, -{0}, GETDATE())".format(days_filter)
            
        # Get overdue tasks for this assignee with time filter
        sql = """
            SELECT 
                tn.TaskNoteId,
                tn.DueDate,
                DATEDIFF(day, tn.DueDate, GETDATE()) AS DaysOverdue,
                tn.Instructions,
                tn.Notes,
                owner.Name AS OwnerName,
                about.Name AS AboutName,
                tn.StatusId,
                (
                    SELECT STRING_AGG(k.Description, ', ')
                    FROM TaskNoteKeyword tnk
                    JOIN Keyword k ON tnk.KeywordId = k.KeywordId
                    WHERE tnk.TaskNoteId = tn.TaskNoteId
                ) AS Keywords
            FROM TaskNote tn
            LEFT JOIN People owner ON owner.PeopleId = tn.OwnerId
            LEFT JOIN People about ON about.PeopleId = tn.AboutPersonId
            WHERE 
                tn.AssigneeId = {0}
                AND tn.DueDate IS NOT NULL  -- Only include tasks with due dates
                AND tn.DueDate < GETDATE() 
                AND tn.StatusId NOT IN (1, 6) -- Not completed or archived
                AND tn.IsNote = 0 -- Only tasks, not notes
                {1}
            ORDER BY tn.DueDate ASC
        """.format(assignee_id, time_filter)
        
        overdue_tasks = q.QuerySql(sql)
        
        # Construct the back URL with the days parameter if it exists
        back_url = "?tab=overdue"
        if days_filter:
            back_url += "&days=" + str(days_filter)
        
        print """
        <div class="dashboard-container">
            <a href="{0}" class="back-button">&larr; Back to Dashboard</a>
            
            <h1>Overdue Tasks for {1}</h1>
            
            <div class="card">
                <h2 class="card-title">Overdue Tasks ({2}){3}</h2>
                
                <div class="overdue-list">
        """.format(
            back_url, 
            assignee_name, 
            len(overdue_tasks),
            " (Last {0} Days)".format(days_filter) if days_filter else ""
        )
        
        for task in overdue_tasks:
            # Format content for display
            content = task.Notes if task.Notes else task.Instructions
            
            # Format days overdue text
            days_text = "{0} day{1} overdue".format(
                task.DaysOverdue, 
                "s" if task.DaysOverdue != 1 else ""
            )
            
            # Format keywords as badges
            keyword_html = ""
            if task.Keywords:
                keywords = task.Keywords.split(", ")
                for keyword in keywords:
                    keyword_html += '<span class="overdue-task-keyword">{0}</span>'.format(keyword)
            
            print """
            <div class="overdue-task">
                <div class="overdue-task-header">
                    <div class="overdue-task-title">Task about {0}</div>
                    <div class="overdue-task-days">{1}</div>
                </div>
                <div class="overdue-task-details">
                    <strong>Due:</strong> {2} | <strong>Owner:</strong> {3}
                </div>
                <div class="overdue-task-details">
                    {4}
                </div>
                <div class="overdue-task-keywords">
                    {5}
                </div>
            </div>
            """.format(
                task.AboutName,
                days_text,
                format_date(task.DueDate),
                task.OwnerName,
                content,
                keyword_html
            )
        
        print """
                </div>
            </div>
        </div>
        """
    except Exception as e:
        print "<h2>Error</h2>"
        print "<p>Error loading overdue task details: {0}</p>".format(str(e))
        print '<a href="javascript:history.back()" class="back-button">&larr; Back to Dashboard</a>'

def render_stats_tab(days_filter):
    """Render the analytics tab with detailed stats"""
    # Determine the label based on filter
    time_period_label = "(Last {0} Days)".format(days_filter) if days_filter else "(All Time)"
    
    print """
        <div class="flex-row">
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Task Status Distribution {0}</h2>
    """.format(time_period_label)
    
    # Show status distribution as simple visual components
    try:
        summary_data = q.QuerySqlTop1(get_task_summary_sql(days_filter))
        
        status_data = [
            ("Pending", summary_data.PendingTasks, "#e67e22", "status-pending"),
            ("Accepted", summary_data.AcceptedTasks, "#3498db", "status-accepted"),
            ("Completed", summary_data.CompletedTasks, "#2ecc71", "status-completed"),
            ("Declined", summary_data.DeclinedTasks, "#e74c3c", "status-declined"),
            ("Archived", summary_data.ArchivedTasks, "#7f8c8d", "status-archived"),
            ("Notes", summary_data.TotalNotes, "#9b59b6", "status-note")
        ]
        
        # Calculate total for percentage
        total_count = sum(count for _, count, _, _ in status_data)
        if total_count == 0:
            total_count = 1  # Prevent division by zero
        
        # Create simple pie chart with colored boxes
        print '<div class="simple-pie-container">'
        
        for label, count, color, class_name in status_data:
            percent = float(count) / total_count * 100
            
            print """
            <div class="simple-pie-segment">
                <div class="simple-pie-color" style="background-color: {0};"></div>
                <div class="simple-pie-label {3}">{1}</div>
                <div class="simple-pie-value {3}">{2} ({4:.1f}%)</div>
            </div>
            """.format(color, label, count, class_name, percent)
        
        print '</div>'
        
        # Also create horizontal bars
        for label, count, color, class_name in status_data:
            print html_bar(count, total_count, color=color, width=300, label=label)
        
        print """
                    <table class="data-table">
                        <tr>
                            <th>Status</th>
                            <th>Count</th>
                            <th>Percentage</th>
                        </tr>
        """
        
        for label, count, _, class_name in status_data:
            percent = float(count) / total_count * 100
            print """
            <tr>
                <td class="{2}">{0}</td>
                <td>{1}</td>
                <td>{3:.1f}%</td>
            </tr>
            """.format(label, count, class_name, percent)
        
        print """
                    </table>
                </div>
            </div>
            
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Task Type Distribution {0}</h2>
        """.format(time_period_label)
        
        # Create simple visual for task types
        task_types = [
            ("Notes", summary_data.TotalNotes, "#9b59b6"),
            ("Tasks", summary_data.TotalActionTasks, "#3498db")
        ]
        
        # Calculate total
        total_type_count = sum(count for _, count, _ in task_types)
        if total_type_count == 0:
            total_type_count = 1  # Prevent division by zero
        
        # Create simple visual
        for label, count, color in task_types:
            percent = float(count) / total_type_count * 100
            print html_bar(count, total_type_count, color=color, width=300, label="{0} ({1:.1f}%)".format(label, percent))
        
        # Also show simple stats
        print """
                    <table class="data-table">
                        <tr>
                            <th>Type</th>
                            <th>Count</th>
                            <th>Percentage</th>
                        </tr>
        """
        
        for label, count, color in task_types:
            percent = float(count) / total_type_count * 100
            print """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2:.1f}%</td>
                </tr>
            """.format(label, count, percent)
        
        print """
                    </table>
                </div>
                
                <div class="card">
                    <h2 class="card-title">Weekly Activity Stats</h2>
                    <div class="stat-grid">
        """
        
        # Weekly stats
        weekly_stats = [
            ("Created Last Week", summary_data.CreatedLast7Days, "bg-blue"),
            ("Completed Last Week", summary_data.CompletedLast7Days, "bg-green"),
            ("Completion Rate", "{0:.1f}%".format(
                (float(summary_data.CompletedLast7Days) / summary_data.CreatedLast7Days * 100) 
                if summary_data.CreatedLast7Days > 0 else 0
            ), "bg-orange")
        ]
        
        for label, value, color_class in weekly_stats:
            print """
                <div class="stat-box {2}">
                    <div class="number">{1}</div>
                    <div class="label">{0}</div>
                </div>
            """.format(label, value, color_class)
        
        print """
                    </div>
                </div>
            </div>
        </div>
        """
        
    except Exception as e:
        # Handle errors gracefully
        print "<p>Error loading stats data: {0}</p>".format(str(e))

def render_completion_kpi_tab(days_filter):
    """Render the task completion KPI tab with metrics and charts"""
    print """
        <div class="flex-row">
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Task Completion KPIs (Last {0} Days)</h2>
                    <div class="stat-grid">
    """.format(days_filter)
    
    # Fetch KPI data
    try:
        kpi_data = q.QuerySqlTop1(get_completion_kpi_sql(days_filter))
        
        # Display summary KPI boxes
        # The issue is in these format strings - fixing them below
        on_time_percentage = kpi_data.OnTimePercentage if hasattr(kpi_data, 'OnTimePercentage') and kpi_data.OnTimePercentage is not None else 0
        avg_days = kpi_data.AvgCompletionDays if hasattr(kpi_data, 'AvgCompletionDays') and kpi_data.AvgCompletionDays is not None else 0
        min_days = kpi_data.MinCompletionDays if hasattr(kpi_data, 'MinCompletionDays') and kpi_data.MinCompletionDays is not None else 0
        max_days = kpi_data.MaxCompletionDays if hasattr(kpi_data, 'MaxCompletionDays') and kpi_data.MaxCompletionDays is not None else 0
        
        kpi_boxes = [
            ("Total Completed", kpi_data.TotalCompleted if hasattr(kpi_data, 'TotalCompleted') else 0, "bg-blue"),
            ("On-Time Rate", "{0:.1f}%".format(float(on_time_percentage)), "bg-green" if on_time_percentage >= 75 else "bg-orange"),
            ("Avg Days to Complete", "{0:.1f}".format(float(avg_days)), "bg-purple"),
            ("Fastest Completion", "{0} day{1}".format(min_days, "s" if min_days != 1 else ""), "bg-green"),
            ("Slowest Completion", "{0} day{1}".format(max_days, "s" if max_days != 1 else ""), "bg-orange"),
        ]
        
        for label, value, color_class in kpi_boxes:
            print """
                <div class="stat-box {2}">
                    <div class="number">{1}</div>
                    <div class="label">{0}</div>
                </div>
            """.format(label, value, color_class)
            
        print """
                    </div>
                </div>
                
                <div class="card">
                    <h2 class="card-title">Task Completion Trends (Last 12 Weeks)</h2>
                    <div class="viz-container">
        """
        
        # Weekly completion trends
        weekly_trends = q.QuerySql(get_completion_trend_sql(12))
        
        # Find max values for scaling the bars
        max_tasks = 1  # Default to prevent division by zero
        for week in weekly_trends:
            tasks_created = week.TasksCreated if hasattr(week, 'TasksCreated') and week.TasksCreated is not None else 0
            tasks_completed = week.TasksCompleted if hasattr(week, 'TasksCompleted') and week.TasksCompleted is not None else 0
            max_tasks = max(max_tasks, tasks_created, tasks_completed)
        
        # Create legend
        print """
            <div class="viz-legend">
                <div class="viz-legend-item">
                    <div class="viz-color-box" style="background-color: #3498db;"></div>
                    <span>Tasks Created</span>
                </div>
                <div class="viz-legend-item">
                    <div class="viz-color-box" style="background-color: #2ecc71;"></div>
                    <span>Tasks Completed</span>
                </div>
                <div class="viz-legend-item">
                    <div class="viz-color-box" style="background-color: #f39c12;"></div>
                    <span>Completion Rate</span>
                </div>
            </div>
        """
        
        # Create bar charts for each week
        for week in weekly_trends:
            # Skip weeks with no activity
            tasks_created = week.TasksCreated if hasattr(week, 'TasksCreated') and week.TasksCreated is not None else 0
            tasks_completed = week.TasksCompleted if hasattr(week, 'TasksCompleted') and week.TasksCompleted is not None else 0
            
            if tasks_created == 0 and tasks_completed == 0:
                continue
                
            print """
            <div style="margin-bottom: 20px;">
                <div style="font-weight: bold; margin-bottom: 5px;">{0}</div>
            """.format(week.WeekLabel if hasattr(week, 'WeekLabel') else "")
            
            # Created tasks bar
            print html_bar(tasks_created, max_tasks, color="#3498db", width=300, label="Created")
            
            # Completed tasks bar
            print html_bar(tasks_completed, max_tasks, color="#2ecc71", width=300, label="Completed")
            
            # Completion rate (scaled to 100%)
            completion_rate = week.CompletionRate if hasattr(week, 'CompletionRate') and week.CompletionRate is not None else 0
            rate_color = "#2ecc71" if completion_rate >= 75 else "#f39c12" if completion_rate >= 50 else "#e74c3c"
            print html_bar(float(completion_rate), 100, color=rate_color, width=300, label="{0:.1f}%".format(float(completion_rate)))
            
            print """</div>"""
        
        print """
                    </div>
                    <table class="data-table">
                        <tr>
                            <th>Week</th>
                            <th>Created</th>
                            <th>Completed</th>
                            <th>On-Time</th>
                            <th>On-Time %</th>
                            <th>Completion %</th>
                        </tr>
        """
        
        # Data table for the trends
        for week in weekly_trends:
            week_label = week.WeekLabel if hasattr(week, 'WeekLabel') else ""
            tasks_created = week.TasksCreated if hasattr(week, 'TasksCreated') and week.TasksCreated is not None else 0
            tasks_completed = week.TasksCompleted if hasattr(week, 'TasksCompleted') and week.TasksCompleted is not None else 0
            tasks_completed_on_time = week.TasksCompletedOnTime if hasattr(week, 'TasksCompletedOnTime') and week.TasksCompletedOnTime is not None else 0
            on_time_percentage = week.OnTimePercentage if hasattr(week, 'OnTimePercentage') and week.OnTimePercentage is not None else 0
            completion_rate = week.CompletionRate if hasattr(week, 'CompletionRate') and week.CompletionRate is not None else 0
            
            print """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>{4:.1f}%</td>
                    <td>{5:.1f}%</td>
                </tr>
            """.format(
                week_label,
                tasks_created,
                tasks_completed,
                tasks_completed_on_time,
                float(on_time_percentage),
                float(completion_rate)
            )
            
        print """
                    </table>
                </div>
            </div>
            
            <div class="flex-col">
                <div class="card">
                    <h2 class="card-title">Completion Efficiency by Assignee</h2>
                    <div id="assignee-efficiency">
        """
        
        # Assignee efficiency data
        assignee_data = q.QuerySql(get_completion_by_assignee_sql(days_filter, 30))
        
        for assignee in assignee_data:
            # Skip if no data
            completed_tasks = assignee.CompletedTasks if hasattr(assignee, 'CompletedTasks') and assignee.CompletedTasks is not None else 0
            on_time_tasks = assignee.OnTimeTasks if hasattr(assignee, 'OnTimeTasks') and assignee.OnTimeTasks is not None else 0
            pending_tasks = assignee.PendingTasks if hasattr(assignee, 'PendingTasks') and assignee.PendingTasks is not None else 0
            overdue_tasks = assignee.OverdueTasks if hasattr(assignee, 'OverdueTasks') and assignee.OverdueTasks is not None else 0
            
            # Double-check the calculation here
            if completed_tasks > 0:
                on_time_pct = (float(on_time_tasks) / completed_tasks) * 100
            else:
                on_time_pct = 0
            
            if not completed_tasks and not pending_tasks and not overdue_tasks:
                continue
                
            print """
            <div style="margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid #eee;">
                <div style="font-weight: bold; margin-bottom: 10px;">
                    {0}
                </div>
            """.format(get_person_link(assignee.AssigneeId, assignee.AssigneeName if hasattr(assignee, 'AssigneeName') else ""))
            
            # Completion rate (total completed / total assigned)
            completion_rate = assignee.CompletionRate if hasattr(assignee, 'CompletionRate') and assignee.CompletionRate is not None else 0
            rate_color = "#2ecc71" if completion_rate >= 75 else "#f39c12" if completion_rate >= 50 else "#e74c3c"
            
            print """
                <div style="margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
                        <span>Completion Rate:</span>
                        <span style="color: {1};">{0:.1f}%</span>
                    </div>
            """.format(float(completion_rate), rate_color)
            
            print html_bar(float(completion_rate), 100, color=rate_color, width=300)
            print """</div>"""
            
            # On-time percentage (completed on time / completed)
            on_time_pct = assignee.OnTimePercentage if hasattr(assignee, 'OnTimePercentage') and assignee.OnTimePercentage is not None else 0
            on_time_color = "#2ecc71" if on_time_pct >= 75 else "#f39c12" if on_time_pct >= 50 else "#e74c3c"
            
            print """
                <div style="margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
                        <span>On-Time Rate:</span>
                        <span style="color: {1};">{0:.1f}%</span>
                    </div>
            """.format(float(on_time_pct), on_time_color)
            
            print html_bar(float(on_time_pct), 100, color=on_time_color, width=300)
            print """</div>"""
            
            # Avg completion time
            avg_days = assignee.AvgCompletionDays if hasattr(assignee, 'AvgCompletionDays') and assignee.AvgCompletionDays is not None else 0
            accepted_tasks = assignee.AcceptedTasks if hasattr(assignee, 'AcceptedTasks') and assignee.AcceptedTasks is not None else 0
            
            print """
                <div style="margin-top: 10px; display: flex; justify-content: space-between;">
                    <div>
                        <span style="font-size: 13px; color: #777;">Completed:</span>
                        <span style="margin-left: 5px; font-weight: bold;">{0}</span>
                    </div>
                    <div>
                        <span style="font-size: 13px; color: #777;">Pending:</span>
                        <span style="margin-left: 5px; font-weight: bold;">{1}</span>
                    </div>
                    <div>
                        <span style="font-size: 13px; color: #777;">Overdue:</span>
                        <span style="margin-left: 5px; font-weight: bold; color: {3};">{2}</span>
                    </div>
                    <div>
                        <span style="font-size: 13px; color: #777;">Avg Time:</span>
                        <span style="margin-left: 5px; font-weight: bold;">{4:.1f} days</span>
                    </div>
                </div>
            """.format(
                completed_tasks,
                pending_tasks + accepted_tasks,
                overdue_tasks,
                "#e74c3c" if overdue_tasks > 0 else "#777",
                float(avg_days)
            )
            
            print """</div>"""
            
        print """
                    </div>
                    
                    <table class="data-table">
                        <tr>
                            <th>Assignee</th>
                            <th>Completed</th>
                            <th>On-Time</th>
                            <th>Late</th>
                            <th>Pending</th>
                            <th>Overdue</th>
                            <th>Avg Days</th>
                            <th>On-Time %</th>
                        </tr>
        """
        
        # Reset the query to get fresh data
        assignee_data = q.QuerySql(get_completion_by_assignee_sql(days_filter))
        
        for assignee in assignee_data:
            assignee_name = assignee.AssigneeName if hasattr(assignee, 'AssigneeName') else ""
            completed_tasks = assignee.CompletedTasks if hasattr(assignee, 'CompletedTasks') else 0
            on_time_tasks = assignee.OnTimeTasks if hasattr(assignee, 'OnTimeTasks') else 0
            late_tasks = assignee.LateTasks if hasattr(assignee, 'LateTasks') else 0
            pending_tasks = assignee.PendingTasks if hasattr(assignee, 'PendingTasks') else 0
            accepted_tasks = assignee.AcceptedTasks if hasattr(assignee, 'AcceptedTasks') else 0
            overdue_tasks = assignee.OverdueTasks if hasattr(assignee, 'OverdueTasks') else 0
            avg_completion_days = assignee.AvgCompletionDays if hasattr(assignee, 'AvgCompletionDays') and assignee.AvgCompletionDays is not None else 0
            on_time_percentage = assignee.OnTimePercentage if hasattr(assignee, 'OnTimePercentage') and assignee.OnTimePercentage is not None else 0
            
            print """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6:.1f}</td>
                    <td>{7:.1f}%</td>
                </tr>
            """.format(
                get_person_link(assignee.AssigneeId if hasattr(assignee, 'AssigneeId') else 0, assignee_name),
                completed_tasks,
                on_time_tasks,
                late_tasks,
                pending_tasks + accepted_tasks,
                overdue_tasks,
                float(avg_completion_days),
                float(on_time_percentage)
            )
            
        print """
                    </table>
                </div>
            </div>
        </div>
        """
        
    except Exception as e:
        # Handle errors gracefully
        import traceback
        print "<h2>Error</h2>"
        print "<p>Error loading completion KPI data: " + str(e) + "</p>"
        print "<pre>"
        traceback.print_exc()
        print "</pre>"

# ---------------- MAIN EXECUTION ----------------

try:
    # Main execution - render the dashboard
    render_dashboard()
except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
