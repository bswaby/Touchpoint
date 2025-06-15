#Role=Edit

"""
--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin > Advanced > Special Content > Python
2. Click New Python Script File
3. Name the Python script "LapsedAttendanceDashboard" and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Lapsed Attendance Dashboard - Complete Working Version
=====================================================
This script identifies people whose attendance patterns have significantly deviated 
from their normal behavior. It uses statistical analysis to find people who haven't
attended in longer than their normal pattern suggests.

Key Features:
- Statistical analysis of attendance gaps
- ag-Grid with visual indicators
- Bulk tagging and task creation
- Contact tracking
- Configurable thresholds
- Working person search for task assignment

The system looks for people whose current absence duration exceeds their normal 
pattern by 2+ standard deviations.

Standard Devitation: Standard deviation is a measure of how spread out or varied a set 
of values is from the average (mean). In this case, it represents how much a person's 
absence duration typically varies from their usual average. If someone's current 
absence duration exceeds their average by more than 2 standard deviations, it's 
considered unusually high or abnormal.

written by: Ben Swaby
email: bswaby@fbchtn.org

"""

import datetime
from datetime import datetime

# =============================================================================
# ::START:: Configuration Variables - CUSTOMIZE FOR YOUR CHURCH
# =============================================================================

# Church-specific settings - MODIFY THESE FOR YOUR SETUP
EXCLUDED_PROGRAM_IDS = "1108,1109,1143,1149"  # Programs to exclude (modify for your church)
MIN_ATTENDANCE_COUNT = 2                       # Minimum attendance records needed
MIN_STD_DEVIATION = 2                         # Minimum std dev between absences
MAX_STD_DEVIATION = 5                         # Maximum std dev between absences  
DEVIATION_THRESHOLD = 2                       # Standard deviations to trigger "lapsed"
ANALYSIS_PERIOD_MONTHS = 12                   # Months to look back for analysis
MAX_RECORDS = 1000                             # Limit records to prevent timeouts

# Display settings
SHOW_PHOTOS = True                            # Show profile photos in grid
ENABLE_BULK_ACTIONS = True                    # Enable bulk tagging/tasks
DEFAULT_TAG_NAME = "Lapsed-Attendance"        # Default tag name for bulk actions

# Contact tracking settings  
RECENT_CONTACT_DAYS = 7                       # Days to consider "recent contact"
SOME_CONTACT_DAYS = 30                        # Days to consider "some contact"

# =============================================================================
# ::END:: Configuration Variables
# =============================================================================

# =============================================================================
# ::START:: Utility Functions
# =============================================================================

def get_form_value(field_name, default_value=''):
    """Safely get form field value with default"""
    try:
        if hasattr(model.Data, field_name):
            value = getattr(model.Data, field_name)
            return str(value) if value is not None else default_value
        return default_value
    except:
        return default_value

def escape_html(text):
    """Escape HTML characters"""
    if not text:
        return ''
    return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def print_error(message):
    """Print user-friendly error message"""
    print '''
    <div style="margin: 20px 0; padding: 15px; background: #fee2e2; border: 1px solid #fecaca; border-radius: 8px; color: #dc2626;">
        <strong>Error:</strong> {0}
    </div>
    '''.format(escape_html(message))

def print_success(message):
    """Print success message"""
    print '''
    <div style="margin: 20px 0; padding: 15px; background: #dcfce7; border: 1px solid #bbf7d0; border-radius: 8px; color: #16a34a;">
        <strong>Success:</strong> {0}
    </div>
    '''.format(escape_html(message))

def print_loading():
    """Print loading indicator"""
    print '''
    <div id="loading" style="text-align: center; padding: 40px; color: #6b7280;">
        <div style="font-size: 24px; margin-bottom: 10px;">⏳</div>
        <div>Loading attendance analysis...</div>
    </div>
    '''

# =============================================================================
# ::END:: Utility Functions
# =============================================================================

# =============================================================================
# ::START:: Data Retrieval Functions
# =============================================================================

def get_attendance_data(age_filter="all"):
    """Get lapsed attendance data with simplified query"""
    try:
        # ::STEP:: Build Age Filter
        age_condition = ""
        if age_filter == "adults":
            age_condition = "AND p.Age >= 18"
        elif age_filter == "children":
            age_condition = "AND p.Age <= 17"
        else:
            age_condition = "AND p.Age IS NOT NULL"
        
        # ::STEP:: Simplified SQL Query with Fixed HAVING Clause
        sql = """
        SELECT TOP {max_records}
            filtered_people.PeopleId,
            filtered_people.Name,
            filtered_people.Age,
            filtered_people.EmailAddress,
            filtered_people.CellPhone,
            filtered_people.pic,
            filtered_people.attendance_count,
            filtered_people.avg_gap_days,
            filtered_people.last_attendance_date,
            filtered_people.days_since_last,
            filtered_people.std_dev_gaps,
            filtered_people.std_devs_ago,
            ISNULL(contact_stats.days_since_contact, 999) as days_since_contact,
            ISNULL(contact_stats.contact_count, 0) as contact_count,
            org_info.org_names
            
        FROM (
            SELECT 
                p.PeopleId,
                CONCAT(COALESCE(p.NickName, p.FirstName), ' ', p.LastName) as Name,
                p.Age,
                p.EmailAddress,
                p.CellPhone,
                pic.ThumbUrl as pic,
                att_stats.attendance_count,
                att_stats.avg_gap_days,
                att_stats.last_attendance_date,
                att_stats.days_since_last,
                att_stats.std_dev_gaps,
                -- Calculate priority based on deviation
                CASE 
                    WHEN att_stats.std_dev_gaps > 0 
                    THEN (att_stats.days_since_last - att_stats.avg_gap_days) / att_stats.std_dev_gaps
                    ELSE 0
                END as std_devs_ago
                
            FROM People p
            
            -- Attendance statistics subquery with filtering built in
            INNER JOIN (
                SELECT 
                    att.PeopleId,
                    COUNT(*) as attendance_count,
                    AVG(CAST(gap_days as FLOAT)) as avg_gap_days,
                    MAX(att.MeetingDate) as last_attendance_date,
                    DATEDIFF(DAY, MAX(att.MeetingDate), GETDATE()) as days_since_last,
                    STDEV(CAST(gap_days as FLOAT)) as std_dev_gaps
                FROM (
                    SELECT 
                        a.PeopleId,
                        a.MeetingDate,
                        DATEDIFF(DAY, 
                            LAG(a.MeetingDate) OVER (PARTITION BY a.PeopleId ORDER BY a.MeetingDate),
                            a.MeetingDate
                        ) as gap_days
                    FROM Attend a
                    INNER JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                    INNER JOIN Division d ON o.DivisionId = d.Id
                    INNER JOIN ProgDiv pd ON d.Id = pd.DivId
                    INNER JOIN Program pr ON pd.ProgId = pr.Id
                    WHERE a.AttendanceFlag = 1
                        AND pr.Id NOT IN ({excluded_programs})
                        AND a.MeetingDate > DATEADD(MONTH, -{analysis_months}, GETDATE())
                ) att
                WHERE att.gap_days IS NOT NULL
                GROUP BY att.PeopleId
                HAVING COUNT(*) >= {min_attendance}
                    AND STDEV(CAST(gap_days as FLOAT)) BETWEEN {min_std} AND {max_std}
                    -- Filter for lapsed people in the subquery
                    AND (
                        CASE 
                            WHEN STDEV(CAST(gap_days as FLOAT)) > 0 
                            THEN (DATEDIFF(DAY, MAX(att.MeetingDate), GETDATE()) - AVG(CAST(gap_days as FLOAT))) / STDEV(CAST(gap_days as FLOAT))
                            ELSE 0
                        END
                    ) >= {deviation_threshold}
            ) att_stats ON p.PeopleId = att_stats.PeopleId
            
            LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
            
            WHERE p.DeceasedDate IS NULL
                AND p.ArchivedFlag = 0
                {age_condition}
        ) filtered_people
        
        -- Contact statistics
        LEFT JOIN (
            SELECT 
                tn.AboutPersonId,
                DATEDIFF(DAY, MAX(tn.CreatedDate), GETDATE()) as days_since_contact,
                COUNT(*) as contact_count
            FROM TaskNote tn
            WHERE tn.CreatedDate > DATEADD(MONTH, -{analysis_months}, GETDATE())
            GROUP BY tn.AboutPersonId
        ) contact_stats ON filtered_people.PeopleId = contact_stats.AboutPersonId
        
        -- Organization info (using STUFF/FOR XML for compatibility)  
        LEFT JOIN (
            SELECT DISTINCT
                om.PeopleId,
                STUFF((
                    SELECT ', ' + o2.OrganizationName
                    FROM OrganizationMembers om2
                    INNER JOIN Organizations o2 ON om2.OrganizationId = o2.OrganizationId
                    WHERE om2.PeopleId = om.PeopleId
                        AND om2.InactiveDate IS NULL
                    FOR XML PATH('')
                ), 1, 2, '') as org_names
            FROM OrganizationMembers om
            WHERE om.InactiveDate IS NULL
        ) org_info ON filtered_people.PeopleId = org_info.PeopleId
        
        ORDER BY filtered_people.std_devs_ago DESC
        """.format(
            max_records=MAX_RECORDS,
            excluded_programs=EXCLUDED_PROGRAM_IDS,
            analysis_months=ANALYSIS_PERIOD_MONTHS,
            min_attendance=MIN_ATTENDANCE_COUNT,
            min_std=MIN_STD_DEVIATION,
            max_std=MAX_STD_DEVIATION,
            age_condition=age_condition,
            deviation_threshold=DEVIATION_THRESHOLD
        )
        
        # ::STEP:: Execute Query with Error Handling
        try:
            data = q.QuerySql(sql)
            return data if data is not None else []
        except Exception as sql_error:
            print_error("Database query failed: " + str(sql_error))
            return []
            
    except Exception as e:
        print_error("Error in get_attendance_data: " + str(e))
        return []

def get_dashboard_summary(data):
    """Calculate summary statistics"""
    try:
        if not data:
            return {
                'total_count': 0,
                'avg_days_ago': 0,
                'avg_std_devs': 0,
                'contacted_count': 0,
                'high_priority_count': 0
            }
        
        total_count = len(data)
        contacted_count = 0
        high_priority_count = 0
        total_days_ago = 0
        total_std_devs = 0
        
        for record in data:
            try:
                contact_count = getattr(record, 'contact_count', 0)
                std_devs_ago = getattr(record, 'std_devs_ago', 0)
                days_since_last = getattr(record, 'days_since_last', 0)
                
                if contact_count and contact_count > 0:
                    contacted_count += 1
                if std_devs_ago and std_devs_ago >= 3:
                    high_priority_count += 1
                    
                total_days_ago += days_since_last if days_since_last else 0
                total_std_devs += std_devs_ago if std_devs_ago else 0
            except:
                continue
        
        return {
            'total_count': total_count,
            'avg_days_ago': int(total_days_ago / total_count) if total_count > 0 else 0,
            'avg_std_devs': round(total_std_devs / total_count, 1) if total_count > 0 else 0,
            'contacted_count': contacted_count,
            'high_priority_count': high_priority_count
        }
    except Exception as e:
        print_error("Error calculating summary: " + str(e))
        return {'total_count': 0, 'avg_days_ago': 0, 'avg_std_devs': 0, 
                'contacted_count': 0, 'high_priority_count': 0}

# =============================================================================
# ::END:: Data Retrieval Functions
# =============================================================================

# =============================================================================
# ::START:: Form Processing Functions
# =============================================================================

def get_current_user_id():
    """Get current user's people ID"""
    try:
        user_query = q.QuerySqlTop1("SELECT PeopleId FROM Users WHERE Username = @p1", model.UserName)
        if user_query and hasattr(user_query, 'PeopleId'):
            return user_query.PeopleId
        return model.UserPeopleId if hasattr(model, 'UserPeopleId') else 1
    except:
        return 1

def find_person_by_name(name):
    """Try to find a person by name for task assignment"""
    try:
        if not name or name.strip() == '':
            return None
            
        # Search for person by name in Users first (for staff)
        sql = """
        SELECT TOP 1 u.PeopleId 
        FROM Users u 
        INNER JOIN People p ON u.PeopleId = p.PeopleId
        WHERE p.Name LIKE @p1 OR (p.FirstName + ' ' + p.LastName) LIKE @p1
        ORDER BY p.Name
        """
        
        search_term = '%' + name.strip() + '%'
        result = q.QuerySqlTop1(sql, search_term)
        
        if result and hasattr(result, 'PeopleId'):
            return result.PeopleId
            
        # If not found in Users, search all People
        sql2 = """
        SELECT TOP 1 p.PeopleId 
        FROM People p
        WHERE p.Name LIKE @p1 OR (p.FirstName + ' ' + p.LastName) LIKE @p1
        ORDER BY p.Name
        """
        
        result2 = q.QuerySqlTop1(sql2, search_term)
        if result2 and hasattr(result2, 'PeopleId'):
            return result2.PeopleId
            
        return None
        
    except:
        return None

def process_bulk_actions():
    """Process bulk actions (tag, task, note)"""
    try:
        # ::STEP:: Get Selected People
        people_ids_str = get_form_value('selected_people_ids', '')
        if not people_ids_str:
            print_error("No people selected")
            return False
            
        people_ids = [int(pid.strip()) for pid in people_ids_str.split(',') if pid.strip()]
        if not people_ids:
            print_error("Invalid people selection")
            return False
        
        # ::STEP:: Get Action Type
        action_type = get_form_value('action_type', 'tag')
        
        if action_type == 'tag':
            return process_tagging(people_ids)
        elif action_type == 'task':
            return process_task_creation(people_ids)
        elif action_type == 'note':
            return process_note_creation(people_ids)
        else:
            print_error("Unknown action type")
            return False
            
    except Exception as e:
        print_error("Error processing bulk actions: " + str(e))
        return False

def process_tagging(people_ids):
    """Process bulk tagging"""
    try:
        tag_name = get_form_value('tag_name', DEFAULT_TAG_NAME)
        
        # ::STEP:: Create Query from People IDs
        people_ids_str = ','.join([str(pid) for pid in people_ids])
        query = "peopleids='{0}'".format(people_ids_str)
        
        # ::STEP:: Apply Tag
        current_user_id = get_current_user_id()
        model.AddTag(query, tag_name, current_user_id, False)
        
        print_success("Successfully tagged {0} people with '{1}'".format(len(people_ids), tag_name))
        return True
        
    except Exception as e:
        print_error("Error during tagging: " + str(e))
        return False

def process_task_creation(people_ids):
    """Process bulk task creation with native TouchPoint fields only"""
    try:
        task_message = get_form_value('task_message', 'Follow up on lapsed attendance')
        assignee_search = get_form_value('assignee_search', '').strip()
        assignee_id = get_form_value('assignee_id', '')
        due_date_str = get_form_value('due_date', '')
        
        # ::STEP:: Handle Assignee
        if assignee_search and assignee_search != '':
            # Try to find person by name
            found_id = find_person_by_name(assignee_search)
            if found_id:
                assignee_id = found_id
            else:
                assignee_id = get_current_user_id()
                # Add note about search failure
                task_message = "[Note: Could not find user '{0}', assigned to you instead]\n\n{1}".format(assignee_search, task_message)
        elif not assignee_id:
            assignee_id = get_current_user_id()
        else:
            try:
                assignee_id = int(assignee_id)
            except:
                assignee_id = get_current_user_id()
        
        # Parse due date if provided
        due_date = None
        if due_date_str:
            try:
                due_date = model.ParseDate(due_date_str)
            except:
                pass
        
        # Get keywords (supports multiple selections)
        keywords = []
        for i in range(20):  # Allow up to 20 keywords
            keyword_field = "keyword_" + str(i)
            if hasattr(model.Data, keyword_field) and getattr(model.Data, keyword_field):
                keywords.append(int(getattr(model.Data, keyword_field)))
        
        current_user_id = get_current_user_id()
        success_count = 0
        
        # ::STEP:: Create Tasks with native TouchPoint parameters only
        for people_id in people_ids:
            try:
                task_id = model.CreateTaskNote(
                    ownerId=current_user_id,
                    aboutPersonId=int(people_id),
                    assigneeId=int(assignee_id),
                    roleId=None,
                    isNote=False,
                    instructions=task_message,
                    notes="",
                    dueDate=due_date,
                    keywordIdList=keywords if keywords else None,
                    sendEmails=True
                )
                
                if task_id:
                    success_count += 1
                    
            except Exception as task_error:
                continue  # Skip failed tasks, don't break the process
        
        print_success("Successfully created {0} of {1} tasks".format(success_count, len(people_ids)))
        return True
        
    except Exception as e:
        print_error("Error creating tasks: " + str(e))
        return False

def process_note_creation(people_ids):
    """Process bulk note creation with native TouchPoint fields only"""
    try:
        note_message = get_form_value('note_message', 'Lapsed attendance pattern identified')
        
        # Get keywords (supports multiple selections)
        keywords = []
        for i in range(20):  # Allow up to 20 keywords
            keyword_field = "keyword_" + str(i)
            if hasattr(model.Data, keyword_field) and getattr(model.Data, keyword_field):
                keywords.append(int(getattr(model.Data, keyword_field)))
        
        current_user_id = get_current_user_id()
        success_count = 0
        
        # ::STEP:: Create Notes
        for people_id in people_ids:
            try:
                note_id = model.CreateTaskNote(
                    ownerId=current_user_id,
                    aboutPersonId=int(people_id),
                    assigneeId=current_user_id,
                    roleId=None,
                    isNote=True,
                    instructions="",
                    notes=note_message,
                    dueDate=None,
                    keywordIdList=keywords if keywords else None,
                    sendEmails=False
                )
                
                if note_id:
                    success_count += 1
                    
            except Exception as note_error:
                continue  # Skip failed notes, don't break the process
        
        print_success("Successfully created {0} notes".format(success_count))
        return True
        
    except Exception as e:
        print_error("Error creating notes: " + str(e))
        return False

# =============================================================================
# ::END:: Form Processing Functions
# =============================================================================

# =============================================================================
# ::START:: Display Functions
# =============================================================================

def render_header(summary, age_filter):
    """Render dashboard header"""
    filter_text = {'all': 'All Ages', 'adults': 'Adults (18+)', 'children': 'Children (Under 18)'}.get(age_filter, 'All Ages')
    
    print '''
    <style>
        .dashboard-header {{ 
            margin: 20px 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
            border-radius: 10px; 
            color: white; 
        }}
        .summary-cards {{ 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
            gap: 15px; 
            margin-top: 20px;
        }}
        .summary-card {{ 
            background: rgba(255,255,255,0.1); 
            padding: 15px; 
            border-radius: 8px; 
            text-align: center;
        }}
        .summary-card .number {{ 
            font-size: 24px; 
            font-weight: bold; 
            margin-bottom: 5px;
        }}
        .summary-card .label {{ 
            opacity: 0.8; 
            font-size: 14px;
        }}
        .controls {{ 
            margin: 20px 0; 
            padding: 20px; 
            background: #f8fafc; 
            border: 1px solid #e2e8f0; 
            border-radius: 8px;
        }}
        .btn {{ 
            padding: 8px 16px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            margin: 0 5px;
        }}
        .btn-primary {{ background: #2563eb; color: white; }}
        .btn-success {{ background: #16a34a; color: white; }}
        .btn-warning {{ background: #ea580c; color: white; }}
        .btn-secondary {{ background: #6b7280; color: white; }}
        .form-group {{ margin-bottom: 15px; }}
        .form-group label {{ display: block; margin-bottom: 5px; font-weight: 600; }}
        .form-control {{ width: 100%; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px; }}
        .ag-row-selected {{ background-color: #e0f2fe !important; }}
        .ag-row-selected .ag-cell {{ background-color: #e0f2fe !important; }}
        .modal {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; }}
        .modal-content {{ position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 30px; border-radius: 10px; min-width: 500px; max-width: 700px; max-height: 90vh; overflow-y: auto; }}
        .search-input-wrapper {{ position: relative; }}
        .search-results {{ position: absolute; top: 100%; left: 0; right: 0; background: white; border: 1px solid #ddd; border-radius: 0 0 4px 4px; max-height: 200px; overflow-y: auto; z-index: 10; box-shadow: 0 4px 8px rgba(0,0,0,0.1); display: none; }}
        .search-result-item {{ padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #eee; }}
        .search-result-item:hover {{ background-color: #f5f5f5; }}
        .search-result-item:last-child {{ border-bottom: none; }}
        .keyword-container {{ max-height: 150px; overflow-y: auto; border: 1px solid #ddd; border-radius: 4px; padding: 10px; background: #f9f9f9; }}
        .keyword-item {{ margin-bottom: 5px; }}
        .keyword-item label {{ font-weight: normal; display: flex; align-items: center; cursor: pointer; }}
        .keyword-item input[type="checkbox"] {{ margin-right: 8px; }}
    </style>
    
    <div class="dashboard-header">
        <h1 style="margin: 0 0 10px 0; font-size: 28px;">📊 Lapsed Attendance Dashboard</h1>
        <p style="margin: 0; opacity: 0.9;">Identifying people whose attendance patterns have significantly deviated from normal behavior ({0})</p>
        
        <div class="summary-cards">
            <div class="summary-card">
                <div class="number">{1}</div>
                <div class="label">Total Lapsed</div>
            </div>
            <div class="summary-card">
                <div class="number">{2}</div>
                <div class="label">Avg Days Since Last</div>
            </div>
            <div class="summary-card">
                <div class="number">{3}</div>
                <div class="label">Have Been Contacted</div>
            </div>
            <div class="summary-card">
                <div class="number">{4}</div>
                <div class="label">High Priority (3+ σ)</div>
            </div>
        </div>
    </div>
    '''.format(
        filter_text,
        summary['total_count'],
        summary['avg_days_ago'],
        summary['contacted_count'],
        summary['high_priority_count']
    )

def render_controls():
    """Render control panel"""
    print '''
    <div class="controls">
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; align-items: end;">
            
            <div class="form-group">
                <label>Age Group Filter:</label>
                <select id="age_filter" class="form-control" onchange="filterData()">
                    <option value="all">All Ages</option>
                    <option value="adults">Adults (18+)</option>
                    <option value="children">Children (Under 18)</option>
                </select>
            </div>
            
            <div class="form-group">
                <label>Bulk Actions:</label>
                <div style="display: flex; gap: 5px;">
                    <select id="action_type" class="form-control" style="flex: 1;">
                        <option value="tag">🏷️ Tag People</option>
                        <option value="task">📋 Create Tasks</option>
                        <option value="note">📝 Add Notes</option>
                    </select>
                    <button onclick="showBulkActionDialog()" class="btn btn-primary">Execute</button>
                </div>
            </div>
            
            <div class="form-group">
                <label>Selected:</label>
                <div id="selection-summary" style="padding: 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; min-height: 45px; display: flex; align-items: center;">
                    <span id="selected-count">0</span> people selected
                </div>
            </div>
            
        </div>
    </div>
    '''

def convert_data_to_js_records(data):
    """Convert SQL result data to JavaScript records for ag-Grid"""
    try:
        if not data:
            return ''
        
        records = []
        for item in data:
            if item is None:
                continue
                
            try:
                # Safely get all values with defaults
                people_id = str(getattr(item, 'PeopleId', ''))
                name = escape_html(str(getattr(item, 'Name', '')))
                age = int(getattr(item, 'Age', 0)) if getattr(item, 'Age', 0) else 0
                std_devs_ago = float(getattr(item, 'std_devs_ago', 0))
                days_since_last = int(getattr(item, 'days_since_last', 0))
                avg_gap_days = float(getattr(item, 'avg_gap_days', 0))
                days_since_contact = int(getattr(item, 'days_since_contact', 999))
                contact_count = int(getattr(item, 'contact_count', 0))
                org_names = escape_html(str(getattr(item, 'org_names', '')))
                pic = str(getattr(item, 'pic', ''))
                attendance_count = int(getattr(item, 'attendance_count', 0))
                
                # Determine contact status
                if days_since_contact <= 7:
                    contact_status = 'Recent Contact'
                elif days_since_contact <= 30:
                    contact_status = 'Some Contact'
                elif contact_count > 0:
                    contact_status = 'Past Contact'
                else:
                    contact_status = 'No Contact'
                
                record_js = '{{peopleid: "{0}",name: "{1}",age: {2},stdevsago: {3},daysago: {4},avgdays: {5},contactdays: {6},contactstatus: "{7}",orgnames: "{8}",pic: "{9}",attendancecount: {10}}}'.format(
                    people_id, name, age, std_devs_ago, days_since_last,
                    round(avg_gap_days, 1), days_since_contact, contact_status,
                    org_names, pic, attendance_count
                )
                
                records.append(record_js)
            except Exception as record_error:
                # Skip problematic records rather than failing completely
                continue
        
        return ',\n            '.join(records)
    except Exception as e:
        print_error("Error converting data to JavaScript: " + str(e))
        return ''

def render_data_grid(data):
    """Render ag-Grid with enhanced features"""
    try:
        # ::STEP:: Convert Data to JavaScript Format
        records = convert_data_to_js_records(data)
        
        print '''
        <script src="https://cdn.jsdelivr.net/npm/ag-grid-enterprise/dist/ag-grid-enterprise.js"></script>
        <link rel="stylesheet" href="https://unpkg.com/ag-grid-community/styles/ag-grid.css" />
        <link rel="stylesheet" href="https://unpkg.com/ag-grid-community/styles/ag-theme-quartz.css" />
        
        <div style="margin: 20px 0;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                <h3 style="margin: 0;">📋 Detailed Analysis</h3>
                <div style="display: flex; gap: 10px; align-items: center;">
                    <button onclick="toggleColumnKey()" class="btn btn-secondary" style="font-size: 12px;">
                        ❓ Column Guide
                    </button>
                </div>
            </div>
            
            <!-- Column Key/Legend -->
            <div id="column-key" style="display: none; margin-bottom: 15px; padding: 15px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;">
                <h4 style="margin: 0 0 15px 0; color: #374151;">📖 Column Guide</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; font-size: 14px;">
                    <div>
                        <strong>Statistical Columns:</strong><br>
                        • <strong>Priority</strong> - How many standard deviations beyond normal absence<br>
                        • <strong>Days Ago</strong> - Days since their last attendance<br>
                        • <strong>Attendance Count</strong> - Number of times attended in the last year<br>
                        • <strong>Avg Gap</strong> - Average days between their attendances<br>
                    </div>
                    <div>
                        <strong>Contact Tracking:</strong><br>
                        • <strong>Contact Status</strong> - Recent contact activity level<br>
                        • <strong>Last Contact</strong> - Days since last recorded contact<br>
                    </div>
                    <div>
                        <strong>Priority Levels:</strong><br>
                        • <span style="background: #dc2626; color: white; padding: 2px 6px; border-radius: 8px;">URGENT</span> - 4+ standard deviations<br>
                        • <span style="background: #ea580c; color: white; padding: 2px 6px; border-radius: 8px;">HIGH</span> - 3+ standard deviations<br>
                        • <span style="background: #16a34a; color: white; padding: 2px 6px; border-radius: 8px;">MEDIUM</span> - 2+ standard deviations<br>
                    </div>
                </div>
            </div>
            
            <div id="attendance-grid" class="ag-theme-quartz" style="height: 600px; border: 1px solid #e2e8f0; border-radius: 8px;"></div>
        </div>

        <script>
        // Global variables and functions
        var selectedRows = [];
        var gridApi = null;
        
        function numberSort(num1, num2) {{
            return num1 - num2;
        }}
        
        function getPriorityLevel(stdevsago) {{
            if (stdevsago >= 4) {{
                return {{ text: 'URGENT', style: 'background: #dc2626; color: white;' }};
            }} else if (stdevsago >= 3) {{
                return {{ text: 'HIGH', style: 'background: #ea580c; color: white;' }};
            }} else {{
                return {{ text: 'MEDIUM', style: 'background: #16a34a; color: white;' }};
            }}
        }}
        
        function getContactStatusBadge(status) {{
            var colors = {{
                'Recent Contact': '#16a34a',
                'Some Contact': '#ea580c', 
                'Past Contact': '#0891b2',
                'No Contact': '#dc2626'
            }};
            var color = colors[status] || '#6b7280';
            return '<span style="background: ' + color + '; color: white; padding: 4px 8px; border-radius: 12px; font-size: 12px;">' + status + '</span>';
        }}
        
        function toggleColumnKey() {{
            var keyEl = document.getElementById('column-key');
            if (keyEl && keyEl.style.display === 'none') {{
                keyEl.style.display = 'block';
            }} else if (keyEl) {{
                keyEl.style.display = 'none';
            }}
        }}
        
        function updateSelectionDisplay() {{
            var count = selectedRows ? selectedRows.length : 0;
            var countEl = document.getElementById('selected-count');
            if (countEl) {{
                countEl.textContent = count;
            }}
            
            var summaryEl = document.getElementById('selection-summary');
            if (summaryEl) {{
                if (count > 0) {{
                    var avgDays = 0;
                    var highPriority = 0;
                    
                    if (selectedRows && selectedRows.length > 0) {{
                        var totalDays = 0;
                        for (var i = 0; i < selectedRows.length; i++) {{
                            var row = selectedRows[i];
                            var daysAgo = row.daysago || 0;
                            var stdDevs = row.stdevsago || 0;
                            
                            if (typeof daysAgo === 'string') daysAgo = parseInt(daysAgo) || 0;
                            if (typeof stdDevs === 'string') stdDevs = parseFloat(stdDevs) || 0;
                            
                            totalDays += daysAgo;
                            if (stdDevs >= 3) {{
                                highPriority++;
                            }}
                        }}
                        avgDays = Math.round(totalDays / count);
                    }}
                    
                    summaryEl.innerHTML = 
                        '<strong>' + count + '</strong> people selected<br>' +
                        '<small>Avg: ' + avgDays + ' days ago, ' + highPriority + ' high priority</small>';
                }} else {{
                    summaryEl.innerHTML = '<span id="selected-count">0</span> people selected';
                }}
            }}
        }}
        
        // Grid configuration
        var gridOptions = {{
            rowData: [
                {0}
            ],
            
            columnDefs: [
                {{ 
                    headerName: '✅', 
                    field: 'selected',
                    checkboxSelection: true,
                    headerCheckboxSelection: true,
                    width: 50,
                    pinned: 'left',
                    sortable: false,
                    filter: false
                }},
                {{ 
                    field: "name", 
                    headerName: "Name",
                    cellRenderer: function(params) {{
                        if(params.data){{
                            return '<a href="/Person2/' + params.data.peopleid + '" target="_blank" style="text-decoration: none;">' + 
                                   params.value + '</a>';
                        }}
                        return '';
                    }}, 
                    width: 220,
                    pinned: 'left'
                }},
                {{ 
                    headerName: 'Photo', 
                    field: "pic", 
                    cellRenderer: function(params) {{
                        if(params.data && params.value && params.value.trim() !== '' && params.value !== 'null'){{
                            return '<img src="' + params.value + '" alt="" style="width: 40px; height: 40px; border-radius: 50%; object-fit: cover;">';
                        }}
                        return '<div style="width: 40px; height: 40px; background: #e2e8f0; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 18px;">👤</div>';
                    }}, 
                    width: 80,
                    sortable: false,
                    filter: false
                }},
                {{ headerName: 'Age', field: "age", width: 50, cellDataType: 'number', comparator: numberSort, filter: 'agNumberColumnFilter' }},
                {{ 
                    headerName: 'Priority', 
                    field: "stdevsago", 
                    cellRenderer: function(params) {{
                        if(params.data){{
                            var priority = getPriorityLevel(params.value);
                            return '<span style="padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold; ' + priority.style + '">' + priority.text + '</span>';
                        }}
                        return '';
                    }},
                    width: 100,
                    cellDataType: 'number',
                    filter: 'agNumberColumnFilter'
                }},
                {{ headerName: 'Days Ago', field: "daysago", width: 90, cellDataType: 'number', filter: 'agNumberColumnFilter' }},
                {{ headerName: 'Attendance Count', field: "attendancecount", width: 130, cellDataType: 'number', filter: 'agNumberColumnFilter' }},
                {{ headerName: 'Avg Gap', field: "avgdays", hide: true, width: 90, cellDataType: 'number', filter: 'agNumberColumnFilter', 
                  valueFormatter: function(params) {{ return Math.round(params.value) + ' days'; }} }},
                {{ headerName: 'Contact Status', field: "contactstatus", width: 130,
                    cellRenderer: function(params) {{
                        if(params.data){{
                            return getContactStatusBadge(params.value);
                        }}
                        return '';
                    }}
                }},
                {{ headerName: 'Last Contact', field: "contactdays", width: 110, 
                  valueFormatter: function(params) {{ return params.value >= 999 ? 'Never' : params.value + ' days ago'; }} }},
                {{ headerName: 'Involvements', field: "orgnames", width: 200 }}
            ],
            
            defaultColDef: {{
                flex: 1,
                minWidth: 100,
                filter: 'agMultiColumnFilter',
                sortable: true,
                resizable: true
            }},
            
            rowSelection: 'multiple',
            suppressRowClickSelection: true,
            rowMultiSelectWithClick: false,
            
            onGridReady: function(params) {{
                gridApi = params.api;
                updateSelectionDisplay();
            }},
            
            onSelectionChanged: function(event) {{
                try {{
                    selectedRows = event.api.getSelectedRows();
                    updateSelectionDisplay();
                }} catch (e) {{
                    console.error('Error in onSelectionChanged:', e);
                }}
            }},
            
            sideBar: {{
                toolPanels: [
                    {{
                        id: 'columns',
                        labelDefault: 'Columns',
                        iconKey: 'columns',
                        toolPanel: 'agColumnsToolPanel'
                    }},
                    {{
                        id: 'filters',
                        labelDefault: 'Filters',
                        iconKey: 'filter',
                        toolPanel: 'agFiltersToolPanel'
                    }}
                ],
                defaultToolPanel: ''
            }}
        }};

        // Initialize grid
        try {{
            var gridElement = document.querySelector('#attendance-grid');
            if (gridElement && typeof agGrid !== 'undefined') {{
                var gridInstance = agGrid.createGrid(gridElement, gridOptions);
                
                if (!gridApi && gridInstance && gridInstance.api) {{
                    gridApi = gridInstance.api;
                }}
            }} else {{
                if (gridElement) {{
                    gridElement.innerHTML = '<div style="padding: 20px; text-align: center; border: 2px dashed #ccc; border-radius: 8px;">Grid loading failed. Please refresh the page to try again.</div>';
                }}
            }}
        }} catch (gridError) {{
            var gridElement = document.getElementById('attendance-grid');
            if (gridElement) {{
                gridElement.innerHTML = '<div style="padding: 20px; text-align: center;">Error loading data grid. Please refresh the page.</div>';
            }}
        }}
        </script>
        '''.format(records)
        
    except Exception as e:
        print_error("Error rendering ag-Grid: " + str(e))

def render_modals():
    """Render simplified action modals"""
    # Load keywords from TouchPoint
    try:
        keywords_sql = """
        SELECT TOP 20 KeywordId, Code, Description 
        FROM Keyword
        WHERE Description IS NOT NULL
            AND Description <> ''
        ORDER BY Description
        """
        keywords = q.QuerySql(keywords_sql)
    except:
        keywords = []
    
    print '''
    <!-- Tag Modal -->
    <div id="tag-modal" class="modal">
        <div class="modal-content">
            <h3>🏷️ Tag Selected People</h3>
            <form id="tag-form" method="POST">
                <input type="hidden" name="form_action" value="bulk_action">
                <input type="hidden" name="action_type" value="tag">
                <input type="hidden" id="tag_people_ids" name="selected_people_ids" value="">
                
                <div class="form-group">
                    <label>Tag Name:</label>
                    <input type="text" name="tag_name" class="form-control" value="{0}">
                </div>
                
                <div style="text-align: right; margin-top: 20px;">
                    <button type="button" onclick="closeModal('tag-modal')" class="btn btn-secondary">Cancel</button>
                    <button type="submit" class="btn btn-primary">Apply Tag</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- Task Modal -->
    <div id="task-modal" class="modal">
        <div class="modal-content">
            <h3>📋 Create Tasks for Selected People</h3>
            <form id="task-form" method="POST">
                <input type="hidden" name="form_action" value="bulk_action">
                <input type="hidden" name="action_type" value="task">
                <input type="hidden" id="task_people_ids" name="selected_people_ids" value="">
                
                <div class="form-group">
                    <label for="assignee_search">Assign To (Search by name):</label>
                    <div class="search-input-wrapper">
                        <input type="text" id="assignee_search" name="assignee_search" class="form-control" placeholder="Search for a person to assign...">
                        <div id="assignee_search_results" class="search-results"></div>
                        <input type="hidden" id="assignee_id" name="assignee_id" value="">
                    </div>
                    <div id="selected_assignee" style="margin-top: 5px; font-style: italic; color: #666;"></div>
                    <small style="color: #888;">Leave blank to assign to yourself</small>
                </div>
                
                <div class="form-group">
                    <label for="task_due_date">Due Date (Optional):</label>
                    <input type="date" id="task_due_date" name="due_date" class="form-control">
                </div>
                
                <div class="form-group">
                    <label>Keywords:</label>
                    <div class="keyword-container">
    '''.format(DEFAULT_TAG_NAME)
    
    # Add keywords to task modal
    if keywords:
        for i, keyword in enumerate(keywords):
            print '''
                        <div class="keyword-item">
                            <label>
                                <input type="checkbox" name="keyword_{0}" value="{1}"> 
                                {2} {3}
                            </label>
                        </div>
            '''.format(i, keyword.KeywordId, keyword.Description, 
                      '(' + keyword.Code + ')' if keyword.Code else '')
    else:
        print '<div class="keyword-item">No keywords available</div>'
    
    print '''
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="task_message">Task Message:</label>
                    <textarea id="task_message" name="task_message" class="form-control" rows="4">Follow up on lapsed attendance pattern</textarea>
                </div>
                
                <div style="text-align: right; margin-top: 20px;">
                    <button type="button" onclick="closeModal('task-modal')" class="btn btn-secondary">Cancel</button>
                    <button type="submit" class="btn btn-success">Create Tasks</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- Note Modal -->
    <div id="note-modal" class="modal">
        <div class="modal-content">
            <h3>📝 Add Notes to Selected People</h3>
            <form id="note-form" method="POST">
                <input type="hidden" name="form_action" value="bulk_action">
                <input type="hidden" name="action_type" value="note">
                <input type="hidden" id="note_people_ids" name="selected_people_ids" value="">
                
                <div class="form-group">
                    <label>Keywords:</label>
                    <div class="keyword-container">
    '''
    
    # Add keywords to note modal
    if keywords:
        for i, keyword in enumerate(keywords):
            print '''
                        <div class="keyword-item">
                            <label>
                                <input type="checkbox" name="keyword_{0}" value="{1}"> 
                                {2} {3}
                            </label>
                        </div>
            '''.format(i, keyword.KeywordId, keyword.Description, 
                      '(' + keyword.Code + ')' if keyword.Code else '')
    else:
        print '<div class="keyword-item">No keywords available</div>'
    
    print '''
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="note_message">Note Message:</label>
                    <textarea id="note_message" name="note_message" class="form-control" rows="4">Lapsed attendance pattern identified through statistical analysis</textarea>
                </div>
                
                <div style="text-align: right; margin-top: 20px;">
                    <button type="button" onclick="closeModal('note-modal')" class="btn btn-secondary">Cancel</button>
                    <button type="submit" class="btn btn-warning">Add Notes</button>
                </div>
            </form>
        </div>
    </div>
    '''

def render_javascript():
    """Render JavaScript functionality"""
    print '''
    <script>
        // Global variables for selection tracking
        var selectedRows = [];
        var gridApi = null;
        
        // Helper function to get the correct form submission URL
        function getPyScriptAddress() {
            var path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        }
        
        function filterData() {
            var ageFilter = document.getElementById('age_filter').value;
            window.location.href = window.location.pathname + '?age_filter=' + ageFilter;
        }
        
        function showBulkActionDialog() {
            if (selectedRows.length === 0) {
                alert('Please select at least one person first.');
                return;
            }
            
            var actionType = document.getElementById('action_type').value;
            var peopleIds = [];
            
            // Extract people IDs from selected rows
            for (var i = 0; i < selectedRows.length; i++) {
                var row = selectedRows[i];
                var peopleId = row.peopleid || row.PeopleId;
                if (peopleId) {
                    peopleIds.push(peopleId);
                }
            }
            
            if (peopleIds.length === 0) {
                alert('No valid people selected.');
                return;
            }
            
            var peopleIdsStr = peopleIds.join(',');
            
            if (actionType === 'tag') {
                document.getElementById('tag_people_ids').value = peopleIdsStr;
                document.getElementById('tag-modal').style.display = 'block';
            } else if (actionType === 'task') {
                document.getElementById('task_people_ids').value = peopleIdsStr;
                // Clear form fields
                document.getElementById('task_due_date').value = '';
                document.getElementById('assignee_search').value = '';
                document.getElementById('assignee_id').value = '';
                document.getElementById('selected_assignee').textContent = '';
                document.getElementById('task_message').value = 'Follow up on lapsed attendance pattern';
                
                // Clear keyword checkboxes
                var taskKeywords = document.querySelectorAll('#task-modal input[type="checkbox"]');
                taskKeywords.forEach(function(checkbox) {
                    checkbox.checked = false;
                });
                
                document.getElementById('task-modal').style.display = 'block';
            } else if (actionType === 'note') {
                document.getElementById('note_people_ids').value = peopleIdsStr;
                document.getElementById('note_message').value = 'Lapsed attendance pattern identified through statistical analysis';
                
                // Clear keyword checkboxes
                var noteKeywords = document.querySelectorAll('#note-modal input[type="checkbox"]');
                noteKeywords.forEach(function(checkbox) {
                    checkbox.checked = false;
                });
                
                document.getElementById('note-modal').style.display = 'block';
            }
        }
        
        // Person search functionality for task assignment
        function searchPeople(term) {
            if (term.length < 2) {
                document.getElementById('assignee_search_results').innerHTML = '';
                document.getElementById('assignee_search_results').style.display = 'none';
                return;
            }
            
            var resultsContainer = document.getElementById('assignee_search_results');
            resultsContainer.innerHTML = '<div style="padding: 10px; text-align: center;">🔍 Searching...</div>';
            resultsContainer.style.display = 'block';
            
            // Use the same URL pattern as working example
            var xhr = new XMLHttpRequest();
            xhr.open('GET', window.location.pathname + '?q=' + encodeURIComponent(term) + '&ajax=1&people_search=1', true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            var peopleTally = 0;
                            var html = '';
                            
                            // Extract people from the HTML response (exactly like working example)
                            var tempContainer = document.createElement('div');
                            tempContainer.innerHTML = xhr.responseText;
                            
                            // Find all result items that look like people
                            var peopleItems = tempContainer.querySelectorAll('.result-item');
                            peopleItems.forEach(function(item) {
                                var nameLink = item.querySelector('.result-name a');
                                if (nameLink) {
                                    // Extract the person ID from the href
                                    var href = nameLink.getAttribute('href');
                                    var idMatch = href.match(/\/Person2\/(\d+)/);
                                    
                                    if (idMatch && idMatch[1]) {
                                        var personId = idMatch[1];
                                        var personName = nameLink.textContent.trim();
                                        
                                        // Extract age if present
                                        var age = "";
                                        var metaSpans = item.querySelectorAll('.result-meta span');
                                        for (var i = 0; i < metaSpans.length; i++) {
                                            if (metaSpans[i].textContent.indexOf('Age:') !== -1) {
                                                age = metaSpans[i].textContent.replace("Age:", "").trim();
                                                break;
                                            }
                                        }
                                        
                                        // Extract address if present
                                        var address = "";
                                        var addressDivs = item.querySelectorAll('.result-meta div');
                                        for (var j = 0; j < addressDivs.length; j++) {
                                            if (addressDivs[j].textContent.indexOf('address') !== -1 || 
                                                addressDivs[j].querySelector('.fa-home')) {
                                                address = addressDivs[j].textContent.trim();
                                                break;
                                            }
                                        }
                                        
                                        // Add this person to our results
                                        html += '<div class="search-result-item" data-id="' + personId + '" data-name="' + personName + '">';
                                        html += personName + (age ? ' (' + age + ')' : '');
                                        if (address) {
                                            html += '<br><small>' + address + '</small>';
                                        }
                                        html += '</div>';
                                        peopleTally++;
                                    }
                                }
                            });
                            
                            if (peopleTally === 0) {
                                html = '<div class="search-result-item">No results found</div>';
                            }
                            
                            resultsContainer.innerHTML = html;
                            
                            // Add click handlers to results
                            var resultItems = document.querySelectorAll('.search-result-item');
                            for (var j = 0; j < resultItems.length; j++) {
                                resultItems[j].addEventListener('click', function() {
                                    var id = this.getAttribute('data-id');
                                    var name = this.getAttribute('data-name');
                                    
                                    // Set the selected person
                                    document.getElementById('assignee_id').value = id;
                                    document.getElementById('selected_assignee').textContent = 'Selected: ' + name;
                                    
                                    // Clear search and hide results
                                    document.getElementById('assignee_search').value = '';
                                    document.getElementById('assignee_search_results').style.display = 'none';
                                });
                            }
                        } catch (e) {
                            console.error('Error processing people search results:', e);
                            resultsContainer.innerHTML = '<div class="search-result-item">Error processing results: ' + e.message + '</div>';
                        }
                    } else {
                        resultsContainer.innerHTML = '<div class="search-result-item">Error loading results (Status: ' + xhr.status + ')</div>';
                    }
                }
            };
            xhr.send();
        }
        
        function closeModal(modalId) {
            document.getElementById(modalId).style.display = 'none';
        }
        
        // Initialize on page load
        document.addEventListener('DOMContentLoaded', function() {
            // Set form action URLs to use PyScriptForm
            var forms = ['tag-form', 'task-form', 'note-form'];
            var correctUrl = getPyScriptAddress();
            
            forms.forEach(function(formId) {
                var form = document.getElementById(formId);
                if (form) {
                    form.action = correctUrl;
                }
            });
            
            // Set current age filter
            var urlParams = new URLSearchParams(window.location.search);
            var ageFilter = urlParams.get('age_filter') || 'all';
            var ageFilterEl = document.getElementById('age_filter');
            if (ageFilterEl) {
                ageFilterEl.value = ageFilter;
            }
            
            // Close modals when clicking outside
            window.addEventListener('click', function(event) {
                var modals = ['tag-modal', 'task-modal', 'note-modal'];
                modals.forEach(function(modalId) {
                    var modal = document.getElementById(modalId);
                    if (event.target === modal) {
                        closeModal(modalId);
                    }
                });
            });
            
            // Setup person search for task assignment
            var assigneeSearch = document.getElementById('assignee_search');
            if (assigneeSearch) {
                var searchTimeout = null;
                assigneeSearch.addEventListener('input', function() {
                    clearTimeout(searchTimeout);
                    var term = this.value.trim();
                    searchTimeout = setTimeout(function() {
                        searchPeople(term);
                    }, 300);
                });
            }
            
            // Close search results when clicking outside
            document.addEventListener('click', function(e) {
                if (!e.target.closest('.search-input-wrapper')) {
                    var searchResults = document.getElementById('assignee_search_results');
                    if (searchResults) {
                        searchResults.style.display = 'none';
                    }
                }
            });
            
            // Hide loading indicator
            var loading = document.getElementById('loading');
            if (loading) {
                loading.style.display = 'none';
            }
        });
    </script>
    '''

def display_main_dashboard():
    """Display the main dashboard"""
    try:
        # ::STEP:: Get Current Filter
        age_filter = get_form_value('age_filter', 'all')
        
        # ::STEP:: Load Data
        print_loading()
        data = get_attendance_data(age_filter)
        
        # ::STEP:: Calculate Summary
        summary = get_dashboard_summary(data)
        
        # ::STEP:: Render Dashboard Components
        render_header(summary, age_filter)
        render_controls()
        render_data_grid(data)
        render_modals()
        render_javascript()
        
    except Exception as e:
        print_error("Error in display_main_dashboard: " + str(e))
        import traceback
        print "<pre style='background: #f8f8f8; padding: 10px; border-radius: 4px; font-size: 12px;'>"
        traceback.print_exc()
        print "</pre>"

# =============================================================================
# ::END:: Display Functions
# =============================================================================

# =============================================================================
# ::START:: Main Controller
# =============================================================================

try:
    # ::START:: Main Controller  
    model.Header = "Lapsed Attendance Dashboard"
    
    # ::STEP:: Get Parameters Using Working Method (from Example Live Search)
    search_term = ""
    ajax_mode = False
    people_search_mode = False
    
    try:
        if hasattr(model.Data, "q"):
            search_term = str(model.Data.q)
        if hasattr(model.Data, "ajax") and model.Data.ajax == "1":
            ajax_mode = True
        if hasattr(model.Data, "people_search") and model.Data.people_search == "1":
            people_search_mode = True
    except Exception as e:
        # Continue with defaults
        pass
    
    # ::STEP:: Handle AJAX People Search
    if ajax_mode and search_term:
        # Check if this is a special people search for assignee selection
        if hasattr(model.Data, "people_search") and model.Data.people_search == "1":
            # People search for assignee
            try:
                people_search_sql = """
                SELECT TOP 15
                    p.PeopleId, p.Name, p.Age, 
                    p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip
                FROM People p
                WHERE p.Name LIKE '%{0}%'
                    AND p.DeceasedDate IS NULL
                    AND p.ArchivedFlag = 0
                ORDER BY 
                    CASE WHEN p.Name LIKE '{0}%' THEN 0 ELSE 1 END,
                    p.Name
                """.format(search_term.replace("'", "''"))
                
                people = q.QuerySql(people_search_sql)
                print "<div class='results-section'>"
                print "<h3 class='section-heading'>People</h3>"
                
                if people and len(people) > 0:
                    for person in people:
                        # Get formatted address
                        address = getattr(person, 'PrimaryAddress', '') or ''
                        
                        print """
                        <div class="result-item">
                            <div class="result-name">
                                <a href="/Person2/{0}" target="_blank">{1}</a>
                            </div>
                            <div class="result-meta">
                                <span>Age: {2}</span>
                                <div><i class="fa fa-home"></i> {3}</div>
                            </div>
                        </div>
                        """.format(person.PeopleId, person.Name, person.Age or "", address)
                else:
                    print "<div class='no-results'>No people found matching your search.</div>"
                
                print "</div>"
            except Exception as search_error:
                print "<div class='no-results'>Search error: {0}</div>".format(str(search_error))
        else:
            # Display main dashboard for regular AJAX requests
            display_main_dashboard()
        
    # ::STEP:: Handle Form Submission (only for non-AJAX requests)  
    elif not ajax_mode:
        form_action = get_form_value('form_action', '')
        
        if form_action == "bulk_action":
            # ::STEP:: Process Bulk Actions
            if process_bulk_actions():
                # Success - redisplay dashboard
                display_main_dashboard()
            else:
                # Error occurred - redisplay dashboard (error already shown)
                display_main_dashboard()
        else:
            # ::STEP:: Display Main Dashboard
            display_main_dashboard()
    
    # ::END:: Main Controller

except Exception as e:
    # ::STEP:: Handle Critical Errors (only for non-AJAX requests)
    if not ajax_mode:
        import traceback
        print "<h2>Critical Error</h2>"
        print "<div style='background: #fee2e2; border: 1px solid #fecaca; padding: 20px; border-radius: 8px; margin: 20px 0;'>"
        print "<p><strong>An unexpected error occurred:</strong> " + escape_html(str(e)) + "</p>"
        print "<details style='margin-top: 15px;'>"
        print "<summary>Technical Details (click to expand)</summary>"
        print "<pre style='background: #f8f8f8; padding: 10px; border-radius: 4px; margin-top: 10px; font-size: 12px; overflow: auto;'>"
        traceback.print_exc()
        print "</pre>"
        print "</details>"
        print "<p style='margin-top: 15px;'><strong>Troubleshooting:</strong></p>"
        print "<ul>"
        print "<li>Check that all required database tables exist</li>"
        print "<li>Verify the excluded program IDs are correct for your system</li>"
        print "<li>Try reducing the MAX_RECORDS limit if you're processing large amounts of data</li>"
        print "<li>Contact your system administrator if the problem persists</li>"
        print "</ul>"
        print "</div>"

# =============================================================================
# ::END:: Main Controller
# =============================================================================
