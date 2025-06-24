#roles=Admin,Staff,MinistryLeader

"""
Involvement Dashboard Widget - Enhanced Debug Version
Shows all involvements in a specific program/division with attendance metrics
Version: 1.4 - Enhanced Schedule Debug
Last Updated: {{DateTime}}


--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python: InvolvementDashboard
4. Paste all this code
5. Test and optionally add to menu
6. Set role permissions in the Python script settings
--Upload Instructions End--

Written By: Ben Swaby
Email: bswaby@fbchtn.org
"""

# Configuration - Update these values for your specific program/division
PROGRAM_ID = 1128  # Replace with your Program ID
DIVISION_ID = 137  # Replace with your Division ID
DAYS_FOR_AVERAGE = 90  # Number of days to calculate average attendance
COMPACT_MODE = True  # Set to True for dashboard widget mode (minimal header)
DEBUG_ROLE = "Developer"  # Role required to see debug output - using Staff so you can definitely see it

# To find your Program and Division IDs, run this SQL query in TouchPoint:
# SELECT DISTINCT ProgId, Program, DivId, Division FROM OrganizationStructure ORDER BY Program, Division

def main():
    """Main entry point for the involvement dashboard"""
    
    # Check if debug mode should be enabled
    debug_mode = model.UserIsInRole(DEBUG_ROLE)
    
    try:
        if debug_mode:
            print "<p>DEBUG: Starting dashboard...</p>"
        
        if debug_mode:
            print "<p>DEBUG: User access OK</p>"
        
        # Get all involvements with their data
        involvements, program_name, division_name = get_involvement_data(debug_mode)
        
        if debug_mode:
            print "<p>DEBUG: Got " + str(len(involvements) if involvements else 0) + " involvements</p>"
        
        # Render the dashboard
        render_dashboard(involvements, program_name, division_name, debug_mode)
        
    except Exception as e:
        print "<div class='alert alert-danger'>"
        print "<h4>Error in main()</h4>"
        print "<p>" + str(e) + "</p>"
        print "</div>"

def check_user_access():
    """Check if current user has required role - handled by TouchPoint roles comment"""
    return True  # Access controlled by #roles comment at top

def get_involvement_data(debug_mode):
    """Get all involvements with attendance data - simplified and more robust"""
    
    program_name = "Program " + str(PROGRAM_ID)
    division_name = "Division " + str(DIVISION_ID)
    
    try:
        if debug_mode:
            print "<p>DEBUG: Getting program/division names...</p>"
        
        # Get program and division names
        names_sql = """
        SELECT DISTINCT Program, Division 
        FROM OrganizationStructure 
        WHERE ProgId = """ + str(PROGRAM_ID) + """ AND DivId = """ + str(DIVISION_ID) + """
        """
        
        names_result = q.QuerySqlTop1(names_sql)
        if names_result:
            if hasattr(names_result, 'Program') and names_result.Program:
                program_name = str(names_result.Program)
            if hasattr(names_result, 'Division') and names_result.Division:
                division_name = str(names_result.Division)
                
        if debug_mode:
            print "<p>DEBUG: Program: " + program_name + ", Division: " + division_name + "</p>"
            
    except Exception as e:
        if debug_mode:
            print "<p>DEBUG: Error getting names: " + str(e) + "</p>"
    
    # Simplified query - just get basic org data first
    try:
        if debug_mode:
            print "<p>DEBUG: Running main organizations query...</p>"
        
        # Ultra-basic organization query - minimal fields
        involvement_sql = """
        SELECT 
            o.OrganizationId,
            o.OrganizationName,
            o.Location,
            o.LeaderName,
            o.MemberCount
        FROM Organizations o
        JOIN Division d ON o.DivisionId = d.Id
        WHERE d.ProgId = """ + str(PROGRAM_ID) + """
          AND o.DivisionId = """ + str(DIVISION_ID) + """
          AND o.OrganizationStatusId = 30
        ORDER BY o.OrganizationName
        """
        
        involvements = q.QuerySql(involvement_sql)
        
        if debug_mode:
            print "<p>DEBUG: Final query returned " + str(len(involvements) if involvements else 0) + " involvements</p>"
        
        if not involvements:
            return [], program_name, division_name
        
        # Process each organization safely
        enhanced_involvements = []
        for i, inv in enumerate(involvements):
            try:
                if debug_mode and i == 0:
                    print "<p>DEBUG: Processing first org...</p>"
                
                # Create enhanced object safely
                enhanced_inv = model.DynamicData()
                
                # Safely copy properties
                enhanced_inv.OrganizationId = getattr(inv, 'OrganizationId', 0)
                enhanced_inv.OrganizationName = str(getattr(inv, 'OrganizationName', ''))
                
                # Handle location safely
                raw_location = getattr(inv, 'Location', None)
                if raw_location and str(raw_location).strip() and str(raw_location).strip().lower() != 'none':
                    enhanced_inv.Location = str(raw_location).strip()
                else:
                    enhanced_inv.Location = ""
                
                # Handle leader safely
                raw_leader = getattr(inv, 'LeaderName', None)
                if raw_leader and str(raw_leader).strip() and str(raw_leader).strip().lower() != 'none':
                    enhanced_inv.LeaderName = str(raw_leader).strip()
                else:
                    enhanced_inv.LeaderName = ""
                    
                enhanced_inv.MemberCount = getattr(inv, 'MemberCount', 0)
                
                if debug_mode and i == 0:
                    print "<p>DEBUG: Enhanced org - Name: '" + enhanced_inv.OrganizationName + "', Location: '" + enhanced_inv.Location + "', Leader: '" + enhanced_inv.LeaderName + "'</p>"
                
                # Get attendance data safely
                try:
                    org_id = enhanced_inv.OrganizationId
                    
                    # Get last meeting info
                    meeting_sql = '''
                    SELECT TOP 1 MeetingDate, NumPresent
                    FROM Meetings 
                    WHERE OrganizationId = @p1 
                      AND DidNotMeet = 0 
                      AND MeetingDate IS NOT NULL
                      AND MeetingDate <= GETDATE()
                    ORDER BY MeetingDate DESC
                    '''
                    last_meeting = q.QuerySqlTop1(meeting_sql, org_id)
                    
                    if last_meeting:
                        enhanced_inv.LastMeetingDate = last_meeting.MeetingDate
                        enhanced_inv.LastMeetingAttendance = getattr(last_meeting, 'NumPresent', None)
                        
                        # Calculate days since
                        try:
                            days_diff = model.DateDiffDays(last_meeting.MeetingDate, model.DateTime)
                            enhanced_inv.DaysSinceLastMeeting = int(days_diff) if days_diff is not None else None
                            enhanced_inv.MeetingStatus = 'PAST'
                        except:
                            enhanced_inv.DaysSinceLastMeeting = None
                            enhanced_inv.MeetingStatus = 'NO_MEETINGS'
                    else:
                        enhanced_inv.LastMeetingDate = None
                        enhanced_inv.LastMeetingAttendance = None
                        enhanced_inv.DaysSinceLastMeeting = None
                        enhanced_inv.MeetingStatus = 'NO_MEETINGS'
                    
                    # Get average attendance
                    avg_sql = """
                    SELECT 
                        COUNT(*) as MeetingCount,
                        AVG(CAST(NumPresent as FLOAT)) as AvgAttendance
                    FROM Meetings 
                    WHERE OrganizationId = @p1 
                      AND DidNotMeet = 0
                      AND NumPresent IS NOT NULL
                      AND NumPresent > 0
                      AND MeetingDate >= DATEADD(day, -""" + str(DAYS_FOR_AVERAGE) + """, GETDATE())
                      AND MeetingDate <= GETDATE()
                    """
                    avg_result = q.QuerySqlTop1(avg_sql, org_id)
                    
                    if avg_result and getattr(avg_result, 'MeetingCount', 0) > 0:
                        enhanced_inv.AvgAttendance = getattr(avg_result, 'AvgAttendance', None)
                        enhanced_inv.MeetingCount = getattr(avg_result, 'MeetingCount', 0)
                    else:
                        enhanced_inv.AvgAttendance = None
                        enhanced_inv.MeetingCount = 0
                    
                    # Enhanced schedule processing with debugging
                    enhanced_inv.Schedule = get_organization_schedule(org_id, debug_mode and i == 0)
                        
                except Exception as e:
                    if debug_mode:
                        print "<p>DEBUG: Error getting attendance data for org " + str(enhanced_inv.OrganizationId) + ": " + str(e) + "</p>"
                    # Set defaults for failed attendance queries
                    enhanced_inv.LastMeetingDate = None
                    enhanced_inv.LastMeetingAttendance = None
                    enhanced_inv.DaysSinceLastMeeting = None
                    enhanced_inv.MeetingStatus = 'NO_MEETINGS'
                    enhanced_inv.AvgAttendance = None
                    enhanced_inv.MeetingCount = 0
                    enhanced_inv.Schedule = 'See Organization'
                
                enhanced_involvements.append(enhanced_inv)
                
            except Exception as e:
                if debug_mode:
                    print "<p>DEBUG: Error processing org " + str(i) + ": " + str(e) + "</p>"
                continue
        
        if debug_mode:
            print "<p>DEBUG: Successfully processed " + str(len(enhanced_involvements)) + " involvements</p>"
        
        return enhanced_involvements, program_name, division_name
        
    except Exception as e:
        if debug_mode:
            print "<p>DEBUG: Major error in get_involvement_data: " + str(e) + "</p>"
        return [], program_name, division_name

def get_organization_schedule(org_id, debug_this_org=False):
    """Get organization schedule with enhanced debugging"""
    
    try:
        if debug_this_org:
            print "<p>DEBUG SCHEDULE: Starting schedule lookup for org " + str(org_id) + "</p>"
        
        # First, let's see what schedule tables exist and have data
        schedule_check_sql = """
        SELECT 
            'OrgSchedule' as TableName,
            COUNT(*) as RecordCount
        FROM OrgSchedule 
        WHERE OrganizationId = @p1
        
        UNION ALL
        
        SELECT 
            'Organizations' as TableName,
            COUNT(*) as RecordCount
        FROM Organizations 
        WHERE OrganizationId = @p1 
          AND (FirstMeetingDate IS NOT NULL OR LastMeetingDate IS NOT NULL)
        """
        
        schedule_check = q.QuerySql(schedule_check_sql, org_id)
        
        if debug_this_org and schedule_check:
            for check in schedule_check:
                print "<p>DEBUG SCHEDULE: " + str(check.TableName) + " has " + str(check.RecordCount) + " records</p>"
        
        # Try OrgSchedule table first - only select columns that exist
        schedule_sql = """
        SELECT SchedDay, SchedTime
        FROM OrgSchedule 
        WHERE OrganizationId = @p1
        ORDER BY SchedDay, SchedTime
        """
        
        schedules = q.QuerySql(schedule_sql, org_id)
        
        if debug_this_org:
            print "<p>DEBUG SCHEDULE: OrgSchedule query returned " + str(len(schedules) if schedules else 0) + " records</p>"
        
        if schedules and len(schedules) > 0:
            schedule_parts = []
            
            for sched in schedules:
                if debug_this_org:
                    print "<p>DEBUG SCHEDULE: Raw schedule - Day: " + str(getattr(sched, 'SchedDay', 'None')) + ", Time: " + str(getattr(sched, 'SchedTime', 'None')) + "</p>"
                
                # Process schedule
                day_name = format_schedule_day(getattr(sched, 'SchedDay', None))
                time_str = format_schedule_time(getattr(sched, 'SchedTime', None), debug_this_org)
                
                if day_name:
                    if time_str:
                        schedule_parts.append(day_name + ' ' + time_str)
                    else:
                        schedule_parts.append(day_name)
            
            if schedule_parts:
                result = ' & '.join(schedule_parts[:3])  # Allow up to 3 schedules for display
                if debug_this_org:
                    print "<p>DEBUG SCHEDULE: Final schedule result: '" + result + "'</p>"
                return result
        
        # Fallback: Try to get meeting pattern from recent meetings
        if debug_this_org:
            print "<p>DEBUG SCHEDULE: No OrgSchedule data, trying meeting pattern...</p>"
        
        meeting_pattern_sql = """
        SELECT TOP 5
            DATEPART(weekday, MeetingDate) as WeekDay,
            DATEPART(hour, MeetingDate) as Hour,
            COUNT(*) as Frequency
        FROM Meetings
        WHERE OrganizationId = @p1
          AND MeetingDate >= DATEADD(month, -6, GETDATE())
          AND DidNotMeet = 0
        GROUP BY DATEPART(weekday, MeetingDate), DATEPART(hour, MeetingDate)
        ORDER BY COUNT(*) DESC
        """
        
        meeting_patterns = q.QuerySql(meeting_pattern_sql, org_id)
        
        if debug_this_org:
            print "<p>DEBUG SCHEDULE: Meeting pattern query returned " + str(len(meeting_patterns) if meeting_patterns else 0) + " patterns</p>"
        
        if meeting_patterns and len(meeting_patterns) > 0:
            pattern = meeting_patterns[0]
            
            if debug_this_org:
                print "<p>DEBUG SCHEDULE: Most common pattern - WeekDay: " + str(getattr(pattern, 'WeekDay', 'None')) + ", Hour: " + str(getattr(pattern, 'Hour', 'None')) + ", Freq: " + str(getattr(pattern, 'Frequency', 'None')) + "</p>"
            
            # Convert SQL weekday (1=Sunday) to our day names
            weekday = getattr(pattern, 'WeekDay', None)
            hour = getattr(pattern, 'Hour', None)
            
            if weekday is not None:
                days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
                day_index = int(weekday) - 1  # SQL weekday is 1-based
                
                if 0 <= day_index < 7:
                    day_name = days[day_index]
                    
                    if hour is not None:
                        hour_int = int(hour)
                        if hour_int > 12:
                            time_display = str(hour_int - 12) + 'PM'
                        elif hour_int == 12:
                            time_display = '12PM'
                        elif hour_int == 0:
                            time_display = '12AM'
                        else:
                            time_display = str(hour_int) + 'AM'
                        
                        result = day_name + ' ' + time_display + ' (pattern)'
                        if debug_this_org:
                            print "<p>DEBUG SCHEDULE: Pattern-based result: '" + result + "'</p>"
                        return result
                    else:
                        result = day_name + ' (pattern)'
                        if debug_this_org:
                            print "<p>DEBUG SCHEDULE: Pattern-based result (no time): '" + result + "'</p>"
                        return result
        
        if debug_this_org:
            print "<p>DEBUG SCHEDULE: No schedule data found, returning default</p>"
        
        return 'See Organization'
        
    except Exception as e:
        if debug_this_org:
            print "<p>DEBUG SCHEDULE: Exception in schedule processing: " + str(e) + "</p>"
        return 'See Organization'

def format_schedule_day(day_value):
    """Format schedule day value"""
    if day_value is None:
        return None
    
    try:
        days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
        day_int = int(day_value)
        
        if 0 <= day_int < 7:
            return days[day_int]
        elif 1 <= day_int <= 7:  # In case it's 1-based
            return days[day_int - 1]
        else:
            return '?'
    except:
        return None

def format_schedule_time(time_value, debug_mode=False):
    """Format schedule time value with enhanced debugging"""
    if time_value is None:
        return None
    
    try:
        time_str = str(time_value)
        
        if debug_mode:
            print "<p>DEBUG TIME: Raw time value: '" + time_str + "' (type: " + str(type(time_value)) + ")</p>"
        
        # Handle datetime strings (like "3/26/2024 4:30:00 PM")
        if '/' in time_str and ' ' in time_str and (':' in time_str or 'AM' in time_str.upper() or 'PM' in time_str.upper()):
            if debug_mode:
                print "<p>DEBUG TIME: Detected datetime format</p>"
            
            # Split on space to separate date and time parts
            parts = time_str.split(' ')
            if len(parts) >= 2:
                # Find the time part (contains : or AM/PM)
                time_part = None
                ampm_part = None
                
                for part in parts:
                    if ':' in part:
                        time_part = part
                    elif part.upper() in ['AM', 'PM']:
                        ampm_part = part.upper()
                
                if debug_mode:
                    print "<p>DEBUG TIME: Extracted time_part: '" + str(time_part) + "', ampm_part: '" + str(ampm_part) + "'</p>"
                
                if time_part and ':' in time_part:
                    time_components = time_part.split(':')
                    if len(time_components) >= 2 and time_components[0].isdigit():
                        hour = int(time_components[0])
                        minute = int(time_components[1]) if time_components[1].isdigit() else 0
                        
                        # If we found AM/PM in the string, use it; otherwise determine from hour
                        if ampm_part:
                            # Hour is already in 12-hour format
                            if hour == 0:
                                hour_display = '12'
                            else:
                                hour_display = str(hour)
                            ampm = ampm_part
                        else:
                            # Convert from 24-hour to 12-hour format
                            if hour > 12:
                                hour_display = str(hour - 12)
                                ampm = 'PM'
                            elif hour == 12:
                                hour_display = '12'
                                ampm = 'PM'
                            elif hour == 0:
                                hour_display = '12'
                                ampm = 'AM'
                            else:
                                hour_display = str(hour)
                                ampm = 'AM'
                        
                        # Format the result
                        if minute > 0:
                            result = hour_display + ':' + str(minute).zfill(2) + ampm
                        else:
                            result = hour_display + ampm
                        
                        if debug_mode:
                            print "<p>DEBUG TIME: Datetime parsing result: '" + result + "'</p>"
                        
                        return result
        
        # Handle standard time formats (like "14:30:00" or "14:30")
        elif ':' in time_str and not '/' in time_str:
            if debug_mode:
                print "<p>DEBUG TIME: Detected standard time format</p>"
            
            time_parts = time_str.split(':')
            if len(time_parts) >= 2 and time_parts[0].isdigit():
                hour = int(time_parts[0])
                minute = int(time_parts[1]) if time_parts[1].isdigit() else 0
                
                # Format hour
                if hour > 12:
                    hour_display = str(hour - 12)
                    ampm = 'PM'
                elif hour == 12:
                    hour_display = '12'
                    ampm = 'PM'
                elif hour == 0:
                    hour_display = '12'
                    ampm = 'AM'
                else:
                    hour_display = str(hour)
                    ampm = 'AM'
                
                # Add minutes if not zero
                if minute > 0:
                    result = hour_display + ':' + str(minute).zfill(2) + ampm
                else:
                    result = hour_display + ampm
                
                if debug_mode:
                    print "<p>DEBUG TIME: Standard time result: '" + result + "'</p>"
                
                return result
        
        # Try to parse as decimal (like 14.5 for 2:30 PM)
        elif '.' in time_str and not '/' in time_str:
            try:
                if debug_mode:
                    print "<p>DEBUG TIME: Trying decimal format</p>"
                
                decimal_time = float(time_str)
                hour = int(decimal_time)
                minute = int((decimal_time - hour) * 60)
                
                if hour > 12:
                    hour_display = str(hour - 12)
                    ampm = 'PM'
                elif hour == 12:
                    hour_display = '12'
                    ampm = 'PM'
                elif hour == 0:
                    hour_display = '12'
                    ampm = 'AM'
                else:
                    hour_display = str(hour)
                    ampm = 'AM'
                
                if minute > 0:
                    result = hour_display + ':' + str(minute).zfill(2) + ampm
                else:
                    result = hour_display + ampm
                
                if debug_mode:
                    print "<p>DEBUG TIME: Decimal time result: '" + result + "'</p>"
                
                return result
            except:
                if debug_mode:
                    print "<p>DEBUG TIME: Decimal parsing failed</p>"
                pass
        
        # Try to parse as integer hour
        elif time_str.isdigit():
            if debug_mode:
                print "<p>DEBUG TIME: Trying integer hour format</p>"
            
            hour = int(time_str)
            
            if hour > 12:
                result = str(hour - 12) + 'PM'
            elif hour == 12:
                result = '12PM'
            elif hour == 0:
                result = '12AM'
            else:
                result = str(hour) + 'AM'
            
            if debug_mode:
                print "<p>DEBUG TIME: Integer hour result: '" + result + "'</p>"
            
            return result
        
        if debug_mode:
            print "<p>DEBUG TIME: Could not parse time format, returning None</p>"
        
        # Return None if we can't parse it (don't return the original long datetime string)
        return None
        
    except Exception as e:
        if debug_mode:
            print "<p>DEBUG TIME: Exception in time formatting: " + str(e) + "</p>"
        return None

def render_dashboard(involvements, program_name, division_name, debug_mode):
    """Render the dashboard HTML"""
    
    try:
        if debug_mode:
            print "<p>DEBUG: Starting render...</p>"
        
        # Calculate summary stats
        total_involvements = len(involvements) if involvements else 0
        total_members = 0
        total_avg_attendance = 0.0
        orgs_with_avg = 0
        need_update_count = 0
        
        if involvements:
            for inv in involvements:
                # Total members
                if hasattr(inv, 'MemberCount') and inv.MemberCount:
                    try:
                        total_members += int(inv.MemberCount)
                    except:
                        pass
                
                # Average attendance
                if hasattr(inv, 'AvgAttendance') and inv.AvgAttendance is not None:
                    try:
                        total_avg_attendance += float(inv.AvgAttendance)
                        orgs_with_avg += 1
                    except:
                        pass
                
                # Need update count - organizations with meetings older than 14 days
                if hasattr(inv, 'DaysSinceLastMeeting') and inv.DaysSinceLastMeeting is not None:
                    try:
                        days_since = int(inv.DaysSinceLastMeeting)
                        if days_since > 14:  # Changed from 7 to 14 days
                            need_update_count += 1
                    except:
                        pass
                elif hasattr(inv, 'MeetingStatus') and inv.MeetingStatus == 'NO_MEETINGS':
                    # Organizations with no meetings also need updates
                    need_update_count += 1
        
        overall_avg_attendance = total_avg_attendance / orgs_with_avg if orgs_with_avg > 0 else 0.0
        
        if debug_mode:
            print "<p>DEBUG: Stats - Total: " + str(total_involvements) + ", Need Updates: " + str(need_update_count) + ", Avg: " + str(overall_avg_attendance) + "</p>"
        
        # Process display properties
        if involvements:
            for inv in involvements:
                # Initialize display properties
                inv.RowClass = ""
                inv.StatusBadge = '<span class="badge badge-secondary">No Meetings</span>'
                inv.MeetingDateDisplay = '<em>No meetings</em>'
                inv.AvgAttendanceDisplay = "-"
                
                # Meeting date display
                if hasattr(inv, 'LastMeetingDate') and inv.LastMeetingDate:
                    try:
                        date_str = str(inv.LastMeetingDate)
                        if ' ' in date_str:
                            inv.MeetingDateDisplay = date_str.split(' ')[0]
                        else:
                            inv.MeetingDateDisplay = date_str
                        
                        # Add days ago
                        if hasattr(inv, 'DaysSinceLastMeeting') and inv.DaysSinceLastMeeting is not None:
                            days = int(inv.DaysSinceLastMeeting)
                            if days > 0:
                                inv.MeetingDateDisplay += ' <small>(' + str(days) + ' days ago)</small>'
                    except:
                        inv.MeetingDateDisplay = '<em>Date error</em>'
                
                # Status badge
                if hasattr(inv, 'DaysSinceLastMeeting') and inv.DaysSinceLastMeeting is not None:
                    try:
                        days = int(inv.DaysSinceLastMeeting)
                        if days > 30:
                            inv.StatusBadge = '<span class="badge badge-danger">Overdue</span>'
                            inv.RowClass = "stale-meeting"
                        elif days > 14:
                            inv.StatusBadge = '<span class="badge badge-warning">Needs Update</span>'
                            inv.RowClass = "past-meeting"
                        elif days > 7:
                            inv.StatusBadge = '<span class="badge badge-info">Recent</span>'
                        else:
                            inv.StatusBadge = '<span class="badge badge-success">Current</span>'
                    except:
                        pass
                elif hasattr(inv, 'MeetingStatus') and inv.MeetingStatus == 'NO_MEETINGS':
                    inv.StatusBadge = '<span class="badge badge-secondary">No Meetings</span>'
                
                # Average attendance display
                if hasattr(inv, 'AvgAttendance') and inv.AvgAttendance is not None:
                    try:
                        avg_val = float(inv.AvgAttendance)
                        inv.AvgAttendanceDisplay = "{:.1f}".format(avg_val)
                        if hasattr(inv, 'MeetingCount') and inv.MeetingCount:
                            inv.AvgAttendanceDisplay += ' <small>(' + str(inv.MeetingCount) + ' meetings)</small>'
                    except:
                        inv.AvgAttendanceDisplay = "-"
                else:
                    inv.AvgAttendanceDisplay = "-"
                
                # Ensure safe display values - less aggressive filtering
                if not hasattr(inv, 'Location'):
                    inv.Location = ""
                if not hasattr(inv, 'LeaderName'):
                    inv.LeaderName = ""
                if not hasattr(inv, 'LastMeetingAttendance'):
                    inv.LastMeetingAttendance = None
                if not hasattr(inv, 'Schedule'):
                    inv.Schedule = "See Organization"
                    
                if debug_mode and inv.OrganizationId == involvements[0].OrganizationId:
                    print "<p>DEBUG: Final values - Location: '" + str(inv.Location) + "', Leader: '" + str(inv.LeaderName) + "', Schedule: '" + str(inv.Schedule) + "'</p>"
        
        # Set up template data
        Data.involvements = involvements
        Data.program_name = program_name
        Data.division_name = division_name
        Data.current_date = model.DateTime
        Data.days_for_average = str(DAYS_FOR_AVERAGE)
        Data.total_involvements = str(total_involvements)
        Data.total_members = str(total_members)
        Data.need_update_count = str(need_update_count)
        Data.has_involvements = total_involvements > 0
        Data.overall_avg_attendance = "{:.1f}".format(overall_avg_attendance) if overall_avg_attendance > 0 else ""
        Data.debug_mode = debug_mode
        Data.compact_mode = COMPACT_MODE
        Data.PROGRAM_ID = str(PROGRAM_ID)
        Data.DIVISION_ID = str(DIVISION_ID)
        
        if debug_mode:
            print "<p>DEBUG: Rendering template with " + str(total_involvements) + " involvements</p>"
        
        # Template with improved responsive layout
        dashboard_template = '''
        <div class="involvement-dashboard {{#if compact_mode}}compact-mode{{/if}}">
            {{#unless compact_mode}}
            <div class="dashboard-header">
                <h2>Involvement Dashboard</h2>
                <p class="lead">{{program_name}} - {{division_name}}</p>
                <p class="text-muted">Average attendance calculated over last {{days_for_average}} days</p>
                {{#if debug_mode}}
                <div class="alert alert-info">
                    <strong>Debug Mode ON</strong> - Program ID: {{PROGRAM_ID}}, Division ID: {{DIVISION_ID}}
                </div>
                {{/if}}
            </div>
            {{else}}
            <div class="compact-header">
                <h4>{{program_name}} - {{division_name}}</h4>
            </div>
            {{/unless}}
            
            <div class="summary-stats">
                <div class="stat-card">
                    <h4>{{total_involvements}}</h4>
                    <p>Involvements</p>
                </div>
                <div class="stat-card">
                    <h4>{{total_members}}</h4>
                    <p>Members</p>
                </div>
                <div class="stat-card">
                    <h4>{{need_update_count}}</h4>
                    <p>Need Updates</p>
                </div>
                {{#if overall_avg_attendance}}
                <div class="stat-card">
                    <h4>{{overall_avg_attendance}}</h4>
                    <p>Avg Attendance</p>
                </div>
                {{/if}}
            </div>
            
            {{#if has_involvements}}
            <div class="table-container">
                <!-- Desktop Table View -->
                <table class="table table-hover involvement-table desktop-table">
                    <thead>
                        <tr>
                            <th style="width: 25%;">Involvement</th>
                            <th style="width: 15%;">Schedule</th>
                            <th style="width: 20%;">Last Meeting</th>
                            <th style="width: 10%;">Last Attend</th>
                            <th style="width: 15%;">Avg Attend</th>
                            <th style="width: 15%;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {{#each involvements}}
                        <tr class="{{RowClass}}">
                            <td class="involvement-name-cell">
                                <div class="org-name">
                                    <a href="/org/{{OrganizationId}}" target="_blank">
                                        <strong>{{OrganizationName}}</strong>
                                    </a>
                                </div>
                                {{#if LeaderName}}
                                <div class="leader-info">Leader: {{LeaderName}}</div>
                                {{/if}}
                                {{#if Location}}
                                <div class="location-info">📍 {{Location}}</div>
                                {{/if}}
                            </td>
                            <td class="schedule-cell">{{Schedule}}</td>
                            <td class="meeting-cell">{{{MeetingDateDisplay}}}</td>
                            <td class="attendance-cell">{{#if LastMeetingAttendance}}{{LastMeetingAttendance}}{{else}}-{{/if}}</td>
                            <td class="avg-cell">{{{AvgAttendanceDisplay}}}</td>
                            <td class="status-cell">{{{StatusBadge}}}</td>
                        </tr>
                        {{/each}}
                    </tbody>
                </table>
                
                <!-- Mobile Card View -->
                <div class="mobile-cards">
                    {{#each involvements}}
                    <div class="involvement-card {{RowClass}}">
                        <div class="card-header">
                            <div class="org-name">
                                <a href="/org/{{OrganizationId}}" target="_blank">
                                    <strong>{{OrganizationName}}</strong>
                                </a>
                            </div>
                            <div class="status-badge">{{{StatusBadge}}}</div>
                        </div>
                        
                        {{#if LeaderName}}
                        <div class="card-detail leader-info">Leader: {{LeaderName}}</div>
                        {{/if}}
                        
                        {{#if Location}}
                        <div class="card-detail location-info">📍 {{Location}}</div>
                        {{/if}}
                        
                        <div class="card-stats">
                            <div class="stat-item">
                                <span class="stat-label">Last:</span>
                                <span class="stat-value">{{#if LastMeetingAttendance}}{{LastMeetingAttendance}}{{else}}-{{/if}}</span>
                            </div>
                            <div class="stat-item">
                                <span class="stat-label">Avg:</span>
                                <span class="stat-value">{{{AvgAttendanceDisplay}}}</span>
                            </div>
                        </div>
                        
                        <div class="card-footer">
                            <span class="schedule-info">{{Schedule}}</span>
                            <span class="meeting-date">{{{MeetingDateDisplay}}}</span>
                        </div>
                    </div>
                    {{/each}}
                </div>
            </div>
            {{else}}
            <div class="alert alert-info">
                <h4>No Active Involvements Found</h4>
                <p>No active involvements found in {{program_name}} - {{division_name}}.</p>
                <p>Check that Program ID (''' + str(PROGRAM_ID) + ''') and Division ID (''' + str(DIVISION_ID) + ''') are correct.</p>
            </div>
            {{/if}}
            
            {{#unless compact_mode}}
            <div class="footer">
                <p class="text-muted">Generated: {{current_date}}</p>
            </div>
            {{/unless}}
        </div>
        
        <style>
        .involvement-dashboard { 
            padding: 20px; 
            max-width: 1400px; 
            margin: 0 auto; 
        }
        
        .involvement-dashboard.compact-mode {
            padding: 10px;
        }
        
        .dashboard-header { 
            margin-bottom: 30px; 
            border-bottom: 2px solid #e9ecef; 
            padding-bottom: 20px; 
        }
        
        .compact-header {
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .compact-header h4 {
            margin: 0;
            color: #495057;
        }
        
        .summary-stats { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); 
            gap: 15px; 
            margin-bottom: 25px; 
        }
        
        .involvement-dashboard.compact-mode .summary-stats {
            gap: 10px;
            margin-bottom: 15px;
        }
        
        .stat-card { 
            background: #f8f9fa; 
            padding: 15px; 
            border-radius: 8px; 
            text-align: center; 
            border: 1px solid #dee2e6; 
        }
        
        .involvement-dashboard.compact-mode .stat-card {
            padding: 10px;
        }
        
        .stat-card h4 { 
            margin: 0; 
            font-size: 1.8em; 
            color: #007bff; 
        }
        
        .involvement-dashboard.compact-mode .stat-card h4 {
            font-size: 1.5em;
        }
        
        .stat-card p { 
            margin: 5px 0 0 0; 
            color: #6c757d; 
            font-size: 13px;
        }
        
        .table-container {
            background: white;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        /* Desktop Table Styles */
        .desktop-table { 
            margin: 0;
            width: 100%;
            table-layout: fixed;
        }
        
        .desktop-table th { 
            background: #f8f9fa; 
            font-weight: 600; 
            border-bottom: 2px solid #dee2e6;
            padding: 12px;
            text-align: left;
        }
        
        .desktop-table td {
            padding: 12px;
            vertical-align: top;
            border-bottom: 1px solid #dee2e6;
        }
        
        .involvement-name-cell {
            word-wrap: break-word;
        }
        
        .org-name a {
            text-decoration: none;
            color: #007bff;
        }
        
        .org-name a:hover {
            text-decoration: underline;
        }
        
        .leader-info, .location-info {
            font-size: 12px;
            color: #6c757d;
            margin-top: 2px;
            line-height: 1.3;
        }
        
        .location-info {
            color: #28a745;
        }
        
        .schedule-cell {
            font-size: 13px;
            color: #495057;
            font-weight: 500;
        }
        
        .meeting-cell {
            font-size: 13px;
        }
        
        .status-cell {
            text-align: center;
        }
        
        /* Mobile Card Styles - Condensed */
        .mobile-cards {
            display: none;
        }
        
        .involvement-card {
            background: white;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            margin-bottom: 8px;
            padding: 10px;
        }
        
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 6px;
        }
        
        .card-header .org-name {
            flex: 1;
            margin-right: 8px;
        }
        
        .card-header .org-name a {
            color: #007bff;
            text-decoration: none;
            font-size: 15px;
            font-weight: 600;
            line-height: 1.1;
            display: block;
        }
        
        .status-badge {
            flex-shrink: 0;
        }
        
        .card-detail {
            font-size: 12px;
            color: #6c757d;
            margin-bottom: 2px;
            line-height: 1.2;
        }
        
        .card-detail.location-info {
            color: #28a745;
        }
        
        .card-stats {
            display: flex;
            gap: 12px;
            margin: 6px 0 4px 0;
            padding: 6px 0 4px 0;
            border-top: 1px solid #e9ecef;
        }
        
        .stat-item {
            display: flex;
            gap: 4px;
            align-items: center;
        }
        
        .stat-label {
            font-size: 12px;
            color: #6c757d;
            font-weight: 500;
        }
        
        .stat-value {
            font-size: 13px;
            color: #495057;
            font-weight: 600;
        }
        
        .card-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 4px;
            font-size: 11px;
            color: #6c757d;
            line-height: 1.2;
        }
        
        .schedule-info {
            flex: 1;
            font-weight: 600;
            color: #495057;
        }
        
        .meeting-date {
            text-align: right;
            flex-shrink: 0;
            margin-left: 8px;
        }
        
        /* Status styling */
        .past-meeting { 
            background-color: rgba(255, 193, 7, 0.1); 
        }
        
        .stale-meeting { 
            background-color: rgba(220, 53, 69, 0.1); 
        }
        
        .badge { 
            padding: 3px 6px; 
            font-size: 11px; 
            border-radius: 3px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            line-height: 1;
        }
        
        .badge-danger { background-color: #dc3545; color: white; }
        .badge-warning { background-color: #ffc107; color: #212529; }
        .badge-info { background-color: #17a2b8; color: white; }
        .badge-success { background-color: #28a745; color: white; }
        .badge-secondary { background-color: #6c757d; color: white; }
        
        .footer { 
            margin-top: 30px; 
            text-align: center; 
        }
        
        /* Responsive Design */
        @media (max-width: 992px) {
            .desktop-table {
                display: none;
            }
            
            .mobile-cards {
                display: block;
            }
            
            .involvement-dashboard {
                padding: 12px;
            }
            
            .summary-stats {
                grid-template-columns: repeat(4, 1fr);
                gap: 8px;
                margin-bottom: 15px;
            }
            
            .stat-card {
                padding: 12px;
            }
        }
        
        @media (max-width: 576px) {
            .involvement-dashboard {
                padding: 8px;
            }
            
            .summary-stats {
                grid-template-columns: repeat(2, 1fr);
                gap: 6px;
                margin-bottom: 10px;
            }
            
            .stat-card {
                padding: 8px;
            }
            
            .stat-card h4 {
                font-size: 1.3em;
            }
            
            .stat-card p {
                font-size: 11px;
            }
            
            .involvement-card {
                padding: 8px;
                margin-bottom: 6px;
            }
            
            .card-header {
                margin-bottom: 4px;
            }
            
            .card-header .org-name a {
                font-size: 14px;
            }
            
            .card-detail {
                font-size: 11px;
            }
            
            .stat-label {
                font-size: 11px;
            }
            
            .stat-value {
                font-size: 12px;
            }
            
            .card-stats {
                gap: 10px;
                margin: 4px 0 3px 0;
                padding: 4px 0;
            }
            
            .card-footer {
                margin-top: 3px;
                font-size: 10px;
            }
            
            .badge {
                padding: 2px 5px;
                font-size: 10px;
            }
        }
        </style>
        '''
        
        print model.RenderTemplate(dashboard_template)
        
        if debug_mode:
            print "<p>DEBUG: Template rendered successfully</p>"
        
    except Exception as e:
        print "<div class='alert alert-danger'>"
        print "<h4>Error in render_dashboard()</h4>"
        print "<p>" + str(e) + "</p>"
        print "</div>"

# Run the script with comprehensive error handling
try:
    main()
except Exception as e:
    import traceback
    print "<div class='alert alert-danger'>"
    print "<h4>Critical Error in Involvement Dashboard</h4>"
    print "<p>Error: " + str(e) + "</p>"
    if model.UserIsInRole(DEBUG_ROLE):
        print "<h5>Traceback:</h5>"
        print "<pre>"
        traceback.print_exc()
        print "</pre>"
    print "</div>"
