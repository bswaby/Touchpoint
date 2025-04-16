#####################################################################
#### REPORT INFORMATION
#####################################################################
# Weekly Attendance Dashboard
# This script creates an attendance dashboard that compares attendance 
# week-over-week and year-over-year with configurable options for fiscal or calendar year.
# The report breaks down attendance by Program, Division, and Organization.
# Also includes fiscal year-to-date comparison.

#written by: Ben Swaby
#email: bswaby@fbchtn.org
#last updated: 04/14/2025

#####################################################################
#### USER CONFIG FIELDS (Modify these settings as needed)
#####################################################################
# Report Title (displayed at top of page)
REPORT_TITLE = "Weekly Attendance"

# Year Type: "fiscal" or "calendar"
YEAR_TYPE = "fiscal"

# If YEAR_TYPE = "fiscal", set the first month and day of fiscal year
FISCAL_YEAR_START_MONTH = 10  # October
FISCAL_YEAR_START_DAY = 1     # 1st

# Number of years to display in comparison (current year + this many previous years)
YEARS_TO_DISPLAY = 2

# Include organization details when expanding divisions
SHOW_ORGANIZATION_DETAILS = True

# Include zero attendance rows for divisions
SHOW_ZERO_ATTENDANCE = True

# Default to collapsed organizations - set to TRUE
DEFAULT_COLLAPSED = True

# Programs to exclude from Average Attendance calculations
EXCLUDED_PROGRAMS =  ['VBS 2025 K-5', 'VBS 2025 Preschool','VBS 2025 Special Friends','VBS DAILY GRAND TOTAL']

# Email settings
EMAIL_FROM_NAME = "Attendance Reports"
EMAIL_FROM_ADDRESS = "attendance@church.org"
EMAIL_SUBJECT = "Weekly Attendance Report"

# Enable or disable performance debugging (set to True to show timing information)
DEBUG_PERFORMANCE = False

#####################################################################
#### START OF CODE - No configuration should be needed beyond this point
#####################################################################

import datetime
import re
import traceback
import time

# Add this at the very beginning of your script, immediately after your config section
# Before you define any classes or do any processing

model.Header = REPORT_TITLE


class PerformanceTimer:
    """Simple class to track and display execution time of code sections"""
    
    def __init__(self, enabled=False):
        """Initialize timer data"""
        self.enabled = enabled
        self.start_times = {}
        self.results = {}
    
    def start(self, section_name):
        """Start timing a section"""
        if not self.enabled:
            return
        self.start_times[section_name] = time.time()
    
    def end(self, section_name):
        """End timing a section and store result"""
        if not self.enabled:
            return 0
        if section_name in self.start_times:
            elapsed = time.time() - self.start_times[section_name]
            self.results[section_name] = elapsed
            return elapsed
        return 0
    
    def log(self, section_name, extra_info=""):
        """Log the timing result immediately"""
        if not self.enabled:
            return
        elapsed = self.end(section_name)
        print "<div class='timing-log' style='margin: 5px; padding: 5px; background-color: #f8f9fa; border-left: 3px solid #4CAF50; font-family: monospace;'>"
        print "TIMING: {} - {:.4f} seconds {}".format(section_name, elapsed, extra_info)
        print "</div>"
    
    def get_report(self):
        """Generate a formatted report of all timing results"""
        if not self.enabled or not self.results:
            return ""
            
        report = "<div class='timing-report' style='margin: 10px; padding: 10px; background-color: #f0f8ff; border: 1px solid #ccc;'>"
        report += "<h3>Performance Report</h3>"
        report += "<table style='width: 100%; border-collapse: collapse;'>"
        report += "<tr><th style='text-align: left; padding: 8px; border-bottom: 1px solid #ddd;'>Section</th>"
        report += "<th style='text-align: right; padding: 8px; border-bottom: 1px solid #ddd;'>Time (seconds)</th></tr>"
        
        # Sort results by time (descending)
        sorted_results = sorted(self.results.items(), key=lambda x: x[1], reverse=True)
        
        for section, elapsed in sorted_results:
            report += "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'>{}</td>".format(section)
            report += "<td style='text-align: right; padding: 8px; border-bottom: 1px solid #eee;'>{:.4f}</td></tr>".format(elapsed)
        
        report += "</table></div>"
        return report
    
    def print_report(self):
        """Print the timing report"""
        if not self.enabled:
            return
        print self.get_report()

# Create a global timer instance with the DEBUG_PERFORMANCE setting
performance_timer = PerformanceTimer(enabled=DEBUG_PERFORMANCE)

class ReportHelper:
    """Helper functions for the attendance dashboard"""
    
    @staticmethod
    def parse_program_rpt_group(rpt_group):
        """Parse the RptGroup column to extract report order and service times."""
        if not rpt_group or isinstance(rpt_group, type(None)) or str(rpt_group).strip() == "":
            return None, []
        
        # Extract the report group number and service times
        match = re.match(r'(\d+)\s*\[(.*?)\]', str(rpt_group))
        if match:
            report_order = match.group(1)
            service_times_text = match.group(2)
            
            # Initialize list to hold service times
            service_times = []
            
            # Split by commas first
            time_segments = [s.strip() for s in service_times_text.split(',')]
            
            for segment in time_segments:
                # Check if this is a combined time format like "(9:20 AM|9:30 AM)=9:20 AM"
                combined_match = re.match(r'\((.*?)\)=(.*?)$', segment)
                if combined_match:
                    # Get the resulting combined time (after the equals sign)
                    combined_result = combined_match.group(2).strip()
                    service_times.append(combined_result)
                else:
                    # Just a regular time
                    service_times.append(segment.strip())
            
            return report_order, service_times
        
        # If the format doesn't match but it's a number, just use it as the order
        # and return "Total" as the service time
        if str(rpt_group).strip().isdigit():
            return str(rpt_group).strip(), ["Total"]
        
        # If the format doesn't match at all, return just the report group as order
        # and "Total" as the service time
        return str(rpt_group).strip(), ["Total"]
    
    @staticmethod
    def get_fiscal_year_dates(year, month=FISCAL_YEAR_START_MONTH, day=FISCAL_YEAR_START_DAY):
        """Get start and end dates for a fiscal year."""
        start_date = datetime.datetime(year, month, day)
        end_date = datetime.datetime(year + 1, month, day) - datetime.timedelta(days=1)
        return start_date, end_date
    
    @staticmethod
    def get_calendar_year_dates(year):
        """Get start and end dates for a calendar year."""
        start_date = datetime.datetime(year, 1, 1)
        end_date = datetime.datetime(year, 12, 31)
        return start_date, end_date
    
    @staticmethod
    def get_year_dates(year, year_type=YEAR_TYPE):
        """Get start and end dates based on year type (fiscal or calendar)."""
        if year_type.lower() == "fiscal":
            return ReportHelper.get_fiscal_year_dates(year)
        else:
            return ReportHelper.get_calendar_year_dates(year)
    
    @staticmethod
    def format_date(date_obj):
        """Format date for SQL queries."""
        return date_obj.strftime('%Y-%m-%d')
    
    @staticmethod
    def format_display_date(date_obj):
        """Format date for display."""
        return date_obj.strftime('%m/%d/%Y')
    
    @staticmethod
    def get_current_year():
        """Get the current year based on year type."""
        today = datetime.datetime.now()
        
        if YEAR_TYPE.lower() == "fiscal":
            fiscal_start = datetime.datetime(today.year, FISCAL_YEAR_START_MONTH, FISCAL_YEAR_START_DAY)
            if today < fiscal_start:
                return today.year - 1
            return today.year
        
        return today.year
    
    @staticmethod
    def format_number(number):
        """Format number with commas and no decimal places."""
        if number is None:
            return "0"
        return "{:,}".format(int(number))
    
    @staticmethod
    def format_float(number, decimals=1):
        """Format float number with decimals."""
        if number is None:
            return "0.0"
        format_str = "{:,." + str(decimals) + "f}"
        return format_str.format(float(number))
    
    @staticmethod
    def get_trend_indicator(current, previous):
        """Generate a trend indicator based on comparison."""
        if previous == 0:
            return '<span style="color: green;">NEW</span>'
        
        percentage = ((current - previous) / float(previous)) * 100
        
        if percentage > 5:
            return '<span style="color: green;">&#9650; {0}%</span>'.format(ReportHelper.format_float(percentage))
        elif percentage < -5:
            return '<span style="color: red;">&#9660; {0}%</span>'.format(ReportHelper.format_float(abs(percentage)))
        else:
            return '<span style="color: #777;">&#8652; {0}%</span>'.format(ReportHelper.format_float(percentage))
    
    @staticmethod
    def get_year_label(year, year_type=YEAR_TYPE):
        """Get human-readable year label based on year type."""
        if year_type.lower() == "fiscal":
            return "FY<br>{}-{}".format(str(year)[-2:], str(year + 1)[-2:])
        return str(year)
    
    @staticmethod
    def get_previous_week_date(date_obj):
        """Get the same weekday from the previous week."""
        return date_obj - datetime.timedelta(days=7)
    
    @staticmethod
    def get_same_weekday_previous_year(date_obj):
        """Get the same weekday from the same week of the previous year."""
        # Go back 364 days (52 weeks) to get the same weekday
        previous_year_same_weekday = date_obj - datetime.timedelta(days=364)
        
        # Adjust if we're in a different fiscal year period
        if YEAR_TYPE.lower() == "fiscal":
            current_fiscal_year = ReportHelper.get_fiscal_year(date_obj)
            prev_fiscal_year = ReportHelper.get_fiscal_year(previous_year_same_weekday)
            
            if prev_fiscal_year != current_fiscal_year - 1:
                # Adjust by adding/subtracting weeks until we're in the correct fiscal year
                while ReportHelper.get_fiscal_year(previous_year_same_weekday) != current_fiscal_year - 1:
                    previous_year_same_weekday -= datetime.timedelta(days=7)
        
        return previous_year_same_weekday
    
    @staticmethod
    def get_fiscal_year(date_obj):
        """Get the fiscal year for a given date."""
        fiscal_year_start = datetime.datetime(date_obj.year, FISCAL_YEAR_START_MONTH, FISCAL_YEAR_START_DAY)
        
        if date_obj < fiscal_year_start:
            return date_obj.year - 1
        return date_obj.year
        
    @staticmethod
    def get_weeks_elapsed_in_fiscal_year(current_date, fiscal_start_date):
        """Calculate number of weeks elapsed in the fiscal year."""
        # Ensure we have datetime objects
        if isinstance(current_date, str):
            current_date = ReportHelper.parse_date_string(current_date)
        if isinstance(fiscal_start_date, str):
            fiscal_start_date = ReportHelper.parse_date_string(fiscal_start_date)
        
        # Calculate difference in days and convert to weeks
        days_elapsed = (current_date - fiscal_start_date).days
        weeks_elapsed = max(1, int(days_elapsed / 7))  # Ensure at least 1 week
        
        return weeks_elapsed

    @staticmethod
    def get_last_sunday():
        """Get the date of the most recent Sunday."""
        today = datetime.datetime.now()
        days_since_sunday = today.weekday() + 1  # Python's weekday() returns 0-6 (Mon-Sun), so add 1 for Sunday
        if days_since_sunday == 7:  # Today is Sunday
            return today
        
        # Go back to the previous Sunday
        last_sunday = today - datetime.timedelta(days=days_since_sunday)
        return last_sunday
    
    @staticmethod
    def parse_date_string(date_string):
        """Parse a date string in YYYY-MM-DD format."""
        try:
            return datetime.datetime.strptime(date_string, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
    
    @staticmethod
    def get_week_start_date(sunday_date):
        """Get the Monday before the given Sunday."""
        # Calculate days to go back to previous Monday (6 days)
        return sunday_date - datetime.timedelta(days=6)
    
    @staticmethod
    def is_sunday(date_obj):
        """Check if a date is a Sunday."""
        return date_obj.weekday() == 6  # Python's weekday() returns 0-6 (Mon-Sun)


class AttendanceReport:
    """Main class to generate the attendance dashboard"""
    
    def __init__(self):
        """Initialize the report with default values."""
        self.current_year = ReportHelper.get_current_year()
        self.year_labels = []
        self.years = []
        
        # Set should_load_data flag to False by default
        self.should_load_data = False
        
        # Use specified report_date or default to the last Sunday
        default_last_sunday = ReportHelper.get_last_sunday()
        
        # Get the report date (must be a Sunday)
        self.report_date = default_last_sunday
        if hasattr(model.Data, 'report_date') and model.Data.report_date:
            temp_date = ReportHelper.parse_date_string(model.Data.report_date)
            if temp_date and ReportHelper.is_sunday(temp_date):
                self.report_date = temp_date
                self.should_load_data = True  # Set to True when report_date is provided
            elif temp_date:
                # If not a Sunday, find the previous Sunday
                days_to_previous_sunday = temp_date.weekday() + 1
                if days_to_previous_sunday == 7:  # It's already Sunday
                    self.report_date = temp_date
                    self.should_load_data = True  # Set to True when report_date is provided
                else:
                    self.report_date = temp_date - datetime.timedelta(days=days_to_previous_sunday)
                    self.should_load_data = True  # Set to True when report_date is provided
        
        # Only continue with other date calculations if we should load data
        if self.should_load_data:
            # Get the week start date (Monday before the Sunday)
            self.week_start_date = ReportHelper.get_week_start_date(self.report_date)
            
            # Get previous week's dates (1 week ago)
            self.previous_week_date = ReportHelper.get_previous_week_date(self.report_date)
            self.previous_week_start = ReportHelper.get_week_start_date(self.previous_week_date)
            
            # Get previous year's same date for comparison
            self.previous_year_date = ReportHelper.get_same_weekday_previous_year(self.report_date)
            self.previous_week_start_date = ReportHelper.get_week_start_date(self.previous_year_date)
            
            # Get fiscal year start dates for YTD comparisons
            self.current_fiscal_year = ReportHelper.get_fiscal_year(self.report_date)
            self.current_fiscal_start, _ = ReportHelper.get_fiscal_year_dates(self.current_fiscal_year)
            self.previous_fiscal_year = self.current_fiscal_year - 1
            self.previous_fiscal_start, _ = ReportHelper.get_fiscal_year_dates(self.previous_fiscal_year)
            
            # Initialize years to display
            for i in range(YEARS_TO_DISPLAY):
                year = self.current_year - i
                self.years.append(year)
                self.year_labels.append(ReportHelper.get_year_label(year))
        
        # Whether to show organizations or collapse them (default to collapsed)
        # Use DEFAULT_COLLAPSED parameter if collapse_orgs not specified
        self.collapse_orgs = True  # Default to collapsed
        if hasattr(model.Data, 'collapse_orgs') and model.Data.collapse_orgs == 'no':
            self.collapse_orgs = False
    
    def get_programs_sql(self):
        """SQL to get all active programs with report groups."""
        return """
            SELECT 
                Id, 
                Name, 
                RptGroup,
                StartHoursOffset,
                EndHoursOffset
            FROM Program
            WHERE RptGroup IS NOT NULL AND RptGroup <> ''
            ORDER BY RptGroup
        """
    
    def get_divisions_sql(self, program_id):
        """SQL to get divisions for a specific program using OrganizationStructure."""
        return '''
            SELECT DISTINCT
                d.Id, 
                d.Name, 
                d.ProgId,
                d.ReportLine,
                d.NoDisplayZero
            FROM Division d
            INNER JOIN OrganizationStructure os ON os.DivId = d.Id
            INNER JOIN Program p ON p.Id = {program_id}
            WHERE d.ProgId = {program_id}
            AND d.ReportLine IS NOT NULL 
            AND d.ReportLine <> ''
            ORDER BY d.ReportLine
        '''.format(program_id=program_id)
    
    def get_organizations_sql(self, division_id):
        """SQL to get organizations for a specific division using OrganizationStructure."""
        return '''
            SELECT DISTINCT
                o.OrganizationId,
                o.OrganizationName,
                {division_id} AS DivisionId,
                o.MemberCount
            FROM Organizations o
            INNER JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
            INNER JOIN Division d ON d.Id = {division_id}
            INNER JOIN Program p ON p.Id = d.ProgId
            WHERE os.DivId = {division_id}
            AND o.OrganizationStatusId = 30
            AND p.RptGroup IS NOT NULL 
            AND p.RptGroup <> ''
            AND d.ReportLine IS NOT NULL 
            AND d.ReportLine <> ''
            ORDER BY o.OrganizationName
        '''.format(division_id=division_id)

    def get_program_specific_date_range(self, program, base_start_date, base_end_date):
        """Calculate program-specific date range using StartHoursOffset and EndHoursOffset."""
        # Default to the provided dates if no offsets are specified
        if program.StartHoursOffset is None and program.EndHoursOffset is None:
            return base_start_date, base_end_date
            
        # Find the Sunday of the week (for reference point)
        sunday_date = base_end_date  # This is already Sunday in your case
        
        # Apply offsets to Sunday at midnight
        sunday_midnight = datetime.datetime(
            sunday_date.year,
            sunday_date.month,
            sunday_date.day,
            0, 0, 0  # Midnight (00:00:00)
        )
        
        # Calculate program-specific start and end dates
        start_hours_offset = program.StartHoursOffset or 0
        end_hours_offset = program.EndHoursOffset or 24  # Default to end of day
        
        # Apply the offsets
        program_start_date = sunday_midnight + datetime.timedelta(hours=start_hours_offset)
        
        # Make sure end_hours_offset is treated as 23:59:59 if it's 24
        if end_hours_offset == 24:
            program_end_date = sunday_midnight.replace(hour=23, minute=59, second=59)
        else:
            program_end_date = sunday_midnight + datetime.timedelta(hours=end_hours_offset)
        
        # Ensure we don't extend beyond the week boundaries
        if program_start_date < base_start_date:
            program_start_date = base_start_date
        
        if program_end_date > base_end_date:
            program_end_date = base_end_date.replace(hour=23, minute=59, second=59)
        
        return program_start_date, program_end_date
    
    def get_week_attendance_sql(self, week_start_date, week_end_date, program_id=None, division_id=None, org_id=None):
        """Get SQL for attendance within a week range, accounting for program-specific offsets."""
        # Format dates for SQL
        start_date_str = ReportHelper.format_date(week_start_date)
        end_date_str = ReportHelper.format_date(week_end_date)
        
        # Add time components to make the dates include the full day
        start_date_time_str = start_date_str + " 00:00:00"
        end_date_time_str = end_date_str + " 23:59:59"
        
        # Base query structure with conditions for meetings
        sql = """
            WITH WeekAttendance AS (
                SELECT DISTINCT
                    m.MeetingId,
                    m.OrganizationId,
                    os.DivId,
                    d.ProgId,
                    p.Name as ProgramName,
                    CONVERT(date, m.MeetingDate) as MeetingDate,
                    DATEPART(HOUR, m.MeetingDate) as MeetingHour,
                    COALESCE(m.MaxCount, 0) as AttendCount
                FROM Meetings m
                JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                JOIN Division d ON os.DivId = d.Id
                JOIN Program p ON p.Id = os.ProgId
                WHERE CONVERT(datetime, m.MeetingDate) BETWEEN '{0}' AND '{1}'
                AND (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL) -- Only include meetings that actually happened
                AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''
                AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''
        """.format(start_date_time_str, end_date_time_str)
        
        # Add filters based on parameters
        if org_id:
            sql += " AND m.OrganizationId = {}".format(org_id)
        elif division_id:
            sql += " AND os.DivId = {}".format(division_id)
        elif program_id:
            sql += " AND d.ProgId = {}".format(program_id)
        
        # Close the CTE and perform aggregation
        sql += """
            )
            SELECT 
                ProgramName,
                MeetingHour,
                SUM(AttendCount) as AttendanceCount,
                COUNT(DISTINCT MeetingId) as MeetingCount
            FROM WeekAttendance
            GROUP BY ProgramName, MeetingHour
        """
        
        return sql
    
    def get_ytd_attendance_sql(self, start_date, end_date, program_id=None, division_id=None, org_id=None):
        """Get SQL for YTD attendance with improved aggregation, using OrganizationStructure."""
        # Format dates for SQL
        start_date_str = ReportHelper.format_date(start_date)
        end_date_str = ReportHelper.format_date(end_date)
        
        # Base query structure with conditions for meetings
        sql = """
            WITH YTDAttendance AS (
                SELECT DISTINCT
                    m.MeetingId,
                    m.OrganizationId,
                    os.DivId,
                    d.ProgId,
                    p.Name as ProgramName,
                    CONVERT(date, m.MeetingDate) as MeetingDate,
                    COALESCE(m.MaxCount, 0) as AttendCount
                FROM Meetings m
                JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                JOIN Division d ON os.DivId = d.Id
                JOIN Program p ON p.Id = os.ProgId
                WHERE CONVERT(date, m.MeetingDate) BETWEEN '{}' AND '{}'
                AND (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL) -- Only include meetings that actually happened
                AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''
                AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''
        """.format(start_date_str, end_date_str)
        
        # Add filters based on parameters
        if org_id:
            sql += " AND m.OrganizationId = {}".format(org_id)
        elif division_id:
            sql += " AND os.DivId = {}".format(division_id)
        elif program_id:
            sql += " AND d.ProgId = {}".format(program_id)
        
        # Close the CTE and perform aggregation
        sql += """
            )
            SELECT 
                ProgramName,
                SUM(AttendCount) as AttendanceCount,
                COUNT(DISTINCT MeetingId) as MeetingCount,
                COUNT(DISTINCT MeetingDate) as DaysWithMeetings
            FROM YTDAttendance
            GROUP BY ProgramName
        """
        
        return sql
    
    def get_specific_programs_sql(self, program_names=None):
        """SQL to get program IDs for specific program names."""
        if not program_names or len(program_names) == 0:
            return None
            
        names_list = ", ".join(["'{}'".format(name.replace("'", "''")) for name in program_names])
        return """
            SELECT 
                Id, 
                Name, 
                RptGroup
            FROM Program
            WHERE Name IN ({})
            AND RptGroup IS NOT NULL AND RptGroup <> ''
        """.format(names_list)
    
    def get_week_attendance_data(self, week_start_date, week_end_date, program_id=None, division_id=None, org_id=None):
        """Get attendance data for a specific week range and entity, with program-specific offsets."""
        
        performance_timer.start("get_week_attendance_data")
        
        attendance_data = {
            'total': 0,
            'meetings': 0,
            'by_hour': {},
            'by_program': {},
            'start_date': week_start_date,
            'end_date': week_end_date
        }
        
        # If this is a program-specific query, adjust dates based on program offsets
        if program_id:
            program_info = q.QuerySqlTop1("""
                SELECT 
                    StartHoursOffset, 
                    EndHoursOffset 
                FROM Program 
                WHERE Id = {}
            """.format(program_id))
            
            if program_info:
                # Create a class-like object from the SQL results if needed
                if not hasattr(program_info, 'StartHoursOffset'):
                    # Create a new object with the properties
                    class ProgramInfo:
                        pass
                    prog_obj = ProgramInfo()
                    prog_obj.StartHoursOffset = program_info[0] if len(program_info) > 0 else None
                    prog_obj.EndHoursOffset = program_info[1] if len(program_info) > 1 else None
                    program_info = prog_obj
                
                # Adjust dates based on program offsets
                week_start_date, week_end_date = self.get_program_specific_date_range(
                    program_info, week_start_date, week_end_date
                )
                attendance_data['start_date'] = week_start_date
                attendance_data['end_date'] = week_end_date
                
        # Now call the SQL with adjusted dates
        sql = self.get_week_attendance_sql(week_start_date, week_end_date, program_id, division_id, org_id)
        
        try:
            results = q.QuerySql(sql)
            #self.debug_sql(sql, "attendance")
            for row in results:
                attendance_data['total'] += row.AttendanceCount
                attendance_data['meetings'] += row.MeetingCount
                attendance_data['by_hour'][row.MeetingHour] = row.AttendanceCount
                
                # Track by program name too
                if row.ProgramName not in attendance_data['by_program']:
                    attendance_data['by_program'][row.ProgramName] = {
                        'total': 0, 
                        'meetings': 0
                    }
                
                attendance_data['by_program'][row.ProgramName]['total'] += row.AttendanceCount
                attendance_data['by_program'][row.ProgramName]['meetings'] += row.MeetingCount
        except Exception as e:
            print "<p>Error getting week attendance data: {}</p>".format(e)
            
        performance_timer.log("get_week_attendance_data", "params: {}".format(program_id or division_id or org_id or "all"))
        
        return attendance_data
        
    def get_ytd_attendance_data(self, start_date, end_date, program_id=None, division_id=None, org_id=None):
        """Get YTD attendance data between start and end dates, with program-specific offsets."""
        ytd_data = {
            'total': 0,
            'meetings': 0,
            'days': 0,
            'by_program': {},
            'start_date': start_date,
            'end_date': end_date
        }
        
        # If this is a program-specific query, adjust dates based on program offsets
        if program_id:
            program_info = q.QuerySqlTop1("""
                SELECT 
                    StartHoursOffset, 
                    EndHoursOffset 
                FROM Program 
                WHERE Id = {}
            """.format(program_id))
            
            if program_info:
                # Create a class-like object from the SQL results if needed
                if not hasattr(program_info, 'StartHoursOffset'):
                    # Create a new object with the properties
                    class ProgramInfo:
                        pass
                    prog_obj = ProgramInfo()
                    prog_obj.StartHoursOffset = program_info[0] if len(program_info) > 0 else None
                    prog_obj.EndHoursOffset = program_info[1] if len(program_info) > 1 else None
                    program_info = prog_obj
                
                # For YTD, we could adjust end date based on program offsets if needed
                _, adjusted_end_date = self.get_program_specific_date_range(
                    program_info, start_date, end_date
                )
                # Only update end_date for YTD calculation
                ytd_data['end_date'] = adjusted_end_date
                end_date = adjusted_end_date
                
        # Now call the SQL with adjusted dates
        sql = self.get_ytd_attendance_sql(start_date, end_date, program_id, division_id, org_id)
        
        try:
            results = q.QuerySql(sql)
            for row in results:
                ytd_data['total'] += row.AttendanceCount or 0
                ytd_data['meetings'] += row.MeetingCount or 0
                ytd_data['days'] += row.DaysWithMeetings or 0
                
                # Track by program
                program_name = row.ProgramName
                if program_name not in ytd_data['by_program']:
                    ytd_data['by_program'][program_name] = {
                        'total': 0,
                        'meetings': 0,
                        'days': 0
                    }
                
                ytd_data['by_program'][program_name]['total'] += row.AttendanceCount or 0
                ytd_data['by_program'][program_name]['meetings'] += row.MeetingCount or 0
                ytd_data['by_program'][program_name]['days'] += row.DaysWithMeetings or 0
            
            # Calculate average attendance if we have meetings
            if ytd_data['meetings'] > 0:
                ytd_data['avg_per_meeting'] = float(ytd_data['total']) / ytd_data['meetings']
            else:
                ytd_data['avg_per_meeting'] = 0
                
            # Calculate average per day if we have days with meetings
            if ytd_data['days'] > 0:
                ytd_data['avg_per_day'] = float(ytd_data['total']) / ytd_data['days']
            else:
                ytd_data['avg_per_day'] = 0
                
            # Calculate program-specific averages
            for program_name in ytd_data['by_program']:
                prog_data = ytd_data['by_program'][program_name]
                
                if prog_data['meetings'] > 0:
                    prog_data['avg_per_meeting'] = float(prog_data['total']) / prog_data['meetings']
                else:
                    prog_data['avg_per_meeting'] = 0
                    
                if prog_data['days'] > 0:
                    prog_data['avg_per_day'] = float(prog_data['total']) / prog_data['days']
                else:
                    prog_data['avg_per_day'] = 0
        except Exception as e:
            print "<p>Error getting YTD attendance data: {}</p>".format(e)
        
        return ytd_data
    
    def get_specific_program_attendance_data(self, program_names, start_date, end_date):
        """Get attendance data specifically for the programs listed in AVG_WEEK_PROGRAMS."""
        program_data = {
            'total': 0,
            'meetings': 0,
            'days': 0,
            'by_program': {},
            'avg_per_day': 0
        }
        
        if not program_names or len(program_names) == 0:
            return program_data
            
        # Get the programs by name
        sql = self.get_specific_programs_sql(program_names)
        if not sql:
            return program_data
            
        try:
            programs = q.QuerySql(sql)
            
            # For each program, get attendance data
            for program in programs:
                # Add to list of found programs
                program_data['by_program'][program.Name] = {
                    'id': program.Id,
                    'total': 0,
                    'meetings': 0,
                    'days': 0,
                    'avg_per_day': 0
                }
                
                # Get program attendance data
                prog_ytd = self.get_ytd_attendance_data(
                    start_date,
                    end_date,
                    program_id=program.Id
                )
                
                # Add to totals
                program_data['total'] += prog_ytd['total']
                program_data['meetings'] += prog_ytd['meetings']
                program_data['days'] += prog_ytd['days']
                
                # Store program-specific data
                prog_specific = program_data['by_program'][program.Name]
                prog_specific['total'] = prog_ytd['total']
                prog_specific['meetings'] = prog_ytd['meetings']
                prog_specific['days'] = prog_ytd['days']
                
                # Calculate averages
                if prog_specific['days'] > 0:
                    prog_specific['avg_per_day'] = float(prog_specific['total']) / prog_specific['days']
            
            # Calculate overall average
            if program_data['days'] > 0:
                program_data['avg_per_day'] = float(program_data['total']) / program_data['days']
                
        except Exception as e:
            print "<p>Error getting specific program data: {}</p>".format(e)
            
        return program_data
    
    def get_multiple_years_attendance_data(self, sunday_date, program_id=None, division_id=None, org_id=None):
        performance_timer.start("multiple_years_attendance_data_{}".format(org_id or division_id or program_id))
        
        years_data = {}
        
        # First get the current year's data
        current_year_data = self.get_week_attendance_data(
            self.week_start_date, 
            self.report_date, 
            program_id, 
            division_id, 
            org_id
        )
        years_data[self.current_year] = current_year_data
        
        # Get data for previous years
        for i in range(1, YEARS_TO_DISPLAY):
            prev_year = self.current_year - i
            prev_year_date = sunday_date - datetime.timedelta(days=364 * i)
            prev_year_week_start = ReportHelper.get_week_start_date(prev_year_date)
            
            prev_year_data = self.get_week_attendance_data(
                prev_year_week_start, 
                prev_year_date, 
                program_id, 
                division_id, 
                org_id
            )
            years_data[prev_year] = prev_year_data
        
        performance_timer.log("multiple_years_attendance_data_{}".format(org_id or division_id or program_id))
        return years_data
    
    def get_multiple_years_ytd_data(self, end_date, program_id=None, division_id=None, org_id=None):
        performance_timer.start("multiple_years_ytd_data_{}".format(org_id or division_id or program_id))
        
        years_ytd_data = {}
        
        # First get the current year's YTD data
        current_ytd_data = self.get_ytd_attendance_data(
            self.current_fiscal_start, 
            self.report_date, 
            program_id, 
            division_id, 
            org_id
        )
        years_ytd_data[self.current_year] = current_ytd_data
        
        # Get data for previous years
        for i in range(1, YEARS_TO_DISPLAY):
            prev_year = self.current_year - i
            prev_year_fiscal_start, _ = ReportHelper.get_fiscal_year_dates(prev_year)
            prev_year_end_date = end_date - datetime.timedelta(days=364 * i)
            
            prev_year_ytd_data = self.get_ytd_attendance_data(
                prev_year_fiscal_start, 
                prev_year_end_date, 
                program_id, 
                division_id, 
                org_id
            )
            years_ytd_data[prev_year] = prev_year_ytd_data
        
        performance_timer.log("multiple_years_ytd_data_{}".format(org_id or division_id or program_id))
        return years_ytd_data

    def normalize_service_time(self, time_str):
        """
        Normalize a service time by checking if it's a combined time expression.
        Handles formats like "(9:20 AM|9:30 AM)=9:20 AM"
        """
        # Check if this is a combined time format
        combined_match = re.match(r'\((.*?)\)=(.*?)$', time_str)
        if combined_match:
            # Return the normalized time (after the equals sign)
            return combined_match.group(2).strip()
        return time_str.strip()

    def parse_service_times(self, rpt_group):
        """Parse the RptGroup field to extract service times."""
        order, service_times = ReportHelper.parse_program_rpt_group(rpt_group)
        
        # If no service times are defined, add a "Total" placeholder
        if not service_times:
            service_times = ["Total"]
        
        # Process any combined times - format like "(9:20 AM|9:30 AM)=9:20 AM"
        processed_times = []
        for time in service_times:
            # Check if this is a combined time format
            combined_match = re.match(r'\((.*?)\)=(.*?)$', time)
            if combined_match:
                # Use the right side of equals sign as the display time
                processed_times.append(combined_match.group(2).strip())
            else:
                processed_times.append(time.strip())
        
        # Remove duplicates
        unique_times = []
        seen = set()
        for time in processed_times:
            normalized = time.strip().upper()
            if normalized not in seen:
                seen.add(normalized)
                unique_times.append(time)
                
        return unique_times
    
    def debug_sql(self, sql, description="SQL Query"):
        """Print SQL for debugging."""
        print "<div style='background-color: #f5f5f5; padding: 10px; margin: 10px 0; border-radius: 5px;'>"
        print "<strong>DEBUG - {}:</strong>".format(description)
        print "<pre style='white-space: pre-wrap;'>{}</pre>".format(sql)
        print "</div>"
    
    def generate_header_row(self, service_times, check_for_other=True):
        """Generate the table header row with service times and years."""
        headers = [""] # First column for entity name
        
        # Add column headers for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            year_label = ReportHelper.get_year_label(year)
            headers.append(year_label)
        
        # Add service time columns
        for time in service_times:
            headers.append(time)
        
        # Always include "Other" column if check_for_other is True
        if check_for_other and "Total" not in service_times:
            headers.append("Other")
        
        # Add year-over-year change columns
        for i in range(1, YEARS_TO_DISPLAY):
            headers.append("{0}-{1} Change".format(
                ReportHelper.get_year_label(self.current_year - i + 1),
                ReportHelper.get_year_label(self.current_year - i)
            ))
        
        # Adjust YTD labels based on year type
        ytd_prefix = "FYTD" if YEAR_TYPE.lower() == "fiscal" else "YTD"
        
        # Add YTD columns for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            year_label = ReportHelper.get_year_label(year)
            headers.append("{0} {1}".format(ytd_prefix, year_label))
        
        # Add YTD change columns
        for i in range(1, YEARS_TO_DISPLAY):
            headers.append("{0} {1}-{2} Change".format(
                ytd_prefix,
                ReportHelper.get_year_label(self.current_year - i + 1),
                ReportHelper.get_year_label(self.current_year - i)
            ))
        
        # Add weekly average column
        headers.append("{0} Avg/Week".format(ytd_prefix))
        
        header_html = "<tr>"
        for header in headers:
            header_html += "<th>{}</th>".format(header)
        header_html += "</tr>"
        
        return header_html
    
    def generate_program_row(self, program, years_data, years_ytd_data):
        """Generate a row for a program with attendance data for multiple years."""
        # Start the row with the program name
        row_html = """
            <tr style="background-color: #e9f7fd;">
                <td><strong>{}</strong></td>
        """.format(program.Name)
        
        # Add attendance data for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            if year in years_data:
                row_html += "<td><strong>{}</strong></td>".format(
                    ReportHelper.format_number(years_data[year]['total'])
                )
            else:
                row_html += "<td><strong>0</strong></td>"
        
        # Add service time data - make sure we get the correct times
        service_times = self.parse_service_times(program.RptGroup)
        
        # Keep track of hours that are assigned to specific service times
        accounted_hours = set()
        
        # If there's only one service time and it's "Total"
        if len(service_times) == 1 and service_times[0] == "Total":
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(years_data[self.current_year]['total'])
            )
        else:
            # Track service times already added
            added_times = set()
            
            for time in service_times:
                # Handle 'Total' case
                if time == "Total":
                    # Just add the total value
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['total'])
                    )
                    continue
                    
                hour = self.get_hour_from_service_time(time)
                
                # Skip duplicates
                if hour in added_times:
                    continue
                    
                added_times.add(hour)
                accounted_hours.add(hour)
                
                if hour in years_data[self.current_year]['by_hour']:
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['by_hour'][hour])
                    )
                else:
                    row_html += "<td>0</td>"
            
            # Calculate "Other" attendance (any attendance not in one of the specified service times)
            other_attendance = 0
            for hour, count in years_data[self.current_year]['by_hour'].items():
                if hour not in accounted_hours and hour != "Total":
                    other_attendance += count
            
            # Always add "Other" column if service times don't include "Total"
            if "Total" not in service_times:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(other_attendance)
                )
        
        # Add year-over-year change columns
        for i in range(1, YEARS_TO_DISPLAY):
            current_year = self.current_year - i + 1
            prev_year = self.current_year - i
            
            current_total = years_data[current_year]['total'] if current_year in years_data else 0
            prev_total = years_data[prev_year]['total'] if prev_year in years_data else 0
            
            trend = ReportHelper.get_trend_indicator(current_total, prev_total)
            row_html += "<td>{}</td>".format(trend)
        
        # Add YTD data for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            if year in years_ytd_data:
                row_html += "<td><strong>{}</strong></td>".format(
                    ReportHelper.format_number(years_ytd_data[year]['total'])
                )
            else:
                row_html += "<td><strong>0</strong></td>"
        
        # Add YTD change columns
        for i in range(1, YEARS_TO_DISPLAY):
            current_year = self.current_year - i + 1
            prev_year = self.current_year - i
            
            current_ytd = years_ytd_data[current_year]['total'] if current_year in years_ytd_data else 0
            prev_ytd = years_ytd_data[prev_year]['total'] if prev_year in years_ytd_data else 0
            
            ytd_trend = ReportHelper.get_trend_indicator(current_ytd, prev_ytd)
            row_html += "<td>{}</td>".format(ytd_trend)
        
        # Add YTD Average - Using weeks elapsed instead of days with meetings
        weeks_elapsed = ReportHelper.get_weeks_elapsed_in_fiscal_year(self.report_date, self.current_fiscal_start)
        avg_attendance = 0
        if weeks_elapsed > 0 and self.current_year in years_ytd_data:
            avg_attendance = years_ytd_data[self.current_year]['total'] / float(weeks_elapsed)
        row_html += "<td><strong>{}</strong></td>".format(ReportHelper.format_float(avg_attendance))
        
        row_html += "</tr>"
        return row_html
    
    def generate_division_row(self, division, years_data, years_ytd_data, indent=1):
        """Generate a row for a division with attendance data for multiple years."""
        # Check if we should show rows with zero attendance
        if not SHOW_ZERO_ATTENDANCE:
            all_zeros = True
            for year in years_data:
                if years_data[year]['total'] > 0:
                    all_zeros = False
                    break
            if all_zeros:
                return ""
        
        # Start the row with the division name
        row_html = """
            <tr style="background-color: #f5f5f5;">
                <td style="padding-left: {}em">{}</td>
        """.format(indent, division.Name)
        
        # Add attendance data for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            if year in years_data:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(years_data[year]['total'])
                )
            else:
                row_html += "<td>0</td>"
        
        # Add service time data - assuming divisions use the same service times as their parent program
        program = q.QuerySqlTop1("SELECT RptGroup FROM Program WHERE Id = {}".format(division.ProgId))
        service_times = self.parse_service_times(program.RptGroup)
        
        # Keep track of hours that are assigned to specific service times
        accounted_hours = set()
        
        # If there's only one service time and it's "Total"
        if len(service_times) == 1 and service_times[0] == "Total":
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(years_data[self.current_year]['total'])
            )
        else:
            # Track service times already added
            added_times = set()
            
            for time in service_times:
                # Handle 'Total' case
                if time == "Total":
                    # Just add the total value
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['total'])
                    )
                    continue
                    
                hour = self.get_hour_from_service_time(time)
                
                # Skip duplicates
                if hour in added_times:
                    continue
                    
                added_times.add(hour)
                accounted_hours.add(hour)
                
                if hour in years_data[self.current_year]['by_hour']:
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['by_hour'][hour])
                    )
                else:
                    row_html += "<td>0</td>"
            
            # Calculate "Other" attendance (any attendance not in one of the specified service times)
            other_attendance = 0
            for hour, count in years_data[self.current_year]['by_hour'].items():
                if hour not in accounted_hours and hour != "Total":
                    other_attendance += count
            
            # Always add "Other" column if service times don't include "Total"
            if "Total" not in service_times:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(other_attendance)
                )
        
        # Add year-over-year change columns
        for i in range(1, YEARS_TO_DISPLAY):
            current_year = self.current_year - i + 1
            prev_year = self.current_year - i
            
            current_total = years_data[current_year]['total'] if current_year in years_data else 0
            prev_total = years_data[prev_year]['total'] if prev_year in years_data else 0
            
            trend = ReportHelper.get_trend_indicator(current_total, prev_total)
            row_html += "<td>{}</td>".format(trend)
        
        # Add YTD data for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            if year in years_ytd_data:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(years_ytd_data[year]['total'])
                )
            else:
                row_html += "<td>0</td>"
        
        # Add YTD change columns
        for i in range(1, YEARS_TO_DISPLAY):
            current_year = self.current_year - i + 1
            prev_year = self.current_year - i
            
            current_ytd = years_ytd_data[current_year]['total'] if current_year in years_ytd_data else 0
            prev_ytd = years_ytd_data[prev_year]['total'] if prev_year in years_ytd_data else 0
            
            ytd_trend = ReportHelper.get_trend_indicator(current_ytd, prev_ytd)
            row_html += "<td>{}</td>".format(ytd_trend)
        
        # Add YTD Average - Using weeks elapsed instead of days with meetings
        weeks_elapsed = ReportHelper.get_weeks_elapsed_in_fiscal_year(self.report_date, self.current_fiscal_start)
        avg_attendance = 0
        if weeks_elapsed > 0 and self.current_year in years_ytd_data:
            avg_attendance = years_ytd_data[self.current_year]['total'] / float(weeks_elapsed)
        row_html += "<td>{}</td>".format(ReportHelper.format_float(avg_attendance))
        
        row_html += "</tr>"
        return row_html
    
    def generate_organization_row(self, org, years_data, years_ytd_data, indent=2):
        """Generate a row for an organization with attendance data for multiple years."""
        # Check if we should show rows with zero attendance
        if not SHOW_ZERO_ATTENDANCE:
            all_zeros = True
            for year in years_data:
                if years_data[year]['total'] > 0:
                    all_zeros = False
                    break
            if all_zeros:
                return ""
        
        # Start the row with the organization name
        row_html = """
            <tr>
                <td style="padding-left: {}em">{}</td>
        """.format(indent, org.OrganizationName)
        
        # Add attendance data for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            if year in years_data:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(years_data[year]['total'])
                )
            else:
                row_html += "<td>0</td>"
        
        # Get the division and program to determine service times
        division = q.QuerySqlTop1("SELECT ProgId FROM Division WHERE Id = {}".format(org.DivisionId))
        program = q.QuerySqlTop1("SELECT RptGroup FROM Program WHERE Id = {}".format(division.ProgId))
        service_times = self.parse_service_times(program.RptGroup)
        
        # Keep track of hours that are assigned to specific service times
        accounted_hours = set()
        
        # If there's only one service time and it's "Total"
        if len(service_times) == 1 and service_times[0] == "Total":
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(years_data[self.current_year]['total'])
            )
        else:
            # Track service times already added
            added_times = set()
            
            for time in service_times:
                # Handle 'Total' case
                if time == "Total":
                    # Just add the total value
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['total'])
                    )
                    continue
                    
                hour = self.get_hour_from_service_time(time)
                
                # Skip duplicates
                if hour in added_times:
                    continue
                    
                added_times.add(hour)
                accounted_hours.add(hour)
                
                if hour in years_data[self.current_year]['by_hour']:
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['by_hour'][hour])
                    )
                else:
                    row_html += "<td>0</td>"
            
            # Calculate "Other" attendance (any attendance not in one of the specified service times)
            other_attendance = 0
            for hour, count in years_data[self.current_year]['by_hour'].items():
                if hour not in accounted_hours and hour != "Total":
                    other_attendance += count
            
            # Always add "Other" column if service times don't include "Total"
            if "Total" not in service_times:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(other_attendance)
                )
        
        # Add year-over-year change columns
        for i in range(1, YEARS_TO_DISPLAY):
            current_year = self.current_year - i + 1
            prev_year = self.current_year - i
            
            current_total = years_data[current_year]['total'] if current_year in years_data else 0
            prev_total = years_data[prev_year]['total'] if prev_year in years_data else 0
            
            trend = ReportHelper.get_trend_indicator(current_total, prev_total)
            row_html += "<td>{}</td>".format(trend)
        
        # Add YTD data for each year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            if year in years_ytd_data:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(years_ytd_data[year]['total'])
                )
            else:
                row_html += "<td>0</td>"
        
        # Add YTD change columns
        for i in range(1, YEARS_TO_DISPLAY):
            current_year = self.current_year - i + 1
            prev_year = self.current_year - i
            
            current_ytd = years_ytd_data[current_year]['total'] if current_year in years_ytd_data else 0
            prev_ytd = years_ytd_data[prev_year]['total'] if prev_year in years_ytd_data else 0
            
            ytd_trend = ReportHelper.get_trend_indicator(current_ytd, prev_ytd)
            row_html += "<td>{}</td>".format(ytd_trend)
        
        # Add YTD Average - Using weeks elapsed instead of days with meetings
        weeks_elapsed = ReportHelper.get_weeks_elapsed_in_fiscal_year(self.report_date, self.current_fiscal_start)
        avg_attendance = 0
        if weeks_elapsed > 0 and self.current_year in years_ytd_data:
            avg_attendance = years_ytd_data[self.current_year]['total'] / float(weeks_elapsed)
        row_html += "<td>{}</td>".format(ReportHelper.format_float(avg_attendance))
        
        row_html += "</tr>"
        return row_html
    
    def get_hour_from_service_time(self, time_str):
        """Convert a service time string (e.g., '8:30 AM') to an hour (8), or handle 'Total'."""
        if time_str == "Total":
            return "Total"  # Return "Total" as a special case
        
        # Check for combined time format like "(9:20 AM|9:30 AM)=9:20 AM"
        combined_match = re.match(r'\((.*?)\)=(.*?)$', time_str)
        if combined_match:
            # Use the right side of equals sign for the hour
            time_str = combined_match.group(2).strip()
            
        try:
            # Extract the hour part
            match = re.search(r'(\d+):', time_str)
            if not match:
                return 0
                
            hour = int(match.group(1))
            
            # Check if PM and adjust hour accordingly (except for 12 PM)
            if 'PM' in time_str.upper() and hour != 12:
                hour += 12
            
            # Adjust for 12 AM
            if 'AM' in time_str.upper() and hour == 12:
                hour = 0
                
            return hour
        except:
            return 0
    
    def generate_date_selector_form(self):
        """Generate a date selector form with improved loading indicators."""
        current_date = self.report_date.strftime('%Y-%m-%d')
        send_email_html = ""
        
        # Add email option with improved button states
        if model.Data.send_email != "yes":
            send_email_html = """
            <div style="margin-top: 10px;">
                <label for="email_to">Send Email To:</label>
                <input type="text" id="email_to" name="email_to" placeholder="email@example.com">
                <button type="submit" name="send_email" value="yes" id="send-email-btn" 
                        style="margin-left: 10px; padding: 5px 10px; background-color: #2196F3; color: white; border: none; border-radius: 3px; cursor: pointer; position: relative;">
                    Send Report
                </button>
            </div>
            """
        
        # Show option to expand organizations (default is collapsed)
        expand_checked = "" if self.collapse_orgs else "checked"
        
        form_html = """
        <form method="get" action="" id="report-form" style="margin-bottom: 20px; padding: 10px; background-color: #f5f5f5; border-radius: 5px;">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 10px; flex-wrap: wrap;">
                <label for="report_date">Report Date (Sunday):</label>
                <input type="date" id="report_date" name="report_date" value="{}" required>
                
                <button type="submit" id="run-report-btn" style="padding: 5px 10px; background-color: #4CAF50; color: white; border: none; border-radius: 3px; cursor: pointer; position: relative;">
                    Run Report
                </button>
            </div>
            
            <div style="margin-top: 10px;">
                <label>
                    <input type="checkbox" name="collapse_orgs" value="no" {}>
                    Expand Involvements (Default is Collapsed).. Note:  This is slow.. 3+ minutes slow at times.
                </label>
            </div>
            
            {}
        </form>
        """.format(current_date, expand_checked, send_email_html)
        
        return form_html
    
    def generate_date_range_summary(self):
        """Generate a summary section for the selected date range."""
        date_range_html = """
        <div style="margin-top: 20px; margin-bottom: 20px; padding: 10px; background-color: #f0f8ff; border-radius: 5px; border-left: 5px solid #4CAF50;">
            <h3 style="margin-top: 0;">Weekly Report</h3>
            <p>Current Week: <strong>{}</strong> to <strong>{}</strong></p>
            <p>Previous Week: <strong>{}</strong> to <strong>{}</strong></p>
            <p>Previous Year: <strong>{}</strong> to <strong>{}</strong></p>
        </div>
        """.format(
            ReportHelper.format_display_date(self.week_start_date),
            ReportHelper.format_display_date(self.report_date),
            ReportHelper.format_display_date(self.previous_week_start),
            ReportHelper.format_display_date(self.previous_week_date),
            ReportHelper.format_display_date(self.previous_week_start_date),
            ReportHelper.format_display_date(self.previous_year_date)
        )
        
        return date_range_html
    
    def generate_fiscal_year_summary(self, current_ytd_total, previous_ytd_total):
        """Generate a fiscal year summary section."""
        # Adjust labels based on year type
        ytd_prefix = "FYTD" if YEAR_TYPE.lower() == "fiscal" else "YTD"
        current_label = ReportHelper.get_year_label(self.current_fiscal_year)
        previous_label = ReportHelper.get_year_label(self.previous_fiscal_year)
        
        # Calculate change percentage and color
        if previous_ytd_total > 0:
            percent_change = ((current_ytd_total - previous_ytd_total) / float(previous_ytd_total)) * 100
            if percent_change > 0:
                change_color = "green"
                change_arrow = "&#9650;"
            elif percent_change < 0:
                change_color = "red"
                change_arrow = "&#9660;"
                percent_change = abs(percent_change)
            else:
                change_color = "#777"
                change_arrow = "&#8652;"
        else:
            percent_change = 0
            change_color = "green"
            change_arrow = "NEW"
        
        summary_html = """
        <div style="margin-top: 20px; margin-bottom: 30px; padding: 15px; background-color: #f3f8ff; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h3 style="margin-top: 0;">{} Summary</h3>
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 20px;">
                <div style="flex: 1; min-width: 200px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 1.2em; font-weight: bold;">{} Current</div>
                    <div style="font-size: 2em; margin: 10px 0;">{}</div>
                    <div>From {} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 200px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 1.2em; font-weight: bold;">{} Previous</div>
                    <div style="font-size: 2em; margin: 10px 0;">{}</div>
                    <div>From {} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 200px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 1.2em; font-weight: bold;">{} Change</div>
                    <div style="font-size: 2em; margin: 10px 0; color: {};">{} {}%</div>
                    <div>{} vs {}</div>
                </div>
            </div>
        </div>
        """.format(
            ytd_prefix,                                     # YTD or FYTD for header
            ytd_prefix,                                     # YTD or FYTD for Current
            ReportHelper.format_number(current_ytd_total),
            ReportHelper.format_display_date(self.current_fiscal_start),
            ReportHelper.format_display_date(self.report_date),
            
            ytd_prefix,                                     # YTD or FYTD for Previous
            ReportHelper.format_number(previous_ytd_total),
            ReportHelper.format_display_date(self.previous_fiscal_start),
            ReportHelper.format_display_date(self.previous_year_date),
            
            ytd_prefix,                                     # YTD or FYTD for Change
            change_color,
            change_arrow,
            ReportHelper.format_float(percent_change),      # Using format_float instead of string formatting
            ReportHelper.format_number(current_ytd_total),
            ReportHelper.format_number(previous_ytd_total)
        )
        
        return summary_html
    
    def generate_selected_program_average_summaries(self, current_week_data, prev_week_data, current_ytd_data, previous_ytd_data):
        """Generate attendance totals summaries for selected programs."""
    
        # Get all programs with their RptGroup for ordering
        all_programs_sql = """
            SELECT 
                Name,
                RptGroup
            FROM Program
            WHERE RptGroup IS NOT NULL AND RptGroup <> ''
            ORDER BY RptGroup
        """
        
        all_program_data = []
        try:
            results = q.QuerySql(all_programs_sql)
            #self.debug_sql(all_programs_sql, "programs")
            for row in results:
                if row.Name not in EXCLUDED_PROGRAMS:
                    # Extract the report order number from RptGroup
                    order, _ = ReportHelper.parse_program_rpt_group(row.RptGroup)
                    all_program_data.append({
                        'name': row.Name,
                        'rpt_group': row.RptGroup,
                        'order': order
                    })
            
            # Sort by the order extracted from RptGroup
            all_program_data.sort(key=lambda x: x['order'] if x['order'] else "999")
            
        except Exception as e:
            print "<p>Error getting programs: {}</p>".format(e)
            return ""
        
        # Extract just the program names for data retrieval
        all_programs = [prog['name'] for prog in all_program_data]
        
        if not all_programs or len(all_programs) == 0:
            return ""
            
        # Get specific program data for current FYTD period
        current_specific_data = self.get_specific_program_attendance_data(
            all_programs,
            self.current_fiscal_start,
            self.report_date
        )
        
        # Get specific program data for previous year's same period (PYTD)
        previous_specific_data = self.get_specific_program_attendance_data(
            all_programs,
            self.previous_fiscal_start,
            self.previous_year_date
        )
        
        # Get previous year's same week data (instead of previous week)
        prev_year_specific_week_data = self.get_specific_program_attendance_data(
            all_programs,
            self.previous_week_start_date,
            self.previous_year_date
        )
        
        summary_html = """
        <div style="margin-top: 20px; margin-bottom: 30px; padding: 15px; background-color: #f0f0f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h3 style="margin-top: 0;">Program Summary</h3>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <thead>
                    <tr style="background-color: #e0e0e8;">
                        <th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Program</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #ddd;">Current Week</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #ddd;">Previous Year</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #ddd;">YoY Change</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #ddd;">FYTD</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #ddd;">PYTD</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #ddd;">FYTD Change</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        # Add a row for each program in order of RptGroup
        for program_data in all_program_data:
            program_name = program_data['name']
            
            # Skip if program isn't found in current data
            if program_name not in current_specific_data['by_program']:
                continue
                
            # Get current program data
            prog_data = current_specific_data['by_program'][program_name]
            
            # Use total instead of average for FYTD
            current_fytd_total = prog_data['total']
            
            # Get previous year FYTD data for the program (PYTD)
            prev_fytd_total = 0
            if program_name in previous_specific_data['by_program']:
                prev_ytd_data = previous_specific_data['by_program'][program_name]
                prev_fytd_total = prev_ytd_data['total']
            
            # Get current week data
            current_week_total = 0
            if program_name in current_week_data['by_program']:
                current_week_total = current_week_data['by_program'][program_name]['total']
            
            # Get previous YEAR'S SAME WEEK data (instead of previous week)
            prev_year_week_total = 0
            if prev_year_specific_week_data and program_name in prev_year_specific_week_data['by_program']:
                prev_year_week_total = prev_year_specific_week_data['by_program'][program_name]['total']
            
            # Calculate trend indicators
            year_to_year_week_trend = ReportHelper.get_trend_indicator(current_week_total, prev_year_week_total)
            year_to_year_trend = ReportHelper.get_trend_indicator(current_fytd_total, prev_fytd_total)
            
            # Add a row for this program
            summary_html += """
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 8px; text-align: left; border: 1px solid #ddd;"><strong>{}</strong></td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>
                    <td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>
                </tr>
            """.format(
                program_name,
                ReportHelper.format_number(current_week_total),
                ReportHelper.format_number(prev_year_week_total),
                year_to_year_week_trend,
                ReportHelper.format_number(current_fytd_total),
                ReportHelper.format_number(prev_fytd_total),
                year_to_year_trend
            )
        
        # Close the table and container (no totals row)
        summary_html += """
                </tbody>
            </table>
        </div>
        """
        
        return summary_html
    
    def generate_overall_summary(self, current_total, previous_total, current_ytd_total, previous_ytd_total):
        """Generate an overall summary section."""
        # Adjust labels based on year type
        ytd_prefix = "FYTD" if YEAR_TYPE.lower() == "fiscal" else "YTD"
        
        trend = ReportHelper.get_trend_indicator(current_total, previous_total)
        ytd_trend = ReportHelper.get_trend_indicator(current_ytd_total, previous_ytd_total)
        
        summary_html = """
        <div style="margin-top: 20px; margin-bottom: 30px; padding: 15px; background-color: #f8f8f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h3 style="margin-top: 0;">Overall Summary</h3>
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 20px;">
                <div style="flex: 1; min-width: 180px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 1.2em; font-weight: bold;">Current Week</div>
                    <div style="font-size: 2em; margin: 10px 0;">{}</div>
                    <div>{} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 180px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 1.2em; font-weight: bold;">Previous Year</div>
                    <div style="font-size: 2em; margin: 10px 0;">{}</div>
                    <div>{} {}</div>
                </div>
            </div>
        </div>
        """.format(
            ReportHelper.format_number(current_total),
            ReportHelper.format_display_date(self.week_start_date),
            ReportHelper.format_display_date(self.report_date),
            
            ReportHelper.format_number(previous_total),
            trend,
            ReportHelper.format_display_date(self.previous_year_date)
        )
        
        return summary_html
    
    def send_email_report(self, email_to, report_content):
        """Send the report via email."""
        try:
            # Add indication that email is being sent
            email_status = """
            <div id="email-status" style="padding: 15px; background-color: #fff8e1; color: #856404; border-radius: 5px; margin-bottom: 20px;">
                <strong>Sending email to {}...</strong>
                <div class="spinner-small" style="display: inline-block; margin-left: 10px;"></div>
            </div>
            <script>
            document.addEventListener('DOMContentLoaded', function() {{
                var emailStatus = document.getElementById('email-status');
                if (emailStatus) {{
                    emailStatus.scrollIntoView({{behavior: 'smooth'}});
                }}
            }});
            </script>
            """.format(email_to)
            
            print email_status
            
            # Set up email parameters
            QueuedBy = model.UserPeopleId
            
            # Set email tracking
            email_content = report_content + "{track}{tracklinks}"
            
            # Use the provided email_to instead of hardcoding
            # Convert to PeopleId if needed or use a query to get PeopleId from email
            try:
                # If email_to is a number (PeopleId), use it directly
                email_id = int(email_to)
            except ValueError:
                # Otherwise, query by email address or use the current user
                # For now, just use the current user as fallback
                email_id = model.UserPeopleId
            
            # Send the email
            model.Email(email_id, QueuedBy, EMAIL_FROM_ADDRESS, EMAIL_FROM_NAME, EMAIL_SUBJECT, email_content)
            
            # Update status to success
            success_message = """
            <script>
            var emailStatus = document.getElementById('email-status');
            if (emailStatus) {{
                emailStatus.className = '';
                emailStatus.style.backgroundColor = '#d4edda';
                emailStatus.style.color = '#155724';
                emailStatus.innerHTML = '<strong>Success!</strong> Email sent successfully to {}';
            }}
            </script>
            """.format(email_to)
            
            print success_message
            
            return True, "Email sent successfully to {}".format(email_to)
        except Exception as e:
            # Update status to error
            error_message = """
            <script>
            var emailStatus = document.getElementById('email-status');
            if (emailStatus) {{
                emailStatus.className = '';
                emailStatus.style.backgroundColor = '#f8d7da';
                emailStatus.style.color = '#721c24';
                emailStatus.innerHTML = '<strong>Error:</strong> {}';
            }}
            </script>
            """.format(str(e))
            
            print error_message
            
            return False, "Error sending email: {}".format(str(e))

    
    def generate_report(self):
        """Generate the complete attendance report."""
        try:
            performance_timer.start("total_report_generation")
            
            # Enhanced CSS for better loading indicators
            report_html = """
            <style>
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
            
            /* Main loading icon - large and centered */
            .loading-container {
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(255, 255, 255, 0.9);
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                z-index: 9999;
                transition: opacity 0.5s;
            }
            
            .spinner-large {
                width: 70px;
                height: 70px;
                border: 8px solid #f3f3f3;
                border-top: 8px solid #3498db;
                border-radius: 50%;
                animation: spin 1s linear infinite;
                margin-bottom: 20px;
            }
            
            /* Button loading states */
            .btn-loading {
                position: relative;
                color: transparent !important;
            }
            
            .btn-loading:after {
                content: '';
                position: absolute;
                left: 50%;
                top: 50%;
                width: 20px;
                height: 20px;
                margin-left: -10px;
                margin-top: -10px;
                border: 3px solid rgba(255, 255, 255, 0.3);
                border-radius: 50%;
                border-top: 3px solid #fff;
                animation: spin 1s linear infinite;
            }
            
            /* Progress reporting */
            .progress-container {
                width: 80%;
                max-width: 300px;
                background-color: #f1f1f1;
                border-radius: 5px;
                margin-top: 20px;
            }
            
            .progress-bar {
                width: 10%;
                height: 20px;
                background-color: #4CAF50;
                border-radius: 5px;
                text-align: center;
                line-height: 20px;
                color: white;
                transition: width 0.3s;
            }
            
            .spinner {
                display: inline-block;
                width: 40px;
                height: 40px;
                border: 4px solid rgba(0, 0, 0, 0.1);
                border-radius: 50%;
                border-top-color: #3498db;
                animation: spin 1s ease-in-out infinite;
            }
            
            .spinner-small {
                display: inline-block;
                width: 20px;
                height: 20px;
                border: 3px solid rgba(0, 0, 0, 0.1);
                border-radius: 50%;
                border-top-color: #3498db;
                animation: spin 1s ease-in-out infinite;
                vertical-align: middle;
            }
            
            /* Processing overlay */
            #processing-overlay {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(255, 255, 255, 0.9);
                z-index: 9999;
                justify-content: center;
                align-items: center;
                flex-direction: column;
            }
            
            #processing-status {
                margin-top: 20px;
                font-size: 18px;
                text-align: center;
            }
            
            #processing-details {
                margin-top: 10px;
                font-size: 14px;
                color: #666;
                max-width: 80%;
                text-align: center;
            }
            </style>
            
            <!-- Processing overlay for when report is generating -->
            <div id="processing-overlay">
                <div class="spinner-large"></div>
                <div id="processing-status">Preparing report...</div>
                <div id="processing-details">Initializing...</div>
                <div class="progress-container" style="margin-top: 20px;">
                    <div id="processing-progress" class="progress-bar" style="width: 0%;">0%</div>
                </div>
            </div>
            
            <!-- Report content container -->
            <div id="report-content">
            """
            
            # Simple CSS included directly in the output
            report_html += "<style>\n"
            report_html += "table {border-collapse: collapse; width: 100%; margin-bottom: 20px;}\n"
            report_html += "th, td {border: 1px solid #ddd; padding: 8px; text-align: center;}\n"
            report_html += "th {background-color: #f2f2f2; font-weight: bold;}\n"
            report_html += "td:first-child {text-align: left;}\n"
            report_html += ".summary-box {margin-bottom: 20px; padding: 15px; background-color: #f9f9f9; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);}\n"
            report_html += ".metric {font-size: 1.5em; font-weight: bold; margin: 5px 0;}\n"
            report_html += ".trend-positive {color: green;}\n"
            report_html += ".trend-negative {color: red;}\n"
            report_html += ".trend-neutral {color: #777;}\n"
            report_html += "@media print {\n"
            report_html += "  form, button, .no-print {display: none !important;}\n"
            report_html += "  body {font-size: 12pt;}\n"
            report_html += "  table {page-break-inside: avoid;}\n"
            report_html += "}\n"
            report_html += "</style>\n"
            
            # Add debug mode indicator if enabled
            if performance_timer.enabled:
                report_html += """
                <div style="padding: 10px; margin-bottom: 20px; background-color: #fff3cd; color: #856404; border-left: 5px solid #ffeeba; border-radius: 3px;">
                    <strong>Performance Debug Mode Enabled</strong> - Timing information will be shown at the end of the report.
                </div>
                """
            
            report_html += """<h2>{}                  
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">i</text>
                  </svg></h2>\n""".format(REPORT_TITLE)
                  
            report_html += "<p>Report Type: {}</p>\n".format(
                "Fiscal Year (Starting {}/{})".format(FISCAL_YEAR_START_MONTH, FISCAL_YEAR_START_DAY) 
                if YEAR_TYPE.lower() == "fiscal" else "Calendar Year"
            )
            
            # Email sending section
            if model.Data.send_email == "yes" and hasattr(model.Data, 'email_to') and model.Data.email_to:
                email_to = model.Data.email_to
                email_report = self.generate_report_content(for_email=True)
                success, message = self.send_email_report(email_to, email_report)
                
                if success:
                    report_html += '<div style="padding: 15px; background-color: #d4edda; color: #155724; border-radius: 5px; margin-bottom: 20px;">'
                    report_html += '<strong>Success!</strong> {}'.format(message)
                    report_html += '</div>'
                else:
                    report_html += '<div style="padding: 15px; background-color: #f8d7da; color: #721c24; border-radius: 5px; margin-bottom: 20px;">'
                    report_html += '<strong>Error:</strong> {}'.format(message)
                    report_html += '</div>'
            
            # Add date selector form 
            report_html += self.generate_date_selector_form()
            
            # Check if we should load data
            if self.should_load_data:
                # Add date range summary
                report_html += self.generate_date_range_summary()
                
                # Generate all report content
                report_html += self.generate_report_content()
            else:
                # Show instructions when no date is selected
                report_html += """
                <div style="margin-top: 30px; padding: 20px; background-color: #f8f9fa; border-radius: 5px; text-align: center;">
                    <h3>Welcome to the Weekly Attendance Dashboard</h3>
                    <p>Please select a report date (Sunday) and click 'Run Report' to generate the attendance report.</p>
                    <div style="margin-top: 20px;">
                        <i class="fa fa-chart-line" title="Attendance List"></i>
                    </div>
                </div>
                """
            
            report_html += "</div> <!-- End report-content -->"
            
            # Add JavaScript to handle loading state and form submission
            report_html += """
            <script>
            // Function to show processing overlay with status updates
            function showProcessingOverlay(message, details, progress) {
                var overlay = document.getElementById('processing-overlay');
                var status = document.getElementById('processing-status');
                var detailsElem = document.getElementById('processing-details');
                var progressBar = document.getElementById('processing-progress');
                
                if (overlay && status && detailsElem && progressBar) {
                    overlay.style.display = 'flex';
                    status.textContent = message || 'Processing...';
                    detailsElem.textContent = details || '';
                    
                    if (typeof progress === 'number') {
                        progressBar.style.width = progress + '%';
                        progressBar.textContent = progress + '%';
                    }
                }
            }
            
            // Function to hide processing overlay
            function hideProcessingOverlay() {
                var overlay = document.getElementById('processing-overlay');
                if (overlay) {
                    overlay.style.display = 'none';
                }
            }
            
            // Enhance form submission with loading states
            document.addEventListener('DOMContentLoaded', function() {
                var reportForm = document.getElementById('report-form');
                if (reportForm) {
                    reportForm.addEventListener('submit', function(e) {
                        // Don't show overlay for email sending
                        if (e.submitter && e.submitter.name === 'send_email') {
                            return;
                        }
                        
                        // Show processing overlay
                        showProcessingOverlay('Preparing report...', 'This may take a moment, please wait.', 10);
                        
                        // Disable submit button
                        var submitBtn = e.submitter;
                        if (submitBtn) {
                            submitBtn.classList.add('btn-loading');
                            submitBtn.disabled = true;
                        }
                        
                        // Animate progress with detailed status updates
                        var progress = 10;
                        var progressInterval = setInterval(function() {
                            progress += 2;
                            if (progress >= 95) {
                                clearInterval(progressInterval);
                            }
                            
                            // Update the detail text based on progress percentage
                            var detailText = '';
                            if (progress < 30) {
                                detailText = 'Loading program data and service times...';
                            } else if (progress < 50) {
                                detailText = 'Calculating weekly attendance metrics...';
                            } else if (progress < 70) {
                                detailText = 'Processing year-to-date comparisons...';
                            } else if (progress < 85) {
                                detailText = 'Generating division and organization data...';
                            } else {
                                detailText = 'Finalizing report and formatting...';
                            }
                            
                            showProcessingOverlay('Generating report...', detailText, progress);
                        }, 300);
                    });
                }
            });
            </script>
            """
            
            performance_timer.end("total_report_generation")
            if performance_timer.enabled:
                report_html += performance_timer.get_report()
    
            return report_html
            
        except Exception as e:
            # Print any errors
            import traceback
            error_html = """
            <h2>Error</h2>
            <p>An error occurred: {}</p>
            <pre>{}</pre>
            """.format(str(e), traceback.format_exc())
            return error_html
            
        except Exception as e:
            # Print any errors
            import traceback
            error_html = """
            <h2>Error</h2>
            <p>An error occurred: {}</p>
            <pre>{}</pre>
            """.format(str(e), traceback.format_exc())
            return error_html
    
    def generate_report_content(self, for_email=False):
        """Generate just the report content (used for both display and email)."""
        performance_timer.start("generate_report_content")
    
        report_content = ""
        
        # Only include progress indicators if not generating for email
        if not for_email:
            report_content += """
            <div id="detailed-progress" style="margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; border-left: 5px solid #4CAF50;">
                <h4 style="margin-top: 0;">Processing Status</h4>
                <div id="current-task" style="margin-bottom: 10px; font-weight: bold;">Initializing report...</div>
                <div style="display: flex; margin-bottom: 5px;">
                    <div style="width: 150px; font-weight: bold;">Programs:</div>
                    <div id="programs-status">Waiting...</div>
                </div>
                <div style="display: flex; margin-bottom: 5px;">
                    <div style="width: 150px; font-weight: bold;">Attendance:</div>
                    <div id="attendance-status">Waiting...</div>
                </div>
                <div style="display: flex; margin-bottom: 5px;">
                    <div style="width: 150px; font-weight: bold;">Divisions:</div>
                    <div id="divisions-status">Waiting...</div>
                </div>
                <div style="display: flex; margin-bottom: 5px;">
                    <div style="width: 150px; font-weight: bold;">Organizations:</div>
                    <div id="organizations-status">Waiting...</div>
                </div>
            </div>
            
            <script>
            // Function to update detailed progress with more specific task information
            function updateDetailedProgress(section, status, isComplete, subtasks) {
                var element = document.getElementById(section + '-status');
                if (element) {
                    if (isComplete) {
                        element.innerHTML = '✓ ' + status;
                        element.style.color = '#28a745';
                    } else {
                        var statusHtml = '<div class="spinner-small" style="margin-right: 10px;"></div> ' + status;
                        
                        // Add subtasks if provided
                        if (subtasks && subtasks.length) {
                            statusHtml += '<ul style="margin: 5px 0 0 25px; font-size: 12px; color: #666;">';
                            subtasks.forEach(function(task) {
                                statusHtml += '<li>' + task + '</li>';
                            });
                            statusHtml += '</ul>';
                        }
                        
                        element.innerHTML = statusHtml;
                        element.style.color = '#007bff';
                    }
                }
                
                // Also update the processing overlay
                var processingDetails = document.getElementById('processing-details');
                if (processingDetails) {
                    processingDetails.textContent = status;
                }
            }
            
            // Function to update current task with more information
            function updateCurrentTask(task, percent) {
                var element = document.getElementById('current-task');
                if (element) {
                    element.textContent = task + ' (' + percent + '% complete)';
                }
                
                // Also update the processing overlay
                var processingStatus = document.getElementById('processing-status');
                if (processingStatus) {
                    processingStatus.textContent = task;
                }
            }
            
            // Start with programs section with detailed subtasks
            updateDetailedProgress('programs', 'Loading program data...', false, [
                'Retrieving program metadata',
                'Processing service times',
                'Sorting by report order'
            ]);
            updateCurrentTask('Retrieving program information', 10);
            </script>
            """
    
        # Track overall totals
        performance_timer.start("calculate_totals")
        overall_totals = {}
        ytd_overall_totals = {}
        
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            overall_totals[year] = 0
            ytd_overall_totals[year] = 0
        performance_timer.end("calculate_totals")
        
        # Get all programs
        performance_timer.start("load_programs")
        programs = q.QuerySql(self.get_programs_sql())
        performance_timer.end("load_programs")
        
        # Update progress after loading programs
        if not for_email:
            report_content += """
            <script>
            // Programs loaded
            updateDetailedProgress('programs', 'Programs loaded successfully', true);
            updateDetailedProgress('attendance', 'Calculating attendance data...', false, [
                'Processing current week attendance',
                'Calculating previous week comparisons',
                'Analyzing year-over-year trends',
                'Computing fiscal year-to-date metrics'
            ]);
            updateCurrentTask('Processing attendance data', 30);
            </script>
            """
            
        # Pre-calculate overall totals
        performance_timer.start("generate_program_tables")
        for program in programs:
            order, service_times = ReportHelper.parse_program_rpt_group(program.RptGroup)
            if not service_times:
                continue
            
            # Get program attendance data for multiple years
            years_program_data = self.get_multiple_years_attendance_data(
                self.report_date,
                program_id=program.Id
            )
            
            # Get YTD data for multiple years
            years_program_ytd = self.get_multiple_years_ytd_data(
                self.report_date,
                program_id=program.Id
            )
            
            # Add to overall totals
            for year in years_program_data:
                overall_totals[year] += years_program_data[year]['total']
                
            for year in years_program_ytd:
                ytd_overall_totals[year] += years_program_ytd[year]['total']
                
        performance_timer.end("generate_program_tables")
        
        # Update progress after calculating attendance
        if not for_email:
            report_content += """
            <script>
            // Attendance calculated
            updateDetailedProgress('attendance', 'Attendance data processed', true);
            updateDetailedProgress('divisions', 'Generating summaries...', false);
            updateCurrentTask('Creating summary sections');
            </script>
            """
        
        # First add the overall summary section
        report_content += self.generate_overall_summary(
            overall_totals[self.current_year],
            overall_totals[self.current_year - 1] if self.current_year - 1 in overall_totals else 0,
            ytd_overall_totals[self.current_year],
            ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0
        )
        
        # Add fiscal year-to-date summary (only once)
        report_content += self.generate_fiscal_year_summary(
            ytd_overall_totals[self.current_year],
            ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0
        )
        
        # Get data for program totals section
        current_week_data = self.get_week_attendance_data(self.week_start_date, self.report_date)
        previous_week_data = self.get_week_attendance_data(
            self.previous_week_start_date, 
            self.previous_year_date
        )
        
        # Add program totals section
        report_content += self.generate_selected_program_average_summaries(
            current_week_data,
            previous_week_data,
            ytd_overall_totals[self.current_year],
            ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0
        )
        
        # Update progress before program tables
        if not for_email:
            report_content += """
            <script>
            // Summaries created
            updateDetailedProgress('divisions', 'Summaries generated', true);
            updateDetailedProgress('organizations', 'Creating program and division tables...', false);
            updateCurrentTask('Generating detailed program tables');
            </script>
            """
        
        # Hide the detailed progress now that we're moving to tables
        if not for_email:
            report_content += """
            <script>
            // Hide detailed progress as we start showing tables
            document.getElementById('detailed-progress').style.display = 'none';
            
            // Update processing overlay
            var processingOverlay = document.getElementById('processing-overlay');
            if (processingOverlay) {
                processingOverlay.style.display = 'none';
            }
            </script>
            """
        
        # Generate program tables
        performance_timer.start("generate_program_tables")
        for program in programs:
            order, service_times = ReportHelper.parse_program_rpt_group(program.RptGroup)
            if not service_times:
                continue
            
            # Start a new table for each program section
            report_content += """
            <h3>{}. {}</h3>
            <table>
            """.format(order, program.Name)
            
            # Add the header row
            report_content += self.generate_header_row(service_times)
            
            # Get program attendance data for multiple years
            years_program_data = self.get_multiple_years_attendance_data(
                self.report_date,
                program_id=program.Id
            )
            
            # Get YTD data for multiple years
            years_program_ytd = self.get_multiple_years_ytd_data(
                self.report_date,
                program_id=program.Id
            )
            
            # Add program row
            report_content += self.generate_program_row(
                program, 
                years_program_data, 
                years_program_ytd
            )
            
            # Get divisions for this program
            divisions = q.QuerySql(self.get_divisions_sql(program.Id))
            
            for division in divisions:
                # Get division attendance data for multiple years
                years_division_data = self.get_multiple_years_attendance_data(
                    self.report_date,
                    division_id=division.Id
                )
                
                # Get YTD data for multiple years
                years_division_ytd = self.get_multiple_years_ytd_data(
                    self.report_date,
                    division_id=division.Id
                )
                
                # Add division row
                division_row = self.generate_division_row(
                    division, 
                    years_division_data, 
                    years_division_ytd
                )
                report_content += division_row
                
                # Add organizations if needed and not collapsed
                if SHOW_ORGANIZATION_DETAILS and not self.collapse_orgs and (division_row != ""):
                    organizations = q.QuerySql(self.get_organizations_sql(division.Id))
                    
                    for org in organizations:
                        # Get organization attendance data for multiple years
                        performance_timer.start("org_processing_{}".format(org.OrganizationId))
                        years_org_data = self.get_multiple_years_attendance_data(
                            self.report_date,
                            org_id=org.OrganizationId
                        )
                        
                        # Get YTD data for multiple years
                        years_org_ytd = self.get_multiple_years_ytd_data(
                            self.report_date,
                            org_id=org.OrganizationId
                        )
                        performance_timer.log("org_processing_{}".format(org.OrganizationId))
                        
                        # Add organization row
                        report_content += self.generate_organization_row(
                            org, 
                            years_org_data, 
                            years_org_ytd
                        )
            
            # Close the table
            report_content += "</table>"
        
        performance_timer.end("generate_program_tables")
        
        # Signal completion
        if not for_email:
            report_content += """
            <script>
            // Update organizations status to complete
            updateDetailedProgress('organizations', 'Tables created successfully', true);
            updateCurrentTask('Report generation complete');
            
            // Hide processing overlay if it's still visible
            hideProcessingOverlay();
            </script>
            """
        
        performance_timer.end("generate_report_content")
        return report_content

# Create and run the report
try:
    # Check for debug parameter in URL
    debug_param = getattr(model.Data, 'debug', None)
    if debug_param == 'performance':
        performance_timer.enabled = True
    
    performance_timer.start("script_execution")
    report = AttendanceReport()
    report_html = report.generate_report()
    print report_html
    performance_timer.end("script_execution")
    
    # Only add this script if we're actually loading data
    if report.should_load_data:
        print """
        <script>
        // Hide the processing overlay when content has loaded
        document.addEventListener('DOMContentLoaded', function() {
            var overlay = document.getElementById('processing-overlay');
            if (overlay) {
                overlay.style.display = 'none';
            }
        });
        
        // Backup in case DOMContentLoaded has already fired
        window.onload = function() {
            var overlay = document.getElementById('processing-overlay');
            if (overlay) {
                overlay.style.display = 'none';
            }
        };
        
        // One more backup approach - after 5 seconds, hide regardless
        setTimeout(function() {
            var overlay = document.getElementById('processing-overlay');
            if (overlay) {
                overlay.style.display = 'none';
            }
        }, 5000);
        </script>
        """
    
    if performance_timer.enabled:
        print "<script>console.log('Total script execution time: {:.4f} seconds');</script>".format(
            performance_timer.results.get("script_execution", 0)
        )
    
except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
