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
#last updated: 04/29/2025

#####################################################################
#### USER CONFIG FIELDS (Modify these settings as needed)
#####################################################################
# Report Title (displayed at top of page)
REPORT_TITLE = "Weekly Attendance"

# Year Type: "fiscal" or "calendar"
YEAR_TYPE = "fiscal"

# Custom labels for year type references (in headers and summaries)
# These will override default "FY" and "FYTD" when YEAR_TYPE = "fiscal"
YEAR_PREFIX_LABEL = "AY"     # Default: "FY" - Used for year labels like "FY24-25"
YTD_PREFIX_LABEL = "AY"    # Default: "FYTD" - Used for year-to-date metrics

# If YEAR_TYPE = "fiscal", set the first month and day of fiscal year
FISCAL_YEAR_START_MONTH = 8  # October
FISCAL_YEAR_START_DAY = 1     # 1st

# Number of years to display in comparison (current year + this many previous years)
YEARS_TO_DISPLAY = 2

# Overall Summary
WORSHIP_PROGRAM = "Worship" 
ENROLLMENT_RATIO_PROGRAM = "Connect Group Attendance"

# Include organization details when expanding divisions
SHOW_ORGANIZATION_DETAILS = True

# Include zero attendance rows for divisions
SHOW_ZERO_ATTENDANCE = True

# Default to collapsed organizations - set to TRUE
DEFAULT_COLLAPSED = True

# Show/hide program summary section
SHOW_PROGRAM_SUMMARY = True

# Show/hide specific columns in program tables
SHOW_CURRENT_WEEK_COLUMN = False
SHOW_PREVIOUS_YEAR_COLUMN = False
SHOW_YOY_CHANGE_COLUMN = True
SHOW_FYTD_COLUMNS = True
SHOW_FYTD_CHANGE_COLUMNS = True
SHOW_FYTD_AVG_COLUMN = True
SHOW_DETAILED_ENROLLMENT = False
SHOW_ENROLLMENT_COLUMN = True

# Show 4-week attendance comparison
SHOW_FOUR_WEEK_COMPARISON = True

# Programs to exclude from Average Attendance calculations
EXCLUDED_PROGRAMS =  ['VBS 2025 K-5', 'VBS 2025 Preschool','VBS 2025 Special Friends','VBS DAILY GRAND TOTAL']

# Email settings
EMAIL_FROM_NAME = "Attendance Reports"
EMAIL_FROM_ADDRESS = "attendance@church.org"
EMAIL_SUBJECT = "Weekly Attendance Report"

# Enable or disable performance debugging (set to True to show timing information)
DEBUG_PERFORMANCE = False

#### ENROLLMENT ANALYSIS CONFIG
# How many weeks to look back for enrollment ratio analysis 
ENROLLMENT_RATIO_WEEKS = 12  # Default to 12 week rolling window
PROSPECT_CONVERSION_WEEKS = 12

# Programs to analyze for enrollment ratios (empty list means analyze all)
ENROLLMENT_ANALYSIS_PROGRAMS = ['Connect Group Attendance']

# Member type ID for prospects
PROSPECT_MEMBER_TYPE_ID = 311

# Enrollment ratio thresholds
ENROLLMENT_RATIO_THRESHOLDS = {
    'needs_inreach': 39,    # 0-44% needs in-reach
    'good_ratio': 59,       # 45-55% is good
    # Above 55% needs outreach
}


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
cache = {}

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
        """Format float number with decimals and handle None or zero values."""
        if number is None or number == 0:
            return "0"
        try:
            format_str = "{:,.{}f}".format(float(number), decimals)
            return format_str
        except (ValueError, TypeError):
            return "0"
    
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
            # Use custom prefix label if defined
            prefix = YEAR_PREFIX_LABEL if YEAR_PREFIX_LABEL else "FY"
            return "{}<br>{}-{}".format(prefix, str(year)[-2:], str(year + 1)[-2:])
        return str(year)
    
    @staticmethod
    def get_ytd_label():
        """Get YTD label based on year type and custom configuration."""
        if YEAR_TYPE.lower() == "fiscal":
            return YTD_PREFIX_LABEL if YTD_PREFIX_LABEL else "FYTD"
        return "YTD"
    
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
    def get_date_from_weeks_ago(date_obj, weeks=4):
        """Get the date that is N weeks ago from the given date."""
        return date_obj - datetime.timedelta(days=7 * weeks)

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
            
            # Get 4-week period dates
            self.four_weeks_ago_date = ReportHelper.get_date_from_weeks_ago(self.report_date, 4)
            self.four_weeks_ago_start_date = ReportHelper.get_week_start_date(self.four_weeks_ago_date)
            
            # Get previous year's 4-week period dates
            self.prev_year_four_weeks_ago_date = ReportHelper.get_date_from_weeks_ago(self.previous_year_date, 4)
            self.prev_year_four_weeks_ago_start_date = ReportHelper.get_week_start_date(self.prev_year_four_weeks_ago_date)
            
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
            
        if hasattr(model.Data, 'show_current_week'):
            global SHOW_CURRENT_WEEK_COLUMN
            SHOW_CURRENT_WEEK_COLUMN = model.Data.show_current_week == 'yes'
        
        if hasattr(model.Data, 'show_previous_year'):
            global SHOW_PREVIOUS_YEAR_COLUMN
            SHOW_PREVIOUS_YEAR_COLUMN = model.Data.show_previous_year == 'yes'
            
        if hasattr(model.Data, 'show_yoy_change'):
            global SHOW_YOY_CHANGE_COLUMN
            SHOW_YOY_CHANGE_COLUMN = model.Data.show_yoy_change == 'yes'
            
        if hasattr(model.Data, 'show_fytd'):
            global SHOW_FYTD_COLUMNS
            SHOW_FYTD_COLUMNS = model.Data.show_fytd == 'yes'
            
        if hasattr(model.Data, 'show_fytd_change'):
            global SHOW_FYTD_CHANGE_COLUMNS
            SHOW_FYTD_CHANGE_COLUMNS = model.Data.show_fytd_change == 'yes'
            
        if hasattr(model.Data, 'show_fytd_avg'):
            global SHOW_FYTD_AVG_COLUMN
            SHOW_FYTD_AVG_COLUMN = model.Data.show_fytd_avg == 'yes'
    
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
            
    def get_division_enrollment(self, division_id, report_date):
        """
        Get enrollment count for a division as of the report date.
        
        This counts people enrolled in organizations in this division
        as of the report date.
        """
        # Format date correctly
        if hasattr(report_date, 'strftime'):
            date_str = report_date.strftime('%Y-%m-%d')
        else:
            date_str = str(report_date)
        
        # Simple SQL query to count enrolled members
        sql = """
            SELECT COUNT(DISTINCT om.PeopleId) AS EnrolledCount
            FROM OrganizationMembers om
            JOIN OrganizationStructure os ON os.OrgId = om.OrganizationId
            WHERE os.DivId = {}
            AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{}')
            AND (om.InactiveDate IS NULL OR om.InactiveDate > '{}')
        """.format(division_id, date_str, date_str)
            
        # Print the SQL for debugging
        #print('Division {} enrollment SQL: {}'.format(division_id, sql))

        try:
            # Execute query and return result
            result = q.QuerySqlTop1(sql)
            if result and hasattr(result, 'EnrolledCount'):
                return result.EnrolledCount
            return 0
        except:
            # On any error, return 0
            return 0
    
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
    
    def get_four_week_attendance_sql(self, start_date, end_date, program_id=None, division_id=None, org_id=None):
        """Get SQL for attendance within a 4-week period."""
        # Format dates for SQL
        start_date_str = ReportHelper.format_date(start_date)
        end_date_str = ReportHelper.format_date(end_date)
        
        # Add time components to make the dates include the full day
        start_date_time_str = start_date_str + " 00:00:00"
        end_date_time_str = end_date_str + " 23:59:59"
        
        # Base query structure with conditions for meetings
        sql = """
            WITH FourWeekAttendance AS (
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
                SUM(AttendCount) as AttendanceCount,
                COUNT(DISTINCT MeetingId) as MeetingCount,
                COUNT(DISTINCT MeetingDate) as DaysWithMeetings
            FROM FourWeekAttendance
            GROUP BY ProgramName
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
    
    def get_specific_programs_sql(self, program_names):
        """
        Generate SQL to retrieve specific programs with robust error handling.
        
        Args:
            program_names (list): List of program names to search for
        
        Returns:
            SQL query string or empty string if no valid names provided
        """
        try:
            # Validate input
            if not program_names or not isinstance(program_names, list):
                print "<p>Warning: No valid program names provided</p>"
                return ""
            
            # Safely escape and quote program names
            safe_program_names = [
                "'{}'".format(str(name).replace("'", "''")) 
                for name in program_names 
                if name  # Ensure non-empty
            ]
            
            # Ensure we have at least one valid name
            if not safe_program_names:
                print "<p>Warning: No safe program names after sanitization</p>"
                return ""
            
            # Construct the SQL query with additional safety checks
            sql = '''
            SELECT 
                Id, 
                Name, 
                COALESCE(RptGroup, '') AS RptGroup 
            FROM Program 
            WHERE 
                Name IN ({}) 
                AND RptGroup IS NOT NULL 
                AND RptGroup <> ''
            '''.format(", ".join(safe_program_names))
            
            return sql
        
        except Exception as e:
            print "<p>Error in get_specific_programs_sql: {}</p>".format(e)
            import traceback
            traceback.print_exc()
            return ""
        
    def get_prospect_conversion_sql(self, program_id=None):
        """Get SQL to analyze prospect conversion rates including attendance history."""
        # Add program filter if specified
        program_filter = ""
        if program_id:
            program_filter = """
            AND EXISTS (
                SELECT 1 
                FROM OrganizationStructure os 
                WHERE os.OrgId = o.OrganizationId 
                AND EXISTS (
                    SELECT 1 
                    FROM Division d 
                    WHERE d.Id = os.DivId 
                    AND d.ProgId = {}
                )
            )""".format(program_id)
        
        return """
        WITH ProspectHistory AS (
            SELECT 
                p.PeopleId,
                p.Name,
                om.OrganizationId,
                o.OrganizationName,
                MIN(om.EnrollmentDate) as FirstProspectDate,
                -- Check both membership change AND attendance
                CASE 
                    WHEN EXISTS (
                        SELECT 1 
                        FROM OrganizationMembers om2 
                        WHERE om2.PeopleId = p.PeopleId
                        AND om2.MemberTypeId != {0}
                        AND om2.OrganizationId = om.OrganizationId
                        AND om2.EnrollmentDate >= om.EnrollmentDate
                    ) 
                    OR EXISTS (
                        -- Look for consistent attendance after becoming a prospect
                        SELECT 1
                        FROM Attend a 
                        WHERE a.PeopleId = p.PeopleId
                        AND a.OrganizationId = om.OrganizationId
                        AND a.MeetingDate >= om.EnrollmentDate
                        AND a.AttendanceFlag = 1  -- Actually attended
                        GROUP BY a.PeopleId
                        HAVING COUNT(*) >= 3  -- Consider "converted" if attended 3+ times
                    )
                    THEN 1 
                    ELSE 0 
                END as Converted,
                -- Add attendance metrics
                (SELECT COUNT(*) 
                 FROM Attend a
                 WHERE a.PeopleId = p.PeopleId 
                 AND a.OrganizationId = om.OrganizationId
                 AND a.MeetingDate >= om.EnrollmentDate
                 AND a.AttendanceFlag = 1) as TimesAttended
            FROM People p
            JOIN OrganizationMembers om ON om.PeopleId = p.PeopleId 
            JOIN Organizations o ON o.OrganizationId = om.OrganizationId
            WHERE om.MemberTypeId = {0}
            AND om.EnrollmentDate >= DATEADD(WEEK, -{1}, GETDATE())
            {2}
            GROUP BY 
                p.PeopleId,
                p.Name,
                om.OrganizationId,
                o.OrganizationName,
                om.EnrollmentDate
        )
        SELECT 
            OrganizationId,
            OrganizationName,
            COUNT(DISTINCT PeopleId) as TotalProspects,
            SUM(Converted) as Conversions,
            CASE 
                WHEN COUNT(DISTINCT PeopleId) > 0 
                THEN (CAST(SUM(Converted) as FLOAT) / COUNT(DISTINCT PeopleId)) * 100 
                ELSE 0 
            END as ConversionRate,
            AVG(CAST(TimesAttended as FLOAT)) as AvgAttendance
        FROM ProspectHistory
        GROUP BY OrganizationId, OrganizationName
        ORDER BY OrganizationName
        """.format(
            PROSPECT_MEMBER_TYPE_ID,
            PROSPECT_CONVERSION_WEEKS,
            program_filter
        )
    
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
    def get_four_week_attendance_data(self, start_date, end_date, program_id=None, division_id=None, org_id=None):
        """Get attendance data for a 4-week period."""
        four_week_data = {
            'total': 0,
            'meetings': 0,
            'days': 0,
            'by_program': {},
            'start_date': start_date,
            'end_date': end_date
        }
        
        # If this is a program-specific query, adjust end date based on program offsets
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
                
                # For 4-week period, we could adjust end date based on program offsets if needed
                _, adjusted_end_date = self.get_program_specific_date_range(
                    program_info, start_date, end_date
                )
                # Only update end_date for calculation
                four_week_data['end_date'] = adjusted_end_date
                end_date = adjusted_end_date
        
        # Now call the SQL with adjusted dates
        sql = self.get_four_week_attendance_sql(start_date, end_date, program_id, division_id, org_id)
        
        try:
            results = q.QuerySql(sql)
            for row in results:
                four_week_data['total'] += row.AttendanceCount or 0
                four_week_data['meetings'] += row.MeetingCount or 0
                four_week_data['days'] += row.DaysWithMeetings or 0
                
                # Track by program
                program_name = row.ProgramName
                if program_name not in four_week_data['by_program']:
                    four_week_data['by_program'][program_name] = {
                        'total': 0,
                        'meetings': 0,
                        'days': 0
                    }
                
                four_week_data['by_program'][program_name]['total'] += row.AttendanceCount or 0
                four_week_data['by_program'][program_name]['meetings'] += row.MeetingCount or 0
                four_week_data['by_program'][program_name]['days'] += row.DaysWithMeetings or 0
            
            # Calculate average attendance if we have meetings
            if four_week_data['meetings'] > 0:
                four_week_data['avg_per_meeting'] = float(four_week_data['total']) / four_week_data['meetings']
            else:
                four_week_data['avg_per_meeting'] = 0
                
            # Calculate average per day if we have days with meetings
            if four_week_data['days'] > 0:
                four_week_data['avg_per_day'] = float(four_week_data['total']) / four_week_data['days']
            else:
                four_week_data['avg_per_day'] = 0
                
            # Calculate program-specific averages
            for program_name in four_week_data['by_program']:
                prog_data = four_week_data['by_program'][program_name]
                
                if prog_data['meetings'] > 0:
                    prog_data['avg_per_meeting'] = float(prog_data['total']) / prog_data['meetings']
                else:
                    prog_data['avg_per_meeting'] = 0
                    
                if prog_data['days'] > 0:
                    prog_data['avg_per_day'] = float(prog_data['total']) / prog_data['days']
                else:
                    prog_data['avg_per_day'] = 0
                    
        except Exception as e:
            print "<p>Error getting 4-week attendance data: {}</p>".format(e)
            
        return four_week_data
        
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
        
        # Initialize data structure for all years with zeros
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            years_data[year] = {
                'total': 0,
                'meetings': 0, 
                'by_hour': {},
                'by_program': {}
            }
        
        # Get current year's data
        current_year_data = self.get_week_attendance_data(
            self.week_start_date, 
            self.report_date, 
            program_id, 
            division_id, 
            org_id
        )
        years_data[self.current_year].update(current_year_data)
        
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
            years_data[prev_year].update(prev_year_data)
        
        performance_timer.log("multiple_years_attendance_data_{}".format(org_id or division_id or program_id))
        return years_data
    
    def get_four_week_attendance_comparison(self, program_id=None, division_id=None, org_id=None):
        """Get attendance data for current and previous year's 4-week periods."""
        four_week_data = {}
        
        # Get current 4-week period data
        current_four_week_data = self.get_four_week_attendance_data(
            self.four_weeks_ago_start_date,
            self.report_date,
            program_id,
            division_id,
            org_id
        )
        four_week_data['current'] = current_four_week_data
        
        # Get previous year's 4-week period data
        prev_year_four_week_data = self.get_four_week_attendance_data(
            self.prev_year_four_weeks_ago_start_date,
            self.previous_year_date,
            program_id,
            division_id,
            org_id
        )
        four_week_data['previous_year'] = prev_year_four_week_data
        
        return four_week_data
    
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
        
        # If no service times are defined, use "Total" only
        if not service_times:
            return ["Total"]
            
        # Process any combined times
        processed_times = []
        has_total = False
        
        for time in service_times:
            if time == "Total":
                has_total = True
            # Check for combined time format
            combined_match = re.match(r'\((.*?)\)=(.*?)$', time)
            if combined_match:
                processed_times.append(combined_match.group(2).strip())
            else:
                processed_times.append(time.strip())
        
        # Always ensure Total is included
        if not has_total:
            processed_times.append("Total")
            
        return processed_times
    
    def debug_sql(self, sql, description="SQL Query"):
        """Print SQL for debugging."""
        print "<div style='background-color: #f5f5f5; padding: 10px; margin: 10px 0; border-radius: 5px;'>"
        print "<strong>DEBUG - {}:</strong>".format(description)
        print "<pre style='white-space: pre-wrap;'>{}</pre>".format(sql)
        print "</div>"
    
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
    
    def generate_organization_row(self, org, years_data, years_ytd_data, four_week_data=None, indent=2):
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
        
        # Add attendance data for each year if showing current week column
        if SHOW_CURRENT_WEEK_COLUMN:
            for i in range(YEARS_TO_DISPLAY):
                year = self.current_year - i
                if year in years_data:
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[year]['total'])
                    )
                else:
                    row_html += "<td>0</td>"
                    
        if SHOW_PREVIOUS_YEAR_COLUMN:
            prev_year = self.current_year - 1
            if prev_year in years_data:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(years_data[prev_year]['total'])
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
        
        # Add year-over-year change columns if enabled
        if SHOW_YOY_CHANGE_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                current_year = self.current_year - i + 1
                prev_year = self.current_year - i
                
                current_total = years_data[current_year]['total'] if current_year in years_data else 0
                prev_total = years_data[prev_year]['total'] if prev_year in years_data else 0
                
                trend = ReportHelper.get_trend_indicator(current_total, prev_total)
                row_html += "<td>{}</td>".format(trend)
        
        # Add 4-week comparison columns if enabled
        if SHOW_FOUR_WEEK_COMPARISON and four_week_data:
            current_four_week_total = four_week_data['current']['total']
            prev_year_four_week_total = four_week_data['previous_year']['total']
            
            # Add current 4-week total
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(current_four_week_total)
            )
            
            # Add previous year's 4-week total
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(prev_year_four_week_total)
            )
            
            # Add 4-week trend indicator
            four_week_trend = ReportHelper.get_trend_indicator(current_four_week_total, prev_year_four_week_total)
            row_html += "<td>{}</td>".format(four_week_trend)
        
        # Add YTD data for each year if enabled
        if SHOW_FYTD_COLUMNS:
            for i in range(YEARS_TO_DISPLAY):
                year = self.current_year - i
                if year in years_ytd_data:
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_ytd_data[year]['total'])
                    )
                else:
                    row_html += "<td>0</td>"
        
        # Add YTD change columns if enabled
        if SHOW_FYTD_CHANGE_COLUMNS:
            for i in range(1, YEARS_TO_DISPLAY):
                current_year = self.current_year - i + 1
                prev_year = self.current_year - i
                
                current_ytd = years_ytd_data[current_year]['total'] if current_year in years_ytd_data else 0
                prev_ytd = years_ytd_data[prev_year]['total'] if prev_year in years_ytd_data else 0
                
                ytd_trend = ReportHelper.get_trend_indicator(current_ytd, prev_ytd)
                row_html += "<td>{}</td>".format(ytd_trend)
        
        # Add YTD Average if enabled
        if SHOW_FYTD_AVG_COLUMN:
            # Using weeks elapsed instead of days with meetings
            weeks_elapsed = ReportHelper.get_weeks_elapsed_in_fiscal_year(self.report_date, self.current_fiscal_start)
            avg_attendance = 0
            if weeks_elapsed > 0 and self.current_year in years_ytd_data:
                avg_attendance = years_ytd_data[self.current_year]['total'] / float(weeks_elapsed)
            row_html += "<td>{}</td>".format(ReportHelper.format_float(avg_attendance))
        
        row_html += "</tr>"
        return row_html
        
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
        
        # Set checked states based on current settings
        show_program_summary_checked = "checked" if SHOW_PROGRAM_SUMMARY else ""
        show_four_week_comparison_checked = "checked" if SHOW_FOUR_WEEK_COMPARISON else ""
        show_current_week_checked = "checked" if SHOW_CURRENT_WEEK_COLUMN else ""
        show_previous_year_checked = "checked" if SHOW_PREVIOUS_YEAR_COLUMN else ""
        show_yoy_change_checked = "checked" if SHOW_YOY_CHANGE_COLUMN else ""
        show_fytd_checked = "checked" if SHOW_FYTD_COLUMNS else ""
        show_fytd_change_checked = "checked" if SHOW_FYTD_CHANGE_COLUMNS else ""
        show_fytd_avg_checked = "checked" if SHOW_FYTD_AVG_COLUMN else ""
        show_detailed_enrollment_checked = "checked" if SHOW_DETAILED_ENROLLMENT else ""
        
        form_html = """
        <form method="get" action="" id="report-form" style="margin-bottom: 20px; padding: 10px; background-color: #f5f5f5; border-radius: 5px;">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 10px; flex-wrap: wrap;">
                <label for="report_date">Report Date (Sunday):</label>
                <input type="date" id="report_date" name="report_date" value="{0:}" required>
                
                <button type="submit" id="run-report-btn" style="padding: 5px 10px; background-color: #4CAF50; color: white; border: none; border-radius: 3px; cursor: pointer; position: relative;">
                    Run Report
                </button>
            </div>
            
            <div style="margin-top: 10px;">
                <label>
                    <input type="checkbox" name="collapse_orgs" value="no" {1:}>
                    Expand Involvements (Default is Collapsed).. Note: This is slow.. 3+ minutes slow at times.
                </label>
            </div>
            
            <div style="margin-top: 15px;">
                <details>
                    <summary style="cursor: pointer; padding: 5px; background-color: #f0f0f0; border-radius: 3px;">Advanced Display Options</summary>
                    <div style="margin-top: 10px; padding: 10px; border: 1px solid #ddd; border-radius: 3px; background-color: #f9f9f9;">
                        <h4 style="margin-top: 0;">Report Content</h4>
                        <div>
                            <label>
                                <input type="checkbox" name="show_program_summary" value="yes" {2:}>
                                Show Program Summary Section
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_four_week" value="yes" {3:}>
                                Show 4-Week Comparison Columns
                            </label>
                        </div>
                        
                        <h4 style="margin-top: 15px;">Column Visibility</h4>
                        <div>
                            <label>
                                <input type="checkbox" name="show_current_week" value="yes" {4:}>
                                Show Current Week Column
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_previous_year" value="yes" {5:}>
                                Show Previous Year Column
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_yoy_change" value="yes" {6:}>
                                Show Year-over-Year Change Columns
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_fytd" value="yes" {7:}>
                                Show {8:} Columns
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_fytd_change" value="yes" {9:}>
                                Show {10:} Change Columns
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_fytd_avg" value="yes" {11:}>
                                Show {12:} Avg/Week Column
                            </label>
                        </div>
                        <div style="margin-top: 5px;">
                            <label>
                                <input type="checkbox" name="show_detailed_enrollment" value="yes" {14:}>
                                Show Detailed Enrollment Analysis
                            </label>
                        </div>
                    </div>
                </details>
            </div>
            
            {13:}
        </form>
        
        <script>
        // Store form settings in localStorage when form is submitted
        document.getElementById('report-form').addEventListener('submit', function(e) {{
            var checkboxes = this.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(function(checkbox) {{
                if (checkbox.name && checkbox.name !== 'collapse_orgs') {{
                    localStorage.setItem('attendance_' + checkbox.name, checkbox.checked ? 'yes' : 'no');
                }}
            }});
        }});
        
        // Load saved settings when page loads
        document.addEventListener('DOMContentLoaded', function() {{
            var checkboxes = document.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(function(checkbox) {{
                if (checkbox.name && checkbox.name !== 'collapse_orgs' && !checkbox.disabled) {{
                    var savedValue = localStorage.getItem('attendance_' + checkbox.name);
                    if (savedValue === 'yes') {{
                        checkbox.checked = true;
                    }} else if (savedValue === 'no') {{
                        checkbox.checked = false;
                    }}
                }}
            }});
        }});
        </script>
        """
        
        # Format with numbered placeholders
        return form_html.format(
            current_date,                         # {0:}
            expand_checked,                       # {1:}
            show_program_summary_checked,         # {2:}
            show_four_week_comparison_checked,    # {3:}
            show_current_week_checked,            # {4:}
            show_previous_year_checked,           # {5:}
            show_yoy_change_checked,              # {6:}
            show_fytd_checked,                    # {7:}
            ReportHelper.get_ytd_label(),         # {8:}
            show_fytd_change_checked,             # {9:}
            ReportHelper.get_ytd_label(),         # {10:}
            show_fytd_avg_checked,                # {11:}
            ReportHelper.get_ytd_label(),         # {12:}
            send_email_html,                      # {13:}
            show_detailed_enrollment_checked      # {14:}  # Add this line
        )
        
        return form_html
        
    def get_enrollment_ratio_sql(self, program_id=None, division_id=None, org_id=None):
        """Get SQL to calculate enrollment to attendance ratio over rolling window."""
        
        # Get date parameters
        selected_date = model.Data.report_date if hasattr(model.Data, 'report_date') and model.Data.report_date else model.DateTime
        weeks_ago_date = model.DateAddDays(selected_date, -7 * ENROLLMENT_RATIO_WEEKS)
        
        # Format dates for SQL 
        current_date_sql = str(selected_date).split(' ')[0] if ' ' in str(selected_date) else str(selected_date)
        weeks_ago_date_sql = str(weeks_ago_date).split(' ')[0] if ' ' in str(weeks_ago_date) else str(weeks_ago_date)
        
        # Find the most recent Sunday (for LastSundayDate)
        last_sunday_date = model.SundayForDate(selected_date)
        last_sunday_sql = str(last_sunday_date).split(' ')[0] if ' ' in str(last_sunday_date) else str(last_sunday_date)
        
        # Build organization/division/program filters
        org_filter = "AND o.OrganizationId = {}".format(org_id) if org_id else ""
        division_filter = "AND d.Id = {}".format(division_id) if division_id else ""
        program_id_filter = "AND p.Id = {}".format(program_id) if program_id else ""
        
        # Get program insertion SQL
        program_insert_sql = self._build_program_insert_sql()
        
        # Format SQL
        sql = """
        -- Define variables at the top
        DECLARE @WeeksCount INT = {weeks_count}; -- Number of weeks between start and end dates
        DECLARE @StartDate DATE = '{start_date}'; -- Start date
        DECLARE @EndDate DATE = '{end_date}'; -- End date
        DECLARE @LastSundayDate DATE = '{last_sunday_date}'; -- Last Sunday date
        DECLARE @NeedsInReachThreshold INT = {needs_inreach}; -- Threshold for "Needs In-reach" category
        DECLARE @GoodRatioThreshold INT = {good_ratio}; -- Threshold for "Good Ratio" category
        DECLARE @IncludeOrgDetails BIT = {include_org_details}; -- Set to 1 to include Organization name and ID in results
    
        -- Temporary table for program names
        CREATE TABLE #ProgramNames (Name NVARCHAR(100));
    
        -- Insert program names to include
        {program_insert}
    
        -- Create temporary table for base organization data
        CREATE TABLE #OrganizationBase (
            OrganizationId INT,
            OrganizationName NVARCHAR(255),
            Enrollment INT,
            ProgramName NVARCHAR(100),
            ProgramId INT,
            DivisionName NVARCHAR(100),
            DivisionId INT
        );
    
        -- Populate the organization base table
        INSERT INTO #OrganizationBase
        SELECT DISTINCT
            o.OrganizationId,
            o.OrganizationName,
            o.MemberCount AS Enrollment,
            p.Name AS ProgramName,
            p.Id AS ProgramId,
            d.Name AS DivisionName,
            d.Id AS DivisionId
        FROM 
            Organizations o
            JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
            JOIN Division d ON d.Id = os.DivId
            JOIN Program p ON p.Id = d.ProgId AND p.Name IN (SELECT Name FROM #ProgramNames)
        WHERE 
            o.OrganizationStatusId = 30
            {org_filter}
            {division_filter}
            {program_id_filter};
    
        -- Create temporary table for organization meetings
        CREATE TABLE #OrganizationMeetings (
            OrganizationId INT,
            OrganizationName NVARCHAR(255),
            Enrollment INT,
            ProgramName NVARCHAR(100),
            ProgramId INT,
            DivisionName NVARCHAR(100),
            DivisionId INT,
            TotalAttendance FLOAT,
            AvgAttendance FLOAT
        );
    
        -- Populate the organization meetings table
        INSERT INTO #OrganizationMeetings
        SELECT 
            ob.OrganizationId,
            ob.OrganizationName,
            ob.Enrollment,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId,
            SUM(CAST(COALESCE(m.MaxCount, 0) AS FLOAT)) AS TotalAttendance,
            AVG(CAST(COALESCE(m.MaxCount, 0) AS FLOAT)) AS AvgAttendance
        FROM 
            #OrganizationBase ob
            JOIN Meetings m ON m.OrganizationId = ob.OrganizationId
        WHERE
            (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL)
            AND m.MeetingDate BETWEEN @StartDate AND @EndDate
        GROUP BY 
            ob.OrganizationId, 
            ob.OrganizationName,
            ob.Enrollment,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId;
    
        -- Create temporary table for last Sunday meetings
        CREATE TABLE #LastSundayMeetings (
            OrganizationId INT,
            OrganizationName NVARCHAR(255),
            Enrollment INT,
            ProgramName NVARCHAR(100),
            ProgramId INT,
            DivisionName NVARCHAR(100),
            DivisionId INT,
            LastSundayAttendance FLOAT
        );
    
        -- Populate the last Sunday meetings table
        INSERT INTO #LastSundayMeetings
        SELECT 
            ob.OrganizationId,
            ob.OrganizationName,
            ob.Enrollment,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId,
            SUM(CAST(COALESCE(m.MaxCount, 0) AS FLOAT)) AS LastSundayAttendance
        FROM 
            #OrganizationBase ob
            JOIN Meetings m ON m.OrganizationId = ob.OrganizationId
        WHERE
            (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL)
            AND CONVERT(date, m.MeetingDate) = @LastSundayDate
        GROUP BY 
            ob.OrganizationId, 
            ob.OrganizationName,
            ob.Enrollment,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId;
    
        -- Create temporary table for organization ratios
        CREATE TABLE #OrganizationRatios (
            OrganizationId INT,
            OrganizationName NVARCHAR(255),
            ProgramName NVARCHAR(100),
            ProgramId INT,
            DivisionName NVARCHAR(100),
            DivisionId INT,
            Enrollment INT,
            RatioCategory NVARCHAR(50),
            LastSundayRatioCategory NVARCHAR(50)
        );
    
        -- Populate the organization ratios table
        INSERT INTO #OrganizationRatios
        SELECT
            ob.OrganizationId,
            ob.OrganizationName,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId,
            ob.Enrollment,
            
            -- Regular period ratio category
            CASE
                WHEN ob.Enrollment = 0 OR COALESCE(om.TotalAttendance, 0) = 0 THEN 'No Data'
                WHEN (COALESCE(om.TotalAttendance, 0) / @WeeksCount) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN (COALESCE(om.TotalAttendance, 0) / @WeeksCount) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END AS RatioCategory,
            
            -- Last Sunday ratio category
            CASE
                WHEN ob.Enrollment = 0 OR COALESCE(lsm.LastSundayAttendance, 0) = 0 THEN 'No Data'
                WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END AS LastSundayRatioCategory
        FROM 
            #OrganizationBase ob
            LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
            LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId;
    
        -- Create temporary table for program category counts
        CREATE TABLE #ProgramCategoryCounts (
            ProgramId INT,
            ProgramName NVARCHAR(100),
            NeedsInReachCount INT,
            GoodRatioCount INT,
            NeedsOutreachCount INT,
            LastSundayNeedsInReachCount INT,
            LastSundayGoodRatioCount INT,
            LastSundayNeedsOutreachCount INT
        );
    
        -- Populate the program category counts table
        INSERT INTO #ProgramCategoryCounts
        SELECT
            ProgramId,
            ProgramName,
            SUM(CASE WHEN RatioCategory = 'Needs In-reach' THEN 1 ELSE 0 END) AS NeedsInReachCount,
            SUM(CASE WHEN RatioCategory = 'Good Ratio' THEN 1 ELSE 0 END) AS GoodRatioCount,
            SUM(CASE WHEN RatioCategory = 'Needs Outreach' THEN 1 ELSE 0 END) AS NeedsOutreachCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Needs In-reach' THEN 1 ELSE 0 END) AS LastSundayNeedsInReachCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Good Ratio' THEN 1 ELSE 0 END) AS LastSundayGoodRatioCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Needs Outreach' THEN 1 ELSE 0 END) AS LastSundayNeedsOutreachCount
        FROM 
            #OrganizationRatios
        GROUP BY
            ProgramId, ProgramName;
    
        -- Create temporary table for division category counts
        CREATE TABLE #DivisionCategoryCounts (
            DivisionId INT,
            DivisionName NVARCHAR(100),
            ProgramId INT,
            ProgramName NVARCHAR(100),
            NeedsInReachCount INT,
            GoodRatioCount INT,
            NeedsOutreachCount INT,
            LastSundayNeedsInReachCount INT,
            LastSundayGoodRatioCount INT,
            LastSundayNeedsOutreachCount INT
        );
    
        -- Populate the division category counts table
        INSERT INTO #DivisionCategoryCounts
        SELECT
            DivisionId,
            DivisionName,
            ProgramId,
            ProgramName,
            SUM(CASE WHEN RatioCategory = 'Needs In-reach' THEN 1 ELSE 0 END) AS NeedsInReachCount,
            SUM(CASE WHEN RatioCategory = 'Good Ratio' THEN 1 ELSE 0 END) AS GoodRatioCount,
            SUM(CASE WHEN RatioCategory = 'Needs Outreach' THEN 1 ELSE 0 END) AS NeedsOutreachCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Needs In-reach' THEN 1 ELSE 0 END) AS LastSundayNeedsInReachCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Good Ratio' THEN 1 ELSE 0 END) AS LastSundayGoodRatioCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Needs Outreach' THEN 1 ELSE 0 END) AS LastSundayNeedsOutreachCount
        FROM 
            #OrganizationRatios
        GROUP BY
            DivisionId, DivisionName, ProgramId, ProgramName;
    
        -- Create temporary table for overall category counts
        CREATE TABLE #OverallCategoryCounts (
            NeedsInReachCount INT,
            GoodRatioCount INT,
            NeedsOutreachCount INT,
            LastSundayNeedsInReachCount INT,
            LastSundayGoodRatioCount INT,
            LastSundayNeedsOutreachCount INT
        );
    
        -- Populate the overall category counts table
        INSERT INTO #OverallCategoryCounts
        SELECT
            SUM(CASE WHEN RatioCategory = 'Needs In-reach' THEN 1 ELSE 0 END) AS NeedsInReachCount,
            SUM(CASE WHEN RatioCategory = 'Good Ratio' THEN 1 ELSE 0 END) AS GoodRatioCount,
            SUM(CASE WHEN RatioCategory = 'Needs Outreach' THEN 1 ELSE 0 END) AS NeedsOutreachCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Needs In-reach' THEN 1 ELSE 0 END) AS LastSundayNeedsInReachCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Good Ratio' THEN 1 ELSE 0 END) AS LastSundayGoodRatioCount,
            SUM(CASE WHEN LastSundayRatioCategory = 'Needs Outreach' THEN 1 ELSE 0 END) AS LastSundayNeedsOutreachCount
        FROM 
            #OrganizationRatios;
    
        -- Create temporary table for overall summary
        CREATE TABLE #OverallSummary (
            TotalEnrollment INT,
            EnrollmentWithMeetings INT,
            TotalAttendance FLOAT,
            AvgAttendance FLOAT,
            LastSundayEnrollment INT,
            LastSundayAttendance FLOAT
        );
    
        -- Populate the overall summary table
        INSERT INTO #OverallSummary
        SELECT
            -- Enrollment for ALL qualifying organizations
            SUM(ob.Enrollment) AS TotalEnrollment,
            
            -- Enrollment ONLY for organizations that had meetings
            SUM(CASE WHEN om.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END) AS EnrollmentWithMeetings,
            
            -- Attendance (only for those with meetings)
            SUM(COALESCE(om.TotalAttendance, 0)) AS TotalAttendance,
            
            -- Average attendance
            AVG(COALESCE(om.AvgAttendance, 0)) AS AvgAttendance,
            
            -- Last Sunday's enrollment (for groups that met)
            SUM(CASE WHEN lsm.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END) AS LastSundayEnrollment,
            
            -- Last Sunday's attendance
            SUM(COALESCE(lsm.LastSundayAttendance, 0)) AS LastSundayAttendance
        FROM 
            #OrganizationBase ob
            LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
            LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId;
    
        -- Create temporary table for program summary
        CREATE TABLE #ProgramSummary (
            ProgramId INT,
            ProgramName NVARCHAR(100),
            TotalEnrollment INT,
            EnrollmentWithMeetings INT,
            TotalAttendance FLOAT,
            AvgAttendance FLOAT,
            LastSundayEnrollment INT,
            LastSundayAttendance FLOAT
        );
    
        -- Populate the program summary table
        INSERT INTO #ProgramSummary
        SELECT
            ob.ProgramId,
            ob.ProgramName,
            
            -- Enrollment for ALL qualifying organizations in this program
            SUM(ob.Enrollment) AS TotalEnrollment,
            
            -- Enrollment ONLY for organizations that had meetings
            SUM(CASE WHEN om.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END) AS EnrollmentWithMeetings,
            
            -- Attendance (only for those with meetings)
            SUM(COALESCE(om.TotalAttendance, 0)) AS TotalAttendance,
            
            -- Average attendance
            AVG(COALESCE(om.AvgAttendance, 0)) AS AvgAttendance,
            
            -- Last Sunday's enrollment (for groups that met)
            SUM(CASE WHEN lsm.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END) AS LastSundayEnrollment,
            
            -- Last Sunday's attendance
            SUM(COALESCE(lsm.LastSundayAttendance, 0)) AS LastSundayAttendance
        FROM 
            #OrganizationBase ob
            LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
            LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId
        GROUP BY
            ob.ProgramId,
            ob.ProgramName;
    
        -- Create temporary table for division summary
        CREATE TABLE #DivisionSummary (
            DivisionId INT,
            DivisionName NVARCHAR(100),
            ProgramId INT,
            ProgramName NVARCHAR(100),
            TotalEnrollment INT,
            EnrollmentWithMeetings INT,
            TotalAttendance FLOAT,
            AvgAttendance FLOAT,
            LastSundayEnrollment INT,
            LastSundayAttendance FLOAT
        );
    
        -- Populate the division summary table
        INSERT INTO #DivisionSummary
        SELECT
            ob.DivisionId,
            ob.DivisionName,
            ob.ProgramId,
            ob.ProgramName,
            
            -- Enrollment for ALL qualifying organizations in this division
            SUM(ob.Enrollment) AS TotalEnrollment,
            
            -- Enrollment ONLY for organizations that had meetings
            SUM(CASE WHEN om.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END) AS EnrollmentWithMeetings,
            
            -- Attendance (only for those with meetings)
            SUM(COALESCE(om.TotalAttendance, 0)) AS TotalAttendance,
            
            -- Average attendance
            AVG(COALESCE(om.AvgAttendance, 0)) AS AvgAttendance,
            
            -- Last Sunday's enrollment (for groups that met)
            SUM(CASE WHEN lsm.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END) AS LastSundayEnrollment,
            
            -- Last Sunday's attendance
            SUM(COALESCE(lsm.LastSundayAttendance, 0)) AS LastSundayAttendance
        FROM 
            #OrganizationBase ob
            LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
            LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId
        GROUP BY
            ob.DivisionId,
            ob.DivisionName,
            ob.ProgramId,
            ob.ProgramName;
    
        -- Create temporary table for results to properly handle ordering
        CREATE TABLE #Results (
            Level NVARCHAR(100),
            DivisionName NVARCHAR(100) NULL,
            ProgramName NVARCHAR(100),
            OrganizationName NVARCHAR(255) NULL,
            LevelSortOrder INT,
            CategorySortOrder INT,
            ResultOrder INT IDENTITY(1,1),
            TotalEnrollment INT NULL,
            EnrollmentWithMeetings INT NULL,
            TotalAttendance FLOAT NULL,
            AvgAttendance FLOAT NULL,
            LastSundayEnrollment INT NULL,
            LastSundayAttendance FLOAT NULL,
            AttendanceRatio FLOAT NULL,
            RatioCategory NVARCHAR(50) NULL,
            LastSundayAttendanceRatio FLOAT NULL,
            LastSundayRatioCategory NVARCHAR(50) NULL,
            NeedsInReachCount INT NULL,
            GoodRatioCount INT NULL,
            NeedsOutreachCount INT NULL,
            LastSundayNeedsInReachCount INT NULL,
            LastSundayGoodRatioCount INT NULL,
            LastSundayNeedsOutreachCount INT NULL
        );
    
        -- Organization details temp table (for @IncludeOrgDetails = 1)
        IF @IncludeOrgDetails = 1
        BEGIN
            CREATE TABLE #OrganizationDetails (
                OrganizationId INT,
                OrganizationName NVARCHAR(255),
                DivisionId INT,
                DivisionName NVARCHAR(100),
                ProgramId INT,
                ProgramName NVARCHAR(100),
                TotalEnrollment INT,
                EnrollmentWithMeetings INT,
                TotalAttendance FLOAT,
                AvgAttendance FLOAT,
                LastSundayEnrollment INT,
                LastSundayAttendance FLOAT,
                AttendanceRatio FLOAT,
                RatioCategory NVARCHAR(50),
                LastSundayAttendanceRatio FLOAT,
                LastSundayRatioCategory NVARCHAR(50)
            );
            
            -- Populate the organization details table
            INSERT INTO #OrganizationDetails
            SELECT
                ob.OrganizationId,
                ob.OrganizationName,
                ob.DivisionId,
                ob.DivisionName,
                ob.ProgramId,
                ob.ProgramName,
                ob.Enrollment AS TotalEnrollment,
                CASE WHEN om.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END AS EnrollmentWithMeetings,
                COALESCE(om.TotalAttendance, 0) AS TotalAttendance,
                COALESCE(om.AvgAttendance, 0) AS AvgAttendance,
                CASE WHEN lsm.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END AS LastSundayEnrollment,
                COALESCE(lsm.LastSundayAttendance, 0) AS LastSundayAttendance,
                
                -- Organization Attendance Ratio (divide by @WeeksCount for multiple weeks)
                CASE 
                    WHEN COALESCE(om.TotalAttendance, 0) > 0 AND ob.Enrollment > 0
                    THEN (COALESCE(om.TotalAttendance, 0) / @WeeksCount) / ob.Enrollment * 100
                    ELSE 0 
                END AS AttendanceRatio,
                
                -- Organization Ratio Category (divide by @WeeksCount for multiple weeks)
                CASE
                    WHEN ob.Enrollment = 0 OR COALESCE(om.TotalAttendance, 0) = 0 THEN 'No Data'
                    WHEN (COALESCE(om.TotalAttendance, 0) / @WeeksCount) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                    WHEN (COALESCE(om.TotalAttendance, 0) / @WeeksCount) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                    ELSE 'Needs Outreach'
                END AS RatioCategory,
                
                -- Last Sunday's Attendance Ratio (NO division by @WeeksCount for single day)
                CASE 
                    WHEN COALESCE(lsm.LastSundayAttendance, 0) > 0 AND ob.Enrollment > 0
                    THEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100
                    ELSE 0 
                END AS LastSundayAttendanceRatio,
                
                -- Last Sunday's Ratio Category (NO division by @WeeksCount for single day)
                CASE
                    WHEN ob.Enrollment = 0 OR COALESCE(lsm.LastSundayAttendance, 0) = 0 THEN 'No Data'
                    WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                    WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                    ELSE 'Needs Outreach'
                END AS LastSundayRatioCategory
            FROM 
                #OrganizationBase ob
                LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
                LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId;
        END
    
        -- Populate results table - Overall Summary
        INSERT INTO #Results (
            Level, DivisionName, ProgramName, OrganizationName, 
            LevelSortOrder, CategorySortOrder,
            TotalEnrollment, EnrollmentWithMeetings, TotalAttendance, AvgAttendance,
            LastSundayEnrollment, LastSundayAttendance, AttendanceRatio, RatioCategory,
            LastSundayAttendanceRatio, LastSundayRatioCategory,
            NeedsInReachCount, GoodRatioCount, NeedsOutreachCount,
            LastSundayNeedsInReachCount, LastSundayGoodRatioCount, LastSundayNeedsOutreachCount
        )
        SELECT 
            'Overall', NULL, 'All Programs', NULL, 
            1, 1,
            os.TotalEnrollment,
            os.EnrollmentWithMeetings,
            os.TotalAttendance,
            os.AvgAttendance,
            os.LastSundayEnrollment,
            os.LastSundayAttendance,
            CASE 
                WHEN os.EnrollmentWithMeetings > 0 
                THEN (os.TotalAttendance / @WeeksCount) / NULLIF(os.EnrollmentWithMeetings, 0) * 100
                ELSE 0 
            END,
            CASE
                WHEN (os.TotalAttendance / @WeeksCount) / NULLIF(os.EnrollmentWithMeetings, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN (os.TotalAttendance / @WeeksCount) / NULLIF(os.EnrollmentWithMeetings, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END,
            CASE 
                WHEN os.LastSundayEnrollment > 0 
                THEN os.LastSundayAttendance / NULLIF(os.LastSundayEnrollment, 0) * 100
                ELSE 0 
            END,
            CASE
                WHEN os.LastSundayAttendance / NULLIF(os.LastSundayEnrollment, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN os.LastSundayAttendance / NULLIF(os.LastSundayEnrollment, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END,
            occ.NeedsInReachCount,
            occ.GoodRatioCount,
            occ.NeedsOutreachCount,
            occ.LastSundayNeedsInReachCount,
            occ.LastSundayGoodRatioCount,
            occ.LastSundayNeedsOutreachCount
        FROM 
            #OverallSummary os
            CROSS JOIN #OverallCategoryCounts occ;
    
        -- Populate results table - Program Summary
        INSERT INTO #Results (
            Level, DivisionName, ProgramName, OrganizationName, 
            LevelSortOrder, CategorySortOrder,
            TotalEnrollment, EnrollmentWithMeetings, TotalAttendance, AvgAttendance,
            LastSundayEnrollment, LastSundayAttendance, AttendanceRatio, RatioCategory,
            LastSundayAttendanceRatio, LastSundayRatioCategory,
            NeedsInReachCount, GoodRatioCount, NeedsOutreachCount,
            LastSundayNeedsInReachCount, LastSundayGoodRatioCount, LastSundayNeedsOutreachCount
        )
        SELECT
            'Program', NULL, ps.ProgramName, NULL, 
            2, 1,
            ps.TotalEnrollment,
            ps.EnrollmentWithMeetings,
            ps.TotalAttendance,
            ps.AvgAttendance,
            ps.LastSundayEnrollment,
            ps.LastSundayAttendance,
            CASE 
                WHEN ps.EnrollmentWithMeetings > 0 
                THEN (ps.TotalAttendance / @WeeksCount) / NULLIF(ps.EnrollmentWithMeetings, 0) * 100
                ELSE 0 
            END,
            CASE
                WHEN (ps.TotalAttendance / @WeeksCount) / NULLIF(ps.EnrollmentWithMeetings, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN (ps.TotalAttendance / @WeeksCount) / NULLIF(ps.EnrollmentWithMeetings, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END,
            CASE 
                WHEN ps.LastSundayEnrollment > 0 
                THEN ps.LastSundayAttendance / NULLIF(ps.LastSundayEnrollment, 0) * 100
                ELSE 0 
            END,
            CASE
                WHEN ps.LastSundayAttendance / NULLIF(ps.LastSundayEnrollment, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN ps.LastSundayAttendance / NULLIF(ps.LastSundayEnrollment, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END,
            pcc.NeedsInReachCount,
            pcc.GoodRatioCount,
            pcc.NeedsOutreachCount,
            pcc.LastSundayNeedsInReachCount,
            pcc.LastSundayGoodRatioCount,
            pcc.LastSundayNeedsOutreachCount
        FROM
            #ProgramSummary ps
            JOIN #ProgramCategoryCounts pcc ON ps.ProgramId = pcc.ProgramId;
    
        -- Populate results table - Division Summary
        INSERT INTO #Results (
            Level, DivisionName, ProgramName, OrganizationName, 
            LevelSortOrder, CategorySortOrder,
            TotalEnrollment, EnrollmentWithMeetings, TotalAttendance, AvgAttendance,
            LastSundayEnrollment, LastSundayAttendance, AttendanceRatio, RatioCategory,
            LastSundayAttendanceRatio, LastSundayRatioCategory,
            NeedsInReachCount, GoodRatioCount, NeedsOutreachCount,
            LastSundayNeedsInReachCount, LastSundayGoodRatioCount, LastSundayNeedsOutreachCount
        )
        SELECT
            'Division', ds.DivisionName, ds.ProgramName, NULL, 
            3, 1,
            ds.TotalEnrollment,
            ds.EnrollmentWithMeetings,
            ds.TotalAttendance,
            ds.AvgAttendance,
            ds.LastSundayEnrollment,
            ds.LastSundayAttendance,
            CASE 
                WHEN ds.EnrollmentWithMeetings > 0 
                THEN (ds.TotalAttendance / @WeeksCount) / NULLIF(ds.EnrollmentWithMeetings, 0) * 100
                ELSE 0 
            END,
            CASE
                WHEN (ds.TotalAttendance / @WeeksCount) / NULLIF(ds.EnrollmentWithMeetings, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN (ds.TotalAttendance / @WeeksCount) / NULLIF(ds.EnrollmentWithMeetings, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END,
            CASE 
                WHEN ds.LastSundayEnrollment > 0 
                THEN ds.LastSundayAttendance / NULLIF(ds.LastSundayEnrollment, 0) * 100
                ELSE 0 
            END,
            CASE
                WHEN ds.LastSundayAttendance / NULLIF(ds.LastSundayEnrollment, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach'
                WHEN ds.LastSundayAttendance / NULLIF(ds.LastSundayEnrollment, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio'
                ELSE 'Needs Outreach'
            END,
            dcc.NeedsInReachCount,
            dcc.GoodRatioCount,
            dcc.NeedsOutreachCount,
            dcc.LastSundayNeedsInReachCount,
            dcc.LastSundayGoodRatioCount,
            dcc.LastSundayNeedsOutreachCount
        FROM
            #DivisionSummary ds
            JOIN #DivisionCategoryCounts dcc ON ds.DivisionId = dcc.DivisionId AND ds.ProgramId = dcc.ProgramId;
    
        -- Add Organization Details if @IncludeOrgDetails = 1
        IF @IncludeOrgDetails = 1
        BEGIN
            -- Populate results table - Organization Details
            INSERT INTO #Results (
                Level, DivisionName, ProgramName, OrganizationName, 
                LevelSortOrder, CategorySortOrder,
                TotalEnrollment, EnrollmentWithMeetings, TotalAttendance, AvgAttendance,
                LastSundayEnrollment, LastSundayAttendance, AttendanceRatio, RatioCategory,
                LastSundayAttendanceRatio, LastSundayRatioCategory,
                NeedsInReachCount, GoodRatioCount, NeedsOutreachCount,
                LastSundayNeedsInReachCount, LastSundayGoodRatioCount, LastSundayNeedsOutreachCount
            )
            SELECT
                'Organization', od.DivisionName, od.ProgramName, od.OrganizationName, 
                4, 1,
                od.TotalEnrollment,
                od.EnrollmentWithMeetings,
                od.TotalAttendance,
                od.AvgAttendance,
                od.LastSundayEnrollment,
                od.LastSundayAttendance,
                od.AttendanceRatio,
                od.RatioCategory,
                od.LastSundayAttendanceRatio,
                od.LastSundayRatioCategory,
                NULL, NULL, NULL, NULL, NULL, NULL
            FROM
                #OrganizationDetails od;
            
            -- Category headers and details
            -- Regular period - Needs In-reach
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE RatioCategory = 'Needs In-reach')
            BEGIN
                -- Header
                INSERT INTO #Results (
                    Level, DivisionName, ProgramName, OrganizationName, 
                    LevelSortOrder, CategorySortOrder,
                    TotalEnrollment, EnrollmentWithMeetings, TotalAttendance, AvgAttendance,
                    LastSundayEnrollment, LastSundayAttendance, AttendanceRatio, RatioCategory,
                    LastSundayAttendanceRatio, LastSundayRatioCategory,
                    NeedsInReachCount, GoodRatioCount, NeedsOutreachCount,
                    LastSundayNeedsInReachCount, LastSundayGoodRatioCount, LastSundayNeedsOutreachCount
                )
                VALUES (
                    'Needs In-reach Organizations', NULL, 'Regular Period', NULL, 
                    5, 1,
                    NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL
                );
                
                -- Details
                INSERT INTO #Results (
                    Level, DivisionName, ProgramName, OrganizationName, 
                    LevelSortOrder, CategorySortOrder,
                    TotalEnrollment, EnrollmentWithMeetings, TotalAttendance, AvgAttendance,
                    LastSundayEnrollment, LastSundayAttendance, AttendanceRatio, RatioCategory,
                    LastSundayAttendanceRatio, LastSundayRatioCategory,
                    NeedsInReachCount, GoodRatioCount, NeedsOutreachCount,
                    LastSundayNeedsInReachCount, LastSundayGoodRatioCount, LastSundayNeedsOutreachCount
                )
                SELECT
                    'Organization', DivisionName, ProgramName, OrganizationName, 
                    5, 2,
                    TotalEnrollment,
                    EnrollmentWithMeetings,
                    TotalAttendance,
                    AvgAttendance,
                    NULL,
                    NULL,
                    AttendanceRatio,
                    RatioCategory,
                    NULL,
                    NULL,
                    NULL, NULL, NULL, NULL, NULL, NULL
                FROM
                    #OrganizationDetails
                WHERE
                    RatioCategory = 'Needs In-reach'
                ORDER BY
                    AttendanceRatio;
            END
        
            -- Add other organization details sections as in the original SQL...
        END
    
        -- Return the final result, sorted properly
        SELECT
            Level,
            DivisionName,
            ProgramName,
            OrganizationName,
            TotalEnrollment,
            EnrollmentWithMeetings,
            TotalAttendance,
            AvgAttendance,
            LastSundayEnrollment,
            LastSundayAttendance,
            AttendanceRatio,
            RatioCategory,
            LastSundayAttendanceRatio,
            LastSundayRatioCategory,
            NeedsInReachCount,
            GoodRatioCount,
            NeedsOutreachCount,
            LastSundayNeedsInReachCount,
            LastSundayGoodRatioCount,
            LastSundayNeedsOutreachCount
        FROM
            #Results
        ORDER BY
            LevelSortOrder,
            CategorySortOrder,
            ProgramName,
            DivisionName,
            OrganizationName;
    
        -- Clean up temporary tables
        DROP TABLE #ProgramNames;
        DROP TABLE #OrganizationBase;
        DROP TABLE #OrganizationMeetings;
        DROP TABLE #LastSundayMeetings;
        DROP TABLE #OrganizationRatios;
        DROP TABLE #ProgramCategoryCounts;
        DROP TABLE #DivisionCategoryCounts;
        DROP TABLE #OverallCategoryCounts;
        DROP TABLE #OverallSummary;
        DROP TABLE #ProgramSummary;
        DROP TABLE #DivisionSummary;
        DROP TABLE #Results;
    
        IF @IncludeOrgDetails = 1
        BEGIN
            DROP TABLE #OrganizationDetails;
        END
        """.format(
            weeks_count=ENROLLMENT_RATIO_WEEKS,
            start_date=weeks_ago_date_sql,
            end_date=current_date_sql,
            last_sunday_date=last_sunday_sql,
            needs_inreach=ENROLLMENT_RATIO_THRESHOLDS['needs_inreach'],
            good_ratio=ENROLLMENT_RATIO_THRESHOLDS['good_ratio'],
            include_org_details=1 if SHOW_DETAILED_ENROLLMENT else 0,
            org_filter=org_filter,
            division_filter=division_filter,
            program_id_filter=program_id_filter,
            program_insert=program_insert_sql
        )
        
        return sql
    
    def _build_program_insert_sql(self):
        """Build the SQL to insert program names into the #ProgramNames table."""
        # Check if the ENROLLMENT_ANALYSIS_PROGRAMS constant is defined and has values
        if not hasattr(self, 'ENROLLMENT_ANALYSIS_PROGRAMS') or not self.ENROLLMENT_ANALYSIS_PROGRAMS:
            return "INSERT INTO #ProgramNames (Name) VALUES ('Connect Group Attendance');"
        
        # Build INSERT statements for all configured programs
        insert_statements = []
        for program_name in self.ENROLLMENT_ANALYSIS_PROGRAMS:
            # Escape single quotes for SQL
            safe_name = program_name.replace("'", "''")
            insert_statements.append("INSERT INTO #ProgramNames (Name) VALUES ('{}');".format(safe_name))
        
        # If no statements were generated, add a default
        if not insert_statements:
            insert_statements.append("INSERT INTO #ProgramNames (Name) VALUES ('Connect Group Attendance');")
        
        # Join all statements
        return "\n".join(insert_statements)
    
    def _build_enrollment_additional_where(self, program_id=None, division_id=None, org_id=None):
        """Build additional WHERE conditions for the main query."""
        selected_date = model.Data.report_date if hasattr(model.Data, 'report_date') and model.Data.report_date else model.DateTime
        weeks_ago_date = model.DateAddDays(selected_date, -7 * ENROLLMENT_RATIO_WEEKS)
        
        # Format dates for SQL using simple string manipulation
        current_date_sql = str(selected_date).split(' ')[0] if ' ' in str(selected_date) else str(selected_date)
        weeks_ago_date_sql = str(weeks_ago_date).split(' ')[0] if ' ' in str(weeks_ago_date) else str(weeks_ago_date)
        conditions = [
            "o.OrganizationStatusId = 30",
            "(m.DidNotMeet = 0 OR m.DidNotMeet IS NULL)",
            "m.MeetingDate BETWEEN '{0}' AND '{1}'".format(weeks_ago_date,current_date_sql)
        ]
    
        if org_id:
            conditions.append("o.OrganizationId = {}".format(org_id))
        
        if division_id:
            conditions.append("os.DivId = {}".format(division_id))
        
        if program_id:
            conditions.append("d.ProgId = {}".format(program_id))
    
        return "AND " + " AND ".join(conditions)
    
    def _build_enrollment_program_filter(self, program_id=None, division_id=None, org_id=None):
        """Build additional program filtering conditions."""
        if org_id or division_id or program_id:
            return ""  # Already handled in the additional_where clause
        
        # If specific programs are configured, filter by those
        if ENROLLMENT_ANALYSIS_PROGRAMS:
            program_names = ", ".join(["'{}'".format(p.replace("'", "''")) for p in ENROLLMENT_ANALYSIS_PROGRAMS])
            return "AND oa.ProgramName IN ({})".format(program_names)
        
        return ""
    

    def generate_enrollment_analysis_section(self):
        """Generate a more compact enrollment ratio analysis section."""
        try:
            # Check if we should show this section at all
            if not ENROLLMENT_ANALYSIS_PROGRAMS:
                return "<p>No programs configured for enrollment analysis.</p>"
            
            # Start with loading indicator
            print "<div id='loading-enrollment' style='text-align: center; padding: 30px;'>"
            print "<p><i class='fa fa-spinner fa-spin fa-3x'></i></p>"
            print "<p>Loading enrollment analysis data...</p>"
            print "</div>"
            
            # Get the enrollment SQL and execute it
            try:
                enrollment_sql = self.get_enrollment_ratio_sql()
                enrollment_data = q.QuerySql(enrollment_sql)
                
                # If no data, return a message
                if not enrollment_data:
                    return "<p>No enrollment data found for the configured programs.</p>"
                    
            except Exception as sql_error:
                # Detailed error handling
                import traceback
                error_html = "<h3>Error Executing Enrollment SQL</h3>"
                error_html += "<p>Detailed Error: {}</p>".format(str(sql_error))
                error_html += "<pre>{}</pre>".format(traceback.format_exc())
                return error_html
            
            # Organize data by level
            organization_levels = {
                'Overall': [],
                'Program': [],
                'Division': [],
                'Organization': []
            }
            
            for item in enrollment_data:
                level = getattr(item, 'Level', 'Unknown')
                if level in organization_levels:
                    organization_levels[level].append(item)
            
            # Set title based on the first program if only one program is defined
            program_title = "Enrollment Analysis"
            if len(organization_levels['Program']) == 1:
                program_title = "{} Attendance".format(
                    getattr(organization_levels['Program'][0], 'ProgramName', 'Group')
                )
            
            # Format the HTML using a more compact style
            
            # <div style="font-style: italic; color: #666; margin-bottom: 15px; font-size: 13px;">
            #     Note: Average attendance is calculated based on the last {weeks} weeks of attendance data.
            # </div>
            
            analysis_html = """
            <div id="enrollment-analysis" style="display: none;">
                <div style="background-color: #f9f9f9; border-radius: 5px; padding: 15px; margin-bottom: 20px;">
                    <h3 style="margin-top: 0;">{title}</h3>

            """.format(
                title=program_title,
                weeks=ENROLLMENT_RATIO_WEEKS
            )
            
            # Create the main metrics section - use Overall or the single Program
            data_source = None
            if organization_levels['Overall']:
                data_source = organization_levels['Overall'][0]
            elif len(organization_levels['Program']) == 1:
                data_source = organization_levels['Program'][0]
                    
            if data_source:
                # Direct access to values from the SQL result
                total_enrollment = getattr(data_source, 'TotalEnrollment', 0) or 0
                last_sunday_attendance = getattr(data_source, 'LastSundayAttendance', 0) or 0
                attendance_ratio = getattr(data_source, 'AttendanceRatio', 0) or 0
                
                # Get correct field names from SQL result
                # Updated field names to match SQL query result
                needs_inreach_count = int(getattr(data_source, 'LastSundayNeedsInReachCount', 0) or 0)
                good_ratio_count = int(getattr(data_source, 'LastSundayGoodRatioCount', 0) or 0) 
                needs_outreach_count = int(getattr(data_source, 'LastSundayNeedsOutreachCount', 0) or 0)
                
                #<h4 style="color: #333; padding-bottom: 10px;">{title}</h4>
                analysis_html += """
                    <div style="border-bottom: 1px solid #4A90E2; margin-bottom: 15px;">

                        <div style="display: flex; flex-wrap: wrap; gap: 3px; margin-bottom: 15px;">
                            <div style="flex: 1; min-width: 110px; background-color: #f0f4f8; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 12px; color: #666; margin-bottom: 3px;">Enrollment</div>
                                <div style="font-size: 24px; font-weight: bold; color: #333;">{enrollment:,}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #f0f4f8; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 12px; color: #666; margin-bottom: 3px;">Avg Attendance</div>
                                <div style="font-size: 24px; font-weight: bold; color: #333;">{attendance:,.1f}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #f0f4f8; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 12px; color: #666; margin-bottom: 3px;">Ratio</div>
                                <div style="font-size: 24px; font-weight: bold; color: #333;">{ratio:.1f}%</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #ffe6e6; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 12px; color: #666; margin-bottom: 3px;">Needs In-reach</div>
                                <div style="font-size: 24px; font-weight: bold; color: #d9534f;">{needs_inreach}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #e6ffe6; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 12px; color: #666; margin-bottom: 3px;">Good Ratio</div>
                                <div style="font-size: 24px; font-weight: bold; color: #5cb85c;">{good_ratio}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #fff3e6; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 12px; color: #666; margin-bottom: 3px;">Needs Outreach</div>
                                <div style="font-size: 24px; font-weight: bold; color: #f0ad4e;">{needs_outreach}</div>
                            </div>
                        </div>
                    </div>
                """.format(
                    title="Connect Group Attendance" if program_title == "Enrollment Analysis" else program_title,
                    enrollment=total_enrollment,
                    attendance=last_sunday_attendance,
                    ratio=attendance_ratio,
                    needs_inreach=needs_inreach_count,
                    good_ratio=good_ratio_count,
                    needs_outreach=needs_outreach_count
                )
            
            # Division breakdown
            if organization_levels['Division']:
                analysis_html += """<h4 style="margin-top: 20px; margin-bottom: 15px;">Division Breakdown</h4>"""
                
                # Process each division
                for division_item in sorted(organization_levels['Division'], 
                                      key=lambda x: getattr(x, 'DivisionName', '')):
                    div_name = getattr(division_item, 'DivisionName', 'Unknown Division')
                    div_enrollment = getattr(division_item, 'TotalEnrollment', 0) or 0
                    div_attendance = getattr(division_item, 'LastSundayAttendance', 0) or 0
                    div_ratio = getattr(division_item, 'AttendanceRatio', 0) or 0
                    
                    # Get division counts directly from the SQL result
                    # Updated field names to match SQL query result
                    div_needs_inreach = int(getattr(division_item, 'NeedsInReachCount', 0) or 0)
                    div_good_ratio = int(getattr(division_item, 'GoodRatioCount', 0) or 0)
                    div_needs_outreach = int(getattr(division_item, 'NeedsOutreachCount', 0) or 0)
                    
                    analysis_html += """
                    <div style="border-left: 4px solid #4A90E2; margin-bottom: 15px;">
                        <div style="font-weight: bold; font-size: 16px; margin-bottom: 5px; padding-left: 10px;">{division}</div>
                        
                        <div style="display: flex; align-items: center; padding: 5px 10px; flex-wrap: wrap;">
                            <div style="min-width: 150px; flex: 1;">
                                <span style="color: #666;">Enrollment:</span> <strong>{enrollment:,}</strong>
                            </div>
                            <div style="min-width: 170px; flex: 1;">
                                <span style="color: #666;">Attendance:</span> <strong>{attendance:,.1f}</strong>
                            </div>
                            
                            <div style="min-width: 200px; flex: 2; display: flex; align-items: center; flex-wrap: wrap;">
                                <div style="margin-right: 10px;">
                                    <span style="display: inline-block; padding: 2px 6px; background-color: #ffe6e6; border-radius: 3px; margin-right: 3px;">
                                        <span style="color: #666;">In-reach:</span> <strong style="color: #d9534f;">{in_reach}</strong>
                                    </span>
                                </div>
                                
                                <div style="margin-right: 10px;">
                                    <span style="display: inline-block; padding: 2px 6px; background-color: #e6ffe6; border-radius: 3px; margin-right: 3px;">
                                        <span style="color: #666;">Good:</span> <strong style="color: #5cb85c;">{good_ratio}</strong>
                                    </span>
                                </div>
                                
                                <div style="margin-right: 10px;">
                                    <span style="display: inline-block; padding: 2px 6px; background-color: #fff3e6; border-radius: 3px; margin-right: 3px;">
                                        <span style="color: #666;">Outreach:</span> <strong style="color: #f0ad4e;">{outreach}</strong>
                                    </span>
                                </div>
                            </div>
                            
                            <div style="min-width: 120px; text-align: right;">
                                <span style="display: inline-block; padding: 2px 6px; background-color: #e6f0ff; border-radius: 3px;">
                                    <span style="color: #666;">Ratio:</span> <strong style="color: #4A90E2;">{ratio:.1f}%</strong>
                                </span>
                            </div>
                        </div>
                    </div>
                    """.format(
                        division=div_name,
                        enrollment=div_enrollment,
                        attendance=div_attendance,
                        ratio=div_ratio,
                        in_reach=div_needs_inreach,
                        good_ratio=div_good_ratio,
                        outreach=div_needs_outreach
                    )
                
            analysis_html += """
                </div>
            </div>
            
            <script>
                document.getElementById('loading-enrollment').style.display = 'none';
                document.getElementById('enrollment-analysis').style.display = 'block';
            </script>
            """
            
            return analysis_html
            
        except Exception as e:
            # Final error handling
            import traceback
            error_html = "<h3>Error in Enrollment Analysis</h3>"
            error_html += "<p>Details: {}</p>".format(str(e))
            error_html += "<pre>{}</pre>".format(traceback.format_exc())
            return error_html
    
    def _get_category_color(self, category, light_version=False):
        """Get the background color for a ratio category."""
        if category == 'Needs In-reach':
            return '#ffebee' if light_version else '#ffcdd2'
        elif category == 'Good Ratio':
            return '#e8f5e9' if light_version else '#c8e6c9'
        elif category == 'Needs Outreach':
            return '#fff3e0' if light_version else '#ffe0b2'
        return '#f5f5f5'  # Default/No Data
    
    def _get_category_text_color(self, category):
        """Get the text color for a ratio category."""
        if category == 'Needs In-reach':
            return '#d32f2f'
        elif category == 'Good Ratio':
            return '#2E7D32'
        elif category == 'Needs Outreach':
            return '#F57C00'
        return '#757575'  # Default/No Data

    def _generate_detailed_division_view(self, division_data):
        """Generate a detailed view of organizations within a division."""
        if not division_data['organizations']:
            return ""
        
        html = """
        <div style="margin-top: 15px; background-color: white; border-radius: 5px; padding: 10px;">
            <h6 style="margin: 0 0 10px 0; color: #333;">Organization Details</h6>
            <table style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr style="background-color: #f1f1f1;">
                        <th style="padding: 8px; text-align: left;">Organization</th>
                        <th style="padding: 8px; text-align: right;">Enrollment</th>
                        <th style="padding: 8px; text-align: right;">Avg Attendance</th>
                        <th style="padding: 8px; text-align: right;">Ratio</th>
                        <th style="padding: 8px; text-align: left;">Status</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for org in sorted(division_data['organizations'], key=lambda x: x.OrganizationName):
            # Determine row color based on ratio category
            row_color = {
                'Needs In-reach': '#ffebee',    # Light red
                'Good Ratio': '#e8f5e9',        # Light green
                'Needs Outreach': '#fff3e0'     # Light orange
            }.get(org.RatioCategory, 'white')
            
            html += """
                <tr style="background-color: {0};">
                    <td style="padding: 8px;">{1}</td>
                    <td style="padding: 8px; text-align: right;">{2}</td>
                    <td style="padding: 8px; text-align: right;">{3}</td>
                    <td style="padding: 8px; text-align: right;">{4}%</td>
                    <td style="padding: 8px;">{5}</td>
                </tr>
            """.format(
                row_color,
                org.OrganizationName,
                ReportHelper.format_number(org.Enrollment),
                ReportHelper.format_float(org.AvgAttendance),
                ReportHelper.format_float(org.AttendanceRatio),
                org.RatioCategory
            )
        
        html += """
                </tbody>
            </table>
        </div>
        """
        
        return html

    def generate_date_range_summary(self):
        """Generate a summary section for the selected date range."""
        if SHOW_FOUR_WEEK_COMPARISON:
            # Use separate format strings for with 4-week data
            date_range_html = """
            <div style="margin-top: 20px; margin-bottom: 20px; padding: 10px; background-color: #f0f8ff; border-radius: 5px; border-left: 5px solid #4CAF50;">
                <h3 style="margin-top: 0;">Weekly Report</h3>
                <p>Current Week: <strong>{0}</strong> to <strong>{1}</strong></p>
                <p>Previous Week: <strong>{2}</strong> to <strong>{3}</strong></p>
                <p>Previous Year: <strong>{4}</strong> to <strong>{5}</strong></p>
                <p>Current 4 Weeks: <strong>{6}</strong> to <strong>{7}</strong></p>
                <p>Previous Year 4 Weeks: <strong>{8}</strong> to <strong>{9}</strong></p>
            </div>
            """
            
            # Format with all dates using index numbers instead of empty {}
            formatted_html = date_range_html.format(
                ReportHelper.format_display_date(self.week_start_date),
                ReportHelper.format_display_date(self.report_date),
                ReportHelper.format_display_date(self.previous_week_start),
                ReportHelper.format_display_date(self.previous_week_date),
                ReportHelper.format_display_date(self.previous_week_start_date),
                ReportHelper.format_display_date(self.previous_year_date),
                ReportHelper.format_display_date(self.four_weeks_ago_start_date),
                ReportHelper.format_display_date(self.report_date),
                ReportHelper.format_display_date(self.prev_year_four_weeks_ago_start_date),
                ReportHelper.format_display_date(self.previous_year_date)
            )
        else:
            # Simpler format string without 4-week data
            date_range_html = """
            <div style="margin-top: 20px; margin-bottom: 20px; padding: 10px; background-color: #f0f8ff; border-radius: 5px; border-left: 5px solid #4CAF50;">
                <h3 style="margin-top: 0;">Weekly Report</h3>
                <p>Current Week: <strong>{0}</strong> to <strong>{1}</strong></p>
                <p>Previous Week: <strong>{2}</strong> to <strong>{3}</strong></p>
                <p>Previous Year: <strong>{4}</strong> to <strong>{5}</strong></p>
            </div>
            """
            
            formatted_html = date_range_html.format(
                ReportHelper.format_display_date(self.week_start_date),
                ReportHelper.format_display_date(self.report_date),
                ReportHelper.format_display_date(self.previous_week_start),
                ReportHelper.format_display_date(self.previous_week_date),
                ReportHelper.format_display_date(self.previous_week_start_date),
                ReportHelper.format_display_date(self.previous_year_date)
            )
        
        return formatted_html
    
    def generate_fiscal_year_summary(self, current_ytd_total, previous_ytd_total):
        """Generate a fiscal year summary section."""
        # Adjust labels based on year type and custom configuration
        ytd_prefix = ReportHelper.get_ytd_label()
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
        
        # Skip this section if not enabled
        if not SHOW_PROGRAM_SUMMARY:
            return ""
        
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
        
        # Get four-week comparison data if enabled
        current_four_week_data = None
        prev_year_four_week_data = None
        if SHOW_FOUR_WEEK_COMPARISON:
            # Get current 4-week period data
            current_four_week_data = self.get_specific_program_attendance_data(
                all_programs,
                self.four_weeks_ago_start_date,
                self.report_date
            )
            
            # Get previous year's 4-week period data
            prev_year_four_week_data = self.get_specific_program_attendance_data(
                all_programs,
                self.prev_year_four_weeks_ago_start_date,
                self.previous_year_date
            )
        
        # Get YTD label based on configuration
        ytd_prefix = ReportHelper.get_ytd_label()
        
        # Build HTML table header
        summary_html = """
        <div style="margin-top: 20px; margin-bottom: 30px; padding: 15px; background-color: #f0f0f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h3 style="margin-top: 0;">Program Summary</h3>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <thead>
                    <tr style="background-color: #e0e0e8;">
                        <th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Program</th>
        """
        
        # Prepare column definitions
        columns = [
            {
                'name': 'Current Week', 
                'data': 'current_week', 
                'show': SHOW_CURRENT_WEEK_COLUMN
            },
            {
                'name': 'Previous Year', 
                'data': 'previous_year', 
                'show': SHOW_PREVIOUS_YEAR_COLUMN
            },
            {
                'name': 'YoY Change', 
                'data': 'yoy_change', 
                'show': SHOW_YOY_CHANGE_COLUMN
            },
            {
                'name': 'Last 4 Weeks', 
                'data': 'four_week_current', 
                'show': SHOW_FOUR_WEEK_COMPARISON
            },
            {
                'name': 'PY 4 Weeks', 
                'data': 'four_week_previous', 
                'show': SHOW_FOUR_WEEK_COMPARISON
            },
            {
                'name': '4-Week Change', 
                'data': 'four_week_change', 
                'show': SHOW_FOUR_WEEK_COMPARISON
            },
            {
                'name': ytd_prefix, 
                'data': 'current_fytd', 
                'show': SHOW_FYTD_COLUMNS
            },
            {
                'name': 'P' + ytd_prefix, 
                'data': 'previous_fytd', 
                'show': SHOW_FYTD_COLUMNS
            },
            {
                'name': ytd_prefix + ' Change', 
                'data': 'fytd_change', 
                'show': SHOW_FYTD_CHANGE_COLUMNS
            }
        ]
        
        # Add headers for visible columns
        for col in columns:
            if col['show']:
                summary_html += '<th style="padding: 10px; text-align: center; border: 1px solid #ddd;">{}</th>'.format(col['name'])
        
        summary_html += """
                    </tr>
                </thead>
                <tbody>
        """
        
        # Add a row for each program in order of RptGroup
        for program_info in all_program_data:
            program_name = program_info['name']
            
            # Skip if program isn't found in current data
            if program_name not in current_specific_data['by_program']:
                continue
            
            # Start the row with program name
            summary_html += """
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 8px; text-align: left; border: 1px solid #ddd;"><strong>{}</strong></td>
            """.format(program_name)
            
            # Get current program data
            prog_data = current_specific_data['by_program'][program_name]
            
            # Use total instead of average for FYTD
            current_fytd_total = prog_data['total']
            
            # Get previous year FYTD data for the program (PYTD)
            prev_fytd_total = 0
            if program_name in previous_specific_data['by_program']:
                prev_ytd_data = previous_specific_data['by_program'][program_name]
                prev_fytd_total = prev_ytd_data['total']
            
            # Add columns based on configuration
            for col in columns:
                if not col['show']:
                    continue
                
                # Determine the value for each column type
                if col['data'] == 'current_week':
                    value = 0
                    if program_name in current_week_data['by_program']:
                        value = current_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(value)
                    )
                
                elif col['data'] == 'previous_year':
                    value = 0
                    if prev_year_specific_week_data and program_name in prev_year_specific_week_data['by_program']:
                        value = prev_year_specific_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(value)
                    )
                
                elif col['data'] == 'yoy_change':
                    current_week_total = 0
                    if program_name in current_week_data['by_program']:
                        current_week_total = current_week_data['by_program'][program_name]['total']
                    
                    prev_year_week_total = 0
                    if prev_year_specific_week_data and program_name in prev_year_specific_week_data['by_program']:
                        prev_year_week_total = prev_year_specific_week_data['by_program'][program_name]['total']
                    
                    trend = ReportHelper.get_trend_indicator(current_week_total, prev_year_week_total)
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(trend)
                
                elif col['data'] == 'four_week_current':
                    value = 0
                    if current_four_week_data and program_name in current_four_week_data['by_program']:
                        value = current_four_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(value)
                    )
                
                elif col['data'] == 'four_week_previous':
                    value = 0
                    if prev_year_four_week_data and program_name in prev_year_four_week_data['by_program']:
                        value = prev_year_four_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(value)
                    )
                
                elif col['data'] == 'four_week_change':
                    current_four_week_total = 0
                    if current_four_week_data and program_name in current_four_week_data['by_program']:
                        current_four_week_total = current_four_week_data['by_program'][program_name]['total']
                    
                    prev_year_four_week_total = 0
                    if prev_year_four_week_data and program_name in prev_year_four_week_data['by_program']:
                        prev_year_four_week_total = prev_year_four_week_data['by_program'][program_name]['total']
                    
                    trend = ReportHelper.get_trend_indicator(current_four_week_total, prev_year_four_week_total)
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(trend)
                
                elif col['data'] == 'current_fytd':
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(current_fytd_total)
                    )
                
                elif col['data'] == 'previous_fytd':
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(prev_fytd_total)
                    )
                
                elif col['data'] == 'fytd_change':
                    trend = ReportHelper.get_trend_indicator(current_fytd_total, prev_fytd_total)
                    summary_html += '<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(trend)
            
            # Close the row
            summary_html += "</tr>"
        
        # Close the table and container
        summary_html += """
                </tbody>
            </table>
        </div>
        """
        
        return summary_html
    
    def generate_overall_summary(self, current_total, previous_total, current_ytd_total, previous_ytd_total, four_week_totals=None):
        """Generate an improved overall summary section with robust error handling and performance tracking."""
        import traceback
        import time  # Add timing functionality
        
        # Add overall function timing
        start_time_total = time.time()
        
        # Create a dictionary to track execution times of different sections
        timing_data = {}
        summary_html = ""
        
        try:
            # Ensure all input parameters have valid types
            current_total = current_total or 0
            previous_total = previous_total or 0
            current_ytd_total = current_ytd_total or 0
            previous_ytd_total = previous_ytd_total or 0
            four_week_totals = four_week_totals or {'current': 0, 'previous_year': 0}
    
            # Adjust labels based on year type and custom configuration
            ytd_prefix = ReportHelper.get_ytd_label()
            
            # Initialize with default values
            worship_program_name = WORSHIP_PROGRAM  # Default fallback
            worship_current_ytd_avg = 0
            worship_previous_ytd_avg = 0
            four_week_program_avg = 0
            four_week_previous_program_avg = 0
            enrollment_ratio = 0
            ratio_category = "No Data"
            
            # Time the worship program section
            start_time = time.time()
            try:
                # Find the specified worship program - keeping this but handling all errors
                try:
                    worship_program_sql = ""
                    start_time_sql = time.time()
                    try:
                        worship_program_sql = self.get_specific_programs_sql([WORSHIP_PROGRAM])
                        timing_data['get_programs_sql'] = time.time() - start_time_sql
                        print "<p class='debug'>SQL generation took: {:.2f} seconds</p>".format(timing_data['get_programs_sql'])
                        if not worship_program_sql:
                            print "<p class='debug'>Warning: get_specific_programs_sql returned empty SQL</p>"
                            worship_program_sql = """
                            SELECT p.Id, p.Name 
                            FROM Program p 
                            WHERE p.Name = '{0}'
                            """.format(WORSHIP_PROGRAM)
                    except Exception as sql_error:
                        print "<p class='debug'>Error generating program SQL: " + str(sql_error) + "</p>"
                        # Fallback SQL if method fails
                        worship_program_sql = """
                        SELECT p.Id, p.Name 
                        FROM Program p 
                        WHERE p.Name = '{0}'
                        """.format(WORSHIP_PROGRAM)
                    
                    print "<p class='debug'>Program SQL: " + worship_program_sql + "</p>"
                    
                    # Initialize worship_programs to empty list before usage
                    worship_programs = []
                    
                    # Try to execute the query with careful error handling
                    start_time_query = time.time()
                    try:
                        worship_programs = q.QuerySql(worship_program_sql) or []
                        timing_data['worship_program_query'] = time.time() - start_time_query
                        print "<p class='debug'>Worship program query took: {:.2f} seconds</p>".format(timing_data['worship_program_query'])
                        print "<p class='debug'>Found " + str(len(worship_programs)) + " worship programs</p>"
                    except Exception as query_error:
                        print "<p class='debug'>Error executing worship program query: " + str(query_error) + "</p>"
                        worship_programs = []
                    
                    # Safely get the first program using a generator expression
                    first_program_match = None
                    if worship_programs and len(worship_programs) > 0:
                        first_program_match = worship_programs[0]
                    
                    if first_program_match:
                        # Safely extract Name with fallback
                        if hasattr(first_program_match, 'Name'):
                            worship_program_name = first_program_match.Name
                    else:
                        worship_program_name = WORSHIP_PROGRAM
                    
                except Exception as worship_error:
                    print "<p class='debug'>Error in worship program section: " + str(worship_error) + "</p>"
                    worship_program_name = WORSHIP_PROGRAM
                
            finally:
                timing_data['worship_program_section'] = time.time() - start_time
                print "<p class='debug'>Worship program section took: {:.2f} seconds</p>".format(timing_data['worship_program_section'])
            
            # Calculate weeks elapsed and YTD averages
            start_time = time.time()
            try:
                # Get weeks elapsed safely
                weeks_elapsed = 1
                start_time_weeks = time.time()
                try:
                    # Use helper function defined in your class
                    weeks_elapsed = max(1, ReportHelper.get_weeks_elapsed_in_fiscal_year(
                        self.report_date, 
                        self.current_fiscal_start
                    ))
                    timing_data['get_weeks_elapsed'] = time.time() - start_time_weeks
                    print "<p class='debug'>Weeks elapsed calculation took: {:.2f} seconds</p>".format(timing_data['get_weeks_elapsed'])
                except Exception as weeks_error:
                    print "<p class='debug'>Error calculating weeks elapsed: " + str(weeks_error) + "</p>"
                    weeks_elapsed = 1  # Fallback to 1 to avoid division by zero
                
                # Get current YTD attendance with proper error handling
                start_time_current = time.time()
                try:
                    print "<p class='debug'>Getting current YTD data from {} to {}</p>".format(
                        str(self.current_fiscal_start), str(self.report_date))
                    
                    current_worship_ytd_data = self.get_specific_program_attendance_data(
                        [worship_program_name],
                        self.current_fiscal_start,
                        self.report_date
                    ) or {'by_program': {}}
                    
                    timing_data['current_ytd_data'] = time.time() - start_time_current
                    print "<p class='debug'>Current YTD data took: {:.2f} seconds</p>".format(timing_data['current_ytd_data'])
                    
                    # Calculate current YTD average safely
                    if (isinstance(current_worship_ytd_data, dict) and 
                        'by_program' in current_worship_ytd_data and 
                        worship_program_name in current_worship_ytd_data['by_program']):
                        
                        program_data = current_worship_ytd_data['by_program'][worship_program_name]
                        if isinstance(program_data, dict) and 'total' in program_data:
                            total = program_data.get('total', 0)
                            if total > 0:
                                worship_current_ytd_avg = float(total) / float(weeks_elapsed)
                except Exception as current_error:
                    print "<p class='debug'>Error calculating current YTD average: " + str(current_error) + "</p>"
                
                # Get previous YTD attendance with proper error handling
                start_time_previous = time.time()
                try:
                    print "<p class='debug'>Getting previous YTD data from {} to {}</p>".format(
                        str(self.previous_fiscal_start), str(self.previous_year_date))
                    
                    previous_worship_ytd_data = self.get_specific_program_attendance_data(
                        [worship_program_name],
                        self.previous_fiscal_start,
                        self.previous_year_date
                    ) or {'by_program': {}}
                    
                    timing_data['previous_ytd_data'] = time.time() - start_time_previous
                    print "<p class='debug'>Previous YTD data took: {:.2f} seconds</p>".format(timing_data['previous_ytd_data'])
                    
                    # Calculate previous YTD average safely
                    if (isinstance(previous_worship_ytd_data, dict) and 
                        'by_program' in previous_worship_ytd_data and 
                        worship_program_name in previous_worship_ytd_data['by_program']):
                        
                        program_data = previous_worship_ytd_data['by_program'][worship_program_name]
                        if isinstance(program_data, dict) and 'total' in program_data:
                            total = program_data.get('total', 0)
                            if total > 0:
                                worship_previous_ytd_avg = float(total) / float(weeks_elapsed)
                except Exception as previous_error:
                    print "<p class='debug'>Error calculating previous YTD average: " + str(previous_error) + "</p>"
                    
            except Exception as ytd_error:
                print "<p class='debug'>Error in YTD section: " + str(ytd_error) + "</p>"
            finally:
                timing_data['ytd_section'] = time.time() - start_time
                print "<p class='debug'>YTD averages section took: {:.2f} seconds</p>".format(timing_data['ytd_section'])
            
            # Process four-week averages
            start_time = time.time()
            try:
                # Get current 4-week data safely
                start_time_current_4wk = time.time()
                try:
                    print "<p class='debug'>Getting current 4-week data from {} to {}</p>".format(
                        str(self.four_weeks_ago_date), str(self.report_date))
                    
                    four_week_program_data = self.get_specific_program_attendance_data(
                        [ENROLLMENT_RATIO_PROGRAM],
                        self.four_weeks_ago_date,
                        self.report_date
                    ) or {'by_program': {}}
                    
                    timing_data['current_4wk_data'] = time.time() - start_time_current_4wk
                    print "<p class='debug'>Current 4-week data took: {:.2f} seconds</p>".format(timing_data['current_4wk_data'])
                    
                    if (isinstance(four_week_program_data, dict) and 
                        'by_program' in four_week_program_data and 
                        ENROLLMENT_RATIO_PROGRAM in four_week_program_data['by_program']):
                        
                        program_data = four_week_program_data['by_program'][ENROLLMENT_RATIO_PROGRAM]
                        if isinstance(program_data, dict) and 'total' in program_data:
                            total = program_data.get('total', 0)
                            # Avoid zero division
                            four_week_program_avg = total / 4.0 if total > 0 else 0
                except Exception as current_4wk_error:
                    print "<p class='debug'>Error calculating current 4-week average: " + str(current_4wk_error) + "</p>"
                
                # Get previous year 4-week data safely
                start_time_prev_4wk = time.time()
                try:
                    print "<p class='debug'>Getting previous 4-week data from {} to {}</p>".format(
                        str(self.prev_year_four_weeks_ago_date), str(self.previous_year_date))
                    
                    four_week_previous_program_data = self.get_specific_program_attendance_data(
                        [ENROLLMENT_RATIO_PROGRAM],
                        self.prev_year_four_weeks_ago_date,
                        self.previous_year_date
                    ) or {'by_program': {}}
                    
                    timing_data['previous_4wk_data'] = time.time() - start_time_prev_4wk
                    print "<p class='debug'>Previous 4-week data took: {:.2f} seconds</p>".format(timing_data['previous_4wk_data'])
                    
                    if (isinstance(four_week_previous_program_data, dict) and 
                        'by_program' in four_week_previous_program_data and 
                        ENROLLMENT_RATIO_PROGRAM in four_week_previous_program_data['by_program']):
                        
                        program_data = four_week_previous_program_data['by_program'][ENROLLMENT_RATIO_PROGRAM]
                        if isinstance(program_data, dict) and 'total' in program_data:
                            total = program_data.get('total', 0)
                            # Avoid zero division
                            four_week_previous_program_avg = total / 4.0 if total > 0 else 0
                except Exception as previous_4wk_error:
                    print "<p class='debug'>Error calculating previous 4-week average: " + str(previous_4wk_error) + "</p>"
                    
            except Exception as four_week_error:
                print "<p class='debug'>Error in 4-week section: " + str(four_week_error) + "</p>"
            finally:
                timing_data['four_week_section'] = time.time() - start_time
                print "<p class='debug'>Four-week averages section took: {:.2f} seconds</p>".format(timing_data['four_week_section'])
            
            # Calculate enrollment ratio with TouchPoint-specific date handling
            start_time = time.time()
            try:
                # Get selected date with minimal processing
                selected_date = model.Data.report_date if hasattr(model.Data, 'report_date') and model.Data.report_date else model.DateTime
                weeks_ago_date = model.DateAddDays(selected_date, -7 * ENROLLMENT_RATIO_WEEKS)
                
                # Format dates for SQL using simple string manipulation
                current_date_sql = str(selected_date).split(' ')[0] if ' ' in str(selected_date) else str(selected_date)
                weeks_ago_date_sql = str(weeks_ago_date).split(' ')[0] if ' ' in str(weeks_ago_date) else str(weeks_ago_date)
                
                print "<p class='debug'>Date range: " + weeks_ago_date_sql + " to " + current_date_sql + "</p>"
 
                verification_sql = """
                SELECT 
                    SUM(o.MemberCount) AS TotalEnrollment,
                    COALESCE(SUM(att.AvgAttendance), 0) AS TotalAvgAttendance,
                    COALESCE(SUM(att.TotalAttendance), 0) AS TotalAttendance
                FROM 
                    Organizations o
                    JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
                    JOIN Division d ON d.Id = os.DivId
                    JOIN Program p ON p.Id = d.ProgId AND p.Name = '{2}'
                    LEFT JOIN (
                        SELECT 
                            OrganizationId,
                            SUM(CAST(COALESCE(MaxCount, 0) AS FLOAT)) AS TotalAttendance,
                            AVG(CAST(COALESCE(MaxCount, 0) AS FLOAT)) AS AvgAttendance
                        FROM 
                            Meetings
                        WHERE 
                            (DidNotMeet = 0 OR DidNotMeet IS NULL)
                            AND MeetingDate BETWEEN '{0}' AND '{1}'
                        GROUP BY 
                            OrganizationId
                    ) att ON att.OrganizationId = o.OrganizationId
                WHERE 
                    o.OrganizationStatusId = 30
                """.format(weeks_ago_date_sql, current_date_sql, ENROLLMENT_RATIO_PROGRAM)
                
                print "<p class='debug'>Enrollment ratio SQL:</p><pre class='debug'>" + verification_sql + "</pre>"
                
                # Execute query
                start_time_ratio_query = time.time()
                verification_result = q.QuerySqlTop1(verification_sql)
                timing_data['enrollment_ratio_query'] = time.time() - start_time_ratio_query
                print "<p class='debug'>Enrollment ratio query took: {:.2f} seconds</p>".format(timing_data['enrollment_ratio_query'])
                
                # Set default values
                total_enrollment = 0
                total_avg_attendance = 0
                enrollment_ratio = 0
                ratio_category = "No Data"
                ratio_color = "#333"
                
                # Process results if available
                if verification_result:
                    total_enrollment = getattr(verification_result, 'TotalEnrollment', 0) or 0
                    total_attendance = getattr(verification_result, 'TotalAttendance', 0) or 0
                    total_avg_attendance = getattr(verification_result, 'TotalAvgAttendance', 0) or 0
                    
                    # Calculate ratio only if enrollment exists
                    #if total_enrollment > 0:
                    #    enrollment_ratio = (total_avg_attendance / float(total_enrollment)) * 100
                    
                    total_weeks = ENROLLMENT_RATIO_WEEKS
                    
                    if total_enrollment > 0 and total_weeks > 0:
                        enrollment_ratio = ((total_attendance / float(total_weeks))/float(total_enrollment)) * 100
                    
                        # Determine category using constants
                        if enrollment_ratio <= ENROLLMENT_RATIO_THRESHOLDS['needs_inreach']:
                            ratio_category = 'Needs In-reach'
                            ratio_color = "#d9534f"
                        elif enrollment_ratio <= ENROLLMENT_RATIO_THRESHOLDS['good_ratio']:
                            ratio_category = 'Good Ratio'
                            ratio_color = "#5cb85c"
                        else:
                            ratio_category = 'Needs Outreach'
                            ratio_color = "#f0ad4e"
                    
                    print "<p class='debug'>Enrollment: " + str(total_enrollment) + " | Avg: " + str(total_avg_attendance) + " | Ratio: " + str(round(enrollment_ratio, 1)) + "%</p>"
                
            except Exception as ratio_error:
                print "<p class='debug'>Error calculating ratio: " + str(ratio_error) + "</p>"
                
                # Fallback values
                enrollment_ratio = 0
                ratio_category = "Error"
                ratio_color = "#333"  # Default gray
            finally:
                timing_data['enrollment_ratio_section'] = time.time() - start_time
                print "<p class='debug'>Enrollment ratio section took: {:.2f} seconds</p>".format(timing_data['enrollment_ratio_section'])
            
            # Format all values for display with error handling
            start_time = time.time()
            try:
                # Use the class formatting methods - replaced with direct formatting if they fail
                try:
                    formatted_worship_current_ytd_avg = ReportHelper.format_float(worship_current_ytd_avg)
                except:
                    formatted_worship_current_ytd_avg = "{0:.1f}".format(worship_current_ytd_avg) if worship_current_ytd_avg else "0"
                    
                try:
                    formatted_worship_previous_ytd_avg = ReportHelper.format_number(worship_previous_ytd_avg)
                except:
                    formatted_worship_previous_ytd_avg = "{0:,}".format(int(worship_previous_ytd_avg)) if worship_previous_ytd_avg else "0"
                    
                try:
                    formatted_four_week_program_avg = ReportHelper.format_number(four_week_program_avg)
                except:
                    formatted_four_week_program_avg = "{0:,}".format(int(four_week_program_avg)) if four_week_program_avg else "0"
                    
                try:
                    formatted_four_week_previous_avg = ReportHelper.format_number(four_week_previous_program_avg)
                except:
                    formatted_four_week_previous_avg = "{0:,}".format(int(four_week_previous_program_avg)) if four_week_previous_program_avg else "0"
                
                # Calculate trend indicator
                trend_indicator = ""
                try:
                    trend_indicator = ReportHelper.get_trend_indicator(worship_current_ytd_avg, worship_previous_ytd_avg)
                except:
                    # Direct calculation if the helper fails
                    if worship_previous_ytd_avg > 0:
                        trend_pct = ((worship_current_ytd_avg - worship_previous_ytd_avg) / worship_previous_ytd_avg) * 100
                        if trend_pct > 1:
                            trend_indicator = '<span style="color: green;">▲ ' + str(round(trend_pct, 1)) + '%</span>'
                        elif trend_pct < -1:
                            trend_indicator = '<span style="color: red;">▼ ' + str(round(abs(trend_pct), 1)) + '%</span>'
                        else:
                            trend_indicator = '<span style="color: gray;">◆ 0.0%</span>'
                    else:
                        trend_indicator = ""
                    
            except Exception as format_error:
                print "<p class='debug'>Error formatting numbers: " + str(format_error) + "</p>"
                # Use very basic formatting as fallback
                formatted_worship_current_ytd_avg = str(int(worship_current_ytd_avg))
                formatted_worship_previous_ytd_avg = str(int(worship_previous_ytd_avg))
                formatted_four_week_program_avg = str(int(four_week_program_avg))
                formatted_four_week_previous_avg = str(int(four_week_previous_program_avg))
                trend_indicator = ""
            finally:
                timing_data['formatting_section'] = time.time() - start_time
                print "<p class='debug'>Formatting section took: {:.2f} seconds</p>".format(timing_data['formatting_section'])
            
            # Generate the summary HTML
            start_time = time.time()
            try:
                # Using traditional string formatting to avoid f-string issues
                enrollment_ratio_rounded = str(round(enrollment_ratio, 1))  # Pre-format as string to avoid format specifier issue
                
                summary_html = """
                <div style="margin-top: 20px; margin-bottom: 30px; padding: 15px; background-color: #f8f8f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                    <h3 style="margin-top: 0;">Overall Summary</h3>
                    <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 20px;">
                        <div style="flex: 1; min-width: 180px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 1.2em; font-weight: bold;">{0} Avg - {1}</div>
                            <div style="font-size: 2em; margin: 10px 0;">{2}</div>
                            <div>{1} Last Year: {3}</div>
                            <div style="margin-top: 5px;">{4}</div>
                        </div>
                        
                        <div style="flex: 1; min-width: 180px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 1.2em; font-weight: bold;">4-Week Avg - {5}</div>
                            <div style="font-size: 2em; margin: 10px 0;">{6}</div>
                            <div>4 Weeks Last Year: {7}</div>
                        </div>
                        
                        <div style="flex: 1; min-width: 180px; padding: 15px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 1.2em; font-weight: bold;">{5} {11}-Week Enrollment Ratio</div>
                            <div style="font-size: 2em; margin: 10px 0;">{8}%</div>
                            <div style="font-size: 1em; color: {9};">{10}</div>
                        </div>
                    </div>
                </div>
                """.format(
                    worship_program_name,              # 0
                    ytd_prefix,                        # 1
                    formatted_worship_current_ytd_avg, # 2
                    formatted_worship_previous_ytd_avg,# 3
                    trend_indicator,                   # 4
                    ENROLLMENT_RATIO_PROGRAM,          # 5
                    formatted_four_week_program_avg,   # 6
                    formatted_four_week_previous_avg,  # 7
                    enrollment_ratio_rounded,          # 8 - using pre-formatted string
                    ratio_color,                       # 9
                    ratio_category,                    # 10
                    ENROLLMENT_RATIO_WEEKS             # 11
                )
                
            except Exception as html_error:
                print "<p class='debug'>Error generating HTML: " + str(html_error) + "</p>"
                
                # Fallback to even simpler HTML if the formatting fails
                summary_html = """
                <div style="margin-top: 20px; margin-bottom: 30px; padding: 15px; background-color: #f8f8f8; border-radius: 5px;">
                    <h3 style="margin-top: 0;">Overall Summary</h3>
                    <div style="display: flex; flex-wrap: wrap; gap: 20px;">
                        <div style="flex: 1; min-width: 180px; padding: 15px; background-color: #fff; border-radius: 5px;">
                            <div><strong>{0} Average</strong></div>
                            <div style="font-size: 1.5em;">{1}</div>
                        </div>
                    </div>
                </div>
                """.format(worship_program_name, formatted_worship_current_ytd_avg)
            finally:
                timing_data['html_generation'] = time.time() - start_time
                print "<p class='debug'>HTML generation took: {:.2f} seconds</p>".format(timing_data['html_generation'])
            
        except Exception as e:
            # Final fallback for any unhandled errors
            print "<h2>Error</h2>"
            print "<p>An error occurred: " + str(e) + "</p>"
            import traceback
            print "<pre>"
            traceback.print_exc()
            print "</pre>"
            
            # Return absolute minimal HTML that won't break
            summary_html = """
            <div style="padding: 15px; background-color: #f8f8f8; border-radius: 5px;">
                <h3>Overall Summary</h3>
                <p>Unable to generate summary due to an error.</p>
            </div>
            """
        finally:
            # Print overall execution time and summary of all timings
            total_time = time.time() - start_time_total
            
            # Create the performance analysis report
            performance_html = """
            <div class='debug' style='margin-top: 20px; padding: 10px; background-color: #f0f0f0; border: 1px solid #ccc;'>
                <h4>Performance Analysis</h4>
                <p>Total execution time: {0:.2f} seconds</p>
                <table style='width: 100%; border-collapse: collapse;'>
                    <tr>
                        <th style='text-align: left; padding: 5px; border-bottom: 1px solid #ccc;'>Section</th>
                        <th style='text-align: right; padding: 5px; border-bottom: 1px solid #ccc;'>Time (seconds)</th>
                        <th style='text-align: right; padding: 5px; border-bottom: 1px solid #ccc;'>Percentage</th>
                    </tr>
            """.format(total_time)
            
            # Sort timings by duration (longest first)
            sorted_timings = sorted(timing_data.items(), key=lambda x: x[1], reverse=True)
            for section, duration in sorted_timings:
                percentage = (duration / total_time) * 100
                bgColor = "#ffdddd" if percentage > 25 else "#ffffff"  # Highlight slow sections
                
                performance_html += """
                    <tr style='background-color: {3};'>
                        <td style='padding: 5px; border-bottom: 1px solid #eee;'>{0}</td>
                        <td style='text-align: right; padding: 5px; border-bottom: 1px solid #eee;'>{1:.2f}</td>
                        <td style='text-align: right; padding: 5px; border-bottom: 1px solid #eee;'>{2:.1f}%</td>
                    </tr>
                """.format(section, duration, percentage, bgColor)
            
            performance_html += """
                </table>
            </div>
            """
            
            print performance_html
            
            # Add a button to toggle debug info
            print """
            <script>
                document.addEventListener('DOMContentLoaded', function() {
                    // Initially hide all debug elements
                    var debugElements = document.querySelectorAll('.debug');
                    for (var i = 0; i < debugElements.length; i++) {
                        debugElements[i].style.display = 'none';
                    }
                    
                    // Create toggle button
                    var button = document.createElement('button');
                    button.textContent = 'Show Debug Info';
                    button.style.margin = '10px 0';
                    button.style.padding = '5px 10px';
                    button.style.backgroundColor = '#337ab7';
                    button.style.color = 'white';
                    button.style.border = 'none';
                    button.style.borderRadius = '4px';
                    button.style.cursor = 'pointer';
                    
                    button.onclick = function() {
                        var isHidden = debugElements[0].style.display === 'none';
                        for (var i = 0; i < debugElements.length; i++) {
                            debugElements[i].style.display = isHidden ? 'block' : 'none';
                        }
                        button.textContent = isHidden ? 'Hide Debug Info' : 'Show Debug Info';
                    };
                    
                    // Insert button at top of page
                    document.body.insertBefore(button, document.body.firstChild);
                });
            </script>
            """
        
        # Return the summary HTML
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
    
    def generate_header_row(self, service_times):
        """Generate the table header row with service times and years."""
        global SHOW_CURRENT_WEEK_COLUMN, SHOW_PREVIOUS_YEAR_COLUMN, SHOW_YOY_CHANGE_COLUMN
        global SHOW_FYTD_COLUMNS, SHOW_FYTD_CHANGE_COLUMNS, SHOW_FYTD_AVG_COLUMN
        global SHOW_FOUR_WEEK_COMPARISON, SHOW_ENROLLMENT_COLUMN
        
        # troubleshoot with debug class
        print('<div class="debug">SHOW_CURRENT_WEEK_COLUMN = {}</div>'.format(SHOW_CURRENT_WEEK_COLUMN))
        print('<div class="debug">SHOW_PREVIOUS_YEAR_COLUMN = {}</div>'.format(SHOW_PREVIOUS_YEAR_COLUMN))
        print('<div class="debug">SHOW_ENROLLMENT_COLUMN = {}</div>'.format(SHOW_ENROLLMENT_COLUMN))
    
        # Make sure SHOW_ENROLLMENT_COLUMN is initialized
        # if not hasattr(self, '_initialized_enrollment_column'):
        #     self._initialized_enrollment_column = True
        #     # Initialize from model.Data or default to False
        #     SHOW_ENROLLMENT_COLUMN = getattr(model.Data, 'show_enrollment', 'no') == 'yes'
        
        # Debug output
        print('<div class="debug">SHOW_ENROLLMENT_COLUMN = {}</div>'.format(SHOW_ENROLLMENT_COLUMN))
    
        headers = [""] # First column for entity name
        
        # Add enrollment column if enabled (do this once in header generation)
        if SHOW_ENROLLMENT_COLUMN:
            headers.append("Enrolled")
        
        # Handle current year column independently
        if SHOW_CURRENT_WEEK_COLUMN:
            current_year_label = ReportHelper.get_year_label(self.current_year)
            headers.append(current_year_label)
        
        # Handle previous year column independently
        if SHOW_PREVIOUS_YEAR_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                year = self.current_year - i
                year_label = ReportHelper.get_year_label(year)
                headers.append(year_label)
        
        # Add service time columns
        if service_times:
            for time in service_times:
                normalized_time = self.normalize_service_time(time)
                headers.append(normalized_time)
                
            # Add Total if not already included
            if "Total" not in service_times:
                headers.append("Total")
        else:
            headers.append("Total")
    
        # Add YoY change columns if enabled
        if SHOW_YOY_CHANGE_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                headers.append("{}-{} Change".format(
                    ReportHelper.get_year_label(self.current_year - i + 1),
                    ReportHelper.get_year_label(self.current_year - i)
                ))
        
        # Add 4-week comparison columns if enabled 
        if SHOW_FOUR_WEEK_COMPARISON:
            headers.extend([
                "Last 4 Weeks",
                "PY 4 Weeks", 
                "4-Week Change"
            ])
        
        # Add FYTD columns if enabled
        if SHOW_FYTD_COLUMNS:
            ytd_prefix = ReportHelper.get_ytd_label()
            for i in range(YEARS_TO_DISPLAY):
                year = self.current_year - i
                headers.append("{} {}".format(
                    ytd_prefix,
                    ReportHelper.get_year_label(year)
                ))
                
        # Add FYTD change columns if enabled
        if SHOW_FYTD_CHANGE_COLUMNS:
            ytd_prefix = ReportHelper.get_ytd_label()
            for i in range(1, YEARS_TO_DISPLAY):
                headers.append("{} {}-{} Change".format(
                    ytd_prefix,
                    ReportHelper.get_year_label(self.current_year - i + 1),
                    ReportHelper.get_year_label(self.current_year - i)
                ))
                
        # Add FYTD average column if enabled
        if SHOW_FYTD_AVG_COLUMN:
            headers.append("{} Avg/Week".format(ReportHelper.get_ytd_label()))
    
        # Generate the HTML
        header_html = "<tr>"
        for header in headers:
            header_html += "<th>{}</th>".format(header)
        header_html += "</tr>"
        
        return header_html
    
    def generate_program_row(self, program, years_data, years_ytd_data, four_week_data=None):
        """Generate a row for a program with attendance data for multiple years."""
        global SHOW_CURRENT_WEEK_COLUMN, SHOW_PREVIOUS_YEAR_COLUMN, SHOW_YOY_CHANGE_COLUMN
        global SHOW_FYTD_COLUMNS, SHOW_FYTD_CHANGE_COLUMNS, SHOW_FYTD_AVG_COLUMN
        global SHOW_FOUR_WEEK_COMPARISON, SHOW_ENROLLMENT_COLUMN, DEFAULT_COLLAPSED
        
        # Start row HTML
        row_html = """
            <tr style="background-color: #e9f7fd;">
                <td><strong>{}</strong></td>
        """.format(program.Name)
        
        # Add enrollment column - should always be included
        if SHOW_ENROLLMENT_COLUMN:
            try:
                # Get all divisions for this program
                divisions_sql = """
                    SELECT Id FROM Division 
                    WHERE ProgId = {} 
                    AND ReportLine IS NOT NULL 
                    AND ReportLine <> ''
                """.format(program.Id)
                
                enrollment_count = 0
                divisions = q.QuerySql(divisions_sql)
                
                # Sum up enrollment for all divisions
                for division in divisions:
                    # Basic division enrollment
                    div_enrollment = self.get_division_enrollment(division.Id, self.report_date)
                    enrollment_count += div_enrollment
                    
                    # Add organization counts if not using DEFAULT_COLLAPSED
                    if not DEFAULT_COLLAPSED:
                        org_sql = self.get_organizations_sql(division.Id)
                        orgs = q.QuerySql(org_sql)
                        for org in orgs:
                            if hasattr(org, 'MemberCount') and org.MemberCount:
                                enrollment_count += org.MemberCount
                                
                row_html += "<td><strong>{}</strong></td>".format(
                    ReportHelper.format_number(enrollment_count)
                )
            except Exception as e:
                # On error, show 0
                row_html += "<td><strong>0</strong></td>"
                print('<!-- Error getting enrollment for Program {}: {} -->'.format(program.Id, str(e)))
        
        # Handle current year column independently
        if SHOW_CURRENT_WEEK_COLUMN:
            current_year_value = years_data[self.current_year]['total'] if self.current_year in years_data else 0
            row_html += "<td><strong>{}</strong></td>".format(
                ReportHelper.format_number(current_year_value)
            )
        
        # Handle current year column independently
        if SHOW_CURRENT_WEEK_COLUMN:
            current_year_value = years_data[self.current_year]['total'] if self.current_year in years_data else 0
            row_html += "<td><strong>{}</strong></td>".format(
                ReportHelper.format_number(current_year_value)
            )
        
        # Handle previous year column independently
        if SHOW_PREVIOUS_YEAR_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                year = self.current_year - i
                value = years_data[year]['total'] if year in years_data else 0
                row_html += "<td><strong>{}</strong></td>".format(
                    ReportHelper.format_number(value)
                )
        
        # Get service times from program
        order, service_times = ReportHelper.parse_program_rpt_group(program.RptGroup)
        
        # Track accounted hours
        accounted_hours = set()
        
        # Handle service times and totals
        if service_times:
            # Process each defined service time
            for time in service_times:
                hour = self.get_hour_from_service_time(time)
                if hour == "Total":
                    value = years_data[self.current_year]['total']
                else:
                    value = years_data[self.current_year]['by_hour'].get(hour, 0)
                    accounted_hours.add(hour)
                row_html += "<td><strong>{}</strong></td>".format(ReportHelper.format_number(value))
            
            # Add Total if not already included
            if "Total" not in service_times:
                total_value = years_data[self.current_year]['total']
                row_html += "<td><strong>{}</strong></td>".format(
                    ReportHelper.format_number(total_value)
                )
        else:
            # Just show total if no service times defined 
            total_value = years_data[self.current_year]['total']
            row_html += "<td><strong>{}</strong></td>".format(
                ReportHelper.format_number(total_value)
            )
        
        # Add YoY change columns if enabled
        if SHOW_YOY_CHANGE_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                current_year = self.current_year - i + 1
                prev_year = self.current_year - i
                
                current_total = years_data[current_year]['total'] if current_year in years_data else 0
                prev_total = years_data[prev_year]['total'] if prev_year in years_data else 0
                
                trend = ReportHelper.get_trend_indicator(current_total, prev_total)
                row_html += "<td><strong>{}</strong></td>".format(trend)
        
        # Add 4-week comparison columns if enabled
        if SHOW_FOUR_WEEK_COMPARISON and four_week_data:
            current_four_week_total = four_week_data['current']['total']
            prev_year_four_week_total = four_week_data['previous_year']['total']
            
            row_html += """
                <td><strong>{}</strong></td>
                <td><strong>{}</strong></td>
                <td><strong>{}</strong></td>
            """.format(
                ReportHelper.format_number(current_four_week_total),
                ReportHelper.format_number(prev_year_four_week_total),
                ReportHelper.get_trend_indicator(current_four_week_total, prev_year_four_week_total)
            )
        
        # Add FYTD columns if enabled
        if SHOW_FYTD_COLUMNS:
            for i in range(YEARS_TO_DISPLAY):
                year = self.current_year - i
                value = years_ytd_data[year]['total'] if year in years_ytd_data else 0
                row_html += "<td><strong>{}</strong></td>".format(ReportHelper.format_number(value))
        
        # Add FYTD change columns if enabled
        if SHOW_FYTD_CHANGE_COLUMNS:
            for i in range(1, YEARS_TO_DISPLAY):
                current_year = self.current_year - i + 1
                prev_year = self.current_year - i
                
                current_ytd = years_ytd_data[current_year]['total'] if current_year in years_ytd_data else 0
                prev_ytd = years_ytd_data[prev_year]['total'] if prev_year in years_ytd_data else 0
                
                ytd_trend = ReportHelper.get_trend_indicator(current_ytd, prev_ytd)
                row_html += "<td><strong>{}</strong></td>".format(ytd_trend)
        
        # Add FYTD average if enabled
        if SHOW_FYTD_AVG_COLUMN:
            weeks_elapsed = ReportHelper.get_weeks_elapsed_in_fiscal_year(
                self.report_date, 
                self.current_fiscal_start
            )
            avg_attendance = 0
            if weeks_elapsed > 0 and self.current_year in years_ytd_data:
                avg_attendance = years_ytd_data[self.current_year]['total'] / float(weeks_elapsed)
            row_html += "<td><strong>{}</strong></td>".format(ReportHelper.format_float(avg_attendance))
        
        row_html += "</tr>"
        return row_html
        
    def generate_division_row(self, division, years_data, years_ytd_data, four_week_data=None, indent=1):
        """Generate a row for a division with attendance data for multiple years."""
    
        performance_timer.start("generate_report_content")
    
        # Handle display options
        global SHOW_CURRENT_WEEK_COLUMN, SHOW_PREVIOUS_YEAR_COLUMN, SHOW_YOY_CHANGE_COLUMN
        global SHOW_FYTD_COLUMNS, SHOW_FYTD_CHANGE_COLUMNS, SHOW_FYTD_AVG_COLUMN
        global SHOW_FOUR_WEEK_COMPARISON, SHOW_ENROLLMENT_COLUMN, DEFAULT_COLLAPSED
        
        # Only toggle optional columns (processing model.Data inputs)
        if hasattr(model.Data, 'show_four_week'):
            SHOW_FOUR_WEEK_COMPARISON = model.Data.show_four_week == 'yes'
        if hasattr(model.Data, 'show_yoy_change'):
            SHOW_YOY_CHANGE_COLUMN = model.Data.show_yoy_change == 'yes'
        if hasattr(model.Data, 'show_fytd'):
            SHOW_FYTD_COLUMNS = model.Data.show_fytd == 'yes'
        if hasattr(model.Data, 'show_fytd_change'):
            SHOW_FYTD_CHANGE_COLUMNS = model.Data.show_fytd_change == 'yes'
        if hasattr(model.Data, 'show_fytd_avg'):
            SHOW_FYTD_AVG_COLUMN = model.Data.show_fytd_avg == 'yes'
        
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
    
        # Add enrollment column - should always be included
        if SHOW_ENROLLMENT_COLUMN:
            try:
                # Get division enrollment count
                enrollment_count = self.get_division_enrollment(division.Id, self.report_date)
                
                # Add organization counts if not using DEFAULT_COLLAPSED
                if not DEFAULT_COLLAPSED:
                    org_sql = self.get_organizations_sql(division.Id)
                    orgs = q.QuerySql(org_sql)
                    for org in orgs:
                        if hasattr(org, 'MemberCount') and org.MemberCount:
                            enrollment_count += org.MemberCount
                            
                row_html += "<td>{}</td>".format(ReportHelper.format_number(enrollment_count))
            except Exception as e:
                # On error, show 0
                row_html += "<td>0</td>"
                print('<!-- Error getting enrollment for Division {}: {} -->'.format(division.Id, str(e)))
        
        # Handle current year column independently  
        if SHOW_CURRENT_WEEK_COLUMN:
            current_year_value = years_data[self.current_year]['total'] if self.current_year in years_data else 0
            row_html += "<td>{}</td>".format(ReportHelper.format_number(current_year_value))
        
        
        # Handle previous year column independently
        if SHOW_PREVIOUS_YEAR_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                year = self.current_year - i
                value = years_data[year]['total'] if year in years_data else 0
                row_html += "<td>{}</td>".format(ReportHelper.format_number(value))
        
        # Get service time data from parent program
        program = q.QuerySqlTop1("SELECT RptGroup FROM Program WHERE Id = {}".format(division.ProgId))
        service_times = self.parse_service_times(program.RptGroup)
        
        # Keep track of hours that are assigned to specific service times
        accounted_hours = set()
        
        # If there are no service times or just "Total", show the total attendance
        if not service_times or (len(service_times) == 1 and service_times[0] == "Total"):
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(years_data[self.current_year]['total'])
            )
        else:
            # Track service times already added
            added_times = set()
            has_total = False
            
            for time in service_times:
                # Handle 'Total' case
                if time == "Total":
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_data[self.current_year]['total'])
                    )
                    has_total = True
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
            if not has_total:
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(other_attendance)
                )
                # Add total column after "Other"
                row_html += "<td>{}</td>".format(
                    ReportHelper.format_number(years_data[self.current_year]['total'])
                )
        
        # Add year-over-year change columns if enabled
        if SHOW_YOY_CHANGE_COLUMN:
            for i in range(1, YEARS_TO_DISPLAY):
                current_year = self.current_year - i + 1
                prev_year = self.current_year - i
                
                current_total = years_data[current_year]['total'] if current_year in years_data else 0
                prev_total = years_data[prev_year]['total'] if prev_year in years_data else 0
                
                trend = ReportHelper.get_trend_indicator(current_total, prev_total)
                row_html += "<td>{}</td>".format(trend)
        
        # Add 4-week comparison columns if enabled
        if SHOW_FOUR_WEEK_COMPARISON and four_week_data:
            current_four_week_total = four_week_data['current']['total']
            prev_year_four_week_total = four_week_data['previous_year']['total']
            
            # Add current 4-week total
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(current_four_week_total)
            )
            
            # Add previous year's 4-week total
            row_html += "<td>{}</td>".format(
                ReportHelper.format_number(prev_year_four_week_total)
            )
            
            # Add 4-week trend indicator
            four_week_trend = ReportHelper.get_trend_indicator(current_four_week_total, prev_year_four_week_total)
            row_html += "<td>{}</td>".format(four_week_trend)
        
        # Add YTD data for each year if enabled
        if SHOW_FYTD_COLUMNS:
            for i in range(YEARS_TO_DISPLAY):
                year = self.current_year - i
                if year in years_ytd_data:
                    row_html += "<td>{}</td>".format(
                        ReportHelper.format_number(years_ytd_data[year]['total'])
                    )
                else:
                    row_html += "<td>0</td>"
        
        # Add YTD change columns if enabled
        if SHOW_FYTD_CHANGE_COLUMNS:
            for i in range(1, YEARS_TO_DISPLAY):
                current_year = self.current_year - i + 1
                prev_year = self.current_year - i
                
                current_ytd = years_ytd_data[current_year]['total'] if current_year in years_ytd_data else 0
                prev_ytd = years_ytd_data[prev_year]['total'] if prev_year in years_ytd_data else 0
                
                ytd_trend = ReportHelper.get_trend_indicator(current_ytd, prev_ytd)
                row_html += "<td>{}</td>".format(ytd_trend)
        
        # Add YTD Average if enabled
        if SHOW_FYTD_AVG_COLUMN:
            # Using weeks elapsed instead of days with meetings
            weeks_elapsed = ReportHelper.get_weeks_elapsed_in_fiscal_year(self.report_date, self.current_fiscal_start)
            avg_attendance = 0
            if weeks_elapsed > 0 and self.current_year in years_ytd_data:
                avg_attendance = years_ytd_data[self.current_year]['total'] / float(weeks_elapsed)
            row_html += "<td>{}</td>".format(ReportHelper.format_float(avg_attendance))
        
        row_html += "</tr>"
        return row_html

    def generate_report_content(self, for_email=False):
        """Generate just the report content (used for both display and email)."""
        
        performance_timer.start("generate_report_content")
    
        global SHOW_ENROLLMENT_COLUMN
        
        # Debug output to verify the setting
        print('<div class="debug">SHOW_ENROLLMENT_COLUMN after init = {}</div>'.format(SHOW_ENROLLMENT_COLUMN))
    
        
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
        four_week_totals = {'current': 0, 'previous_year': 0}  # For 4-week comparison
        
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
            
            # Get four-week comparison data if enabled
            if SHOW_FOUR_WEEK_COMPARISON:
                four_week_data = self.get_four_week_attendance_comparison(program_id=program.Id)
                four_week_totals['current'] += four_week_data['current']['total']
                four_week_totals['previous_year'] += four_week_data['previous_year']['total']
            
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
            ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0,
            four_week_totals if SHOW_FOUR_WEEK_COMPARISON else None
        )
        
        # Add fiscal year-to-date summary (only once)
        report_content += self.generate_fiscal_year_summary(
            ytd_overall_totals[self.current_year],
            ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0
        )
                
        # Add enrollment analysis section if any configured programs are present
        if any(prog in ENROLLMENT_ANALYSIS_PROGRAMS for prog in [p.Name for p in programs]):
            report_content += self.generate_enrollment_analysis_section()
        
        # Get data for program totals section
        current_week_data = self.get_week_attendance_data(self.week_start_date, self.report_date)
        previous_week_data = self.get_week_attendance_data(
            self.previous_week_start_date, 
            self.previous_year_date
        )
        
        # Add program totals section if enabled
        if SHOW_PROGRAM_SUMMARY:
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
            
            # Get four-week comparison data if enabled
            four_week_data = None
            if SHOW_FOUR_WEEK_COMPARISON:
                four_week_data = self.get_four_week_attendance_comparison(program_id=program.Id)
            
            # Add program row
            report_content += self.generate_program_row(
                program, 
                years_program_data, 
                years_program_ytd,
                four_week_data
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
                
                # Get four-week comparison data for division if enabled
                division_four_week_data = None
                if SHOW_FOUR_WEEK_COMPARISON:
                    division_four_week_data = self.get_four_week_attendance_comparison(division_id=division.Id)
                
                # # Add division row
                division_row = self.generate_division_row(
                    division, 
                    years_division_data, 
                    years_division_ytd,
                    division_four_week_data
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
                        
                        # Get four-week comparison data for organization if enabled
                        org_four_week_data = None
                        if SHOW_FOUR_WEEK_COMPARISON:
                            org_four_week_data = self.get_four_week_attendance_comparison(org_id=org.OrganizationId)
                            
                        performance_timer.log("org_processing_{}".format(org.OrganizationId))
                        
                        # Add organization row
                        report_content += self.generate_organization_row(
                            org, 
                            years_org_data, 
                            years_org_ytd,
                            org_four_week_data
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
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
                        
                        // Store form values in localStorage to persist settings
                        var formInputs = reportForm.querySelectorAll('input[type="checkbox"]');
                        for (var i = 0; i < formInputs.length; i++) {
                            var input = formInputs[i];
                            if (input.name && input.name !== 'collapse_orgs') { // Don't save collapse_orgs state
                                localStorage.setItem('attendance_' + input.name, input.checked ? 'yes' : 'no');
                            }
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
                    
                    // Load saved settings from localStorage if available
                    var formInputs = reportForm.querySelectorAll('input[type="checkbox"]');
                    for (var i = 0; i < formInputs.length; i++) {
                        var input = formInputs[i];
                        if (input.name && input.name !== 'collapse_orgs') { // Don't load collapse_orgs state
                            var savedValue = localStorage.getItem('attendance_' + input.name);
                            if (savedValue === 'yes' && !input.disabled) {
                                input.checked = true;
                            } else if (savedValue === 'no' && !input.disabled) {
                                input.checked = false;
                            }
                        }
                    }
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
    
# Create and run the report
try:
    # Check for display options from URL parameters
    show_program_summary = getattr(model.Data, 'show_program_summary', None)
    if show_program_summary == 'yes':
        SHOW_PROGRAM_SUMMARY = True
    elif show_program_summary == 'no':
        SHOW_PROGRAM_SUMMARY = False
    
    show_four_week = getattr(model.Data, 'show_four_week', None)
    if show_four_week == 'yes':
        SHOW_FOUR_WEEK_COMPARISON = True
    elif show_four_week == 'no':
        SHOW_FOUR_WEEK_COMPARISON = False
    
    show_current_week = getattr(model.Data, 'show_current_week', None)
    if show_current_week == 'yes':
        SHOW_CURRENT_WEEK_COLUMN = True
    elif show_current_week == 'no':
        SHOW_CURRENT_WEEK_COLUMN = False
        
    show_previous_year = getattr(model.Data, 'show_previous_year', None)
    if show_previous_year == 'yes':
        SHOW_PREVIOUS_YEAR_COLUMN = True
    elif show_previous_year == 'no':
        SHOW_PREVIOUS_YEAR_COLUMN = False
    
    show_yoy_change = getattr(model.Data, 'show_yoy_change', None)
    if show_yoy_change == 'yes':
        SHOW_YOY_CHANGE_COLUMN = True
    elif show_yoy_change == 'no':
        SHOW_YOY_CHANGE_COLUMN = False
    
    show_fytd = getattr(model.Data, 'show_fytd', None)
    if show_fytd == 'yes':
        SHOW_FYTD_COLUMNS = True
    elif show_fytd == 'no':
        SHOW_FYTD_COLUMNS = False
    
    show_fytd_change = getattr(model.Data, 'show_fytd_change', None)
    if show_fytd_change == 'yes':
        SHOW_FYTD_CHANGE_COLUMNS = True
    elif show_fytd_change == 'no':
        SHOW_FYTD_CHANGE_COLUMNS = False
    
    show_fytd_avg = getattr(model.Data, 'show_fytd_avg', None)
    if show_fytd_avg == 'yes':
        SHOW_FYTD_AVG_COLUMN = True
    elif show_fytd_avg == 'no':
        SHOW_FYTD_AVG_COLUMN = False
        
    show_detailed_enrollment = getattr(model.Data, 'show_detailed_enrollment', None)
    if show_detailed_enrollment == 'yes':
        SHOW_DETAILED_ENROLLMENT = True
    elif show_detailed_enrollment == 'no':
        SHOW_DETAILED_ENROLLMENT = False
    
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
