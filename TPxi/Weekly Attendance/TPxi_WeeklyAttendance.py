#####################################################################
#### REPORT INFORMATION
#####################################################################
# Weekly Attendance Dashboard
#####################################################################
#  Key Features

#  - Real-time attendance tracking across multiple programs/divisions
#  - Interactive date range selector with quick presets (This Week, Last Week, Last 30/90 Days)
#  - Multi-level filtering - Campus, Program, Division, and Organization
#  - Visual charts showing attendance trends over time
#  - Sortable data table with attendance counts and percentages
#  - Export capabilities - Download data as CSV
#  - Responsive design - Works on desktop and mobile
#  - Dynamic organization hierarchy - Filters update based on selections
#  - Attendance metrics including:
#    - Total attendance per period
#    - Average attendance
#    - Attendance trends
#    - Organization-level breakdowns

#  Data Points Tracked

#  - Meeting dates
#  - Organization names and IDs
#  - Attendance counts
#  - Program/Division associations
#  - Campus assignments
#  - Historical attendance patterns

#written by: Ben Swaby
#email: bswaby@fbchtn.org

# Update Notes 11/13/2025:
# - Fixed bug in unique attendance counting.  It was not honoring week-at-a-glance reporting groups.    
# - Added prospects and guests () to unique attendance counting.

# Update Notes 11/12/2025:
# - Added ability to send to self or saved search.
# - Fixed report not sending enrollment metrics
# - Fixed report not sending week summary
# - Added date to email subject
# - Cleaned up some spinner and duplicate notications after email is sent

# Update Notes 10/26/2025:
# - Changed to allow 0 attendance count for enrollment
# - Changed from unique counts to sum count for enrollment metrics, program metrics, and total enrollment

# Update Notes 8/28/2025:
# - Added exception to add notes on things that affect attendance (weather, Easter, etc..).  these exceptions will automatically display if they are within the 4 week period 
# - Added actual and YTD KPI metrics
# - Added unique counts for involvements, enrollment, and 4 week attendance
# - Improved performance
# - Several misc bug fixes.

# Update Notes 12/24/2024:
# - Updated unique attendance counting to use program-specific time windows (StartHoursOffset/EndHoursOffset)
# - Ensures accurate unique people counts for programs like Adults (-144 to 24 hours) and Preschool (-24 to 24 hours)
# - Fixed enrollment queries to use DivOrg table for correct organization-division relationships
# - All enrollment counts properly exclude prospects (MemberTypeId = 230)
# - Active organization filtering (OrganizationStatusId = 30) applied consistently
# - Added exceptions feature to track and display special Sundays (holidays, weather, etc.)

#####################################################################
#### USER CONFIG FIELDS (Modify these settings as needed)
#####################################################################
# Report Title (displayed at top of page)
REPORT_TITLE = "Weekly at a Glance 2.0"

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
# Default email recipient: "self" to send to current user, or email address/PeopleId
# Can also use "attendance_group" to send to a specific group (configure in TouchPoint)
EMAIL_TO = "SwabyTest"  # Options: "self", email address, or PeopleId

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

# Include all enrollments in metrics regardless of attendance data
# Set to True to count ALL enrolled members (true enrollment count)
# Set to False to only count members who had attendance during the period (legacy behavior)
INCLUDE_ALL_ENROLLMENTS = True

# Exceptions Configuration
# Name of the TextContent that stores exception dates
EXCEPTIONS_CONTENT_NAME = "WeeklyAttendance_Exceptions"
# Enable/disable showing exceptions in the report
SHOW_EXCEPTIONS = True
# Set to True to require Admin role for managing (adding/deleting) exceptions
# Non-admins will still be able to VIEW exceptions if SHOW_EXCEPTIONS is True
REQUIRE_ADMIN_FOR_EXCEPTION_MANAGEMENT = True

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
        report += "<tr><th style='text-align: left; padding: 4px 8px; border-bottom: 1px solid #ddd;'>Section</th>"
        report += "<th style='text-align: right; padding: 4px 8px; border-bottom: 1px solid #ddd;'>Time (seconds)</th></tr>"
        
        # Sort results by time (descending)
        sorted_results = sorted(self.results.items(), key=lambda x: x[1], reverse=True)
        
        for section, elapsed in sorted_results:
            report += "<tr><td style='padding: 4px 8px; border-bottom: 1px solid #eee;'>{}</td>".format(section)
            report += "<td style='text-align: right; padding: 4px 8px; border-bottom: 1px solid #eee;'>{:.4f}</td></tr>".format(elapsed)
        
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
        # Round up to nearest week since worship services happen weekly
        # Add 6 to round up (e.g., 9 days = (9+6)/7 = 2.14 -> int = 2 weeks)
        weeks_elapsed = max(1, int((days_elapsed + 6) / 7))  # Ensure at least 1 week
        
        return weeks_elapsed
    
    @staticmethod
    def count_sundays_between(start_date, end_date):
        """Count the number of Sundays between two dates (inclusive)."""
        from datetime import timedelta
        
        # Ensure we have datetime objects
        if isinstance(start_date, str):
            start_date = ReportHelper.parse_date_string(start_date)
        if isinstance(end_date, str):
            end_date = ReportHelper.parse_date_string(end_date)
            
        # Start from the first Sunday on or after start_date
        # If start_date is a Sunday, start from that date
        if start_date.weekday() == 6:  # Sunday
            current_sunday = start_date
        else:
            # Find next Sunday
            days_until_sunday = (6 - start_date.weekday()) % 7
            if days_until_sunday == 0:
                days_until_sunday = 7
            current_sunday = start_date + timedelta(days=days_until_sunday)
        
        # Count Sundays
        sunday_count = 0
        while current_sunday <= end_date:
            sunday_count += 1
            current_sunday += timedelta(days=7)
            
        return max(1, sunday_count)  # Return at least 1
    
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
    def get_previous_year_sunday(sunday_date):
        """Get the Sunday from the same week in the previous year."""
        # Go back exactly 52 weeks (364 days) to maintain Sunday alignment
        prev_year = sunday_date - datetime.timedelta(weeks=52)
        
        # Make sure it's a Sunday
        if prev_year.weekday() != 6:  # 6 is Sunday
            # Adjust to the nearest Sunday
            days_to_sunday = (6 - prev_year.weekday()) % 7
            if days_to_sunday == 0 and prev_year.weekday() != 6:
                days_to_sunday = 7
            prev_year = prev_year + datetime.timedelta(days=days_to_sunday)
        
        return prev_year
    
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

class ExceptionsManager:
    """Manager for handling attendance exceptions (holidays, special events, etc.)"""
    
    def __init__(self):
        """Initialize the exceptions manager."""
        self.exceptions = []
        self.load_exceptions()
    
    def load_exceptions(self):
        """Load exceptions from TextContent."""
        try:
            # Get the exceptions content
            content = model.TextContent(EXCEPTIONS_CONTENT_NAME)
            if content:
                # Parse the content (enhanced format: YYYY-MM-DD | Description | Flags)
                # Flags: W=Worship, G=Groups, B=Both
                lines = content.split('\n')
                for line in lines:
                    line = line.strip()
                    # Skip empty lines and comments
                    if line and not line.startswith('#') and '|' in line:
                        parts = line.split('|')
                        if len(parts) >= 2:
                            date_str = parts[0].strip()
                            description = parts[1].strip()
                            # Get flags if present (default to Both)
                            flags = parts[2].strip() if len(parts) >= 3 else 'B'
                            
                            try:
                                # Parse the date
                                exception_date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
                                self.exceptions.append({
                                    'date': exception_date,
                                    'description': description,
                                    'affects_worship': 'W' in flags.upper() or 'B' in flags.upper(),
                                    'affects_groups': 'G' in flags.upper() or 'B' in flags.upper()
                                })
                            except ValueError:
                                # Skip invalid date formats
                                continue
        except:
            # If content doesn't exist or error loading, continue without exceptions
            pass
    
    def save_exceptions(self):
        """Save exceptions back to TextContent (sorted by date)."""
        try:
            # Add header comments
            lines = [
                "# Weekly Attendance Exceptions",
                "# Format: YYYY-MM-DD | Description | Flags",
                "# Flags: W=Worship only, G=Groups only, B=Both (default)",
                "#",
                ""
            ]
            
            # Format exceptions for saving (sorted by date)
            for exception in sorted(self.exceptions, key=lambda x: x['date']):
                date_str = exception['date'].strftime('%Y-%m-%d')
                # Determine flag
                affects_worship = exception.get('affects_worship', True)
                affects_groups = exception.get('affects_groups', True)
                
                if affects_worship and affects_groups:
                    flag = 'B'
                elif affects_worship:
                    flag = 'W'
                elif affects_groups:
                    flag = 'G'
                else:
                    flag = 'B'  # Default to both if neither
                
                lines.append("{} | {} | {}".format(date_str, exception['description'], flag))
            
            content = '\n'.join(lines)
            
            # Save to TextContent using the TouchPoint API
            # WriteContentText parameters: name, text, keyword (optional)
            model.WriteContentText(EXCEPTIONS_CONTENT_NAME, content)
            
            # Debug: Verify the save worked by reading it back
            saved_content = model.TextContent(EXCEPTIONS_CONTENT_NAME)
            if saved_content and saved_content.strip() != content.strip():
                print "Warning: Content mismatch after save"
                return False
            
            return True
        except Exception as e:
            print "Error saving exceptions: {}".format(str(e))
            return False
    
    def add_exception(self, date_obj, description, affects_worship=True, affects_groups=True):
        """Add a new exception."""
        self.exceptions.append({
            'date': date_obj,
            'description': description,
            'affects_worship': affects_worship,
            'affects_groups': affects_groups
        })
        return self.save_exceptions()
    
    def remove_exception(self, date_obj):
        """Remove an exception by date."""
        self.exceptions = [e for e in self.exceptions if e['date'] != date_obj]
        return self.save_exceptions()
    
    def get_exceptions_in_range(self, start_date, end_date):
        """Get all exceptions within a date range."""
        exceptions_in_range = []
        for exception in self.exceptions:
            if start_date <= exception['date'] <= end_date:
                exceptions_in_range.append(exception)
        return sorted(exceptions_in_range, key=lambda x: x['date'])
    
    def get_exception_for_date(self, date_obj):
        """Get exception for a specific date if it exists."""
        for exception in self.exceptions:
            if exception['date'].date() == date_obj.date():
                return exception
        return None
    
    def format_exceptions_html(self, exceptions):
        """Format exceptions list as HTML for display."""
        if not exceptions:
            return ""
        
        html = '<div style="background-color: #fff9e6; border: 1px solid #ffc107; padding: 10px; margin: 10px 0; border-radius: 5px;">'
        html += '<strong style="color: #856404;">⚠️ Note: The following exceptions occurred during this period:</strong><ul style="margin: 5px 0;">'
        
        for exception in sorted(exceptions, key=lambda x: x['date']):
            date_str = exception['date'].strftime('%m/%d/%Y')
            affects_worship = exception.get('affects_worship', True)
            affects_groups = exception.get('affects_groups', True)
            
            # Add visual indicators for what's affected
            indicators = []
            if affects_worship:
                indicators.append('<span style="color: #e91e63;">⛪ Worship</span>')
            if affects_groups:
                indicators.append('<span style="color: #2196f3;">👥 Groups</span>')
            
            affects_text = ' [{}]'.format(', '.join(indicators)) if indicators else ''
            html += '<li><strong>{}</strong>: {}{}</li>'.format(date_str, exception['description'], affects_text)
        
        html += '</ul></div>'
        return html
    
    def get_sample_content(self):
        """Get sample content for the exceptions file."""
        return """# Weekly Attendance Exceptions
# Format: YYYY-MM-DD | Description | Flags
# Flags: W=Worship only, G=Groups only, B=Both (default)
#

2024-12-25 | Christmas Day - No services | B
2024-12-29 | Combined Service - Single service only | W
2025-01-01 | New Year's Day - Modified schedule | B
2025-02-16 | Winter Storm - Services cancelled | B
2025-04-20 | Easter Sunday - Higher than normal attendance expected | W
2025-05-25 | Memorial Day Weekend - Lower attendance expected | B
2025-07-04 | July 4th - No evening activities | G
2025-07-06 | July 4th Weekend - Lower attendance expected | W
2025-09-07 | Labor Day Weekend - Lower attendance expected | B
2025-11-27 | Thanksgiving - No groups meeting | G
2025-11-30 | Thanksgiving Weekend - Lower attendance expected | W
2025-12-24 | Christmas Eve - Special services | W
2025-12-28 | Combined Service - Single service only | W"""

class AttendanceReport:
    """Main class to generate the attendance dashboard"""
    
    def __init__(self):
        """Initialize the report with default values."""
        self.current_year = ReportHelper.get_current_year()
        self.year_labels = []
        self.years = []
        
        # Initialize exceptions manager if enabled
        self.exceptions_manager = None
        if SHOW_EXCEPTIONS:
            try:
                self.exceptions_manager = ExceptionsManager()
            except:
                # If exceptions manager fails to load, continue without it
                pass
        
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
            
            # Get 4-week period dates (exactly 28 days before report date)
            self.four_weeks_ago_date = ReportHelper.get_date_from_weeks_ago(self.report_date, 4)
            # For 4-week average, use exactly 4 weeks (28 days) from the report date
            # Don't adjust to week start to avoid including extra days
            self.four_weeks_ago_start_date = self.report_date - datetime.timedelta(days=27)
            
            # Get previous year's 4-week period dates
            self.prev_year_four_weeks_ago_date = ReportHelper.get_date_from_weeks_ago(self.previous_year_date, 4)
            # For 4-week average, use exactly 4 weeks (28 days) from the previous year date
            # Don't adjust to week start to avoid including extra days  
            self.prev_year_four_weeks_ago_start_date = self.previous_year_date - datetime.timedelta(days=27)
            
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

        This counts ALL enrollments in organizations in this division.
        People in multiple organizations are counted multiple times (once per enrollment).
        Only includes organizations that are part of the reporting structure.
        """
        # Format date correctly
        if hasattr(report_date, 'strftime'):
            date_str = report_date.strftime('%Y-%m-%d')
        else:
            date_str = str(report_date)

        # SQL query to count ALL enrolled members (not DISTINCT - count each enrollment)
        sql = """
            SELECT COUNT(*) AS EnrolledCount
            FROM OrganizationMembers om
            JOIN Organizations o ON o.OrganizationId = om.OrganizationId
            JOIN OrganizationStructure os ON os.OrgId = om.OrganizationId
            JOIN Division d ON d.Id = os.DivId
            JOIN Program p ON p.Id = d.ProgId
            JOIN People pe ON pe.PeopleId = om.PeopleId
            WHERE os.DivId = {0}
            AND o.OrganizationStatusId = 30  -- Only active organizations
            AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
            AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''  -- Only divisions in reporting
            AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{1}')  -- Enrolled by report date
            AND (om.InactiveDate IS NULL OR om.InactiveDate > '{1}')  -- Still active on report date
            AND om.MemberTypeId NOT IN (230,311)  -- Exclude prospects (230) and other non-members (311)
            AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > '{1}')  -- Not deceased as of report date
            AND (pe.DropDate IS NULL OR pe.DropDate > '{1}')  -- Not dropped/archived as of report date
        """.format(division_id, date_str)
            
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
    
    def get_all_programs_divisions_organizations(self):
        """Fetch all programs, divisions, and organizations in a single query."""
        sql = """
        SELECT 
            p.Id AS ProgramId, 
            p.Name AS ProgramName,
            p.RptGroup AS ProgramRptGroup,
            p.StartHoursOffset,
            p.EndHoursOffset,
            d.Id AS DivisionId,
            d.Name AS DivisionName,
            d.ReportLine AS DivisionReportLine,
            d.NoDisplayZero AS DivisionNoDisplayZero,
            o.OrganizationId,
            o.OrganizationName,
            o.MemberCount
        FROM Program p
        LEFT JOIN Division d ON d.ProgId = p.Id AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''
        LEFT JOIN OrganizationStructure os ON os.DivId = d.Id
        LEFT JOIN Organizations o ON os.OrgId = o.OrganizationId AND o.OrganizationStatusId = 30
        WHERE p.RptGroup IS NOT NULL AND p.RptGroup <> ''
        ORDER BY p.RptGroup, d.ReportLine, o.OrganizationName
        """
        return q.QuerySql(sql)
    
    def get_batch_attendance_data(self, date_ranges):
        """
        Fetch attendance data for multiple date ranges in a single query.
        
        Parameters:
            date_ranges: Dictionary with keys as range identifiers and values as (start_date, end_date) tuples
        """
        # Build UNION ALL query for all date ranges
        query_parts = []
        
        for range_id, (start_date, end_date) in date_ranges.items():
            # Format dates for SQL
            start_date_str = ReportHelper.format_date(start_date)
            end_date_str = ReportHelper.format_date(end_date)
            
            # Create query part for this date range
            query_part = """
            SELECT 
                '{range_id}' AS RangeId,
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
            JOIN Program p ON p.Id = d.ProgId
            WHERE CONVERT(date, m.MeetingDate) BETWEEN '{start_date}' AND '{end_date}'
            AND (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL)
            AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''
            AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''
            """.format(
                range_id=range_id,
                start_date=start_date_str,
                end_date=end_date_str
            )
            
            query_parts.append(query_part)
        
        # Combine all query parts with UNION ALL
        full_query = " UNION ALL ".join(query_parts)
        
        # Add final select to aggregate results
        final_query = """
        WITH CombinedAttendance AS (
            {0}
        )
        SELECT 
            RangeId,
            OrganizationId,
            DivId AS DivisionId,
            ProgId AS ProgramId,
            ProgramName,
            MeetingHour,
            SUM(AttendCount) as TotalAttendance,
            COUNT(*) as MeetingCount
        FROM CombinedAttendance
        GROUP BY RangeId, OrganizationId, DivId, ProgId, ProgramName, MeetingHour
        """.format(full_query)
        
        return q.QuerySql(final_query)

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
    
    def get_unique_involvements_count(self, program_names):
        """Get distinct count of involvements (enrollments) for specified programs."""
        involvement_data = {
            'total': 0,
            'by_program': {},
            'org_count': 0,
            'org_by_program': {}
        }
        
        if not program_names or len(program_names) == 0:
            return involvement_data
        
        # Format report date for SQL
        report_date_str = self.report_date.strftime('%Y-%m-%d') if hasattr(self.report_date, 'strftime') else str(self.report_date)
        
        # Get overall total unique across all programs first
        safe_names = []
        for name in program_names:
            safe_name = name.replace("'", "''")
            safe_names.append("'{}'".format(safe_name))
        
        # Get per-program unique counts (people may be in multiple orgs within a program)
        # Using same filters as other enrollment queries for consistency
        sql = """
        SELECT
            p.Name AS ProgramName,
            COUNT(DISTINCT om.PeopleId) AS UniqueInvolvements,
            COUNT(DISTINCT o.OrganizationId) AS OrgCount
        FROM Program p
        INNER JOIN Division d ON d.ProgId = p.Id
        INNER JOIN OrganizationStructure os ON os.DivId = d.Id
        INNER JOIN Organizations o ON o.OrganizationId = os.OrgId
        INNER JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
        INNER JOIN People pe ON pe.PeopleId = om.PeopleId
        WHERE p.Name IN ({0})
            AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
            AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''  -- Only divisions in reporting
            AND o.OrganizationStatusId = 30  -- Active organizations
            AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{1}')  -- Enrolled by report date
            AND (om.InactiveDate IS NULL OR om.InactiveDate > '{1}')  -- Still active on report date
            AND om.MemberTypeId NOT IN (230, 311)  -- Exclude prospects and non-members
            AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > '{1}')  -- Not deceased as of report date
            AND (pe.DropDate IS NULL OR pe.DropDate > '{1}')  -- Not dropped/archived as of report date
        GROUP BY p.Name
        """.format(','.join(safe_names), report_date_str)
        
        try:
            results = q.QuerySql(sql)
            for row in results:
                program_name = row.ProgramName
                unique_count = row.UniqueInvolvements or 0
                org_count = row.OrgCount or 0
                involvement_data['by_program'][program_name] = unique_count
                involvement_data['org_by_program'][program_name] = org_count
            
            # Calculate total unique across all programs (someone might be in multiple programs)
            # Using same filters as other enrollment queries for consistency
            total_sql = """
            SELECT
                COUNT(DISTINCT om.PeopleId) AS TotalUnique,
                COUNT(DISTINCT o.OrganizationId) AS TotalOrgs
            FROM Program p
            INNER JOIN Division d ON d.ProgId = p.Id
            INNER JOIN OrganizationStructure os ON os.DivId = d.Id
            INNER JOIN Organizations o ON o.OrganizationId = os.OrgId
            INNER JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
            INNER JOIN People pe ON pe.PeopleId = om.PeopleId
            WHERE p.Name IN ({0})
                AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
                AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''  -- Only divisions in reporting
                AND o.OrganizationStatusId = 30  -- Active organizations
                AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{1}')  -- Enrolled by report date
                AND (om.InactiveDate IS NULL OR om.InactiveDate > '{1}')  -- Still active on report date
                AND om.MemberTypeId NOT IN (230, 311)  -- Exclude prospects and non-members
                AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > '{1}')  -- Not deceased as of report date
                AND (pe.DropDate IS NULL OR pe.DropDate > '{1}')  -- Not dropped/archived as of report date
            """.format(','.join(safe_names), report_date_str)
            
            total_results = q.QuerySql(total_sql)
            if total_results and len(total_results) > 0:
                involvement_data['total'] = total_results[0].TotalUnique or 0
                involvement_data['org_count'] = total_results[0].TotalOrgs or 0
            
        except Exception as e:
            print "<!-- Error getting unique involvements: {} -->".format(str(e))
        
        return involvement_data
    
    def get_unique_attendance_data(self, program_names, start_date, end_date):
        """Get distinct count of people who attended for specified programs across last 4 Sundays.
        Returns both total attendees and guest count (not enrolled in OrganizationMembers)."""
        unique_data = {
            'total': 0,
            'by_program': {},
            'guests': 0,
            'guests_by_program': {}
        }

        if not program_names or len(program_names) == 0:
            return unique_data

        # Build safe SQL for program names
        safe_names = []
        for name in program_names:
            safe_name = name.replace("'", "''")
            safe_names.append("'{}'".format(safe_name))

        # Calculate the last 4 Sundays from end_date
        # Note: SQL Server DATEPART(dw, date) returns 1 for Sunday
        end_date_str = str(end_date).split(' ')[0] if ' ' in str(end_date) else str(end_date)

        # Get unique attendance AND guest counts using program-specific time windows
        # This respects each program's StartHoursOffset and EndHoursOffset
        sql = """
        WITH ProgramWindows AS (
            -- Generate 4 Sunday dates with program-specific time windows
            SELECT p.Id AS ProgramId, p.Name AS ProgramName,
                   p.StartHoursOffset, p.EndHoursOffset, s.SundayDate,
                   DATEADD(HOUR, ISNULL(p.StartHoursOffset, 0), s.SundayDate) AS WindowStart,
                   DATEADD(HOUR, ISNULL(p.EndHoursOffset, 24), s.SundayDate) AS WindowEnd
            FROM Program p
            CROSS JOIN (
                SELECT DATEADD(dd, 1-DATEPART(dw, '{1}'), '{1}') AS SundayDate  -- Most recent Sunday
                UNION ALL SELECT DATEADD(week, -1, DATEADD(dd, 1-DATEPART(dw, '{1}'), '{1}'))
                UNION ALL SELECT DATEADD(week, -2, DATEADD(dd, 1-DATEPART(dw, '{1}'), '{1}'))
                UNION ALL SELECT DATEADD(week, -3, DATEADD(dd, 1-DATEPART(dw, '{1}'), '{1}'))
            ) s
            WHERE p.Name IN ({0})
                AND (p.StartHoursOffset IS NOT NULL OR p.EndHoursOffset IS NOT NULL)
        )
        SELECT
            pw.ProgramName,
            -- Count only enrolled members (excluding guests and prospects)
            -- Use Attend.MemberTypeId which captures their status at time of attendance
            COUNT(DISTINCT CASE
                WHEN a.MemberTypeId IS NOT NULL
                    AND a.MemberTypeId NOT IN (230, 311)
                THEN a.PeopleId
            END) AS UniquePeople,
            -- Count guests (not enrolled) + prospects (MemberTypeId 230, 311)
            COUNT(DISTINCT CASE
                WHEN a.MemberTypeId IS NULL
                    OR a.MemberTypeId IN (230, 311)
                THEN a.PeopleId
            END) AS GuestCount
        FROM ProgramWindows pw
        INNER JOIN Division d ON d.ProgId = pw.ProgramId
        INNER JOIN OrganizationStructure os ON os.DivId = d.Id
        INNER JOIN Organizations o ON o.OrganizationId = os.OrgId
        INNER JOIN Attend a ON a.OrganizationId = o.OrganizationId
        INNER JOIN People pe ON pe.PeopleId = a.PeopleId
        WHERE a.MeetingDate >= pw.WindowStart
            AND a.MeetingDate < pw.WindowEnd
            AND a.AttendanceFlag = 1  -- Present
            AND o.OrganizationStatusId = 30  -- Active organizations
            AND (pe.IsDeceased = 0 OR pe.IsDeceased IS NULL)  -- Exclude deceased
            AND pe.ArchivedFlag = 0  -- Exclude archived
            AND d.ReportLine IS NOT NULL  -- Only divisions included in reporting
        GROUP BY pw.ProgramName
        """.format(','.join(safe_names), end_date_str)

        try:
            results = q.QuerySql(sql)
            found_programs = set()

            print "<!-- Debug: get_unique_attendance_data received {} results -->".format(len(list(results)) if results else 0)

            for row in results:
                program_name = row.ProgramName
                unique_count = row.UniquePeople or 0
                guest_count = row.GuestCount or 0
                print "<!-- Debug: Program={}, Unique={}, Guests={} -->".format(program_name, unique_count, guest_count)
                unique_data['by_program'][program_name] = unique_count
                unique_data['guests_by_program'][program_name] = guest_count
                unique_data['total'] += unique_count
                unique_data['guests'] += guest_count
                found_programs.add(program_name)
            
            # For programs without individual attendance records, estimate based on average attendance
            # This provides a reasonable approximation when individual records aren't available
            missing_programs = set(program_names) - found_programs
            if missing_programs:
                for program_name in missing_programs:
                    # Try to estimate based on meeting data - only count enrolled members
                    # For programs using aggregate counts, we need to estimate enrolled attendance
                    estimate_sql = """
                    SELECT 
                        p.Name AS ProgramName,
                        -- Count unique enrolled members who could have attended
                        COUNT(DISTINCT om.PeopleId) AS EstimatedUnique
                    FROM Program p
                    JOIN Division d ON d.ProgId = p.Id
                    JOIN DivOrg do ON do.DivId = d.Id
                    JOIN Organizations o ON o.OrganizationId = do.OrgId
                    JOIN OrganizationMembers om ON om.OrganizationId = o.OrganizationId
                    JOIN People pe ON pe.PeopleId = om.PeopleId
                    -- Check if there were meetings in this period
                    JOIN (
                        SELECT DISTINCT m2.OrganizationId
                        FROM Meetings m2
                        WHERE CONVERT(date, m2.MeetingDate) >= '{}'
                            AND CONVERT(date, m2.MeetingDate) <= '{}'
                            AND (m2.DidNotMeet = 0 OR m2.DidNotMeet IS NULL)
                            AND (m2.HeadCount > 0 OR m2.NumPresent > 0 OR m2.MaxCount > 0)
                    ) met ON met.OrganizationId = o.OrganizationId
                    WHERE p.Name = '{}'
                        AND o.OrganizationStatusId = 30  -- Active organizations
                        AND om.Pending = 0  -- Not pending
                        AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())  -- Active members
                        AND om.MemberTypeId <> 230  -- Exclude prospects
                        AND (pe.IsDeceased = 0 OR pe.IsDeceased IS NULL)  -- Exclude deceased people
                        AND pe.ArchivedFlag = 0  -- Exclude archived people
                    GROUP BY p.Name
                    """.format(end_date_str, end_date_str, program_name.replace("'", "''"))
                    
                    try:
                        estimate_results = q.QuerySql(estimate_sql)
                        if estimate_results and len(estimate_results) > 0:
                            estimated_count = estimate_results[0].EstimatedUnique or 0
                            unique_data['by_program'][program_name] = estimated_count
                            unique_data['guests_by_program'][program_name] = 0  # No guest data for estimates
                            unique_data['total'] += estimated_count
                        else:
                            unique_data['by_program'][program_name] = 0
                            unique_data['guests_by_program'][program_name] = 0
                    except:
                        unique_data['by_program'][program_name] = 0
                        unique_data['guests_by_program'][program_name] = 0
                        
        except Exception as e:
            print "<!-- Error getting unique attendance: {} -->".format(str(e))
        
        return unique_data
    
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
            # Determine default recipient based on EMAIL_TO configuration
            default_self_checked = 'checked' if EMAIL_TO == "self" else ''
            default_group_checked = 'checked' if EMAIL_TO != "self" else ''
            default_email_value = '' if EMAIL_TO == "self" else EMAIL_TO

            send_email_html = """
            <div style="margin-top: 10px; padding: 10px; background-color: #e8f4f8; border-radius: 5px;">
                <strong>Send Email Report:</strong>
                <div style="margin-top: 5px;">
                    <label style="margin-right: 15px;">
                        <input type="radio" name="email_to" value="self" {0}> Send to Self
                    </label>
                    <label style="margin-right: 15px;">
                        <input type="radio" name="email_to" value="{1}" {2}> Send to Attendance Report Group
                    </label>
                </div>
                <button type="submit" name="send_email" value="yes" id="send-email-btn"
                        style="margin-top: 10px; padding: 5px 15px; background-color: #2196F3; color: white; border: none; border-radius: 3px; cursor: pointer; position: relative;">
                    📧 Send Report
                </button>
            </div>
            """.format(default_self_checked, default_email_value, default_group_checked)
        
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
                
                <button type="button" onclick="toggleExceptionsPanel()" style="padding: 5px 10px; background-color: #FF9800; color: white; border: none; border-radius: 3px; cursor: pointer;">
                    ⚠️ Manage Exceptions
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

        # Determine HAVING clause based on INCLUDE_ALL_ENROLLMENTS setting
        if INCLUDE_ALL_ENROLLMENTS:
            # Include all enrollments, even those with zero attendance
            having_clause_meetings = ""
            having_clause_last_sunday = ""
        else:
            # Legacy behavior: only include organizations with attendance
            having_clause_meetings = "HAVING SUM(CAST(COALESCE(m.MaxCount, 0) AS FLOAT)) > 0"
            having_clause_last_sunday = "HAVING SUM(CAST(COALESCE(m.MaxCount, 0) AS FLOAT)) > 0"

        # Format SQL
        sql = """
        -- Define variables at the top
        DECLARE @WeeksCount INT = {weeks_count}; -- Number of weeks between start and end dates
        DECLARE @StartDate DATE = '{start_date}'; -- Start date
        DECLARE @EndDate DATE = '{end_date}'; -- End date
        DECLARE @LastSundayDate DATE = '{last_sunday_date}'; -- Last Sunday date
        DECLARE @SelectedDate DATE = '{selected_date}'; -- The selected report date
        DECLARE @NeedsInReachThreshold INT = {needs_inreach}; -- Threshold for "Needs In-reach" category
        DECLARE @GoodRatioThreshold INT = {good_ratio}; -- Threshold for "Good Ratio" category
        DECLARE @IncludeOrgDetails BIT = {include_org_details}; -- Set to 1 to include Organization name and ID in results

        -- Calculate the Sunday for the current week
        DECLARE @LastSunday DATE = DATEADD(DAY, 
            CASE 
                WHEN DATEPART(WEEKDAY, @SelectedDate) = 1 THEN 0 -- If already Sunday, use current date
                ELSE -(DATEPART(WEEKDAY, @SelectedDate) - 1) -- Otherwise go back to current week's Sunday
            END, 
            @SelectedDate);
        
        -- Debug output
        PRINT 'Selected Date: ' + CAST(@SelectedDate AS VARCHAR(20));
        PRINT 'Last Sunday: ' + CAST(@LastSunday AS VARCHAR(20));
        
        -- Temporary table for program names
        CREATE TABLE #ProgramNames (
            Name NVARCHAR(100),
            StartHoursOffset INT,  -- Added column to store offset
            EndHoursOffset INT     -- Added column to store offset
        ); 
        
        -- Insert program names to include with their respective offsets
        INSERT INTO #ProgramNames (Name, StartHoursOffset, EndHoursOffset) 
        SELECT Name, StartHoursOffset, EndHoursOffset
        FROM Program
        WHERE Name = 'Connect Group Attendance'; 
        
        -- Create temporary table for base organization data 
        CREATE TABLE #OrganizationBase (
            OrganizationId INT,
            OrganizationName NVARCHAR(255),
            Enrollment INT,
            ProgramName NVARCHAR(100),
            ProgramId INT,
            ProgramStartHoursOffset INT,
            ProgramEndHoursOffset INT,
            DivisionName NVARCHAR(100),
            DivisionId INT
        ); 
        
        -- Populate the organization base table with program offsets
        -- Match the breakout logic: include all org-division combinations from OrganizationStructure
        -- but ensure each unique organization is only counted once in the overall total
        INSERT INTO #OrganizationBase
        SELECT
            o.OrganizationId,
            o.OrganizationName,
            -- Count members who were active as of the selected date
            -- This will count the same enrollment multiple times if org is in multiple divisions,
            -- but we'll deduplicate when calculating the overall total
            (
                SELECT COUNT(*)
                FROM OrganizationMembers om
                JOIN People pe ON pe.PeopleId = om.PeopleId
                WHERE om.OrganizationId = o.OrganizationId
                AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= @SelectedDate)
                AND (om.InactiveDate IS NULL OR om.InactiveDate > @SelectedDate)
                AND om.MemberTypeId NOT IN (230, 311) -- Exclude specific member types
                AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > @SelectedDate) -- Not deceased
                AND (pe.DropDate IS NULL OR pe.DropDate > @SelectedDate) -- Not dropped/archived
            ) AS Enrollment,
            p.Name AS ProgramName,
            p.Id AS ProgramId,
            p.StartHoursOffset,
            p.EndHoursOffset,
            d.Name AS DivisionName,
            d.Id AS DivisionId
        FROM Organizations o
        JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
        JOIN Division d ON d.Id = os.DivId
        JOIN Program p ON p.Id = d.ProgId AND p.Name IN (SELECT Name FROM #ProgramNames)
        WHERE (
            o.OrganizationStatusId = 30  -- Currently active
            OR o.LastMeetingDate >= @StartDate  -- Had meetings during/after report period
        )
        AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
        AND d.ReportLine IS NOT NULL AND d.ReportLine <> '';  -- Only divisions in reporting 
        
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
        
        -- Populate the organization meetings table with program-specific date ranges
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
            SUM(CAST(COALESCE(m.MaxCount, 0) AS FLOAT)) AS AvgAttendance -- Changed to SUM since we are only looking at a single day
        FROM #OrganizationBase ob
        LEFT JOIN Meetings m ON m.OrganizationId = ob.OrganizationId
            AND (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL)
            -- Apply program-specific date range using DATETIME for hour offsets
            AND m.MeetingDate BETWEEN
                DATEADD(HOUR, ob.ProgramStartHoursOffset, CAST(DATEADD(DAY, -DATEPART(WEEKDAY, @SelectedDate) + 1, @SelectedDate) AS DATETIME))
                AND
                DATEADD(HOUR, ob.ProgramEndHoursOffset, CAST(DATEADD(DAY, -DATEPART(WEEKDAY, @SelectedDate) + 1, @SelectedDate) AS DATETIME))
        GROUP BY
            ob.OrganizationId,
            ob.OrganizationName,
            ob.Enrollment,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId
        {having_clause_meetings}; 
        
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
        
        -- Populate the last Sunday meetings table with ONLY the Last Sunday
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
        FROM #OrganizationBase ob
        LEFT JOIN Meetings m ON m.OrganizationId = ob.OrganizationId
            AND (m.DidNotMeet = 0 OR m.DidNotMeet IS NULL)
            -- Use the calculated LastSunday date and time range
            AND CONVERT(DATE, m.MeetingDate) = @LastSunday
            -- Apply proper hour range for the specific day
            AND m.MeetingDate BETWEEN
                DATEADD(HOUR, ob.ProgramStartHoursOffset, CAST(@LastSunday AS DATETIME))
                AND
                DATEADD(HOUR, ob.ProgramEndHoursOffset, CAST(@LastSunday AS DATETIME))
        GROUP BY
            ob.OrganizationId,
            ob.OrganizationName,
            ob.Enrollment,
            ob.ProgramName,
            ob.ProgramId,
            ob.DivisionName,
            ob.DivisionId
        {having_clause_last_sunday}; 
        
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
            -- Regular period ratio category - removed division by @WeeksCount since we're only looking at a single day
            CASE 
                WHEN ob.Enrollment = 0 OR COALESCE(om.TotalAttendance, 0) = 0 THEN 'No Data' 
                WHEN (COALESCE(om.TotalAttendance, 0)) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                WHEN (COALESCE(om.TotalAttendance, 0)) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
                ELSE 'Needs Outreach' 
            END AS RatioCategory,
            -- Last Sunday ratio category 
            CASE 
                WHEN ob.Enrollment = 0 OR COALESCE(lsm.LastSundayAttendance, 0) = 0 THEN 'No Data' 
                WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
                ELSE 'Needs Outreach' 
            END AS LastSundayRatioCategory
        FROM #OrganizationBase ob
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
        FROM #OrganizationRatios
        GROUP BY ProgramId, ProgramName; 
        
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
        FROM #OrganizationRatios
        GROUP BY DivisionId, DivisionName, ProgramId, ProgramName; 
        
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
        FROM #OrganizationRatios; 
        
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

        -- Debug: Check for duplicate org-division pairs in #OrganizationBase
        SELECT
            OrganizationId,
            DivisionId,
            Enrollment,
            COUNT(*) as [RowCount]
        INTO #OrgDivDuplicates
        FROM #OrganizationBase
        GROUP BY OrganizationId, DivisionId, Enrollment
        HAVING COUNT(*) > 1;

        -- Populate the division summary table
        -- Deduplicate organizations within each division (org can appear in multiple divs)
        INSERT INTO #DivisionSummary
        SELECT
            DivisionId,
            DivisionName,
            ProgramId,
            ProgramName,
            SUM(Enrollment) AS TotalEnrollment,
            SUM(EnrollmentWithMeetings) AS EnrollmentWithMeetings,
            SUM(TotalAttendance) AS TotalAttendance,
            AVG(AvgAttendance) AS AvgAttendance,
            SUM(LastSundayEnrollment) AS LastSundayEnrollment,
            SUM(LastSundayAttendance) AS LastSundayAttendance
        FROM (
            -- Get unique organizations per division
            SELECT DISTINCT
                ob.OrganizationId,
                ob.DivisionId,
                ob.DivisionName,
                ob.ProgramId,
                ob.ProgramName,
                ob.Enrollment,
                CASE WHEN om.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END AS EnrollmentWithMeetings,
                COALESCE(om.TotalAttendance, 0) AS TotalAttendance,
                COALESCE(om.AvgAttendance, 0) AS AvgAttendance,
                CASE WHEN lsm.OrganizationId IS NOT NULL THEN ob.Enrollment ELSE 0 END AS LastSundayEnrollment,
                COALESCE(lsm.LastSundayAttendance, 0) AS LastSundayAttendance
            FROM #OrganizationBase ob
            LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
            LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId
        ) unique_org_divs
        GROUP BY DivisionId, DivisionName, ProgramId, ProgramName;

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
        -- Sum from division summary to match hierarchical rollup (divisions sum to program)
        -- This ensures the program total matches the sum of all division breakouts shown below
        INSERT INTO #ProgramSummary
        SELECT
            ds.ProgramId,
            ds.ProgramName,
            SUM(ds.TotalEnrollment) AS TotalEnrollment,
            SUM(ds.EnrollmentWithMeetings) AS EnrollmentWithMeetings,
            SUM(ds.TotalAttendance) AS TotalAttendance,
            AVG(ds.AvgAttendance) AS AvgAttendance,
            SUM(ds.LastSundayEnrollment) AS LastSundayEnrollment,
            SUM(ds.LastSundayAttendance) AS LastSundayAttendance
        FROM #DivisionSummary ds
        GROUP BY ds.ProgramId, ds.ProgramName;

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
        -- Use the SAME deduplication logic as division summary to ensure consistency
        INSERT INTO #OverallSummary
        SELECT
            -- Sum enrollment from division summary (already deduplicated by division grouping)
            SUM(ds.TotalEnrollment) AS TotalEnrollment,
            SUM(ds.EnrollmentWithMeetings) AS EnrollmentWithMeetings,
            SUM(ds.TotalAttendance) AS TotalAttendance,
            AVG(ds.AvgAttendance) AS AvgAttendance,
            SUM(ds.LastSundayEnrollment) AS LastSundayEnrollment,
            SUM(ds.LastSundayAttendance) AS LastSundayAttendance
        FROM #DivisionSummary ds;

        -- Create temporary table for results to properly handle ordering 
        CREATE TABLE #Results (
            Level NVARCHAR(100),
            DivisionId INT NULL,
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
                -- Organization Attendance Ratio - removed division by @WeeksCount
                CASE 
                    WHEN COALESCE(om.TotalAttendance, 0) > 0 AND ob.Enrollment > 0 
                    THEN (COALESCE(om.TotalAttendance, 0)) / ob.Enrollment * 100 
                    ELSE 0 
                END AS AttendanceRatio,
                -- Organization Ratio Category - removed division by @WeeksCount
                CASE 
                    WHEN ob.Enrollment = 0 OR COALESCE(om.TotalAttendance, 0) = 0 THEN 'No Data' 
                    WHEN (COALESCE(om.TotalAttendance, 0)) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                    WHEN (COALESCE(om.TotalAttendance, 0)) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
                    ELSE 'Needs Outreach' 
                END AS RatioCategory,
                -- Last Sunday's Attendance Ratio
                CASE 
                    WHEN COALESCE(lsm.LastSundayAttendance, 0) > 0 AND ob.Enrollment > 0 
                    THEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 
                    ELSE 0 
                END AS LastSundayAttendanceRatio,
                -- Last Sunday's Ratio Category
                CASE 
                    WHEN ob.Enrollment = 0 OR COALESCE(lsm.LastSundayAttendance, 0) = 0 THEN 'No Data' 
                    WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                    WHEN COALESCE(lsm.LastSundayAttendance, 0) / ob.Enrollment * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
                    ELSE 'Needs Outreach' 
                END AS LastSundayRatioCategory
            FROM #OrganizationBase ob
            LEFT JOIN #OrganizationMeetings om ON om.OrganizationId = ob.OrganizationId
            LEFT JOIN #LastSundayMeetings lsm ON lsm.OrganizationId = ob.OrganizationId; 
        END 
        
        -- Populate results table - Overall Summary 
        INSERT INTO #Results (
            Level, 
            DivisionId,
            DivisionName, 
            ProgramName, 
            OrganizationName, 
            LevelSortOrder, 
            CategorySortOrder, 
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
        ) 
        SELECT 
            'Overall', 
            NULL,
            NULL, 
            'All Programs', 
            NULL, 
            1, 
            1,
            os.TotalEnrollment,
            os.EnrollmentWithMeetings,
            os.TotalAttendance,
            os.AvgAttendance,
            os.LastSundayEnrollment,
            os.LastSundayAttendance,
            -- Removed division by @WeeksCount
            CASE 
                WHEN os.EnrollmentWithMeetings > 0 
                THEN (os.TotalAttendance) / NULLIF(os.EnrollmentWithMeetings, 0) * 100 
                ELSE 0 
            END,
            -- Removed division by @WeeksCount
            CASE 
                WHEN (os.TotalAttendance) / NULLIF(os.EnrollmentWithMeetings, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                WHEN (os.TotalAttendance) / NULLIF(os.EnrollmentWithMeetings, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
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
        FROM #OverallSummary os
        CROSS JOIN #OverallCategoryCounts occ; 
        
        -- Populate results table - Program Summary 
        INSERT INTO #Results (
            Level, 
            DivisionId,
            DivisionName, 
            ProgramName, 
            OrganizationName, 
            LevelSortOrder, 
            CategorySortOrder, 
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
        ) 
        SELECT 
            'Program', 
            NULL,
            NULL, 
            ps.ProgramName, 
            NULL, 
            2, 
            1,
            ps.TotalEnrollment,
            ps.EnrollmentWithMeetings,
            ps.TotalAttendance,
            ps.AvgAttendance,
            ps.LastSundayEnrollment,
            ps.LastSundayAttendance,
            -- Removed division by @WeeksCount
            CASE 
                WHEN ps.EnrollmentWithMeetings > 0 
                THEN (ps.TotalAttendance) / NULLIF(ps.EnrollmentWithMeetings, 0) * 100 
                ELSE 0 
            END,
            -- Removed division by @WeeksCount
            CASE 
                WHEN (ps.TotalAttendance) / NULLIF(ps.EnrollmentWithMeetings, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                WHEN (ps.TotalAttendance) / NULLIF(ps.EnrollmentWithMeetings, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
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
        FROM #ProgramSummary ps
        JOIN #ProgramCategoryCounts pcc ON ps.ProgramId = pcc.ProgramId; 
        
        -- Populate results table - Division Summary 
        INSERT INTO #Results (
            Level, 
            DivisionId,
            DivisionName, 
            ProgramName, 
            OrganizationName, 
            LevelSortOrder, 
            CategorySortOrder, 
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
        ) 
        SELECT 
            'Division', 
            ds.DivisionId,
            ds.DivisionName, 
            ds.ProgramName, 
            NULL, 
            3, 
            1,
            ds.TotalEnrollment,
            ds.EnrollmentWithMeetings,
            ds.TotalAttendance,
            ds.AvgAttendance,
            ds.LastSundayEnrollment,
            ds.LastSundayAttendance,
            -- Removed division by @WeeksCount
            CASE 
                WHEN ds.EnrollmentWithMeetings > 0 
                THEN (ds.TotalAttendance) / NULLIF(ds.EnrollmentWithMeetings, 0) * 100 
                ELSE 0 
            END,
            -- Removed division by @WeeksCount
            CASE 
                WHEN (ds.TotalAttendance) / NULLIF(ds.EnrollmentWithMeetings, 0) * 100 <= @NeedsInReachThreshold THEN 'Needs In-reach' 
                WHEN (ds.TotalAttendance) / NULLIF(ds.EnrollmentWithMeetings, 0) * 100 <= @GoodRatioThreshold THEN 'Good Ratio' 
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
        FROM #DivisionSummary ds
        JOIN #DivisionCategoryCounts dcc ON ds.DivisionId = dcc.DivisionId AND ds.ProgramId = dcc.ProgramId; 
        
        -- Add Organization Details if @IncludeOrgDetails = 1 
        IF @IncludeOrgDetails = 1 
        BEGIN 
            -- Populate results table - Organization Details 
            INSERT INTO #Results (
                Level, 
                DivisionId,
                DivisionName, 
                ProgramName, 
                OrganizationName, 
                LevelSortOrder, 
                CategorySortOrder, 
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
            ) 
            SELECT 
                'Organization', 
                od.DivisionId,
                od.DivisionName, 
                od.ProgramName, 
                od.OrganizationName, 
                4, 
                1,
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
                NULL,
                NULL,
                NULL,
                NULL,
                NULL,
                NULL
            FROM #OrganizationDetails od; 
        
            -- Category headers and details 
            -- Regular period - Needs In-reach 
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE RatioCategory = 'Needs In-reach') 
            BEGIN 
                -- Header 
                INSERT INTO #Results (
                    Level, 
                    DivisionId,
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                VALUES (
                    'Needs In-reach Organizations', 
                    NULL,
                    NULL, 
                    'Regular Period', 
                    NULL, 
                    5, 
                    1,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                ); 
        
                -- Details 
                INSERT INTO #Results (
                    Level, 
                    DivisionId,
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                SELECT 
                    'Organization', 
                    DivisionId,
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    5, 
                    2,
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
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                FROM #OrganizationDetails
                WHERE RatioCategory = 'Needs In-reach'
                ORDER BY AttendanceRatio; 
            END 
        
            -- Regular period - Good Ratio 
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE RatioCategory = 'Good Ratio') 
            BEGIN 
                -- Header 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                VALUES (
                    'Good Ratio Organizations', 
                    NULL, 
                    'Regular Period', 
                    NULL, 
                    6, 
                    1,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                ); 
        
                -- Details 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                SELECT 
                    'Organization', 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    6, 
                    2,
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
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                FROM #OrganizationDetails
                WHERE RatioCategory = 'Good Ratio'
                ORDER BY AttendanceRatio; 
            END 
        
            -- Regular period - Needs Outreach 
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE RatioCategory = 'Needs Outreach') 
            BEGIN 
                -- Header 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                VALUES (
                    'Needs Outreach Organizations', 
                    NULL, 
                    'Regular Period', 
                    NULL, 
                    7, 
                    1,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                ); 
        
                -- Details 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                SELECT 
                    'Organization', 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    7, 
                    2,
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
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                FROM #OrganizationDetails
                WHERE RatioCategory = 'Needs Outreach'
                ORDER BY AttendanceRatio DESC; 
            END 
        
            -- Last Sunday - Needs In-reach 
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE LastSundayRatioCategory = 'Needs In-reach') 
            BEGIN 
                -- Header 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                VALUES (
                    'Needs In-reach Organizations', 
                    NULL, 
                    'Last Sunday', 
                    NULL, 
                    8, 
                    1,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                ); 
        
                -- Details 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                SELECT 
                    'Organization', 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    8, 
                    2,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    LastSundayEnrollment,
                    LastSundayAttendance,
                    NULL,
                    NULL,
                    LastSundayAttendanceRatio,
                    LastSundayRatioCategory,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                FROM #OrganizationDetails
                WHERE LastSundayRatioCategory = 'Needs In-reach'
                ORDER BY LastSundayAttendanceRatio; 
            END 
        
            -- Last Sunday - Good Ratio 
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE LastSundayRatioCategory = 'Good Ratio') 
            BEGIN 
                -- Header 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                VALUES (
                    'Good Ratio Organizations', 
                    NULL, 
                    'Last Sunday', 
                    NULL, 
                    9, 
                    1,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                ); 
        
                -- Details 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                SELECT 
                    'Organization', 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    9, 
                    2,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    LastSundayEnrollment,
                    LastSundayAttendance,
                    NULL,
                    NULL,
                    LastSundayAttendanceRatio,
                    LastSundayRatioCategory,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                FROM #OrganizationDetails
                WHERE LastSundayRatioCategory = 'Good Ratio'
                ORDER BY LastSundayAttendanceRatio; 
            END 
        
            -- Last Sunday - Needs Outreach 
            IF EXISTS (SELECT 1 FROM #OrganizationDetails WHERE LastSundayRatioCategory = 'Needs Outreach') 
            BEGIN 
                -- Header 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                VALUES (
                    'Needs Outreach Organizations', 
                    NULL, 
                    'Last Sunday', 
                    NULL, 
                    10, 
                    1,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                ); 
        
                -- Details 
                INSERT INTO #Results (
                    Level, 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    LevelSortOrder, 
                    CategorySortOrder, 
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
                ) 
                SELECT 
                    'Organization', 
                    DivisionName, 
                    ProgramName, 
                    OrganizationName, 
                    10, 
                    2,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    LastSundayEnrollment,
                    LastSundayAttendance,
                    NULL,
                    NULL,
                    LastSundayAttendanceRatio,
                    LastSundayRatioCategory,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                FROM #OrganizationDetails
                WHERE LastSundayRatioCategory = 'Needs Outreach'
                ORDER BY LastSundayAttendanceRatio DESC; 
            END
        END

        -- Return the final result, sorted properly
        SELECT
            Level, 
            DivisionId,
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
        FROM #Results 
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
            weeks_count=1,    #ENROLLMENT_RATIO_WEEKS,
            start_date=weeks_ago_date_sql,
            end_date=current_date_sql,
            last_sunday_date=last_sunday_sql,
            selected_date=current_date_sql,
            needs_inreach=ENROLLMENT_RATIO_THRESHOLDS['needs_inreach'],
            good_ratio=ENROLLMENT_RATIO_THRESHOLDS['good_ratio'],
            include_org_details=1 if SHOW_DETAILED_ENROLLMENT else 0,
            org_filter=org_filter,
            division_filter=division_filter,
            program_id_filter=program_id_filter,
            program_insert=program_insert_sql,
            having_clause_meetings=having_clause_meetings,
            having_clause_last_sunday=having_clause_last_sunday
        )
        
        return sql
    
    def get_batch_enrollment_data(self, division_ids, report_date):
        """Get enrollment data for multiple divisions in a single query.

        Counts ALL enrollments - people in multiple organizations are counted multiple times.
        Only includes organizations that are part of the reporting structure.
        """
        # Format date for SQL
        date_str = report_date.strftime('%Y-%m-%d')

        # Create division list for SQL IN clause
        division_list = ", ".join(str(div_id) for div_id in division_ids)

        # Query all division enrollments at once - count each enrollment (not DISTINCT)
        sql = """
        SELECT
            os.DivId AS DivisionId,
            COUNT(*) AS EnrolledCount
        FROM OrganizationMembers om
        JOIN Organizations o ON o.OrganizationId = om.OrganizationId
        JOIN OrganizationStructure os ON os.OrgId = om.OrganizationId
        JOIN Division d ON d.Id = os.DivId
        JOIN Program p ON p.Id = d.ProgId
        JOIN People pe ON pe.PeopleId = om.PeopleId
        WHERE os.DivId IN ({0})
        AND o.OrganizationStatusId = 30  -- Only active organizations
        AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
        AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''  -- Only divisions in reporting
        AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{1}')  -- Enrolled by report date
        AND (om.InactiveDate IS NULL OR om.InactiveDate > '{1}')  -- Still active on report date
        AND om.MemberTypeId NOT IN (230,311)  -- Exclude prospects (230) and other non-members (311)
        AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > '{1}')  -- Not deceased as of report date
        AND (pe.DropDate IS NULL OR pe.DropDate > '{1}')  -- Not dropped/archived as of report date
        GROUP BY os.DivId
        """.format(division_list, date_str)
        
        results = q.QuerySql(sql)
        
        # Convert to dictionary for easy lookup
        enrollment_data = {}
        for row in results:
            enrollment_data[row.DivisionId] = row.EnrolledCount
        
        return enrollment_data
    
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
    

    def generate_enrollment_analysis_section(self, for_email=False):
        """Generate a more compact enrollment ratio analysis section."""
        try:
            # Check if we should show this section at all
            if not ENROLLMENT_ANALYSIS_PROGRAMS:
                return "<p>No programs configured for enrollment analysis.</p>"
            
            # Start with loading indicator (only for browser display, not email)
            if not for_email:
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
            program_title = "Enrollment Metrics"
            if len(organization_levels['Program']) == 1:
                program_title = "{} Enrollment Metrics".format(
                    getattr(organization_levels['Program'][0], 'ProgramName', 'Group')
                )
            
            # Format the HTML using a more compact style
            # For emails, make visible; for browser, hide by default (JavaScript will show it)
            display_style = "display: block;" if for_email else "display: none;"

            analysis_html = """
            <div id="enrollment-analysis" style="{display_style}">
                <div style="background-color: #f9f9f9; border-radius: 5px; padding: 15px; margin-bottom: 20px;">
                    <h3 style="margin-top: 0;">{title}</h3>

                    <!-- Add explanation about enrollment count -->
                    <div style="margin-bottom: 15px; font-size: 13px; font-style: italic; color: #666; padding: 5px 10px; border-left: 3px solid #4A90E2; background-color: #f0f4f8;">
                        <strong>Note:</strong> Enrollment numbers include {enrollment_note}.
                    </div>

            """.format(
                display_style=display_style,
                title=program_title,
                enrollment_note="all involvements regardless of attendance" if INCLUDE_ALL_ENROLLMENTS else "only involvements that have attendance"
            )
            
            # Create the main metrics section - find the row where Level='Program' and DivisionName is NULL
            data_source = None
            program_data_source = None
            attendance_ratio = 0
            
            # Create the main metrics section - find the row where Level='Program' and DivisionName is NULL
            data_source = None
            attendance_ratio = 0
            
            # Look through all data sources to find the Program level row with NULL DivisionName
            for level_type, level_data in organization_levels.items():
                for row in level_data:
                    # Look specifically for the row with Level='Program' and NULL DivisionName
                    if (getattr(row, 'Level', '') == 'Program' and 
                        (getattr(row, 'DivisionName', None) is None or getattr(row, 'DivisionName', '') == 'NULL')):
                        # We found the Program level row - use it for attendance ratio
                        attendance_ratio = getattr(row, 'AttendanceRatio', 0) or 0
                        # But keep looking for a data_source if we don't have one yet
                        if not data_source:
                            data_source = row
            
            # If we couldn't find a Program level row, look for the Overall level for other metrics
            if not data_source and organization_levels['Overall']:
                data_source = organization_levels['Overall'][0]
                    
            # Get the actual weekly attendance from the same source as weekly actuals
            actual_weekly_attendance = 0
            try:
                current_week_connect = self.get_specific_program_attendance_data(
                    [ENROLLMENT_RATIO_PROGRAM],
                    self.week_start_date,
                    self.report_date
                )
                if current_week_connect and 'by_program' in current_week_connect:
                    if ENROLLMENT_RATIO_PROGRAM in current_week_connect['by_program']:
                        actual_weekly_attendance = current_week_connect['by_program'][ENROLLMENT_RATIO_PROGRAM].get('total', 0)
            except:
                actual_weekly_attendance = 0
                
            if data_source:
                # Direct access to values from the SQL result
                # Use TotalEnrollment when INCLUDE_ALL_ENROLLMENTS is True, otherwise use EnrollmentWithMeetings
                if INCLUDE_ALL_ENROLLMENTS:
                    total_enrollment = getattr(data_source, 'TotalEnrollment', 0) or 0
                else:
                    total_enrollment = getattr(data_source, 'EnrollmentWithMeetings', 0) or 0
                # Use the actual weekly attendance from the same calculation as weekly actuals
                last_sunday_attendance = float(actual_weekly_attendance) if actual_weekly_attendance else getattr(data_source, 'LastSundayAttendance', 0) or 0
                #attendance_ratio = getattr(data_source, 'AttendanceRatio', 0) or 0
                
                
                # Ensure values are of the correct type
                total_enrollment = int(total_enrollment)
                # Convert attendance to float to ensure decimal formatting works
                last_sunday_attendance = float(last_sunday_attendance)
                attendance_ratio = float(attendance_ratio)
                
                # Get correct field names from SQL result
                # Updated field names to match SQL query result
                needs_inreach_count = int(getattr(data_source, 'NeedsInReachCount', 0) or 0)
                good_ratio_count = int(getattr(data_source, 'GoodRatioCount', 0) or 0) 
                needs_outreach_count = int(getattr(data_source, 'NeedsOutreachCount', 0) or 0)
                
                analysis_html += """
                    <div style="border-bottom: 1px solid #4A90E2; margin-bottom: 15px;">
    
                        <div style="display: flex; flex-wrap: wrap; gap: 3px; margin-bottom: 15px;">
                            <div style="flex: 1; min-width: 110px; background-color: #f0f4f8; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 0.9em; font-weight: bold; color: #666; margin-bottom: 3px;">Enrollment</div>
                                <div style="font-size: 1.8em; font-weight: bold; color: #333;">{enrollment}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #f0f4f8; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 0.9em; font-weight: bold; color: #666; margin-bottom: 3px;">Attendance</div>
                                <div style="font-size: 1.8em; font-weight: bold; color: #333;">{attendance}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #f0f4f8; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 0.9em; font-weight: bold; color: #666; margin-bottom: 3px;">Ratio</div>
                                <div style="font-size: 1.8em; font-weight: bold; color: #333;">{ratio}%</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #ffe6e6; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 0.9em; font-weight: bold; color: #666; margin-bottom: 3px;">Needs In-reach</div>
                                <div style="font-size: 1.8em; font-weight: bold; color: #d9534f;">{needs_inreach}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #e6ffe6; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 0.9em; font-weight: bold; color: #666; margin-bottom: 3px;">Good Ratio</div>
                                <div style="font-size: 1.8em; font-weight: bold; color: #5cb85c;">{good_ratio}</div>
                            </div>
                            
                            <div style="flex: 1; min-width: 110px; background-color: #fff3e6; border-radius: 3px; text-align: center; padding: 8px;">
                                <div style="font-size: 0.9em; font-weight: bold; color: #666; margin-bottom: 3px;">Needs Outreach</div>
                                <div style="font-size: 1.8em; font-weight: bold; color: #f0ad4e;">{needs_outreach}</div>
                            </div>
                        </div>
                        
                        <!-- Legend/Explanation for the metrics -->
                        <div style="margin-top: 15px; padding: 10px; background-color: #f8f9fa; border-radius: 5px; font-size: 13px;">
                            <div style="font-weight: bold; margin-bottom: 8px; color: #333;">Understanding the Metrics:</div>
                            <div style="display: flex; flex-wrap: wrap; gap: 15px;">
                                <div style="flex: 1; min-width: 200px;">
                                    <span style="color: #d9534f; font-weight: bold;">● Needs In-reach (0-39%):</span>
                                    <span style="color: #666;">Groups with low attendance relative to enrollment. Focus on engaging enrolled members.</span>
                                </div>
                                <div style="flex: 1; min-width: 200px;">
                                    <span style="color: #5cb85c; font-weight: bold;">● Good Ratio (40-59%):</span>
                                    <span style="color: #666;">Healthy balance between attendance and enrollment.</span>
                                </div>
                                <div style="flex: 1; min-width: 200px;">
                                    <span style="color: #f0ad4e; font-weight: bold;">● Needs Outreach (60%+):</span>
                                    <span style="color: #666;">High attendance ratio. Consider inviting more people to join these groups.</span>
                                </div>
                            </div>
                        </div>
                    </div>
                """.format(
                    title="Connect Group Attendance" if program_title == "Enrollment Analysis" else program_title,
                    enrollment="{:,}".format(total_enrollment),
                    attendance="{:.1f}".format(last_sunday_attendance),
                    ratio="{:.1f}".format(attendance_ratio),
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

                    # Ensure values are of the correct type
                    # Use TotalEnrollment when INCLUDE_ALL_ENROLLMENTS is True, otherwise use EnrollmentWithMeetings
                    if INCLUDE_ALL_ENROLLMENTS:
                        div_enrollment = int(getattr(division_item, 'TotalEnrollment', 0) or 0)
                    else:
                        div_enrollment = int(getattr(division_item, 'EnrollmentWithMeetings', 0) or 0)
                    
                    # Get the actual weekly attendance from the same source as the main table
                    div_id = getattr(division_item, 'DivisionId', None)
                    
                    # Try to get actual weekly attendance for this division
                    if div_id:
                        try:
                            div_week_data = self.get_week_attendance_data(
                                self.week_start_date,
                                self.report_date,
                                division_id=div_id
                            )
                            if div_week_data and 'total' in div_week_data:
                                div_attendance = float(div_week_data['total'])
                            else:
                                # Fall back to LastSundayAttendance if weekly data not available
                                div_attendance = float(getattr(division_item, 'LastSundayAttendance', 0) or 0)
                        except Exception as e:
                            # Fall back to LastSundayAttendance on error
                            div_attendance = float(getattr(division_item, 'LastSundayAttendance', 0) or 0)
                    else:
                        # No DivisionId, use LastSundayAttendance
                        div_attendance = float(getattr(division_item, 'LastSundayAttendance', 0) or 0)
                    
                    # Recalculate ratio with the correct attendance
                    if div_enrollment > 0:
                        div_ratio = (div_attendance / float(div_enrollment)) * 100
                    else:
                        div_ratio = 0.0
                    
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
                                <span style="color: #666;">Enrollment:</span> <strong>{enrollment}</strong>
                            </div>
                            <div style="min-width: 170px; flex: 1;">
                                <span style="color: #666;">Attendance:</span> <strong>{attendance}</strong>
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
                                    <span style="color: #666;">Ratio:</span> <strong style="color: #4A90E2;">{ratio}%</strong>
                                </span>
                            </div>
                        </div>
                    </div>
                    """.format(
                        division=div_name,
                        enrollment="{:,}".format(div_enrollment),
                        attendance="{:.1f}".format(div_attendance),
                        ratio="{:.1f}".format(div_ratio),
                        in_reach=div_needs_inreach,
                        good_ratio=div_good_ratio,
                        outreach=div_needs_outreach
                    )
                
            analysis_html += """
                </div>
            </div>
            """

            # Add JavaScript to hide loading spinner (only for browser display)
            if not for_email:
                analysis_html += """
            <script>
                // Hide the loading spinner and show the enrollment section
                var loadingDiv = document.getElementById('loading-enrollment');
                if (loadingDiv) {
                    loadingDiv.style.display = 'none';
                }
                var enrollmentDiv = document.getElementById('enrollment-analysis');
                if (enrollmentDiv) {
                    enrollmentDiv.style.display = 'block';
                }
            </script>
            """
            
            return analysis_html
        except Exception as e:
            # Detailed error handling
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
                        <th style="padding: 4px 8px; text-align: left;">Organization</th>
                        <th style="padding: 4px 8px; text-align: right;">Enrollment</th>
                        <th style="padding: 4px 8px; text-align: right;">Avg Attendance</th>
                        <th style="padding: 4px 8px; text-align: right;">Ratio</th>
                        <th style="padding: 4px 8px; text-align: left;">Status</th>
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
                    <td style="padding: 4px 8px;">{1}</td>
                    <td style="padding: 4px 8px; text-align: right;">{2}</td>
                    <td style="padding: 4px 8px; text-align: right;">{3}</td>
                    <td style="padding: 4px 8px; text-align: right;">{4}%</td>
                    <td style="padding: 4px 8px;">{5}</td>
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
        # Format the selected date for the header
        selected_date_str = ReportHelper.format_display_date(self.report_date)
        
        if SHOW_FOUR_WEEK_COMPARISON:
            # Use separate format strings for with 4-week data
            date_range_html = """
            <div style="margin-top: 20px; margin-bottom: 20px; padding: 10px; background-color: #f0f8ff; border-radius: 5px; border-left: 5px solid #4CAF50;">
                <h3 style="margin-top: 0; color: #2c3e50;">Weekly Report - <span style="color: #27ae60; font-weight: bold;">{10}</span></h3>
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
                ReportHelper.format_display_date(self.previous_year_date),
                selected_date_str
            )
        else:
            # Simpler format string without 4-week data
            date_range_html = """
            <div style="margin-top: 20px; margin-bottom: 20px; padding: 10px; background-color: #f0f8ff; border-radius: 5px; border-left: 5px solid #4CAF50;">
                <h3 style="margin-top: 0; color: #2c3e50;">Weekly Report - <span style="color: #27ae60; font-weight: bold;">{6}</span></h3>
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
                ReportHelper.format_display_date(self.previous_year_date),
                selected_date_str
            )
        
        return formatted_html
    
    def generate_fiscal_year_summary(self, current_ytd_total, previous_ytd_total):
        """Generate a fiscal year summary section with both Worship and Connect Group averages."""
        # Adjust labels based on year type and custom configuration
        ytd_prefix = ReportHelper.get_ytd_label()
        current_label = ReportHelper.get_year_label(self.current_fiscal_year)
        previous_label = ReportHelper.get_year_label(self.previous_fiscal_year)
        
        # Calculate number of Sundays for averaging
        from datetime import timedelta
        current_sundays = ReportHelper.count_sundays_between(self.current_fiscal_start, self.report_date)
        previous_sundays = ReportHelper.count_sundays_between(self.previous_fiscal_start, self.previous_year_date)
        
        # Convert totals to averages
        current_ytd_avg = float(current_ytd_total) / max(current_sundays, 1)
        previous_ytd_avg = float(previous_ytd_total) / max(previous_sundays, 1)
        
        # Calculate change percentage based on averages for accurate trend comparison
        if previous_ytd_avg > 0:
            percent_change = ((current_ytd_avg - previous_ytd_avg) / float(previous_ytd_avg)) * 100
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
            change_color = "green" if current_ytd_avg > 0 else "#777"
            change_arrow = "NEW" if current_ytd_avg > 0 else "&#8652;"
        
        # Get Connect Group Attendance YTD data
        connect_current_ytd_total = 0
        connect_previous_ytd_total = 0
        connect_current_ytd_avg = 0
        connect_previous_ytd_avg = 0
        connect_percent_change = 0
        connect_change_color = "#777"
        connect_change_arrow = "&#8652;"
        
        try:
            # Get current YTD for Connect Group
            current_connect_data = self.get_specific_program_attendance_data(
                [ENROLLMENT_RATIO_PROGRAM],
                self.current_fiscal_start,
                self.report_date
            )
            if current_connect_data and 'by_program' in current_connect_data:
                if ENROLLMENT_RATIO_PROGRAM in current_connect_data['by_program']:
                    connect_current_ytd_total = current_connect_data['by_program'][ENROLLMENT_RATIO_PROGRAM].get('total', 0)
            # Convert to average
            connect_current_ytd_avg = float(connect_current_ytd_total) / max(current_sundays, 1)
            
            # Get previous YTD for Connect Group
            previous_connect_data = self.get_specific_program_attendance_data(
                [ENROLLMENT_RATIO_PROGRAM],
                self.previous_fiscal_start,
                self.previous_year_date
            )
            if previous_connect_data and 'by_program' in previous_connect_data:
                if ENROLLMENT_RATIO_PROGRAM in previous_connect_data['by_program']:
                    connect_previous_ytd_total = previous_connect_data['by_program'][ENROLLMENT_RATIO_PROGRAM].get('total', 0)
            # Convert to average
            connect_previous_ytd_avg = float(connect_previous_ytd_total) / max(previous_sundays, 1)
            
            # Calculate Connect Group change percentage using averages for accurate trend comparison
            if connect_previous_ytd_avg > 0:
                connect_percent_change = ((connect_current_ytd_avg - connect_previous_ytd_avg) / float(connect_previous_ytd_avg)) * 100
                if connect_percent_change > 0:
                    connect_change_color = "green"
                    connect_change_arrow = "&#9650;"
                elif connect_percent_change < 0:
                    connect_change_color = "red"
                    connect_change_arrow = "&#9660;"
                    connect_percent_change = abs(connect_percent_change)
                else:
                    connect_change_color = "#777"
                    connect_change_arrow = "&#8652;"
            else:
                connect_percent_change = 0
                connect_change_color = "green" if connect_current_ytd_avg > 0 else "#777"
                connect_change_arrow = "NEW" if connect_current_ytd_avg > 0 else "&#8652;"
                
        except Exception as e:
            print "<!-- Error getting Connect Group YTD data: {} -->".format(str(e))
        
        summary_html = """
        <div style="margin-top: 10px; margin-bottom: 15px; padding: 10px; background-color: #f3f8ff; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h3 style="margin: 0 0 10px 0; font-size: 1.1em;">{} Year-to-Date Averages - Worship</h3>
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 15px;">
                <div style="flex: 1; min-width: 200px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 0.9em; font-weight: bold; color: #666;">{} Current Average</div>
                    <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                    <div style="font-size: 0.85em; color: #888;">From {} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 200px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 0.9em; font-weight: bold; color: #666;">{} Previous Average</div>
                    <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                    <div style="font-size: 0.85em; color: #888;">From {} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 200px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 0.9em; font-weight: bold; color: #666;">{} Change</div>
                    <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold; color: {};">{} {}%</div>
                    <div style="font-size: 0.85em; color: #888;">{} vs {}</div>
                </div>
            </div>
            
            <h3 style="margin: 15px 0 10px 0; font-size: 1.1em;">{} Year-to-Date Averages - {}</h3>
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 15px;">
                <div style="flex: 1; min-width: 200px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 0.9em; font-weight: bold; color: #666;">{} Current Average</div>
                    <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                    <div style="font-size: 0.85em; color: #888;">From {} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 200px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 0.9em; font-weight: bold; color: #666;">{} Previous Average</div>
                    <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                    <div style="font-size: 0.85em; color: #888;">From {} to {}</div>
                </div>
                
                <div style="flex: 1; min-width: 200px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 0.9em; font-weight: bold; color: #666;">{} Change</div>
                    <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold; color: {};">{} {}%</div>
                    <div style="font-size: 0.85em; color: #888;">{} vs {}</div>
                </div>
            </div>
        </div>
        """.format(
            # Worship section
            ytd_prefix,                                     # YTD or FYTD for header
            ytd_prefix,                                     # YTD or FYTD for Current
            ReportHelper.format_number(int(current_ytd_avg)),  # Show average as integer
            ReportHelper.format_display_date(self.current_fiscal_start),
            ReportHelper.format_display_date(self.report_date),
            
            ytd_prefix,                                     # YTD or FYTD for Previous
            ReportHelper.format_number(int(previous_ytd_avg)),  # Show average as integer
            ReportHelper.format_display_date(self.previous_fiscal_start),
            ReportHelper.format_display_date(self.previous_year_date),
            
            ytd_prefix,                                     # YTD or FYTD for Change
            change_color,
            change_arrow,
            ReportHelper.format_float(percent_change),      # Using format_float instead of string formatting
            ReportHelper.format_number(int(current_ytd_avg)),
            ReportHelper.format_number(int(previous_ytd_avg)),
            
            # Connect Group section
            ytd_prefix,                                     # YTD or FYTD for header
            ENROLLMENT_RATIO_PROGRAM,                       # Program name
            ytd_prefix,                                     # YTD or FYTD for Current
            ReportHelper.format_number(int(connect_current_ytd_avg)),  # Show average as integer
            ReportHelper.format_display_date(self.current_fiscal_start),
            ReportHelper.format_display_date(self.report_date),
            
            ytd_prefix,                                     # YTD or FYTD for Previous
            ReportHelper.format_number(int(connect_previous_ytd_avg)),  # Show average as integer
            ReportHelper.format_display_date(self.previous_fiscal_start),
            ReportHelper.format_display_date(self.previous_year_date),
            
            ytd_prefix,                                     # YTD or FYTD for Change
            connect_change_color,
            connect_change_arrow,
            ReportHelper.format_float(connect_percent_change),
            ReportHelper.format_number(int(connect_current_ytd_avg)),
            ReportHelper.format_number(int(connect_previous_ytd_avg))
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
        
        # Get unique attendance metrics
        involvement_data = self.get_unique_involvements_count(all_programs)
        
        # Get unique attendance for current 4 weeks
        current_unique_data = self.get_unique_attendance_data(
            all_programs,
            self.four_weeks_ago_start_date,
            self.report_date
        )
        
        # Get unique attendance for previous year's 4 weeks
        prev_year_unique_data = self.get_unique_attendance_data(
            all_programs,
            self.prev_year_four_weeks_ago_start_date,
            self.previous_year_date
        )
        
        # Get YTD label based on configuration
        ytd_prefix = ReportHelper.get_ytd_label()
        
        # Build HTML with combined table and section headers
        summary_html = """
        <div style="margin-top: 20px; margin-bottom: 30px; padding: 0; background-color: #f0f0f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; padding: 15px 15px 10px 15px; border-bottom: 2px solid #d0d0e0;">Program Summary</h3>
            
            <div style="padding: 15px;">
                <table style="width: 100%; border-collapse: collapse;">
                    <thead>
                        <tr style="background-color: #e0e0e8;">
                            <th style="padding: 8px 10px; text-align: left; border: 1px solid #ddd;" rowspan="2">Program</th>
                            <th style="padding: 8px 10px; text-align: center; border: 1px solid #ddd; background-color: #d0d0e0;" colspan="{}">Attendance Totals</th>
                            <th style="padding: 8px 10px; text-align: center; border: 1px solid #ddd; background-color: #c0c0d8;" colspan="5">Unique People & Involvements</th>
                        </tr>
                        <tr style="background-color: #e0e0e8;">
        """
        
        # Prepare column definitions
        attendance_columns = [
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
        
        # Count visible columns for colspan
        visible_attendance_cols = sum(1 for col in attendance_columns if col['show'])
        
        # Format the colspan in the header
        summary_html = summary_html.format(visible_attendance_cols)
        
        # Add column headers for attendance metrics
        for col in attendance_columns:
            if col['show']:
                summary_html += '<th style="padding: 6px 10px; text-align: center; border: 1px solid #ddd; background-color: #f0f0f0;">{}</th>'.format(col['name'])
        
        # Add headers for unique metrics (always show these 5 columns)
        summary_html += """
                            <th style="padding: 6px 10px; text-align: center; border: 1px solid #ddd; background-color: #e8e8f0;" title="Number of organizations/involvements">Inv</th>
                            <th style="padding: 6px 10px; text-align: center; border: 1px solid #ddd; background-color: #e8e8f0;" title="Total unique people enrolled">Enrolled</th>
                            <th style="padding: 6px 10px; text-align: center; border: 1px solid #ddd; background-color: #e8e8f0;" title="Enrolled members attended (Guests/Prospects)">Last 4 Weeks</th>
                            <th style="padding: 6px 10px; text-align: center; border: 1px solid #ddd; background-color: #e8e8f0;" title="Enrolled members attended (Guests/Prospects) - Previous Year">PY 4 Weeks</th>
                            <th style="padding: 6px 10px; text-align: center; border: 1px solid #ddd; background-color: #e8e8f0;" title="Percentage change from previous year">Change</th>
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
                    <td style="padding: 4px 8px; text-align: left; border: 1px solid #ddd;"><strong>{}</strong></td>
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
            
            # Add attendance columns based on configuration
            for col in attendance_columns:
                if not col['show']:
                    continue
                
                # Determine the value for each column type
                if col['data'] == 'current_week':
                    value = 0
                    if program_name in current_week_data['by_program']:
                        value = current_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(value)
                    )
                
                elif col['data'] == 'previous_year':
                    value = 0
                    if prev_year_specific_week_data and program_name in prev_year_specific_week_data['by_program']:
                        value = prev_year_specific_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
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
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(trend)
                
                elif col['data'] == 'four_week_current':
                    value = 0
                    if current_four_week_data and program_name in current_four_week_data['by_program']:
                        value = current_four_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(value)
                    )
                
                elif col['data'] == 'four_week_previous':
                    value = 0
                    if prev_year_four_week_data and program_name in prev_year_four_week_data['by_program']:
                        value = prev_year_four_week_data['by_program'][program_name]['total']
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
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
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(trend)
                
                elif col['data'] == 'current_fytd':
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(current_fytd_total)
                    )
                
                elif col['data'] == 'previous_fytd':
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(
                        ReportHelper.format_number(prev_fytd_total)
                    )
                
                elif col['data'] == 'fytd_change':
                    trend = ReportHelper.get_trend_indicator(current_fytd_total, prev_fytd_total)
                    summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd;">{}</td>'.format(trend)
            
            # Add unique metrics columns for the same row
            org_count = involvement_data['org_by_program'].get(program_name, 0)
            unique_enrolled = involvement_data['by_program'].get(program_name, 0)
            current_unique = current_unique_data.get('by_program', {}).get(program_name, 0)
            current_guests = current_unique_data.get('guests_by_program', {}).get(program_name, 0)
            prev_unique = prev_year_unique_data.get('by_program', {}).get(program_name, 0)
            prev_guests = prev_year_unique_data.get('guests_by_program', {}).get(program_name, 0)
            unique_trend = ReportHelper.get_trend_indicator(current_unique, prev_unique)

            # Add unique columns with different background
            summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd; background-color: #f8f8fc;">{}</td>'.format(
                ReportHelper.format_number(org_count)
            )
            summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd; background-color: #f8f8fc;">{}</td>'.format(
                ReportHelper.format_number(unique_enrolled)
            )
            # Display attendance with guest count in parentheses: "376 (38)"
            current_display = ReportHelper.format_number(current_unique)
            if current_guests > 0:
                current_display += ' <span style="color: #666;">({0})</span>'.format(ReportHelper.format_number(current_guests))
            summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd; background-color: #f8f8fc;">{}</td>'.format(
                current_display
            )
            # Display previous year attendance with guest count in parentheses
            prev_display = ReportHelper.format_number(prev_unique)
            if prev_guests > 0:
                prev_display += ' <span style="color: #666;">({0})</span>'.format(ReportHelper.format_number(prev_guests))
            summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd; background-color: #f8f8fc;">{}</td>'.format(
                prev_display
            )
            summary_html += '<td style="padding: 4px 8px; text-align: center; border: 1px solid #ddd; background-color: #f8f8fc;">{}</td>'.format(
                unique_trend
            )
            
            # Close the row
            summary_html += "</tr>"
        
        # Close the table and container
        summary_html += """
                    </tbody>
                </table>
            </div>
        </div>
        """
        
        return summary_html
    
    def generate_overall_summary(self, current_total, previous_total, current_ytd_total, previous_ytd_total, four_week_totals=None, for_email=False):
        """Generate an improved overall summary section with robust error handling and performance tracking."""
        import traceback
        import time  # Add timing functionality
        
        # Add overall function timing
        start_time_total = time.time()
        
        # Create a dictionary to track execution times of different sections
        timing_data = {}
        summary_html = ""
        weekly_kpi_html = ""  # Initialize to prevent errors if exception occurs

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
            worship_current_4week_avg = 0  # Changed to 4-week average
            worship_previous_4week_avg = 0  # Changed to 4-week average
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
                        if not for_email:
                            print "<p class='debug'>SQL generation took: {:.2f} seconds</p>".format(timing_data['get_programs_sql'])
                        if not worship_program_sql:
                            if not for_email:
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
                
                # Get current 4-week worship average (last 4 Sundays)
                start_time_current = time.time()
                try:
                    # For worship, we need total attendance across 4 Sundays
                    # We'll use the four_week_attendance_comparison which gets actual Sunday data
                    print "<p class='debug'>Getting current 4-week worship attendance</p>"
                    
                    # Get the worship program info first
                    worship_prog_sql = """
                        SELECT Id FROM Program WHERE Name = '{}'
                    """.format(worship_program_name.replace("'", "''"))
                    worship_prog = q.QuerySqlTop1(worship_prog_sql)
                    
                    worship_4week_data = {'current': 0, 'previous': 0}
                    if worship_prog and hasattr(worship_prog, 'Id'):
                        # Get 4-week comparison data for worship
                        worship_4week_data = self.get_four_week_attendance_comparison(program_id=worship_prog.Id)
                    
                    current_worship_4week_data = worship_4week_data
                    
                    print "<p class='debug'>Current worship 4-week data: {}</p>".format(str(current_worship_4week_data))
                    
                    timing_data['current_4week_worship'] = time.time() - start_time_current
                    print "<p class='debug'>Current 4-week worship data took: {:.2f} seconds</p>".format(timing_data['current_4week_worship'])
                    
                    # Calculate current and previous 4-week averages
                    if isinstance(current_worship_4week_data, dict):
                        # get_four_week_attendance_comparison returns {'current': {'total': x}, 'previous_year': {'total': y}}
                        if 'current' in current_worship_4week_data and 'total' in current_worship_4week_data['current']:
                            total = current_worship_4week_data['current']['total']
                            if total > 0:
                                # Divide by 4 for the 4-week average (4 Sundays)
                                print "<p class='debug'>Current 4-week worship: total={}, avg={:.1f}</p>".format(
                                    total, float(total) / 4.0)
                                worship_current_4week_avg = float(total) / 4.0
                        
                        # Also get previous year from same data structure
                        if 'previous_year' in current_worship_4week_data and 'total' in current_worship_4week_data['previous_year']:
                            prev_total = current_worship_4week_data['previous_year']['total']
                            if prev_total > 0:
                                print "<p class='debug'>Previous 4-week worship: total={}, avg={:.1f}</p>".format(
                                    prev_total, float(prev_total) / 4.0)
                                worship_previous_4week_avg = float(prev_total) / 4.0
                except Exception as current_error:
                    print "<p class='debug'>Error calculating current YTD average: " + str(current_error) + "</p>"
                
                # Previous year 4-week average already calculated above from same data structure
                    
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
                        str(self.four_weeks_ago_start_date), str(self.report_date))
                    
                    four_week_program_data = self.get_specific_program_attendance_data(
                        [ENROLLMENT_RATIO_PROGRAM],
                        self.four_weeks_ago_start_date,
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
                            # Calculate weekly average for the 4-week period
                            four_week_program_avg = total / 4.0 if total > 0 else 0
                except Exception as current_4wk_error:
                    print "<p class='debug'>Error calculating current 4-week average: " + str(current_4wk_error) + "</p>"
                
                # Get previous year 4-week data safely
                start_time_prev_4wk = time.time()
                try:
                    print "<p class='debug'>Getting previous 4-week data from {} to {}</p>".format(
                        str(self.prev_year_four_weeks_ago_start_date), str(self.previous_year_date))
                    
                    four_week_previous_program_data = self.get_specific_program_attendance_data(
                        [ENROLLMENT_RATIO_PROGRAM],
                        self.prev_year_four_weeks_ago_start_date,
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
                            # Calculate weekly average for the 4-week period
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
                    formatted_worship_current_4week_avg = ReportHelper.format_number(int(worship_current_4week_avg))
                except:
                    formatted_worship_current_4week_avg = "{0:,}".format(int(worship_current_4week_avg)) if worship_current_4week_avg else "0"
                    
                try:
                    formatted_worship_previous_4week_avg = ReportHelper.format_number(int(worship_previous_4week_avg))
                except:
                    formatted_worship_previous_4week_avg = "{0:,}".format(int(worship_previous_4week_avg)) if worship_previous_4week_avg else "0"
                    
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
                    trend_indicator = ReportHelper.get_trend_indicator(worship_current_4week_avg, worship_previous_4week_avg)
                except:
                    # Direct calculation if the helper fails
                    if worship_previous_4week_avg > 0:
                        trend_pct = ((worship_current_4week_avg - worship_previous_4week_avg) / worship_previous_4week_avg) * 100
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
                formatted_worship_current_4week_avg = str(int(worship_current_4week_avg))
                formatted_worship_previous_4week_avg = str(int(worship_previous_4week_avg))
                formatted_four_week_program_avg = str(int(four_week_program_avg))
                formatted_four_week_previous_avg = str(int(four_week_previous_program_avg))
                trend_indicator = ""
            finally:
                timing_data['formatting_section'] = time.time() - start_time
                print "<p class='debug'>Formatting section took: {:.2f} seconds</p>".format(timing_data['formatting_section'])
            
            # Get weekly actuals for the new KPIs
            weekly_worship_current = 0
            weekly_worship_previous = 0
            weekly_connect_current = 0
            weekly_connect_previous = 0
            weekly_enrollment_current = 0
            weekly_enrollment_previous = 0
            
            try:
                # Get current week worship attendance
                current_week_worship = self.get_specific_program_attendance_data(
                    [worship_program_name],
                    self.week_start_date,
                    self.report_date
                )
                if current_week_worship and 'by_program' in current_week_worship:
                    if worship_program_name in current_week_worship['by_program']:
                        weekly_worship_current = current_week_worship['by_program'][worship_program_name].get('total', 0)
                
                # Get previous year same week worship attendance
                previous_week_worship = self.get_specific_program_attendance_data(
                    [worship_program_name],
                    self.previous_week_start_date,
                    self.previous_year_date
                )
                if previous_week_worship and 'by_program' in previous_week_worship:
                    if worship_program_name in previous_week_worship['by_program']:
                        weekly_worship_previous = previous_week_worship['by_program'][worship_program_name].get('total', 0)
                
                # Get current week connect group attendance
                current_week_connect = self.get_specific_program_attendance_data(
                    [ENROLLMENT_RATIO_PROGRAM],
                    self.week_start_date,
                    self.report_date
                )
                if current_week_connect and 'by_program' in current_week_connect:
                    if ENROLLMENT_RATIO_PROGRAM in current_week_connect['by_program']:
                        weekly_connect_current = current_week_connect['by_program'][ENROLLMENT_RATIO_PROGRAM].get('total', 0)
                
                # Get previous year same week connect group attendance
                previous_week_connect = self.get_specific_program_attendance_data(
                    [ENROLLMENT_RATIO_PROGRAM],
                    self.previous_week_start_date,
                    self.previous_year_date
                )
                if previous_week_connect and 'by_program' in previous_week_connect:
                    if ENROLLMENT_RATIO_PROGRAM in previous_week_connect['by_program']:
                        weekly_connect_previous = previous_week_connect['by_program'][ENROLLMENT_RATIO_PROGRAM].get('total', 0)
                
                # Get enrollment data for current week - count ALL enrollments (not unique people)
                enrollment_sql_current = """
                    SELECT COUNT(*) as TotalEnrollment
                    FROM OrganizationMembers om
                    JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                    JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
                    JOIN Division d ON d.Id = os.DivId
                    JOIN Program p ON p.Id = d.ProgId AND p.Name = '{0}'
                    JOIN People pe ON pe.PeopleId = om.PeopleId
                    WHERE o.OrganizationStatusId = 30
                    AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
                    AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''  -- Only divisions in reporting
                    AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{1}')  -- Enrolled by report date
                    AND (om.InactiveDate IS NULL OR om.InactiveDate > '{1}')  -- Still active on report date
                    AND om.MemberTypeId NOT IN (230, 311)  -- Exclude prospects and non-members
                    AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > '{1}')  -- Not deceased as of report date
                    AND (pe.DropDate IS NULL OR pe.DropDate > '{1}')  -- Not dropped/archived as of report date
                """.format(ENROLLMENT_RATIO_PROGRAM,
                          ReportHelper.format_date(self.report_date))
                
                enrollment_result = q.QuerySqlTop1(enrollment_sql_current)
                if enrollment_result:
                    weekly_enrollment_current = enrollment_result.TotalEnrollment or 0
                
                # Get enrollment data for previous year same week - count ALL enrollments (not unique people)
                enrollment_sql_previous = """
                    SELECT COUNT(*) as TotalEnrollment
                    FROM OrganizationMembers om
                    JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                    JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
                    JOIN Division d ON d.Id = os.DivId
                    JOIN Program p ON p.Id = d.ProgId AND p.Name = '{0}'
                    JOIN People pe ON pe.PeopleId = om.PeopleId
                    WHERE o.OrganizationStatusId = 30
                    AND p.RptGroup IS NOT NULL AND p.RptGroup <> ''  -- Only programs in reporting
                    AND d.ReportLine IS NOT NULL AND d.ReportLine <> ''  -- Only divisions in reporting
                    AND (om.EnrollmentDate IS NULL OR om.EnrollmentDate <= '{1}')  -- Enrolled by report date
                    AND (om.InactiveDate IS NULL OR om.InactiveDate > '{1}')  -- Still active on report date
                    AND om.MemberTypeId NOT IN (230, 311)  -- Exclude prospects and non-members
                    AND (pe.DeceasedDate IS NULL OR pe.DeceasedDate > '{1}')  -- Not deceased as of report date
                    AND (pe.DropDate IS NULL OR pe.DropDate > '{1}')  -- Not dropped/archived as of report date
                """.format(ENROLLMENT_RATIO_PROGRAM,
                          ReportHelper.format_date(self.previous_year_date))
                
                enrollment_result_prev = q.QuerySqlTop1(enrollment_sql_previous)
                if enrollment_result_prev:
                    weekly_enrollment_previous = enrollment_result_prev.TotalEnrollment or 0
                    
            except Exception as weekly_error:
                print "<p class='debug'>Error getting weekly actuals: " + str(weekly_error) + "</p>"
            
            # Format weekly actuals
            formatted_worship_current = ReportHelper.format_number(weekly_worship_current)
            formatted_worship_previous = ReportHelper.format_number(weekly_worship_previous)
            formatted_connect_current = ReportHelper.format_number(weekly_connect_current)
            formatted_connect_previous = ReportHelper.format_number(weekly_connect_previous)
            formatted_enrollment_current = ReportHelper.format_number(weekly_enrollment_current)
            formatted_enrollment_previous = ReportHelper.format_number(weekly_enrollment_previous)
            
            # Calculate trends for weekly actuals
            worship_weekly_trend = ReportHelper.get_trend_indicator(weekly_worship_current, weekly_worship_previous)
            connect_weekly_trend = ReportHelper.get_trend_indicator(weekly_connect_current, weekly_connect_previous)
            enrollment_weekly_trend = ReportHelper.get_trend_indicator(weekly_enrollment_current, weekly_enrollment_previous)
            
            # Generate weekly actuals KPI section
            weekly_kpi_html = """
            <div style="margin-top: 10px; margin-bottom: 10px; padding: 10px; background-color: #f8f8f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <h3 style="margin: 0 0 10px 0; font-size: 1.1em;">Weekly Actuals</h3>
                <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 15px;">
                    <div style="flex: 1; min-width: 180px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <div style="font-size: 0.9em; font-weight: bold; color: #666;">Worship Total</div>
                        <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                        <div style="font-size: 0.85em; color: #888;">Last Year: {}</div>
                        <div style="margin-top: 3px; font-size: 0.9em;">{}</div>
                    </div>
                    
                    <div style="flex: 1; min-width: 180px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <div style="font-size: 0.9em; font-weight: bold; color: #666;">Connect Group Attendance</div>
                        <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                        <div style="font-size: 0.85em; color: #888;">Last Year: {}</div>
                        <div style="margin-top: 3px; font-size: 0.9em;">{}</div>
                    </div>
                    
                    <div style="flex: 1; min-width: 180px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        <div style="font-size: 0.9em; font-weight: bold; color: #666;">Connect Group Enrollment</div>
                        <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{}</div>
                        <div style="font-size: 0.85em; color: #888;">Last Year: {}</div>
                        <div style="margin-top: 3px; font-size: 0.9em;">{}</div>
                    </div>
                </div>
            </div>
            """.format(
                formatted_worship_current,
                formatted_worship_previous,
                worship_weekly_trend,
                formatted_connect_current,
                formatted_connect_previous,
                connect_weekly_trend,
                formatted_enrollment_current,
                formatted_enrollment_previous,
                enrollment_weekly_trend
            )
            
            # Generate the summary HTML
            start_time = time.time()
            try:
                # Using traditional string formatting to avoid f-string issues
                enrollment_ratio_rounded = str(round(enrollment_ratio, 1))  # Pre-format as string to avoid format specifier issue
                
                summary_html = """
                <div style="margin-top: 10px; margin-bottom: 15px; padding: 10px; background-color: #f8f8f8; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                    <h3 style="margin: 0 0 10px 0; font-size: 1.1em;">Rolling Averages</h3>
                    <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 15px;">
                        <div style="flex: 1; min-width: 180px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 0.9em; font-weight: bold; color: #666;">{0} 4-Week Average</div>
                            <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{2}</div>
                            <div style="font-size: 0.85em; color: #888;">4 Weeks Last Year: {3}</div>
                            <div style="margin-top: 3px; font-size: 0.9em;">{4}</div>
                        </div>
                        
                        <div style="flex: 1; min-width: 180px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 0.9em; font-weight: bold; color: #666;">4-Week Avg - {5}</div>
                            <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{6}</div>
                            <div style="font-size: 0.85em; color: #888;">4 Weeks Last Year: {7}</div>
                        </div>
                        
                        <div style="flex: 1; min-width: 180px; padding: 10px; background-color: #fff; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 0.9em; font-weight: bold; color: #666;">{5} {11}-Week Enrollment Ratio</div>
                            <div style="font-size: 1.8em; margin: 5px 0; font-weight: bold;">{8}%</div>
                            <div style="font-size: 0.85em; color: {9};">{10}</div>
                        </div>
                    </div>
                </div>
                """.format(
                    worship_program_name,              # 0
                    ytd_prefix,                        # 1 (not used in first box anymore)
                    formatted_worship_current_4week_avg, # 2
                    formatted_worship_previous_4week_avg,# 3
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
                """.format(worship_program_name, formatted_worship_current_4week_avg)
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
            
            # Only print performance data if not generating for email
            if not for_email:
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

        # Return both the weekly KPIs and the summary HTML
        return weekly_kpi_html + summary_html
        
    def send_email_report(self, email_to, report_content):
        """Send the report via email."""
        try:
            # Resolve email recipient based on EMAIL_TO configuration
            if not email_to or email_to == "self":
                email_id = model.UserPeopleId
                email_display = "yourself"
            else:
                # Try to parse as PeopleId first
                try:
                    email_id = int(email_to)
                    # Get person name for display
                    try:
                        person = model.GetPerson(email_id)
                        email_display = person.Name if person else "PeopleId: {}".format(email_id)
                    except:
                        email_display = "PeopleId: {}".format(email_id)
                except ValueError:
                    # Try to find by email address
                    try:
                        email_id = model.FindPersonId(email_to)
                        if not email_id:
                            # Fallback to current user
                            email_id = model.UserPeopleId
                            email_display = "yourself (email lookup failed)"
                        else:
                            email_display = email_to
                    except:
                        # Final fallback
                        email_id = model.UserPeopleId
                        email_display = "yourself (lookup failed)"

            # Add indication that email is being sent
            email_status = """
            <div id="email-status" style="padding: 15px; background-color: #fff8e1; color: #856404; border-radius: 5px; margin-bottom: 20px;">
                <strong>Sending email to {}...</strong>
                <span id="email-spinner" class="spinner-small" style="display: inline-block; margin-left: 10px;"></span>
            </div>
            <script>
            document.addEventListener('DOMContentLoaded', function() {{
                var emailStatus = document.getElementById('email-status');
                if (emailStatus) {{
                    emailStatus.scrollIntoView({{behavior: 'smooth'}});
                }}
            }});
            </script>
            """.format(email_display)

            print email_status

            # Set up email parameters
            QueuedBy = model.UserPeopleId

            # Set email tracking
            email_content = report_content + "{track}{tracklinks}"

            # Add report date to email subject
            report_date_str = self.report_date.strftime('%m/%d/%Y')
            email_subject = "{} - {}".format(EMAIL_SUBJECT, report_date_str)

            # Send the email
            model.Email(email_id, QueuedBy, EMAIL_FROM_ADDRESS, EMAIL_FROM_NAME, email_subject, email_content)
            
            # Update status to success
            success_message = """
            <script>
            // Hide the spinner first
            var spinner = document.getElementById('email-spinner');
            if (spinner) {{
                spinner.style.display = 'none';
            }}

            // Update the status box
            var emailStatus = document.getElementById('email-status');
            if (emailStatus) {{
                emailStatus.className = '';
                emailStatus.style.backgroundColor = '#d4edda';
                emailStatus.style.color = '#155724';
                emailStatus.innerHTML = '<strong>Success!</strong> Email sent successfully to {0}';
            }}
            </script>
            """.format(email_display)

            print success_message

            return True, "Email sent successfully to {}".format(email_display)
        except Exception as e:
            # Update status to error
            error_message = """
            <script>
            // Hide the spinner first
            var spinner = document.getElementById('email-spinner');
            if (spinner) {{
                spinner.style.display = 'none';
            }}

            // Update the status box
            var emailStatus = document.getElementById('email-status');
            if (emailStatus) {{
                emailStatus.className = '';
                emailStatus.style.backgroundColor = '#f8d7da';
                emailStatus.style.color = '#721c24';
                emailStatus.innerHTML = '<strong>Error:</strong> {0}';
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
                # For program/division breakdown tables, sum the division enrollments
                # (people in multiple divisions count in each)
                
                # Get all divisions for this program and sum their enrollments
                divisions_sql = """
                    SELECT Id 
                    FROM Division d
                    WHERE d.ProgId = {}
                    AND d.ReportLine IS NOT NULL 
                    AND d.ReportLine <> ''
                """.format(program.Id)
                
                divisions = q.QuerySql(divisions_sql)
                enrollment_count = 0
                
                for division in divisions:
                    division_enrollment = self.get_division_enrollment(division.Id, self.report_date)
                    enrollment_count += division_enrollment
                                
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

    def generate_report_content_optimized(self):
        """Generate report content with optimized database queries."""
        # Get all structural data in a single query
        all_data = self.get_all_programs_divisions_organizations()
        
        # Create organized data structures to hold the hierarchy
        programs = {}
        divisions = {}
        organizations = {}
        
        # Process all data into a hierarchical structure
        for row in all_data:
            # Handle program data
            if row.ProgramId not in programs:
                programs[row.ProgramId] = {
                    'id': row.ProgramId,
                    'name': row.ProgramName,
                    'rpt_group': row.ProgramRptGroup,
                    'start_hours_offset': row.StartHoursOffset,
                    'end_hours_offset': row.EndHoursOffset,
                    'divisions': {}
                }
            
            # Skip if no division data
            if not row.DivisionId:
                continue
                
            # Handle division data
            if row.DivisionId not in divisions:
                divisions[row.DivisionId] = {
                    'id': row.DivisionId, 
                    'name': row.DivisionName,
                    'report_line': row.DivisionReportLine,
                    'program_id': row.ProgramId,
                    'organizations': {}
                }
                programs[row.ProgramId]['divisions'][row.DivisionId] = divisions[row.DivisionId]
            
            # Skip if no organization data
            if not row.OrganizationId:
                continue
                
            # Handle organization data
            if row.OrganizationId not in organizations:
                organizations[row.OrganizationId] = {
                    'id': row.OrganizationId,
                    'name': row.OrganizationName,
                    'division_id': row.DivisionId,
                    'member_count': row.MemberCount
                }
                divisions[row.DivisionId]['organizations'][row.OrganizationId] = organizations[row.OrganizationId]
        
        # Build date ranges for all the attendance data we need
        date_ranges = {
            'current_week': (self.week_start_date, self.report_date),
            'previous_week': (self.previous_week_start, self.previous_week_date),
            'previous_year_week': (self.previous_week_start_date, self.previous_year_date),
            'current_ytd': (self.current_fiscal_start, self.report_date),
            'previous_ytd': (self.previous_fiscal_start, self.previous_year_date)
        }
        
        # Add four-week ranges if enabled
        if SHOW_FOUR_WEEK_COMPARISON:
            date_ranges['current_four_week'] = (self.four_weeks_ago_start_date, self.report_date)
            date_ranges['previous_four_week'] = (self.prev_year_four_weeks_ago_start_date, self.previous_year_date)
        
        # Get all attendance data in a single query
        all_attendance = self.get_batch_attendance_data(date_ranges)
        
        # Process attendance data into our hierarchical structure
        attendance_data = {}
        for range_id in date_ranges.keys():
            attendance_data[range_id] = {
                'programs': {},
                'divisions': {},
                'organizations': {}
            }
        
        for row in all_attendance:
            range_id = row.RangeId
            program_id = row.ProgramId
            division_id = row.DivisionId
            org_id = row.OrganizationId
            
            # Add to program attendance
            if program_id not in attendance_data[range_id]['programs']:
                attendance_data[range_id]['programs'][program_id] = {
                    'total': 0,
                    'meetings': 0,
                    'by_hour': {}
                }
            
            attendance_data[range_id]['programs'][program_id]['total'] += row.TotalAttendance
            attendance_data[range_id]['programs'][program_id]['meetings'] += row.MeetingCount
            if row.MeetingHour not in attendance_data[range_id]['programs'][program_id]['by_hour']:
                attendance_data[range_id]['programs'][program_id]['by_hour'][row.MeetingHour] = 0
            attendance_data[range_id]['programs'][program_id]['by_hour'][row.MeetingHour] += row.TotalAttendance
            
            # Add to division attendance
            if division_id not in attendance_data[range_id]['divisions']:
                attendance_data[range_id]['divisions'][division_id] = {
                    'total': 0,
                    'meetings': 0,
                    'by_hour': {}
                }
            
            attendance_data[range_id]['divisions'][division_id]['total'] += row.TotalAttendance
            attendance_data[range_id]['divisions'][division_id]['meetings'] += row.MeetingCount
            if row.MeetingHour not in attendance_data[range_id]['divisions'][division_id]['by_hour']:
                attendance_data[range_id]['divisions'][division_id]['by_hour'][row.MeetingHour] = 0
            attendance_data[range_id]['divisions'][division_id]['by_hour'][row.MeetingHour] += row.TotalAttendance
            
            # Add to organization attendance
            if org_id not in attendance_data[range_id]['organizations']:
                attendance_data[range_id]['organizations'][org_id] = {
                    'total': 0,
                    'meetings': 0,
                    'by_hour': {}
                }
            
            attendance_data[range_id]['organizations'][org_id]['total'] += row.TotalAttendance
            attendance_data[range_id]['organizations'][org_id]['meetings'] += row.MeetingCount
            if row.MeetingHour not in attendance_data[range_id]['organizations'][org_id]['by_hour']:
                attendance_data[range_id]['organizations'][org_id]['by_hour'][row.MeetingHour] = 0
            attendance_data[range_id]['organizations'][org_id]['by_hour'][row.MeetingHour] += row.TotalAttendance
        
        # Structure data by years for easier report generation
        years_attendance = {}
        years_ytd_attendance = {}
        
        # Initialize attendance data structures by year
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            years_attendance[year] = {'programs': {}, 'divisions': {}, 'organizations': {}}
            years_ytd_attendance[year] = {'programs': {}, 'divisions': {}, 'organizations': {}}
        
        # Map the data from date ranges to year structures
        years_attendance[self.current_year] = attendance_data.get('current_week', {'programs': {}, 'divisions': {}, 'organizations': {}})
        years_attendance[self.current_year - 1] = attendance_data.get('previous_year_week', {'programs': {}, 'divisions': {}, 'organizations': {}})
        years_ytd_attendance[self.current_year] = attendance_data.get('current_ytd', {'programs': {}, 'divisions': {}, 'organizations': {}})
        years_ytd_attendance[self.current_year - 1] = attendance_data.get('previous_ytd', {'programs': {}, 'divisions': {}, 'organizations': {}})
        
        # Set up four-week data if enabled
        four_week_attendance = {}
        if SHOW_FOUR_WEEK_COMPARISON:
            four_week_attendance = {
                'current': attendance_data.get('current_four_week', {'programs': {}, 'divisions': {}, 'organizations': {}}),
                'previous_year': attendance_data.get('previous_four_week', {'programs': {}, 'divisions': {}, 'organizations': {}})
            }
        
        # Get all division enrollments in batch if needed
        division_ids = list(divisions.keys())
        division_enrollment = {}
        if SHOW_ENROLLMENT_COLUMN and division_ids:
            division_enrollment = self.get_batch_enrollment_data(division_ids, self.report_date)
        
        # Generate HTML from the collected data
        report_content = ""
        
        # Process programs in order
        for program_id, program_info in sorted(programs.items(), 
                                            key=lambda x: ReportHelper.parse_program_rpt_group(x[1]['rpt_group'])[0]):
            # Create a program object for compatibility with existing methods
            program_obj = type('Program', (), {
                'Id': program_info['id'],
                'Name': program_info['name'],
                'RptGroup': program_info['rpt_group'],
                'StartHoursOffset': program_info['start_hours_offset'],
                'EndHoursOffset': program_info['end_hours_offset']
            })
            
            # Start a new table for each program
            report_content += "<h3>{}</h3>".format(program_info['name'])
            report_content += "<table>"
            
            # Add the header row
            service_times = self.parse_service_times(program_info['rpt_group'])
            report_content += self.generate_header_row(service_times)
            
            # Prepare program attendance data for all years
            years_program_data = {}
            for year in years_attendance:
                if program_id in years_attendance[year]['programs']:
                    years_program_data[year] = years_attendance[year]['programs'][program_id]
                else:
                    years_program_data[year] = {'total': 0, 'meetings': 0, 'by_hour': {}}
            
            # Prepare program YTD data for all years
            years_program_ytd = {}
            for year in years_ytd_attendance:
                if program_id in years_ytd_attendance[year]['programs']:
                    years_program_ytd[year] = years_ytd_attendance[year]['programs'][program_id]
                else:
                    years_program_ytd[year] = {'total': 0, 'meetings': 0, 'by_hour': {}}
            
            # Get four-week data for this program
            program_four_week_data = None
            if SHOW_FOUR_WEEK_COMPARISON:
                current_data = {'total': 0, 'meetings': 0}
                prev_year_data = {'total': 0, 'meetings': 0}
                
                if program_id in four_week_attendance['current']['programs']:
                    current_data = four_week_attendance['current']['programs'][program_id]
                
                if program_id in four_week_attendance['previous_year']['programs']:
                    prev_year_data = four_week_attendance['previous_year']['programs'][program_id]
                
                program_four_week_data = {
                    'current': current_data,
                    'previous_year': prev_year_data
                }
            
            # Generate program row
            report_content += self.generate_program_row(
                program_obj,
                years_program_data,
                years_program_ytd,
                program_four_week_data
            )
            
            # Process divisions for this program
            for division_id, division_info in sorted(
                program_info['divisions'].items(),
                key=lambda x: x[1]['report_line'] if x[1]['report_line'] else '999'
            ):
                # Create a division object for compatibility with existing methods
                division_obj = type('Division', (), {
                    'Id': division_id,
                    'Name': division_info['name'],
                    'ProgId': program_id
                })
                
                # Prepare division attendance data for all years
                years_division_data = {}
                for year in years_attendance:
                    if division_id in years_attendance[year]['divisions']:
                        years_division_data[year] = years_attendance[year]['divisions'][division_id]
                    else:
                        years_division_data[year] = {'total': 0, 'meetings': 0, 'by_hour': {}}
                
                # Prepare division YTD data for all years
                years_division_ytd = {}
                for year in years_ytd_attendance:
                    if division_id in years_ytd_attendance[year]['divisions']:
                        years_division_ytd[year] = years_ytd_attendance[year]['divisions'][division_id]
                    else:
                        years_division_ytd[year] = {'total': 0, 'meetings': 0, 'by_hour': {}}
                
                # Get four-week data for this division
                division_four_week_data = None
                if SHOW_FOUR_WEEK_COMPARISON:
                    current_data = {'total': 0, 'meetings': 0}
                    prev_year_data = {'total': 0, 'meetings': 0}
                    
                    if division_id in four_week_attendance['current']['divisions']:
                        current_data = four_week_attendance['current']['divisions'][division_id]
                    
                    if division_id in four_week_attendance['previous_year']['divisions']:
                        prev_year_data = four_week_attendance['previous_year']['divisions'][division_id]
                    
                    division_four_week_data = {
                        'current': current_data,
                        'previous_year': prev_year_data
                    }
                
                # Generate division row
                division_row = self.generate_division_row(
                    division_obj,
                    years_division_data,
                    years_division_ytd,
                    division_four_week_data
                )
                report_content += division_row
                
                # Process organizations if needed
                if SHOW_ORGANIZATION_DETAILS and not self.collapse_orgs and division_row != "":
                    for org_id, org_info in sorted(
                        division_info['organizations'].items(),
                        key=lambda x: x[1]['name']
                    ):
                        # Create an organization object for compatibility
                        org_obj = type('Organization', (), {
                            'OrganizationId': org_id,
                            'OrganizationName': org_info['name'],
                            'DivisionId': division_id
                        })
                        
                        # Prepare organization attendance data for all years
                        years_org_data = {}
                        for year in years_attendance:
                            if org_id in years_attendance[year]['organizations']:
                                years_org_data[year] = years_attendance[year]['organizations'][org_id]
                            else:
                                years_org_data[year] = {'total': 0, 'meetings': 0, 'by_hour': {}}
                        
                        # Prepare organization YTD data for all years
                        years_org_ytd = {}
                        for year in years_ytd_attendance:
                            if org_id in years_ytd_attendance[year]['organizations']:
                                years_org_ytd[year] = years_ytd_attendance[year]['organizations'][org_id]
                            else:
                                years_org_ytd[year] = {'total': 0, 'meetings': 0, 'by_hour': {}}
                        
                        # Get four-week data for this organization
                        org_four_week_data = None
                        if SHOW_FOUR_WEEK_COMPARISON:
                            current_data = {'total': 0, 'meetings': 0}
                            prev_year_data = {'total': 0, 'meetings': 0}
                            
                            if org_id in four_week_attendance['current']['organizations']:
                                current_data = four_week_attendance['current']['organizations'][org_id]
                            
                            if org_id in four_week_attendance['previous_year']['organizations']:
                                prev_year_data = four_week_attendance['previous_year']['organizations'][org_id]
                            
                            org_four_week_data = {
                                'current': current_data,
                                'previous_year': prev_year_data
                            }
                        
                        # Generate organization row
                        report_content += self.generate_organization_row(
                            org_obj,
                            years_org_data,
                            years_org_ytd,
                            org_four_week_data
                        )
            
            # Close the table for this program
            report_content += "</table>"
        
        # Return the complete HTML report
        return report_content

    def generate_report_content(self, for_email=False):
        """Generate just the report content (used for both display and email)."""
        
        performance_timer.start("generate_report_content")

        global SHOW_ENROLLMENT_COLUMN

        report_content = ""

        # Add date header for emails
        if for_email:
            # Format dates for display
            current_week_start = self.week_start_date.strftime('%m/%d/%Y')
            current_week_end = self.report_date.strftime('%m/%d/%Y')
            prev_week_start = self.previous_week_start.strftime('%m/%d/%Y')
            prev_week_end = self.previous_week_date.strftime('%m/%d/%Y')
            prev_year_start = ReportHelper.get_week_start_date(self.previous_year_date).strftime('%m/%d/%Y')
            prev_year_end = self.previous_year_date.strftime('%m/%d/%Y')
            four_weeks_start = self.four_weeks_ago_start_date.strftime('%m/%d/%Y')
            prev_four_weeks_start = (self.four_weeks_ago_start_date - datetime.timedelta(days=364)).strftime('%m/%d/%Y')
            prev_four_weeks_end = (self.report_date - datetime.timedelta(days=364)).strftime('%m/%d/%Y')

            report_content += """
            <div style="background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 5px; padding: 20px; margin-bottom: 20px; font-family: Arial, sans-serif;">
                <h2 style="margin-top: 0; color: #2c3e50;">Weekly Report - <span style="color: #27ae60;">{}</span></h2>
                <div style="line-height: 1.8; color: #34495e;">
                    <div><strong>Current Week:</strong> {} to {}</div>
                    <div><strong>Previous Week:</strong> {} to {}</div>
                    <div><strong>Previous Year:</strong> {} to {}</div>
                    <div><strong>Current 4 Weeks:</strong> {} to {}</div>
                    <div><strong>Previous Year 4 Weeks:</strong> {} to {}</div>
                </div>
            </div>
            """.format(
                self.report_date.strftime('%m/%d/%Y'),
                current_week_start,
                current_week_end,
                prev_week_start,
                prev_week_end,
                prev_year_start,
                prev_year_end,
                four_weeks_start,
                current_week_end,
                prev_four_weeks_start,
                prev_four_weeks_end
            )

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
        worship_ytd_totals = {}  # Track worship YTD separately
        four_week_totals = {'current': 0, 'previous_year': 0}  # For 4-week comparison
        
        for i in range(YEARS_TO_DISPLAY):
            year = self.current_year - i
            overall_totals[year] = 0
            ytd_overall_totals[year] = 0
            worship_ytd_totals[year] = 0
            
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
                # Track worship YTD separately
                if program.Name == WORSHIP_PROGRAM:
                    worship_ytd_totals[year] = years_program_ytd[year]['total']
                
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
        try:
            overall_summary = self.generate_overall_summary(
                overall_totals[self.current_year],
                overall_totals[self.current_year - 1] if self.current_year - 1 in overall_totals else 0,
                ytd_overall_totals[self.current_year],
                ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0,
                four_week_totals if SHOW_FOUR_WEEK_COMPARISON else None,
                for_email=for_email
            )
            report_content += overall_summary
        except Exception as e:
            # If there's an error, add a placeholder and log it
            error_msg = "<!-- Error generating overall summary: {} -->".format(str(e))
            report_content += error_msg
            if not for_email:
                print "<div style='color: red;'>Error: {}</div>".format(str(e))
        
        # Add exceptions display if enabled
        if SHOW_EXCEPTIONS and self.exceptions_manager:
            # Get exceptions for the current report period (4 weeks)
            exceptions_in_period = self.exceptions_manager.get_exceptions_in_range(
                self.four_weeks_ago_start_date,
                self.report_date
            )
            if exceptions_in_period:
                report_content += self.exceptions_manager.format_exceptions_html(exceptions_in_period)
        
        # Add fiscal year-to-date summary (only once) - pass worship-specific YTD
        report_content += self.generate_fiscal_year_summary(
            worship_ytd_totals.get(self.current_year, 0),
            worship_ytd_totals.get(self.current_year - 1, 0)
        )
                
        # Add enrollment analysis section if any configured programs are present
        program_names = [p.Name for p in programs]
        has_enrollment_programs = any(prog in ENROLLMENT_ANALYSIS_PROGRAMS for prog in program_names)

        if has_enrollment_programs:
            enrollment_section = self.generate_enrollment_analysis_section(for_email=for_email)
            report_content += enrollment_section
        
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
    
    def generate_exceptions_management_section(self):
        """Generate an interactive section to manage attendance exceptions."""
        # Check if user has permission to manage exceptions
        is_admin = False
        try:
            # Check if user has Admin role using built-in TouchPoint method
            is_admin = model.UserIsInRole("Admin")
        except:
            is_admin = False
        
        # Override permission check if config says not to require admin
        if not REQUIRE_ADMIN_FOR_EXCEPTION_MANAGEMENT:
            is_admin = True
        
        # Generate the management UI
        html = """
        <div id="exceptions-management" style="display: none; margin-top: 40px; padding: 20px; background-color: #f9f9f9; border: 1px solid #ddd; border-radius: 5px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h3 style="margin-top: 0;">📅 Attendance Exceptions</h3>
                <button type="button" onclick="toggleExceptionsPanel()" style="padding: 3px 8px; background-color: #666; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;">✖ Close</button>
            </div>
            <p style="color: #666;">Special dates that affect attendance patterns (holidays, weather events, etc.)</p>
        """
        
        # Only show add/delete forms if user has permission
        if is_admin:
            html += """
            <!-- Add New Exception Form -->
            <div style="background-color: #fff; padding: 15px; border-radius: 5px; margin: 15px 0;">
                <h4>Add New Exception</h4>
                <form id="add-exception-form" style="display: flex; gap: 10px; align-items: flex-end; flex-wrap: wrap;">
                    <div style="display: flex; flex-direction: column;">
                        <label style="font-size: 12px; color: #666; margin-bottom: 2px;">Date</label>
                        <input type="date" id="exception_date" required 
                               style="padding: 5px; border: 1px solid #ddd; border-radius: 3px;">
                    </div>
                    
                    <div style="display: flex; flex-direction: column; flex: 1; min-width: 200px;">
                        <label style="font-size: 12px; color: #666; margin-bottom: 2px;">Description</label>
                        <input type="text" id="exception_description" required 
                               placeholder="e.g., Easter Sunday, Winter Storm"
                               style="padding: 5px; border: 1px solid #ddd; border-radius: 3px;">
                    </div>
                    
                    <div style="display: flex; gap: 15px; align-items: center;">
                        <label style="display: flex; align-items: center; font-size: 14px;">
                            <input type="checkbox" id="affects_worship" checked style="margin-right: 5px;">
                            ⛪ Worship
                        </label>
                        <label style="display: flex; align-items: center; font-size: 14px;">
                            <input type="checkbox" id="affects_groups" checked style="margin-right: 5px;">
                            👥 Groups
                        </label>
                    </div>
                    
                    <button type="button" onclick="addException()" style="padding: 5px 15px; background-color: #4CAF50; color: white; border: none; border-radius: 3px; cursor: pointer;">
                        Add Exception
                    </button>
                </form>
                <div id="exception-message" style="margin-top: 10px; padding: 10px; border-radius: 3px; display: none;"></div>
            </div>
            """
        
        # Always show current exceptions list (for both admins and viewers)
        html += """
            <!-- Current Exceptions List -->
            <div style="background-color: #fff; padding: 15px; border-radius: 5px; margin: 15px 0;">
                <h4>Current Exceptions</h4>
        """
        
        if self.exceptions_manager and self.exceptions_manager.exceptions:
            html += """
                <div style="max-height: 300px; overflow-y: auto;">
                    <table style="width: 100%; border-collapse: collapse;">
                        <thead>
                            <tr style="background-color: #f0f0f0;">
                                <th style="padding: 8px; text-align: left; border-bottom: 2px solid #ddd;">Date</th>
                                <th style="padding: 8px; text-align: left; border-bottom: 2px solid #ddd;">Description</th>
                                <th style="padding: 8px; text-align: center; border-bottom: 2px solid #ddd;">Affects</th>
            """
            
            # Only show Action column header if admin
            if is_admin:
                html += '<th style="padding: 8px; text-align: center; border-bottom: 2px solid #ddd;">Action</th>'
            
            html += """
                            </tr>
                        </thead>
                        <tbody>
            """
            
            for exc in sorted(self.exceptions_manager.exceptions, key=lambda x: x['date']):
                date_str = exc['date'].strftime('%Y-%m-%d')
                day_name = exc['date'].strftime('%A')
                
                # Determine icons based on flags
                flags = exc.get('flags', 'B')
                if flags == 'W':
                    affects_icons = '⛪'
                    affects_text = 'Worship'
                elif flags == 'G':
                    affects_icons = '👥'
                    affects_text = 'Groups'
                else:
                    affects_icons = '⛪👥'
                    affects_text = 'Both'
                
                html += """
                    <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #eee;">
                            <code>{}</code><br>
                            <span style="font-size: 11px; color: #666;">{}</span>
                        </td>
                        <td style="padding: 8px; border-bottom: 1px solid #eee;">{}</td>
                        <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">
                            <span title="{}">{}</span>
                        </td>
                """.format(date_str, day_name, exc['description'], affects_text, affects_icons)
                
                # Only show delete button if admin
                if is_admin:
                    html += """
                        <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">
                            <button type="button" 
                                    onclick="deleteException('{}')"
                                    style="padding: 2px 8px; background-color: #f44336; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;">
                                Delete
                            </button>
                        </td>
                    """.format(date_str)
                
                html += "</tr>"
            
            html += """
                        </tbody>
                    </table>
                </div>
            """
        else:
            html += """
                <p style="color: #666; font-style: italic;">No exceptions currently defined.</p>
            """
        
        html += """
            </div>
            
            <!-- Instructions -->
            <div style="background-color: #e8f4fd; padding: 15px; border-radius: 5px; border-left: 4px solid #2196F3;">
                <h4 style="margin-top: 0;">ℹ️ About Exceptions</h4>
                <ul style="margin: 10px 0; line-height: 1.6;">
                    <li><strong>Purpose:</strong> Exceptions help explain unusual attendance patterns in reports</li>
                    <li><strong>Worship (⛪):</strong> Affects Sunday morning worship services</li>
                    <li><strong>Groups (👥):</strong> Affects small groups, classes, and midweek programs</li>
                    <li><strong>Auto-sorting:</strong> Exceptions are automatically sorted by date</li>
                    <li><strong>Storage:</strong> Saved in Special Content as <code>{}</code></li>
                </ul>
                
                <details style="margin-top: 10px;">
                    <summary style="cursor: pointer; color: #1976D2;">Advanced: Direct Content Editing</summary>
                    <div style="margin-top: 10px; padding: 10px; background-color: #fff; border-radius: 3px;">
                        <p>You can also manage exceptions directly:</p>
                        <ol>
                            <li>Go to <strong>Administration → Special Content → Text</strong></li>
                            <li>Edit content named: <code>{}</code></li>
                            <li>Format: <code>YYYY-MM-DD | Description | Flags</code></li>
                            <li>Flags: W=Worship, G=Groups, B=Both (default)</li>
                        </ol>
                    </div>
                </details>
            </div>
        </div>
        """
        
        # Add JavaScript functions if user is admin (for add/delete functionality)
        if is_admin:
            html += '''
        <script>
        function getPyScriptAddress() {
            // Use the current path - this works regardless of script name
            // For POST requests, must use /PyScriptForm/ instead of /PyScript/
            let path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        }
        
        function addException() {
            var date = document.getElementById('exception_date').value;
            var description = document.getElementById('exception_description').value;
            var affectsWorship = document.getElementById('affects_worship').checked;
            var affectsGroups = document.getElementById('affects_groups').checked;
            
            if (!date || !description) {
                showMessage('Please provide both date and description', 'error');
                return;
            }
            
            // Make AJAX request
            var xhr = new XMLHttpRequest();
            xhr.open('POST', getPyScriptAddress(), true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    console.log('Response status:', xhr.status);
                    console.log('Response text:', xhr.responseText);
                    
                    if (xhr.status === 200) {
                        try {
                            var response = JSON.parse(xhr.responseText);
                            console.log('Parsed response:', response);
                            
                            if (response.success) {
                                showMessage('Exception added successfully! Refreshing...', 'success');
                                setTimeout(function() {
                                    window.location.reload();
                                }, 1500);
                            } else {
                                var errorMsg = 'Error: ' + response.message;
                                if (response.traceback) {
                                    console.error('Traceback:', response.traceback);
                                }
                                showMessage(errorMsg, 'error');
                            }
                        } catch(e) {
                            console.error('Failed to parse JSON:', e);
                            console.error('Raw response:', xhr.responseText);
                            showMessage('Error parsing response. Check console for details.', 'error');
                        }
                    } else {
                        showMessage('Error adding exception (HTTP ' + xhr.status + ')', 'error');
                    }
                }
            };
            
            var params = 'ajax_exception_action=add' +
                        '&exception_date=' + encodeURIComponent(date) +
                        '&exception_description=' + encodeURIComponent(description) +
                        '&affects_worship=' + affectsWorship +
                        '&affects_groups=' + affectsGroups;
            
            console.log('Sending AJAX request with params:', params);
            xhr.send(params);
        }
        
        function deleteException(date) {
            if (!confirm('Delete exception for ' + date + '?')) {
                return;
            }
            
            // Make AJAX request
            var xhr = new XMLHttpRequest();
            xhr.open('POST', getPyScriptAddress(), true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    console.log('Delete response status:', xhr.status);
                    console.log('Delete response text:', xhr.responseText);
                    
                    if (xhr.status === 200) {
                        try {
                            var response = JSON.parse(xhr.responseText);
                            console.log('Delete parsed response:', response);
                            
                            if (response.success) {
                                showMessage('Exception deleted successfully! Refreshing...', 'success');
                                setTimeout(function() {
                                    window.location.reload();
                                }, 1500);
                            } else {
                                var errorMsg = 'Error: ' + response.message;
                                if (response.traceback) {
                                    console.error('Delete traceback:', response.traceback);
                                }
                                showMessage(errorMsg, 'error');
                            }
                        } catch(e) {
                            console.error('Failed to parse delete JSON:', e);
                            console.error('Raw delete response:', xhr.responseText);
                            showMessage('Error parsing delete response. Check console for details.', 'error');
                        }
                    } else {
                        showMessage('Error deleting exception (HTTP ' + xhr.status + ')', 'error');
                    }
                }
            };
            
            var params = 'ajax_exception_action=delete&delete_date=' + encodeURIComponent(date);
            console.log('Sending delete request with params:', params);
            xhr.send(params);
        }
        
        function showMessage(message, type) {
            var msgDiv = document.getElementById('exception-message');
            if (msgDiv) {
                msgDiv.style.display = 'block';
                msgDiv.style.backgroundColor = type === 'success' ? '#d4edda' : '#f8d7da';
                msgDiv.style.color = type === 'success' ? '#155724' : '#721c24';
                msgDiv.style.border = '1px solid ' + (type === 'success' ? '#c3e6cb' : '#f5c6cb');
                msgDiv.textContent = message;
                
                // Auto-hide after 5 seconds
                setTimeout(function() {
                    msgDiv.style.display = 'none';
                }, 5000);
            }
        }
        </script>
        '''
        
        # Add the toggle function that's always available (for all users)
        html += '''
        <script>
        function toggleExceptionsPanel() {
            var panel = document.getElementById('exceptions-management');
            if (panel) {
                if (panel.style.display === 'none' || panel.style.display === '') {
                    panel.style.display = 'block';
                    // Scroll to the panel
                    panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
                } else {
                    panel.style.display = 'none';
                }
            }
        }
        </script>
        '''
        
        return html
    
    def generate_four_week_exceptions_display(self):
        """Generate a display of exceptions that fall within the 4-week comparison window."""
        if not SHOW_FOUR_WEEK_COMPARISON or not self.exceptions_manager:
            return ""
            
        # Calculate the date range for 4 weeks back and previous year 4 weeks
        current_date = self.report_date
        four_weeks_back = current_date - datetime.timedelta(weeks=4)
        
        # Previous year same period
        prev_year_date = ReportHelper.get_previous_year_sunday(current_date)
        prev_year_four_weeks_back = prev_year_date - datetime.timedelta(weeks=4)
        
        # Get exceptions in both ranges
        current_exceptions = []
        prev_year_exceptions = []
        
        for exc in self.exceptions_manager.exceptions:
            exc_date = exc['date']
            # Check current 4-week window
            if four_weeks_back <= exc_date <= current_date:
                current_exceptions.append(exc)
            # Check previous year 4-week window
            if prev_year_four_weeks_back <= exc_date <= prev_year_date:
                prev_year_exceptions.append(exc)
        
        # If no exceptions in either period, return empty
        if not current_exceptions and not prev_year_exceptions:
            return ""
        
        # Build the display
        html = """
        <div style="margin-top: 30px; padding: 15px; background-color: #fff3cd; border: 1px solid #ffeeba; border-radius: 5px;">
            <h4 style="margin-top: 0; color: #856404;">⚠️ Exceptions in 4-Week Comparison Period</h4>
        """
        
        if current_exceptions:
            html += """
            <div style="margin-bottom: 15px;">
                <strong>Current Period ({} to {}):</strong>
                <ul style="margin-top: 5px;">
            """.format(four_weeks_back.strftime('%m/%d'), current_date.strftime('%m/%d'))
            
            for exc in sorted(current_exceptions, key=lambda x: x['date']):
                date_str = exc['date'].strftime('%m/%d')
                day_name = exc['date'].strftime('%a')
                flags = exc.get('flags', 'B')
                icons = '⛪' if flags == 'W' else '👥' if flags == 'G' else '⛪👥'
                html += '<li>{} ({}) - {} {}</li>'.format(
                    date_str, day_name, exc['description'], icons
                )
            html += "</ul></div>"
        
        if prev_year_exceptions:
            html += """
            <div>
                <strong>Previous Year Period ({} to {}):</strong>
                <ul style="margin-top: 5px;">
            """.format(prev_year_four_weeks_back.strftime('%m/%d/%Y'), prev_year_date.strftime('%m/%d/%Y'))
            
            for exc in sorted(prev_year_exceptions, key=lambda x: x['date']):
                date_str = exc['date'].strftime('%m/%d/%Y')
                day_name = exc['date'].strftime('%a')
                flags = exc.get('flags', 'B')
                icons = '⛪' if flags == 'W' else '👥' if flags == 'G' else '⛪👥'
                html += '<li>{} ({}) - {} {}</li>'.format(
                    date_str, day_name, exc['description'], icons
                )
            html += "</ul></div>"
        
        html += """
            <p style="margin: 10px 0 0 0; font-size: 12px; color: #666;">
                Icons: ⛪ = Affects Worship | 👥 = Affects Groups | ⛪👥 = Affects Both
            </p>
        </div>
        """
        
        return html
    
    def generate_all_exceptions_display(self):
        """Generate a collapsible display of all exceptions."""
        if not SHOW_EXCEPTIONS or not self.exceptions_manager or not self.exceptions_manager.exceptions:
            return ""
        
        # Build the display with all exceptions
        html = """
        <div style="margin-top: 20px; padding: 15px; background-color: #f9f9f9; border: 1px solid #ddd; border-radius: 5px;">
            <details style="cursor: pointer;">
                <summary style="font-weight: bold; color: #333; padding: 5px;">
                    📅 View All Exceptions ({} total)
                </summary>
                <div style="margin-top: 15px; max-height: 400px; overflow-y: auto;">
                    <table style="width: 100%; border-collapse: collapse; background-color: white;">
                        <thead>
                            <tr style="background-color: #f0f0f0; position: sticky; top: 0;">
                                <th style="padding: 8px; text-align: left; border-bottom: 2px solid #ddd;">Date</th>
                                <th style="padding: 8px; text-align: left; border-bottom: 2px solid #ddd;">Description</th>
                                <th style="padding: 8px; text-align: center; border-bottom: 2px solid #ddd;">Affects</th>
                                <th style="padding: 8px; text-align: left; border-bottom: 2px solid #ddd;">Year</th>
                            </tr>
                        </thead>
                        <tbody>
        """.format(len(self.exceptions_manager.exceptions))
        
        current_year = self.report_date.year
        
        for exc in sorted(self.exceptions_manager.exceptions, key=lambda x: x['date'], reverse=True):
            date_str = exc['date'].strftime('%m/%d/%Y')
            day_name = exc['date'].strftime('%a')
            year = exc['date'].year
            
            # Determine icons based on flags
            flags = exc.get('flags', 'B')
            if flags == 'W':
                affects_icons = '⛪'
                affects_text = 'Worship'
            elif flags == 'G':
                affects_icons = '👥'
                affects_text = 'Groups'
            else:
                affects_icons = '⛪👥'
                affects_text = 'Both'
            
            # Highlight current year exceptions
            row_style = 'background-color: #fffbf0;' if year == current_year else ''
            
            html += """
                        <tr style="{}">
                            <td style="padding: 8px; border-bottom: 1px solid #eee;">
                                <strong>{}</strong> ({})
                            </td>
                            <td style="padding: 8px; border-bottom: 1px solid #eee;">{}</td>
                            <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">
                                <span title="{}">{}</span>
                            </td>
                            <td style="padding: 8px; border-bottom: 1px solid #eee;">
                                {}
                            </td>
                        </tr>
            """.format(row_style, date_str, day_name, exc['description'], affects_text, affects_icons, year)
        
        html += """
                        </tbody>
                    </table>
                    <p style="margin: 10px 0 0 0; font-size: 11px; color: #666;">
                        <strong>Note:</strong> Exceptions are sorted by date (newest first). Current year exceptions are highlighted.
                    </p>
                </div>
            </details>
        </div>
        """
        
        return html
    
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
            report_html += "th, td {border: 1px solid #ddd; padding: 4px 8px; text-align: center;}\n"
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
            if model.Data.send_email == "yes":
                # Use email_to from form if provided, otherwise use EMAIL_TO config
                email_to = model.Data.email_to if hasattr(model.Data, 'email_to') and model.Data.email_to else EMAIL_TO
                email_report = self.generate_report_content(for_email=True)
                success, message = self.send_email_report(email_to, email_report)

                # The send_email_report method already prints the status message via JavaScript
                # No need to add duplicate message here
            
            # Add date selector form 
            report_html += self.generate_date_selector_form()
            
            # Add exceptions management panel (initially hidden)
            if SHOW_EXCEPTIONS and self.exceptions_manager:
                report_html += self.generate_exceptions_management_section()
            
            # Check if we should load data
            if self.should_load_data:
                # Add date range summary
                report_html += self.generate_date_range_summary()
                
                # Calculate overall totals for summaries
                performance_timer.start("calculate_totals")
                overall_totals = {}
                ytd_overall_totals = {}
                worship_ytd_totals = {}  # Track worship YTD separately
                four_week_totals = {'current': 0, 'previous_year': 0}
                
                for i in range(YEARS_TO_DISPLAY):
                    year = self.current_year - i
                    overall_totals[year] = 0
                    ytd_overall_totals[year] = 0
                    worship_ytd_totals[year] = 0
                
                # Get all programs
                performance_timer.start("load_programs")
                programs = q.QuerySql(self.get_programs_sql())
                performance_timer.end("load_programs")
                
                # Pre-calculate overall totals
                for program in programs:
                    years_program_data = self.get_multiple_years_attendance_data(
                        self.report_date,
                        program_id=program.Id
                    )
                    
                    years_program_ytd = self.get_multiple_years_ytd_data(
                        self.report_date,
                        program_id=program.Id
                    )
                    
                    if SHOW_FOUR_WEEK_COMPARISON:
                        four_week_data = self.get_four_week_attendance_comparison(program_id=program.Id)
                        four_week_totals['current'] += four_week_data['current']['total']
                        four_week_totals['previous_year'] += four_week_data['previous_year']['total']
                    
                    for year in years_program_data:
                        overall_totals[year] += years_program_data[year]['total']
                        
                    for year in years_program_ytd:
                        ytd_overall_totals[year] += years_program_ytd[year]['total']
                        # Track worship YTD separately
                        if program.Name == WORSHIP_PROGRAM:
                            worship_ytd_totals[year] = years_program_ytd[year]['total']
                
                performance_timer.end("calculate_totals")
                
                # Add the overall summary section
                report_html += self.generate_overall_summary(
                    overall_totals[self.current_year],
                    overall_totals[self.current_year - 1] if self.current_year - 1 in overall_totals else 0,
                    ytd_overall_totals[self.current_year],
                    ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0,
                    four_week_totals if SHOW_FOUR_WEEK_COMPARISON else None
                )
                
                # Add fiscal year-to-date summary - pass worship-specific YTD
                report_html += self.generate_fiscal_year_summary(
                    worship_ytd_totals.get(self.current_year, 0),
                    worship_ytd_totals.get(self.current_year - 1, 0)
                )
                
                # Add enrollment analysis section if configured
                if any(prog in ENROLLMENT_ANALYSIS_PROGRAMS for prog in [p.Name for p in programs]):
                    report_html += self.generate_enrollment_analysis_section()
                
                # Get data for program totals section
                current_week_data = self.get_week_attendance_data(self.week_start_date, self.report_date)
                previous_week_data = self.get_week_attendance_data(
                    self.previous_week_start_date, 
                    self.previous_year_date
                )
                
                # Add program totals section if enabled
                if SHOW_PROGRAM_SUMMARY:
                    report_html += self.generate_selected_program_average_summaries(
                        current_week_data,
                        previous_week_data,
                        ytd_overall_totals[self.current_year],
                        ytd_overall_totals[self.current_year - 1] if self.current_year - 1 in ytd_overall_totals else 0
                    )
                
                # Now generate the detailed program tables using the optimized method
                report_html += self.generate_report_content_optimized()
                
                # Add 4-week exceptions display at the bottom if applicable
                if SHOW_EXCEPTIONS and self.exceptions_manager:
                    report_html += self.generate_four_week_exceptions_display()
                    
                    # Add the full exceptions list (collapsible)
                    report_html += self.generate_all_exceptions_display()
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
    
# Create and run the report
try:
    # Handle AJAX exception management requests first
    # TouchPoint may pass form data with or without 'Data.' prefix
    ajax_action = getattr(model.Data, 'ajax_exception_action', '')
    if not ajax_action:
        # Try without Data prefix
        ajax_action = getattr(model, 'ajax_exception_action', '')
    
    # Also check for common TouchPoint form parameter patterns
    if not ajax_action and hasattr(model, 'Data'):
        for attr in dir(model.Data):
            if not attr.startswith('_'):
                try:
                    if attr == 'ajax_exception_action':
                        ajax_action = str(getattr(model.Data, attr, ''))
                        break
                except:
                    pass
    
    if ajax_action:
        import json
        response = {'success': False, 'message': ''}
        
        try:
            # Check if user has permission to manage exceptions (for AJAX operations)
            is_admin = False
            if REQUIRE_ADMIN_FOR_EXCEPTION_MANAGEMENT:
                try:
                    # Use built-in TouchPoint method to check Admin role
                    is_admin = model.UserIsInRole("Admin")
                except:
                    is_admin = False
            else:
                is_admin = True  # If admin not required, allow all users
            
            # Check permission before allowing any modifications
            if not is_admin:
                response = {'success': False, 'message': 'Permission denied. Admin access required to manage exceptions.'}
                model.Header = "application/json"
                print json.dumps(response)
            else:
                # Load exceptions manager
                exceptions_manager = ExceptionsManager()
                action = ajax_action
                
                if action == 'add':
                    # Add new exception - check both model.Data and model for parameters
                    exc_date = getattr(model.Data, 'exception_date', '')
                    if not exc_date:
                        exc_date = getattr(model, 'exception_date', '')
                        
                    exc_desc = getattr(model.Data, 'exception_description', '')
                    if not exc_desc:
                        exc_desc = getattr(model, 'exception_description', '')
                        
                    affects_worship_str = getattr(model.Data, 'affects_worship', '')
                    if not affects_worship_str:
                        affects_worship_str = getattr(model, 'affects_worship', 'false')
                    affects_worship = affects_worship_str == 'true'
                    
                    affects_groups_str = getattr(model.Data, 'affects_groups', '')
                    if not affects_groups_str:
                        affects_groups_str = getattr(model, 'affects_groups', 'false')
                    affects_groups = affects_groups_str == 'true'
                    
                    if exc_date and exc_desc:
                        # Determine flags
                        if affects_worship and affects_groups:
                            flags = 'B'
                        elif affects_worship:
                            flags = 'W'
                        elif affects_groups:
                            flags = 'G'
                        else:
                            flags = 'B'  # Default to both if neither selected
                        
                        # Add to exceptions
                        exceptions_manager.exceptions.append({
                            'date': datetime.datetime.strptime(exc_date, '%Y-%m-%d'),
                            'description': exc_desc,
                            'flags': flags
                        })
                        
                        # Save and report result
                        save_result = exceptions_manager.save_exceptions()
                        if save_result:
                            # Verify the save by checking the content
                            saved_content = model.TextContent(EXCEPTIONS_CONTENT_NAME)
                            exception_count = len(exceptions_manager.exceptions)
                            response = {
                                'success': True, 
                                'message': 'Exception added successfully. Total exceptions: {}'.format(exception_count)
                            }
                        else:
                            response = {'success': False, 'message': 'Failed to save exception to Special Content'}
                    else:
                        response = {'success': False, 'message': 'Date and description are required'}
                
                elif action == 'delete':
                    # Delete exception - check both model.Data and model for parameters
                    delete_date = getattr(model.Data, 'delete_date', '')
                    if not delete_date:
                        delete_date = getattr(model, 'delete_date', '')
                    if delete_date:
                        delete_dt = datetime.datetime.strptime(delete_date, '%Y-%m-%d')
                        original_count = len(exceptions_manager.exceptions)
                        exceptions_manager.exceptions = [
                            exc for exc in exceptions_manager.exceptions 
                            if exc['date'] != delete_dt
                        ]
                        new_count = len(exceptions_manager.exceptions)
                        
                        if original_count > new_count:
                            save_result = exceptions_manager.save_exceptions()
                            if save_result:
                                response = {
                                    'success': True, 
                                    'message': 'Exception deleted successfully. Remaining exceptions: {}'.format(new_count)
                                }
                            else:
                                response = {'success': False, 'message': 'Failed to save changes to Special Content'}
                        else:
                            response = {'success': False, 'message': 'Exception not found for the specified date'}
                    else:
                        response = {'success': False, 'message': 'Date is required'}
                else:
                    response = {'success': False, 'message': 'Unknown action: ' + str(action)}
                    
        except Exception as e:
            import traceback
            response = {'success': False, 'message': 'Error: ' + str(e), 'traceback': traceback.format_exc()}
        
        # Return JSON response only - no HTML
        print json.dumps(response)
    
    else:
        # Only run the regular report if not an AJAX request
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
