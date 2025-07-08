#####################################################################
#### COMMUNICATION DASHBOARD
#####################################################################
# This dashboard provides comprehensive communication analytics including:
# - Email delivery metrics with success/failure breakdown
# - SMS/text message statistics with delivery performance
# - Top senders analytics with delivery rates
# - Program-specific communication metrics
#
#####################################################################
#### UPLOAD INSTRUCTIONS
#####################################################################
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python script file with your preferred name and paste all this code
# 4. Test and optionally add to menu

# Written By: Ben Swaby
# Email: bswaby@fbchtn.org

#####################################################################
#### IMPORT LIBRARIES AND CONFIGURATION
#####################################################################
from datetime import datetime, timedelta
import traceback

# Initialize page title
model.Header = 'Communication Dashboard'

#####################################################################
#### CONFIGURATION OPTIONS
#####################################################################

# Configuration class to make the dashboard customizable
class Config:
    # Dashboard title that appears at the top of the page
    DASHBOARD_TITLE = "Communication Dashboard"
    
    # Default number of days to look back for data when first loading
    DEFAULT_LOOKBACK_DAYS = 30
    
    # Maximum number of records to show in tables
    MAX_ROWS_PER_TABLE = 150  # Show up to 150 campaigns without pagination
    
    # Default tab to show when no tab is selected
    DEFAULT_TAB = "dashboard"  # Options: dashboard, email, sms, senders, programs
    
    # Color scheme for the dashboard (Bootstrap classes)
    PRIMARY_COLOR = "primary"     # Blue
    SUCCESS_COLOR = "success"     # Green
    INFO_COLOR = "info"           # Light blue
    WARNING_COLOR = "warning"     # Yellow
    DANGER_COLOR = "danger"       # Red
    
    # Table styles
    TABLE_STRIPED = True
    TABLE_BORDERED = True
    TABLE_HOVER = True
    
    # Debug mode - set to True to display more detailed error messages
    DEBUG_MODE = True
    
    # Pagination debug mode - no longer used (pagination removed)
    PAGINATION_DEBUG = False
    
    # Performance mode - set to True to use simplified queries for large datasets
    PERFORMANCE_MODE = False  # Disabled actual metrics

#####################################################################
#### INITIALIZATION
#####################################################################

# Helper function to safely get form data with defaults
def get_form_data(attr_name, default_value):
    try:
        value = getattr(model.Data, attr_name)
        # Check if value is empty or None
        if value is None or str(value).strip() == '':
            return default_value
        return value
    except:
        return default_value

# Helper function to safely get integer form data
def get_form_data_int(attr_name, default_value):
    try:
        value = getattr(model.Data, attr_name)
        # Check if value is empty or None
        if value is None or str(value).strip() == '':
            return default_value
        # Try to convert to int
        try:
            return int(value)
        except ValueError:
            return default_value
    except:
        return default_value

# Helper function to get the current script URL
def get_current_script_url():
    """Get the current script URL dynamically"""
    # Check if we have script name from model
    if hasattr(model, 'ScriptName'):
        return "/PyScript/" + model.ScriptName
    # Otherwise return empty string to post to current URL
    return ""

# Initialize data model with default values if not already set
today = datetime.now()
default_start = today - timedelta(days=Config.DEFAULT_LOOKBACK_DAYS)

# Set default date range
model.Data.sDate = get_form_data('sDate', default_start.strftime('%Y-%m-%d'))
model.Data.eDate = get_form_data('eDate', today.strftime('%Y-%m-%d'))

# Set default filters
model.Data.program = get_form_data('program', '999999')
model.Data.failclassification = get_form_data('failclassification', '999999')
model.Data.HideSuccess = get_form_data('HideSuccess', 'yes')

# Set active tab
model.Data.activeTab = get_form_data('activeTab', Config.DEFAULT_TAB)

# Current date for reference
current_date = datetime.now().strftime("%B %d, %Y")

#####################################################################
#### HELPER FUNCTIONS
#####################################################################

class FormatHelper:
    """Helper class for formatting data values"""
    
    @staticmethod
    def format_percent(value):
        """Format a value as a percentage string"""
        try:
            return "{0:.1f}%".format(float(value))
        except:
            return "0.0%"
    
    @staticmethod
    def safe_div(a, b):
        """Safely divide a by b, returning 0 if b is 0"""
        try:
            # Convert inputs to float and handle potential None values
            a_val = float(a or 0)
            b_val = float(b or 0)
            
            if b_val == 0:
                return 0
            return a_val / b_val
        except:
            return 0
    
    @staticmethod
    def format_number(value):
        """Format a number with commas"""
        try:
            return "{:,}".format(int(value))
        except:
            return str(value)

    @staticmethod
    def format_date(date_str):
        """Format a date string for display"""
        try:
            if isinstance(date_str, datetime):
                return date_str.strftime("%Y-%m-%d")
            elif isinstance(date_str, str):
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                return dt.strftime("%b %d, %Y")
            return str(date_str)
        except:
            return str(date_str)

class TableGenerator:
    """Class for generating HTML tables"""
    
    @staticmethod
    def generate_html_table(data, hide_columns=None, column_order=None, header_labels=None, 
                          url_columns=None, extra_classes="", id=None):
        """
        Generate an HTML table with various options for formatting
        
        Args:
            data: List of dictionaries containing the data
            hide_columns: List of column names to hide
            column_order: List of column names to specify order
            header_labels: Dictionary mapping column names to display labels
            url_columns: Dictionary mapping column names to URL formats
            extra_classes: Additional CSS classes for the table
            id: Optional ID for the table element
        """
        if not data:
            return "<div class='alert alert-info'>No data available for the selected filters.</div>"
        
        # Set defaults
        hide_columns = hide_columns or []
        column_order = column_order or []
        header_labels = header_labels or {}
        url_columns = url_columns or {}
        
        # Get headers, respecting column_order
        headers = data[0].keys()
        if column_order:
            headers = [h for h in column_order if h in headers] + [h for h in headers if h not in column_order and h not in hide_columns]
        else:
            headers = [h for h in headers if h not in hide_columns]
        
        # Build table CSS classes
        table_classes = "table"
        if Config.TABLE_STRIPED:
            table_classes += " table-striped"
        if Config.TABLE_BORDERED:
            table_classes += " table-bordered"
        if Config.TABLE_HOVER:
            table_classes += " table-hover"
        if extra_classes:
            table_classes += " " + extra_classes
            
        # Table ID attribute
        id_attr = ' id="{}"'.format(id) if id else ''
        
        # Start generating the HTML table
        html = '<table class="{}"{}>\n'.format(table_classes, id_attr)
        
        # Table header
        html += '  <thead>\n    <tr>\n'
        for header in headers:
            display_name = header_labels.get(header, header)
            html += '      <th>{}</th>\n'.format(display_name)
        html += '    </tr>\n  </thead>\n'
        
        # Table body
        html += '  <tbody>\n'
        for row in data:
            html += '    <tr>\n'
            for header in headers:
                cell_value = row.get(header, '')
                
                # Handle URL columns
                if header in url_columns:
                    url_format = url_columns[header]
                    # Replace placeholders with values from the row
                    url = url_format
                    for key, value in row.items():
                        placeholder = '{' + key + '}'
                        url = url.replace(placeholder, str(value))
                    cell_value = '<a href="{}">{}</a>'.format(url, cell_value)
                
                html += '      <td>{}</td>\n'.format(cell_value)
            html += '    </tr>\n'
        html += '  </tbody>\n'
        
        # Finish the table
        html += '</table>\n'
        return html

class ErrorHandler:
    """Class for handling and displaying errors"""
    
    @staticmethod
    def handle_error(error, section_name=""):
        """Format an error message for display"""
        if Config.DEBUG_MODE:
            return """
            <div class="alert alert-danger">
                <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Error in {0}</h4>
                <p>{1}</p>
                <pre style="max-height: 200px; overflow-y: auto;">{2}</pre>
            </div>
            """.format(section_name, str(error), traceback.format_exc())
        else:
            return """
            <div class="alert alert-danger">
                <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Error</h4>
                <p>An error occurred while loading {0}. Please try again later or contact support.</p>
            </div>
            """.format(section_name)

class DatabaseHelper:
    """Helper class for database operations"""
    
    @staticmethod
    def table_exists(table_name):
        """Check if a table exists in the database"""
        try:
            check_sql = """
            SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_NAME = '{}'
            """.format(table_name)
            
            count_result = q.QuerySqlScalar(check_sql)
            return count_result is not None and int(count_result) > 0
        except:
            return False
    
    @staticmethod
    def get_person_name(people_id):
        """Get a person's name from their PeopleId"""
        if not people_id or people_id <= 0:
            return "Unknown"
            
        try:
            sql = "SELECT Name FROM People WHERE PeopleId = {}".format(people_id)
            result = q.QuerySqlScalar(sql)
            return result if result else "Unknown"
        except:
            return "Unknown"

#####################################################################
#### DATA RETRIEVAL FUNCTIONS
#####################################################################

class DataRetrieval:
    """Class for retrieving data from the database"""
    
    @staticmethod
    def get_dashboard_summary(sDate, eDate):
        """Get summary data for the main dashboard with correct SMS campaign counting"""
        try:
            # Initialize with default values
            email_summary = {
                'total_emails': 0,
                'sent_emails': 0,
                'failed_emails': 0,
                'delivery_rate': 0,
                'failed_types': []
            }
            
            sms_stats = {
                'total_count': 0,
                'sent_count': 0,
                'delivery_rate': 0
            }
            
            email_campaign_count = 0
            sms_campaign_count = 0
            
            # Try to get email summary - continue even if this fails
            try:
                email_summary = DataRetrieval.get_email_summary(sDate, eDate, 'yes', '999999', '999999')
            except Exception as e:
                # Log but continue - we'll use default values
                print("<div class='alert alert-warning'>Warning: Unable to retrieve email statistics. {0}</div>".format(str(e)))
            
            # Try to get SMS summary - continue even if this fails
            try:
                sms_stats = DataRetrieval.get_sms_stats(sDate, eDate)
            except Exception as e:
                # Log but continue - we'll use default values
                print("<div class='alert alert-warning'>Warning: Unable to retrieve SMS statistics. {0}</div>".format(str(e)))
            
            # Get email campaign count with extreme robustness
            try:
                # Try multiple methods to get email campaign count
                try:
                    # Simplest possible SQL to avoid nulls completely
                    campaign_sql = """
                    SELECT COUNT(DISTINCT Id) FROM EmailQueue
                    WHERE Id IN (
                        SELECT DISTINCT Id FROM EmailQueueTo 
                        WHERE Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    )
                    """.format(sDate, eDate)
                    
                    # Try direct scalar first for simplicity
                    scalar_result = q.QuerySqlScalar(campaign_sql)
                    if scalar_result is not None:
                        try:
                            email_campaign_count = int(scalar_result)
                        except:
                            # Try string conversion
                            try:
                                email_campaign_count = int(float(str(scalar_result)))
                            except:
                                email_campaign_count = 0
                except Exception as inner_e:
                    print("<div class='alert alert-warning'>Warning: Inner email campaign count error: {0}</div>".format(str(inner_e)))
            except Exception as e:
                # Log but continue - we'll use default values
                print("<div class='alert alert-warning'>Warning: Unable to retrieve email campaign counts. {0}</div>".format(str(e)))
            
            # Get SMS campaign count - count of distinct campaigns in SMSList, not total messages
            try:
                if DatabaseHelper.table_exists('SMSList'):
                    # Simple count of campaigns
                    sms_count_sql = """
                    SELECT COUNT(*) FROM SMSList
                    WHERE Created BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    """.format(sDate, eDate)
                    
                    sms_result = q.QuerySqlScalar(sms_count_sql)
                    if sms_result is not None:
                        try:
                            sms_campaign_count = int(sms_result)
                        except:
                            # Try string conversion
                            try:
                                sms_campaign_count = int(float(str(sms_result)))
                            except:
                                sms_campaign_count = 0
            except Exception as e:
                # Log but continue - we'll use default values
                print("<div class='alert alert-warning'>Warning: Unable to retrieve SMS campaign counts. {0}</div>".format(str(e)))
            
            # Double-check all values for sanity
            if not isinstance(email_campaign_count, (int, long)):
                email_campaign_count = 0
                
            if not isinstance(sms_campaign_count, (int, long)):
                sms_campaign_count = 0
            
            # Combine results
            return {
                'email_summary': email_summary,
                'sms_stats': sms_stats,
                'email_campaign_count': email_campaign_count,
                'sms_campaign_count': sms_campaign_count,
                'date_range': {
                    'start': sDate,
                    'end': eDate
                }
            }
        except Exception as e:
            # Log but continue with defaults
            print("<div class='alert alert-danger'>Dashboard summary error: {0}</div>".format(str(e)))
            return {
                'email_summary': {
                    'total_emails': 0,
                    'sent_emails': 0,
                    'failed_emails': 0,
                    'delivery_rate': 0,
                    'failed_types': []
                },
                'sms_stats': {
                    'total_count': 0,
                    'sent_count': 0,
                    'delivery_rate': 0
                },
                'email_campaign_count': 0,
                'sms_campaign_count': 0,
                'date_range': {
                    'start': sDate,
                    'end': eDate
                }
            }
            
    @staticmethod
    def get_email_summary(sDate, eDate, hide_success, program_filter, failure_filter):
        """Get email delivery summary stats with ultra-robust null reference protection"""
        try:
            # Start with default values
            total_emails = 0
            sent_emails = 0
            failed_emails = 0
            failed_types = []
            opens = 0
            clicks = 0
            unsubscribes = 0
            bounces = 0
            
            # Format filter conditions - but only if needed
            sql_hide_success = '' if hide_success != 'yes' else ' AND fe.Fail IS NOT NULL '
            filter_program = '' if program_filter == str(999999) else ' AND pro.Id = {0}'.format(program_filter)
            filter_fail = '' if failure_filter == str(999999) else " AND fe.Fail = '{0}'".format(failure_filter)
            
            # Get total email count with absolute minimal SQL (no joins)
            try:
                sql_total = """
                SELECT COUNT(*) AS TotalCount 
                FROM EmailQueueTo 
                WHERE Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                """.format(sDate, eDate)
                
                # Direct SQL execution with null checks
                total_result = q.QuerySqlScalar(sql_total)
                
                if total_result is not None and str(total_result).isdigit():
                    total_emails = int(total_result)
                else:
                    # Try a different approach with QuerySql
                    total_results = q.QuerySql(sql_total)
                    if total_results and len(total_results) > 0:
                        try:
                            total_val = getattr(total_results[0], 'TotalCount', 0)
                            if total_val is not None:
                                total_emails = int(total_val)
                        except:
                            # Keep default
                            pass
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get total email count. {0}</div>".format(str(e)))
            
            # Get successful email count with absolute minimal SQL - use EXISTS to avoid nulls
            try:
                sql_success = """
                SELECT COUNT(*) AS SuccessCount
                FROM EmailQueueTo eqt
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND NOT EXISTS (
                    SELECT 1 FROM FailedEmails fe 
                    WHERE fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
                )
                """.format(sDate, eDate)
                
                # Direct SQL execution with null checks
                success_result = q.QuerySqlScalar(sql_success)
                
                if success_result is not None and str(success_result).isdigit():
                    sent_emails = int(success_result)
                else:
                    # Try a different approach with QuerySql
                    success_results = q.QuerySql(sql_success)
                    if success_results and len(success_results) > 0:
                        try:
                            success_val = getattr(success_results[0], 'SuccessCount', 0)
                            if success_val is not None:
                                sent_emails = int(success_val)
                        except:
                            # Keep default
                            pass
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get success email count. {0}</div>".format(str(e)))
            
            # Calculate failed emails if we have both total and success counts
            if total_emails >= sent_emails:
                failed_emails = total_emails - sent_emails
            
            # Get email opens count - count unique people who opened emails
            try:
                sql_opens = """
                SELECT COUNT(DISTINCT er.PeopleId) AS OpenCount
                FROM EmailResponses er
                INNER JOIN EmailQueueTo eqt ON er.EmailQueueId = eqt.Id AND er.PeopleId = eqt.PeopleId
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND er.Type = 'o'
                """.format(sDate, eDate)
                
                opens_result = q.QuerySqlScalar(sql_opens)
                if opens_result is not None:
                    opens = int(opens_result)
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get email opens. {0}</div>".format(str(e)))
            
            # Get email clicks count - use EmailLinks table since EmailResponses doesn't track clicks
            try:
                # Since EmailResponses doesn't have Type='c', use EmailLinks table
                # This gives total clicks across all campaigns in the date range
                sql_clicks = """
                SELECT ISNULL(SUM(el.Count), 0) AS ClickCount
                FROM EmailLinks el
                INNER JOIN EmailQueue eq ON el.EmailId = eq.Id
                WHERE eq.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                """.format(sDate, eDate)
                
                clicks_result = q.QuerySqlScalar(sql_clicks)
                if clicks_result is not None:
                    clicks = int(clicks_result)
                    
                # Note: This is total clicks, not unique people clicking
                # These are aggregate click counts per link
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get email clicks. {0}</div>".format(str(e)))
            
            # Get unsubscribe count using EmailOptOut table
            try:
                if DatabaseHelper.table_exists('EmailOptOut'):
                    # First, detect which date column exists
                    column_check_sql = """
                    SELECT COLUMN_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = 'EmailOptOut'
                    AND COLUMN_NAME IN ('Date', 'OptOutDate', 'CreatedDate', 'DateX')
                    """
                    columns = q.QuerySql(column_check_sql)
                    
                    date_column = None
                    for col in columns:
                        if hasattr(col, 'COLUMN_NAME'):
                            # Use the first available date column
                            if col.COLUMN_NAME in ['Date', 'OptOutDate', 'CreatedDate', 'DateX']:
                                date_column = col.COLUMN_NAME
                                break
                    
                    if date_column:
                        sql_unsubscribes = """
                        SELECT COUNT(DISTINCT eo.ToPeopleId) AS UnsubscribeCount
                        FROM EmailOptOut eo
                        INNER JOIN People p ON eo.ToPeopleId = p.PeopleId
                        WHERE eo.{2} BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                        AND p.IsDeceased = 0
                        AND p.ArchivedFlag = 0
                        """.format(sDate, eDate, date_column)
                        
                        unsubscribes_result = q.QuerySqlScalar(sql_unsubscribes)
                        if unsubscribes_result is not None:
                            unsubscribes = int(unsubscribes_result)
                    else:
                        # No date column found, raise exception to use fallback
                        raise Exception("No date column found in EmailOptOut table")
                else:
                    # Table doesn't exist, raise exception to use fallback
                    raise Exception("EmailOptOut table does not exist")
            except Exception as e:
                # If EmailOptOut table doesn't exist or has issues, try alternative approaches
                try:
                    # Check for DoNotMailFlag as fallback
                    sql_unsubscribes_alt = """
                    SELECT COUNT(DISTINCT p.PeopleId) AS UnsubscribeCount
                    FROM People p
                    WHERE (p.DoNotMailFlag = 1 OR p.DoNotCallFlag = 1)
                    AND p.ModifiedDate BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    """.format(sDate, eDate)
                    
                    unsubscribes_result = q.QuerySqlScalar(sql_unsubscribes_alt)
                    if unsubscribes_result is not None:
                        unsubscribes = int(unsubscribes_result)
                except Exception as e2:
                    print("<div class='alert alert-warning'>Warning: Failed to get unsubscribes. {0}</div>".format(str(e)))
            
            # Get bounce count from failed emails
            # Based on user's failure types, the following should be considered bounces:
            # - bouncedaddress: Direct bounce indicator
            # - Mailbox Unavailable: Mailbox doesn't exist or is full
            # - Invalid Address: Email address is invalid
            # - invalid: Another form of invalid address
            try:
                # Count total bounce events (not unique people) to match failure breakdown
                # Using the same join pattern as the failure breakdown query
                sql_bounces = """
                SELECT COUNT(*) AS BounceCount
                FROM (
                    SELECT DISTINCT fe.Fail, eqt.Id, eqt.PeopleId
                    FROM EmailQueueTo eqt
                    JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
                    WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    AND fe.Fail IS NOT NULL
                    AND fe.Fail IN ('bounce', 'hardbounce', 'blocked', 'invalid', 
                                   'bouncedaddress', 'Mailbox Unavailable', 'Invalid Address')
                ) AS BounceList
                """.format(sDate, eDate)
                
                bounces_result = q.QuerySqlScalar(sql_bounces)
                if bounces_result is not None:
                    bounces = int(bounces_result)
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get bounce count. {0}</div>".format(str(e)))
            
            # Get failed email breakdown with minimal SQL - use temp table approach to reduce null issues
            try:
                sql_failures = """
                SELECT Fail AS Status, COUNT(*) AS FailCount
                FROM (
                    SELECT DISTINCT fe.Fail, eqt.Id, eqt.PeopleId
                    FROM EmailQueueTo eqt
                    JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
                    WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    AND fe.Fail IS NOT NULL
                ) AS FailedEmailList
                GROUP BY Fail
                ORDER BY FailCount DESC
                """.format(sDate, eDate)
                
                failure_results = q.QuerySql(sql_failures)
                
                if failure_results and len(failure_results) > 0:
                    failure_total = 0
                    for row in failure_results:
                        # Handle each row with null protection
                        try:
                            status = getattr(row, 'Status', 'Unknown')
                            if status is None:
                                status = 'Unknown'
                                
                            count = getattr(row, 'FailCount', 0)
                            if count is None or not str(count).isdigit():
                                count = 0
                            else:
                                count = int(count)
                                
                            failed_types.append({
                                'Status': status,
                                'Count': count
                            })
                            
                            failure_total += count
                        except:
                            # Skip this row
                            continue
                            
                    # If we have a valid failure count from breakdown, use it
                    if failure_total > 0:
                        failed_emails = failure_total
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get failure breakdown. {0}</div>".format(str(e)))
            
            # Double-check calculations to ensure consistency
            if total_emails < (sent_emails + failed_emails):
                # If our counts don't add up, recalculate total
                total_emails = sent_emails + failed_emails
            
            # Calculate delivery rate with extra safety
            try:
                if total_emails > 0:
                    delivery_rate = (float(sent_emails) / float(total_emails)) * 100
                else:
                    delivery_rate = 0
            except:
                delivery_rate = 0
            
            # Calculate open rate
            open_rate = 0
            if sent_emails > 0:
                open_rate = (float(opens) / float(sent_emails)) * 100
            
            # Calculate click-through rate
            # Note: clicks are total clicks (not unique), opens are unique people
            # So this rate might exceed 100% if people clicked multiple times
            click_rate = 0
            if sent_emails > 0:
                click_rate = (float(clicks) / float(sent_emails)) * 100
            
            # Calculate bounce rate
            bounce_rate = 0
            if total_emails > 0:
                bounce_rate = (float(bounces) / float(total_emails)) * 100
            
            # Calculate unsubscribe rate
            unsubscribe_rate = 0
            if sent_emails > 0:
                unsubscribe_rate = (float(unsubscribes) / float(sent_emails)) * 100
            
            # Return final stats with guaranteed values
            return {
                'total_emails': total_emails,
                'sent_emails': sent_emails,
                'failed_emails': failed_emails,
                'delivery_rate': delivery_rate,
                'failed_types': failed_types,
                'opens': opens,
                'clicks': clicks,
                'open_rate': open_rate,
                'click_rate': click_rate,
                'bounces': bounces,
                'bounce_rate': bounce_rate,
                'unsubscribes': unsubscribes,
                'unsubscribe_rate': unsubscribe_rate
            }
        except Exception as e:
            # Log error but return a valid structure with defaults
            print("<div class='alert alert-danger'>Error in email summary: {0}</div>".format(str(e)))
            return {
                'total_emails': 0,
                'sent_emails': 0,
                'failed_emails': 0,
                'delivery_rate': 0,
                'failed_types': [],
                'opens': 0,
                'clicks': 0,
                'open_rate': 0,
                'click_rate': 0,
                'bounces': 0,
                'bounce_rate': 0,
                'unsubscribes': 0,
                'unsubscribe_rate': 0
            }
            
    @staticmethod
    def get_subscriber_growth(sDate, eDate, interval='daily'):
        """Get subscriber growth over time based on email activity"""
        try:
            # First check if we can use EmailOptOut table effectively
            can_use_optout = False
            optout_date_column = None
            
            if DatabaseHelper.table_exists('EmailOptOut'):
                # Try to detect which date column exists
                try:
                    # First, get column names from the table
                    column_check_sql = """
                    SELECT COLUMN_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = 'EmailOptOut'
                    AND COLUMN_NAME IN ('Date', 'OptOutDate', 'CreatedDate', 'DateX')
                    """
                    columns = q.QuerySql(column_check_sql)
                    
                    # Check which columns exist
                    available_columns = []
                    for col in columns:
                        if hasattr(col, 'COLUMN_NAME'):
                            available_columns.append(col.COLUMN_NAME)
                    
                    # Use the first available date column
                    if 'Date' in available_columns:
                        optout_date_column = 'Date'
                        can_use_optout = True
                    elif 'OptOutDate' in available_columns:
                        optout_date_column = 'OptOutDate'
                        can_use_optout = True
                    elif 'CreatedDate' in available_columns:
                        optout_date_column = 'CreatedDate'
                        can_use_optout = True
                    elif 'DateX' in available_columns:
                        optout_date_column = 'DateX'
                        can_use_optout = True
                    else:
                        can_use_optout = False
                except:
                    # If the column check fails, we'll use the fallback
                    can_use_optout = False
            # Determine date grouping based on interval
            if interval == 'daily':
                date_format = 'YYYY-MM-DD'
                date_group = "CONVERT(VARCHAR(10), DatePoint, 101)"  # MM/DD/YYYY format
            elif interval == 'weekly':
                date_format = 'YYYY-WW'
                date_group = "DATEPART(YEAR, DatePoint) * 100 + DATEPART(WEEK, DatePoint)"
            else:  # monthly
                date_format = 'YYYY-MM'
                date_group = "CONVERT(VARCHAR(7), DatePoint, 126)"
            
            # Build date group expressions
            date_group_sent = date_group.replace('DatePoint', 'eqt.Sent')
            date_group_mod = date_group.replace('DatePoint', 'p.ModifiedDate')
            
            # Only set optout date group if we have a valid column
            if can_use_optout and optout_date_column:
                date_group_optout = date_group.replace('DatePoint', 'eo.' + optout_date_column)
            else:
                date_group_optout = None
                
            # Track email activity-based growth
            # Focus on actual email recipients and opt-outs during the period
            
            if can_use_optout and optout_date_column:
                # Use the detected date column
                date_column = "eo." + optout_date_column
                
                sql_growth = """
            WITH DateRange AS (
                SELECT CAST('{0}' AS DATE) AS StartDate, CAST('{1}' AS DATE) AS EndDate
            ),
            -- Get unique email recipients per period (active subscribers)
            EmailActivity AS (
                SELECT 
                    {2} AS DateGroup,
                    COUNT(DISTINCT eqt.PeopleId) AS ActiveRecipients
                FROM EmailQueueTo eqt
                INNER JOIN People p ON eqt.PeopleId = p.PeopleId
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND p.IsDeceased = 0
                AND p.EmailAddress IS NOT NULL
                AND NOT EXISTS (
                    SELECT 1 FROM EmailOptOut eo
                    WHERE eo.ToPeopleId = p.PeopleId
                )
                GROUP BY {2}
            ),
            -- Count first-time email recipients (new to email list)
            FirstTimeRecipients AS (
                SELECT 
                    {2} AS DateGroup,
                    COUNT(DISTINCT eqt.PeopleId) AS NewRecipients
                FROM EmailQueueTo eqt
                INNER JOIN People p ON eqt.PeopleId = p.PeopleId
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND p.IsDeceased = 0
                -- This is their first email
                AND NOT EXISTS (
                    SELECT 1 FROM EmailQueueTo eqt2
                    WHERE eqt2.PeopleId = eqt.PeopleId
                    AND eqt2.Sent < '{0} 00:00:00'
                )
                GROUP BY {2}
            ),
            -- Count people who opted out during the period
            OptOuts AS (
                SELECT 
                    {3} AS DateGroup,
                    COUNT(DISTINCT eo.ToPeopleId) AS OptOutCount
                FROM EmailOptOut eo
                INNER JOIN People p ON eo.ToPeopleId = p.PeopleId
                WHERE {4} BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND p.IsDeceased = 0
                AND p.ArchivedFlag = 0
                -- They received emails before opting out
                AND EXISTS (
                    SELECT 1 FROM EmailQueueTo eqt
                    WHERE eqt.PeopleId = eo.ToPeopleId
                    AND eqt.Sent < {4}
                )
                GROUP BY {3}
            ),
            -- Combine all date groups
            AllDates AS (
                SELECT DateGroup FROM EmailActivity
                UNION SELECT DateGroup FROM FirstTimeRecipients  
                UNION SELECT DateGroup FROM OptOuts
            )
            SELECT 
                ad.DateGroup,
                ISNULL(ftr.NewRecipients, 0) AS NewSubscribers,
                ISNULL(oo.OptOutCount, 0) AS Unsubscribes,
                ISNULL(ftr.NewRecipients, 0) - ISNULL(oo.OptOutCount, 0) AS NetGrowth,
                ISNULL(ea.ActiveRecipients, 0) AS TotalActive
            FROM AllDates ad
            LEFT JOIN EmailActivity ea ON ad.DateGroup = ea.DateGroup
            LEFT JOIN FirstTimeRecipients ftr ON ad.DateGroup = ftr.DateGroup
            LEFT JOIN OptOuts oo ON ad.DateGroup = oo.DateGroup
            ORDER BY ad.DateGroup
            """.format(sDate, eDate, date_group_sent, 
                      date_group_optout,
                      date_column)
            else:
                # Fallback query when EmailOptOut table doesn't exist
                # Use DoNotMailFlag as proxy for opt-outs
                sql_growth = """
            WITH DateRange AS (
                SELECT CAST('{0}' AS DATE) AS StartDate, CAST('{1}' AS DATE) AS EndDate
            ),
            -- Get unique email recipients per period (active subscribers)
            EmailActivity AS (
                SELECT 
                    {2} AS DateGroup,
                    COUNT(DISTINCT eqt.PeopleId) AS ActiveRecipients
                FROM EmailQueueTo eqt
                INNER JOIN People p ON eqt.PeopleId = p.PeopleId
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND p.IsDeceased = 0
                AND p.EmailAddress IS NOT NULL
                AND p.DoNotMailFlag = 0
                GROUP BY {2}
            ),
            -- Count first-time email recipients (new to email list)
            FirstTimeRecipients AS (
                SELECT 
                    {2} AS DateGroup,
                    COUNT(DISTINCT eqt.PeopleId) AS NewRecipients
                FROM EmailQueueTo eqt
                INNER JOIN People p ON eqt.PeopleId = p.PeopleId
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND p.IsDeceased = 0
                -- This is their first email
                AND NOT EXISTS (
                    SELECT 1 FROM EmailQueueTo eqt2
                    WHERE eqt2.PeopleId = eqt.PeopleId
                    AND eqt2.Sent < '{0} 00:00:00'
                )
                GROUP BY {2}
            ),
            -- Count people who were marked DoNotMail during the period
            OptOuts AS (
                SELECT 
                    {3} AS DateGroup,
                    COUNT(DISTINCT p.PeopleId) AS OptOutCount
                FROM People p
                WHERE p.ModifiedDate BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                AND p.DoNotMailFlag = 1
                AND p.IsDeceased = 0
                AND p.ArchivedFlag = 0
                -- They received emails before being marked DoNotMail
                AND EXISTS (
                    SELECT 1 FROM EmailQueueTo eqt
                    WHERE eqt.PeopleId = p.PeopleId
                    AND eqt.Sent < p.ModifiedDate
                )
                GROUP BY {3}
            ),
            -- Combine all date groups
            AllDates AS (
                SELECT DateGroup FROM EmailActivity
                UNION SELECT DateGroup FROM FirstTimeRecipients  
                UNION SELECT DateGroup FROM OptOuts
            )
            SELECT 
                ad.DateGroup,
                ISNULL(ftr.NewRecipients, 0) AS NewSubscribers,
                ISNULL(oo.OptOutCount, 0) AS Unsubscribes,
                ISNULL(ftr.NewRecipients, 0) - ISNULL(oo.OptOutCount, 0) AS NetGrowth,
                ISNULL(ea.ActiveRecipients, 0) AS TotalActive
            FROM AllDates ad
            LEFT JOIN EmailActivity ea ON ad.DateGroup = ea.DateGroup
            LEFT JOIN FirstTimeRecipients ftr ON ad.DateGroup = ftr.DateGroup
            LEFT JOIN OptOuts oo ON ad.DateGroup = oo.DateGroup
            ORDER BY ad.DateGroup
            """.format(sDate, eDate, date_group_sent, date_group_mod)
            
            results = q.QuerySql(sql_growth)
            
            # Convert to list of dictionaries
            data = []
            for row in results:
                data.append({
                    'DateGroup': str(getattr(row, 'DateGroup', '')),
                    'NewSubscribers': getattr(row, 'NewSubscribers', 0),
                    'Unsubscribes': getattr(row, 'Unsubscribes', 0),
                    'NetGrowth': getattr(row, 'NetGrowth', 0)
                })
                
            return data
        except Exception as e:
            print("<div class='alert alert-danger'>Error in subscriber growth: {0}</div>".format(str(e)))
            return []
    
    @staticmethod
    def get_failure_recipients(sDate, eDate, hide_success, program_filter, failure_filter):
        """Get recipients with failed emails"""
        # Format filter conditions
        sql_hide_success = '' if hide_success != 'yes' else ' AND fe.Fail IS NOT NULL '
        filter_program = '' if program_filter == str(999999) else ' AND pro.Id = {}'.format(program_filter)
        filter_fail = '' if failure_filter == str(999999) else " AND fe.Fail = '{}'".format(failure_filter)
        
        # SQL for user failure stats
        sql = """
        SELECT 
            p.PeopleId,
            p.Name,
            p.EmailAddress,
            fe.Fail AS FailureType,  
            COUNT(*) AS FailureCount
        FROM EmailQueueTo eqt
        JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
        LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
        LEFT JOIN Division d ON d.Id = o.DivisionId
        LEFT JOIN Program pro ON pro.Id = d.ProgId
        LEFT JOIN People p ON p.PeopleId = eqt.PeopleId
        WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2} {3} {4}
        GROUP BY p.PeopleId, p.Name, p.EmailAddress, fe.Fail
        ORDER BY FailureCount DESC
        """
        
        try:
            # Format SQL with parameters
            formatted_sql = sql.format(sDate, eDate, sql_hide_success, filter_program, filter_fail)
            results = q.QuerySql(formatted_sql)
            
            # Convert results to list of dictionaries
            data = []
            for row in results:
                data.append({
                    'PeopleId': getattr(row, 'PeopleId', 0),
                    'Name': getattr(row, 'Name', ''),
                    'EmailAddress': getattr(row, 'EmailAddress', ''),
                    'FailureType': getattr(row, 'FailureType', ''),
                    'FailureCount': getattr(row, 'FailureCount', 0)
                })
            
            return data
        except Exception as e:
            raise Exception("Error retrieving failure recipients: {}".format(str(e)))

    @staticmethod
    def get_recent_campaigns(sDate, eDate, program_filter, page=1, page_size=None, exclude_single_recipient=False):
        """Get recent email campaigns with engagement metrics - simplified version"""
        # Always use max rows setting, ignore pagination
        page_size = Config.MAX_ROWS_PER_TABLE
        offset = 0  # No pagination - always start from beginning
        
        # Format filter conditions
        filter_program = '' if program_filter == str(999999) else ' AND pro.Id = {}'.format(program_filter)
        # For HAVING clause, we need the condition without the leading AND
        filter_single_recipient = ' AND COUNT(DISTINCT eqt.PeopleId) > 1' if exclude_single_recipient else ''
        
        # Choose query based on performance mode
        if Config.PERFORMANCE_MODE:
            # Super simplified query for maximum performance
            sql = """
            -- Performance mode: Basic campaign info only, no metrics
            WITH CampaignList AS (
                SELECT 
                    eq.Id AS EmailQueueId,
                    eq.Subject AS CampaignName,
                    eq.FromName AS Sender,
                    eq.QueuedBy AS SenderPeopleId,
                    COUNT(DISTINCT eqt.PeopleId) AS RecipientCount,
                    CONVERT(VARCHAR(10), MAX(eqt.Sent), 101) AS SentDate,
                    MAX(eqt.Sent) AS SentDateTime
                FROM EmailQueue eq WITH (NOLOCK)
                INNER JOIN EmailQueueTo eqt WITH (NOLOCK) ON eq.Id = eqt.Id
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                GROUP BY eq.Id, eq.Subject, eq.FromName, eq.QueuedBy
                HAVING 1=1 {2}
                ORDER BY MAX(eqt.Sent) DESC
                OFFSET 0 ROWS
                FETCH NEXT {3} ROWS ONLY
            )
            SELECT 
                EmailQueueId,
                CampaignName,
                Sender,
                SenderPeopleId,
                SentDate,
                RecipientCount,
                0 AS OpenCount,
                0 AS ClickCount,
                0 AS FailureCount,
                0 AS UnsubscribeCount,
                0.0 AS OpenRate,
                0.0 AS ClickRate,
                0.0 AS FailureRate
            FROM CampaignList
            ORDER BY SentDateTime DESC
            """.format(sDate, eDate, filter_single_recipient, page_size)
        else:
            # Ultra-optimized SQL - Get basic campaign info first, metrics are loaded separately
            sql = """
        -- Get campaign list with basic info only
        WITH CampaignList AS (
            SELECT 
                eq.Id AS EmailQueueId,
                eq.Subject AS CampaignName,
                eq.FromName AS Sender,
                eq.QueuedBy AS SenderPeopleId,
                COUNT(DISTINCT eqt.PeopleId) AS RecipientCount,
                CONVERT(VARCHAR(10), MAX(eqt.Sent), 101) AS SentDate,
                MAX(eqt.Sent) AS SentDateTime
            FROM EmailQueue eq WITH (NOLOCK)
            INNER JOIN EmailQueueTo eqt WITH (NOLOCK) ON eq.Id = eqt.Id
            LEFT JOIN Organizations o WITH (NOLOCK) ON o.OrganizationId = eqt.OrgId
            LEFT JOIN Division d WITH (NOLOCK) ON d.Id = o.DivisionId
            LEFT JOIN Program pro WITH (NOLOCK) ON pro.Id = d.ProgId
            WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2}
            GROUP BY eq.Id, eq.Subject, eq.FromName, eq.QueuedBy
            HAVING 1=1 {3}
            ORDER BY MAX(eqt.Sent) DESC
            OFFSET 0 ROWS
            FETCH NEXT {4} ROWS ONLY
        )
        SELECT 
            cl.*,
            -- Simplified metrics - just counts, calculate rates in Python
            (SELECT COUNT(DISTINCT PeopleId) FROM EmailResponses WITH (NOLOCK) WHERE EmailQueueId = cl.EmailQueueId AND Type = 'o') AS OpenCount,
            (SELECT ISNULL(SUM(Count), 0) FROM EmailLinks WITH (NOLOCK) WHERE EmailId = cl.EmailQueueId) AS ClickCount,
            (SELECT COUNT(DISTINCT fe.PeopleId) FROM FailedEmails fe WITH (NOLOCK) 
             INNER JOIN EmailQueueTo eqt2 WITH (NOLOCK) ON fe.Id = eqt2.Id AND fe.PeopleId = eqt2.PeopleId 
             WHERE eqt2.Id = cl.EmailQueueId) AS FailureCount,
            0 AS UnsubscribeCount, -- Skip unsubscribes for performance, they're rarely used
            0.0 AS OpenRate,
            0.0 AS ClickRate,
            0.0 AS FailureRate
        FROM CampaignList cl
        ORDER BY cl.SentDateTime DESC
        """
        
        try:
            # Debug: Log parameter values BEFORE formatting
            if Config.DEBUG_MODE and False:  # Temporarily disabled
                print "<div class='alert alert-warning'>Debug: Parameters BEFORE SQL formatting:</div>"
                print "<ul>"
                print "<li>exclude_single_recipient parameter: {0}</li>".format(exclude_single_recipient)
                print "<li>filter_single_recipient value: '{0}'</li>".format(filter_single_recipient)
                print "<li>Length of filter_single_recipient: {0}</li>".format(len(filter_single_recipient))
                print "</ul>"
                
                # Count placeholders in SQL template
                import re
                placeholder_count = len(re.findall(r'\{(\d+)\}', sql))
                print "<div class='alert alert-info'>Debug: SQL template has {0} placeholders</div>".format(placeholder_count)
                print "<div class='alert alert-info'>Debug: Passing {0} parameters to format()</div>".format(6)
            
            # Format SQL with parameters including pagination
            if Config.PERFORMANCE_MODE:
                # Performance mode query is already formatted
                formatted_sql = sql
            else:
                # Note: The {3} placeholder in the HAVING clause needs the filter_single_recipient value
                formatted_sql = sql.format(sDate, eDate, filter_program, filter_single_recipient, page_size)
            
            # Debug output - remove after testing
            if Config.DEBUG_MODE and False:  # Temporarily disabled
                print "<div class='alert alert-info'>Debug: Page={0}, PageSize={1}, Offset={2}</div>".format(page, page_size, offset)
                print "<div class='alert alert-info'>Debug: Date range: {0} to {1}</div>".format(sDate, eDate)
                print "<div class='alert alert-info'>Debug: Program filter: {0}</div>".format(program_filter if program_filter else "All Programs")
                print "<div class='alert alert-info'>Debug: Filter program SQL: {0}</div>".format(filter_program if filter_program else "None")
                print "<div class='alert alert-info'>Debug: Exclude single recipient: {0}</div>".format(exclude_single_recipient)
                print "<div class='alert alert-info'>Debug: Single recipient filter SQL: {0}</div>".format(filter_single_recipient if filter_single_recipient else "None")
                
                # Show the actual SQL query being executed
                print "<div class='alert alert-warning'>Debug: SQL Query being executed:</div>"
                print "<pre style='font-size: 11px; background: #f8f9fa; padding: 10px; border-radius: 4px;'>{0}</pre>".format(formatted_sql.replace('<', '&lt;').replace('>', '&gt;'))
                
                # Check if HAVING clause was properly formatted
                if 'HAVING' in formatted_sql:
                    having_start = formatted_sql.find('HAVING')
                    having_end = formatted_sql.find('\n', having_start)
                    having_clause = formatted_sql[having_start:having_end] if having_end > -1 else formatted_sql[having_start:]
                    print "<div class='alert alert-warning'>Debug: Extracted HAVING clause: <code>{0}</code></div>".format(having_clause.strip())
                    
                    # Explicitly check if the filter is present
                    if exclude_single_recipient and 'COUNT(DISTINCT eqt.PeopleId) > 1' not in formatted_sql:
                        print "<div class='alert alert-danger'>ERROR: Single-recipient filter was NOT added to SQL despite exclude_single_recipient=True!</div>"
                    elif exclude_single_recipient:
                        print "<div class='alert alert-success'>SUCCESS: Single-recipient filter IS present in SQL</div>"
                
                # Additional debug to verify parameter substitution
                print "<div class='alert alert-info'>Debug: Format parameters:</div>"
                print "<ul>"
                print "<li>sDate: {0}</li>".format(sDate)
                print "<li>eDate: {0}</li>".format(eDate)
                print "<li>filter_program: {0}</li>".format(filter_program)
                print "<li>filter_single_recipient: {0}</li>".format(filter_single_recipient)
                print "<li>offset: {0}</li>".format(offset)
                print "<li>page_size: {0}</li>".format(page_size)
                print "</ul>"
                
                # Check total count without pagination for debugging
                count_sql = sql.replace('OFFSET {4} ROWS\n        FETCH NEXT {5} ROWS ONLY', '').format(sDate, eDate, filter_program, filter_single_recipient)
                total_results = q.QuerySql(count_sql)
                total_count = len(list(total_results)) if total_results else 0
                print "<div class='alert alert-info'>Debug: Total campaigns available: {0}</div>".format(total_count)
                
                # Check if there are any EmailQueueTo records in the date range at all
                basic_count_sql = """
                SELECT COUNT(*) as EmailCount
                FROM EmailQueueTo eqt
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                """.format(sDate, eDate)
                
                basic_result = q.QuerySqlTop1(basic_count_sql)
                basic_count = getattr(basic_result, 'EmailCount', 0) if basic_result else 0
                print "<div class='alert alert-info'>Debug: Total emails in date range: {0}</div>".format(basic_count)
                
                # Check distinct campaigns in date range
                distinct_campaigns_sql = """
                SELECT COUNT(DISTINCT eq.Id) as CampaignCount
                FROM EmailQueue eq
                JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
                WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                """.format(sDate, eDate)
                
                distinct_result = q.QuerySqlTop1(distinct_campaigns_sql)
                distinct_count = getattr(distinct_result, 'CampaignCount', 0) if distinct_result else 0
                print "<div class='alert alert-info'>Debug: Distinct campaigns in date range: {0}</div>".format(distinct_count)
                
                # Test if the program filter is affecting results
                if filter_program:
                    no_filter_sql = """
                    SELECT COUNT(DISTINCT eq.Id) as CampaignCount
                    FROM EmailQueue eq
                    JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
                    WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    """.format(sDate, eDate)
                    
                    no_filter_result = q.QuerySqlTop1(no_filter_sql)
                    no_filter_count = getattr(no_filter_result, 'CampaignCount', 0) if no_filter_result else 0
                    print "<div class='alert alert-info'>Debug: Campaigns without program filter: {0}</div>".format(no_filter_count)
                    
                    # Also check how many campaigns have program associations
                    with_program_sql = """
                    SELECT COUNT(DISTINCT eq.Id) as CampaignCount
                    FROM EmailQueue eq
                    JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
                    LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
                    LEFT JOIN Division d ON d.Id = o.DivisionId
                    LEFT JOIN Program pro ON pro.Id = d.ProgId
                    WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    AND pro.Id IS NOT NULL
                    """.format(sDate, eDate)
                    
                    with_program_result = q.QuerySqlTop1(with_program_sql)
                    with_program_count = getattr(with_program_result, 'CampaignCount', 0) if with_program_result else 0
                    print "<div class='alert alert-info'>Debug: Campaigns with program associations: {0}</div>".format(with_program_count)
            
            # Get all campaigns up to the limit
            results = q.QuerySql(formatted_sql)
            
            # Convert results to list of dictionaries
            data = []
            
            for row in results:
                
                email_queue_id = getattr(row, 'EmailQueueId', 0)
                sender_people_id = getattr(row, 'SenderPeopleId', 0)
                recipient_count = getattr(row, 'RecipientCount', 0)
                open_count = getattr(row, 'OpenCount', 0)
                click_count = getattr(row, 'ClickCount', 0)
                failure_count = getattr(row, 'FailureCount', 0)
                unsubscribe_count = getattr(row, 'UnsubscribeCount', 0)
                open_rate = getattr(row, 'OpenRate', 0)
                click_rate = getattr(row, 'ClickRate', 0)
                failure_rate = getattr(row, 'FailureRate', 0)
                
                # Calculate rates in Python for better performance
                if recipient_count > 0:
                    calc_open_rate = (float(open_count) / float(recipient_count)) * 100
                    calc_click_rate = (float(click_count) / float(recipient_count)) * 100
                    calc_failure_rate = (float(failure_count) / float(recipient_count)) * 100
                else:
                    calc_open_rate = 0
                    calc_click_rate = 0
                    calc_failure_rate = 0
                
                data.append({
                    'EmailQueueId': email_queue_id,
                    'CampaignName': getattr(row, 'CampaignName', ''),
                    'Sender': getattr(row, 'Sender', ''),
                    'SenderPeopleId': sender_people_id,
                    'SentDate': getattr(row, 'SentDate', ''),
                    'RecipientCount': recipient_count,
                    'RecipientCountFormatted': FormatHelper.format_number(recipient_count),
                    'OpenCount': FormatHelper.format_number(open_count),
                    'OpenRate': FormatHelper.format_percent(calc_open_rate),
                    'ClickCount': FormatHelper.format_number(click_count),
                    'ClickRate': FormatHelper.format_percent(calc_click_rate),
                    'FailureCount': FormatHelper.format_number(failure_count),
                    'FailureRate': FormatHelper.format_percent(calc_failure_rate),
                    'UnsubscribeCount': FormatHelper.format_number(unsubscribe_count),
                    'IsSingleRecipient': recipient_count == 1
                })
            
            # No pagination - no need for has_more flag
            
            # Debug output
            if Config.PAGINATION_DEBUG:
                print '<!-- Debug: get_recent_campaigns() Results -->'
                print '<!-- Page: {0}, Page Size: {1}, Offset: {2} -->'.format(page, page_size, offset)
                print '<!-- Exclude Single Recipient: {0} -->'.format(exclude_single_recipient)
                print '<!-- Rows found: {0} -->'.format(len(data))
                print '<div class="alert alert-info" style="margin: 10px;">'
                print '<strong>Debug:</strong> '
                print 'Found {0} campaigns'.format(len(data))
                print '</div>'
            
            return data
        except Exception as e:
            raise Exception("Error retrieving campaign data: {}".format(str(e)))

    @staticmethod
    def get_active_senders(sDate, eDate, program_filter):
        """Get most active email senders"""
        # Format filter conditions
        filter_program = '' if program_filter == str(999999) else ' AND pro.Id = {}'.format(program_filter)
        
        # SQL for top senders with unsubscribe counts
        sql = """
        WITH SenderStats AS (
            SELECT 
                eq.FromName AS SenderName,
                eq.Id AS EmailQueueId,
                COUNT(eqt.PeopleId) AS Recipients
            FROM EmailQueue eq
            JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
            LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2}
            GROUP BY eq.FromName, eq.Id
        ),
        UnsubscribeStats AS (
            SELECT 
                eq.FromName AS SenderName,
                COUNT(DISTINCT p.PeopleId) AS UnsubscribeCount
            FROM EmailQueue eq
            JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
            JOIN People p ON p.PeopleId = eqt.PeopleId
            WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
            AND p.DoNotMailFlag = 1
            AND p.ModifiedDate >= eqt.Sent
            AND p.ModifiedDate <= DATEADD(day, 7, eqt.Sent)
            GROUP BY eq.FromName
        ),
        SenderTotals AS (
            SELECT 
                SenderName,
                COUNT(DISTINCT EmailQueueId) AS CampaignCount,
                SUM(Recipients) AS TotalRecipients
            FROM SenderStats
            GROUP BY SenderName
        ),
        FailureCounts AS (
            SELECT 
                eq.FromName AS SenderName,
                COUNT(DISTINCT fe.Id) AS FailureCount
            FROM EmailQueue eq
            JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
            JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
            WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
            GROUP BY eq.FromName
        )
        SELECT TOP {3}
            st.SenderName,
            st.CampaignCount,
            st.TotalRecipients,
            st.TotalRecipients - ISNULL(fc.FailureCount, 0) AS SuccessfulDeliveries,
            ISNULL(us.UnsubscribeCount, 0) AS Unsubscribes
        FROM SenderTotals st
        LEFT JOIN FailureCounts fc ON fc.SenderName = st.SenderName
        LEFT JOIN UnsubscribeStats us ON us.SenderName = st.SenderName
        ORDER BY st.CampaignCount DESC
        """
        
        try:
            # Format SQL with parameters
            formatted_sql = sql.format(sDate, eDate, filter_program, Config.MAX_ROWS_PER_TABLE)
            results = q.QuerySql(formatted_sql)
            
            # Convert results to list of dictionaries
            data = []
            for row in results:
                total_recipients = getattr(row, 'TotalRecipients', 0)
                successful_deliveries = getattr(row, 'SuccessfulDeliveries', 0)
                unsubscribes = getattr(row, 'Unsubscribes', 0)
                delivery_rate = FormatHelper.safe_div(successful_deliveries, total_recipients) * 100
                
                data.append({
                    'SenderName': getattr(row, 'SenderName', ''),
                    'CampaignCount': getattr(row, 'CampaignCount', 0),
                    'TotalRecipients': FormatHelper.format_number(total_recipients),
                    'SuccessfulDeliveries': FormatHelper.format_number(successful_deliveries),
                    'DeliveryRate': FormatHelper.format_percent(delivery_rate),
                    'Unsubscribes': FormatHelper.format_number(unsubscribes)
                })
            
            return data
        except Exception as e:
            raise Exception("Error retrieving active senders: {}".format(str(e)))

    @staticmethod
    def get_program_email_stats(sDate, eDate):
        """Get email stats broken down by program"""
        # SQL for program breakdown
        sql = """
        SELECT 
            pro.Name AS ProgramName,
            COUNT(DISTINCT eq.Id) AS CampaignCount,
            COUNT(eqt.PeopleId) AS TotalRecipients,
            SUM(CASE WHEN fe.Id IS NULL THEN 1 ELSE 0 END) AS SuccessfulDeliveries,
            COUNT(DISTINCT fe.Id) AS FailureCount
        FROM EmailQueueTo eqt
        JOIN EmailQueue eq ON eq.Id = eqt.Id
        LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
        LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
        LEFT JOIN Division d ON d.Id = o.DivisionId
        LEFT JOIN Program pro ON pro.Id = d.ProgId
        WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
        GROUP BY pro.Name
        HAVING pro.Name IS NOT NULL
        ORDER BY COUNT(eqt.PeopleId) DESC
        """
        
        try:
            # Format SQL with parameters
            formatted_sql = sql.format(sDate, eDate)
            results = q.QuerySql(formatted_sql)
            
            # Convert results to list of dictionaries
            data = []
            for row in results:
                program_name = getattr(row, 'ProgramName', '')
                if not program_name:  # Skip rows with no program name
                    continue
                    
                total_recipients = getattr(row, 'TotalRecipients', 0)
                successful_deliveries = getattr(row, 'SuccessfulDeliveries', 0)
                delivery_rate = FormatHelper.safe_div(successful_deliveries, total_recipients) * 100
                
                data.append({
                    'ProgramName': program_name,
                    'CampaignCount': getattr(row, 'CampaignCount', 0),
                    'TotalRecipients': FormatHelper.format_number(total_recipients),
                    'SuccessfulDeliveries': FormatHelper.format_number(successful_deliveries),
                    'FailureCount': getattr(row, 'FailureCount', 0),
                    'DeliveryRate': FormatHelper.format_percent(delivery_rate)
                })
            
            return data
        except Exception as e:
            raise Exception("Error retrieving program stats: {}".format(str(e)))

    @staticmethod
    def get_sms_stats(sDate, eDate):
        """Get SMS stats with simpler, more reliable counting"""
        # Default return values
        default_stats = {
            'total_count': 0,
            'sent_count': 0,
            'delivery_rate': 0
        }
        
        try:
            # First check if the SMS table exists
            if not DatabaseHelper.table_exists('SMSList'):
                return default_stats
            
            # Total SMS Count - simple SUM of SentSMS column
            total_count = 0
            try:
                total_sql = """
                SELECT SUM(ISNULL(SentSMS, 0)) AS TotalCount
                FROM SMSList
                WHERE Created BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                """.format(sDate, eDate)
                
                scalar_result = q.QuerySqlScalar(total_sql)
                if scalar_result is not None:
                    try:
                        total_count = int(scalar_result)
                    except:
                        # Not a valid integer, try string conversion
                        try:
                            total_count = int(float(str(scalar_result)))
                        except:
                            # Keep default
                            pass
            except Exception as e:
                print("<div class='alert alert-warning'>Warning: Failed to get SMS total count. {0}</div>".format(str(e)))
            
            # Now check if SMSItems exists for delivered count
            items_table_exists = DatabaseHelper.table_exists('SMSItems')
            
            # Get delivered count if SMSItems exists
            delivered_count = 0
            if items_table_exists and total_count > 0:
                try:
                    delivered_sql = """
                    SELECT COUNT(*) AS DeliveredCount
                    FROM SMSItems
                    WHERE ListID IN (
                        SELECT ID FROM SMSList
                        WHERE Created BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
                    )
                    AND Sent = 1
                    """.format(sDate, eDate)
                    
                    delivered_result = q.QuerySqlScalar(delivered_sql)
                    if delivered_result is not None:
                        try:
                            delivered_count = int(delivered_result)
                        except:
                            # Not a valid integer, try string conversion
                            try:
                                delivered_count = int(float(str(delivered_result)))
                            except:
                                # Keep default
                                pass
                except Exception as e:
                    print("<div class='alert alert-warning'>Warning: Failed to get SMS delivered count. {0}</div>".format(str(e)))
            else:
                # If SMSItems doesn't exist, assume all sent messages were delivered
                delivered_count = total_count
            
            # Calculate delivery rate
            delivery_rate = 0
            if total_count > 0:
                delivery_rate = (float(delivered_count) / float(total_count)) * 100
            
            return {
                'total_count': total_count,
                'sent_count': delivered_count,
                'delivery_rate': delivery_rate
            }
        except Exception as e:
            # Log error but always return valid default structure
            print("<div class='alert alert-danger'>Error in SMS stats: {0}</div>".format(str(e)))
            return default_stats
    
    @staticmethod
    def get_sms_campaigns(sDate, eDate):
        """Get SMS campaigns with correct relationship to SMSItems table"""
        try:
            # First check if the table exists
            if not DatabaseHelper.table_exists('SMSList'):
                return []
                
            # Simple query to get campaigns with basic counts
            sql = """
            SELECT TOP {2}
                ID,
                COALESCE(Title, 'SMS Message') AS Title,
                COALESCE(Message, '') AS Message,
                Created AS SentDate,
                COALESCE(SentSMS, 0) AS TotalSent,
                0 AS DeliveredCount,  -- Will calculate this later if SMSItems exists
                0 AS FailedCount,  -- Will calculate this later
                SenderID
            FROM SMSList
            WHERE Created BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
            ORDER BY Created DESC
            """.format(sDate, eDate, Config.MAX_ROWS_PER_TABLE)
            
            # Execute the query and get results
            results = q.QuerySql(sql)
            if not results or len(results) == 0:
                return []
                
            # Check if SMSItems table exists for getting delivery statuses
            items_exists = DatabaseHelper.table_exists('SMSItems')
                
            # Convert results to list of dictionaries with error handling
            data = []
            for row in results:
                try:
                    # Extract basic campaign data
                    list_id = getattr(row, 'ID', 0)
                    title = getattr(row, 'Title', 'SMS Message') or 'SMS Message'
                    message = getattr(row, 'Message', '') or ''
                    sent_date = getattr(row, 'SentDate', '')
                    
                    try:
                        total_sent = int(getattr(row, 'TotalSent', 0) or 0)
                    except:
                        total_sent = 0
                    
                    # Get detailed delivery counts if SMSItems exists
                    delivered_count = 0
                    failed_count = 0
                    
                    if items_exists and list_id > 0:
                        try:
                            # Get delivered count
                            delivered_sql = """
                            SELECT COUNT(*) FROM SMSItems
                            WHERE ListID = {0} AND Sent = 1
                            """.format(list_id)
                            
                            delivered_result = q.QuerySqlScalar(delivered_sql)
                            if delivered_result is not None:
                                try:
                                    delivered_count = int(delivered_result)
                                except:
                                    delivered_count = 0
                                    
                            # Calculate failed count from total and delivered
                            if total_sent > 0:
                                failed_count = max(0, total_sent - delivered_count)
                        except:
                            # If query fails, assume all were delivered
                            delivered_count = total_sent
                            failed_count = 0
                    else:
                        # If SMSItems doesn't exist, assume all were delivered
                        delivered_count = total_sent
                        failed_count = 0
                        
                    # Ensure we have valid counts
                    if total_sent <= 0 and (delivered_count + failed_count) > 0:
                        total_sent = delivered_count + failed_count
                        
                    if total_sent > 0 and delivered_count + failed_count != total_sent:
                        # Adjust counts to match total
                        failed_count = max(0, total_sent - delivered_count)
                    
                    # Calculate delivery rate
                    delivery_rate = 0
                    if total_sent > 0:
                        delivery_rate = (float(delivered_count) / float(total_sent)) * 100
                        
                    # Get sender name
                    sender_id = getattr(row, 'SenderID', 0)
                    sender_name = "Unknown"
                    
                    if sender_id and sender_id > 0:
                        sender_name = DatabaseHelper.get_person_name(sender_id)
                    
                    # Truncate long messages for display
                    display_message = message
                    if len(display_message) > 100:
                        display_message = display_message[:100] + '...'
                    
                    # Add to data
                    data.append({
                        'Title': title,
                        'Message': display_message,
                        'SentDate': sent_date,
                        'SentCount': FormatHelper.format_number(delivered_count),
                        'FailedCount': FormatHelper.format_number(failed_count),
                        'DeliveryRate': FormatHelper.format_percent(delivery_rate),
                        'Sender': sender_name
                    })
                except Exception as row_error:
                    # Skip problematic rows
                    continue
            
            return data
        except Exception as e:
            # Log error but return empty list
            print("<div class='alert alert-danger'>Error in SMS campaigns: {0}</div>".format(str(e)))
            return []
    
    @staticmethod
    def get_sms_failure_breakdown(sDate, eDate):
        """Get SMS failure type breakdown"""
        try:
            # Check if tables exist
            if not DatabaseHelper.table_exists('SMSList') or not DatabaseHelper.table_exists('SMSItems'):
                return []
                
            # Query to get breakdown of failure types
            sql = """
            SELECT 
                COALESCE(ResultStatus, 'Unknown') AS Status,
                COUNT(*) AS FailCount
            FROM SMSItems 
            WHERE ListID IN (
                SELECT ID FROM SMSList
                WHERE Created BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
            )
            AND (Sent = 0 OR Sent IS NULL)
            AND ResultStatus IS NOT NULL
            AND ResultStatus <> 'Delivered'
            GROUP BY ResultStatus
            ORDER BY COUNT(*) DESC
            """.format(sDate, eDate)
            
            # Execute query
            results = q.QuerySql(sql)
            
            # Process results with error handling
            data = []
            
            if results and len(results) > 0:
                for row in results:
                    try:
                        status = getattr(row, 'Status', 'Unknown')
                        if status is None or status == '':
                            status = 'Unknown'
                            
                        try:
                            count = int(getattr(row, 'FailCount', 0) or 0)
                        except:
                            count = 0
                            
                        if count > 0:
                            data.append({
                                'Status': status,
                                'Count': count
                            })
                    except:
                        # Skip problematic rows
                        continue
                        
            # If no data found, add default "No failures" entry
            if len(data) == 0:
                data.append({
                    'Status': 'No failures found',
                    'Count': 0
                })
                
            return data
        except Exception as e:
            # Log error but return empty array
            print("<div class='alert alert-danger'>Error retrieving SMS failure breakdown: {0}</div>".format(str(e)))
            return []
    
    @staticmethod
    def get_sms_top_senders(sDate, eDate):
        """Get top SMS senders based on number of campaigns and message volume"""
        try:
            # Check if SMSList table exists
            if not DatabaseHelper.table_exists('SMSList'):
                return []
                
            # Query to get top senders with campaign and message counts
            sql = """
            SELECT TOP 10
                sl.SenderID,
                COUNT(DISTINCT sl.ID) AS CampaignCount,
                SUM(COALESCE(sl.SentSMS, 0)) AS MessageCount
            FROM SMSList sl
            WHERE sl.Created BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
            AND sl.SenderID IS NOT NULL
            GROUP BY sl.SenderID
            ORDER BY COUNT(DISTINCT sl.ID) DESC, SUM(COALESCE(sl.SentSMS, 0)) DESC
            """.format(sDate, eDate)
            
            # Execute query
            results = q.QuerySql(sql)
            
            # Process results with error handling
            data = []
            
            if results and len(results) > 0:
                for row in results:
                    try:
                        sender_id = getattr(row, 'SenderID', 0)
                        
                        try:
                            campaign_count = int(getattr(row, 'CampaignCount', 0) or 0)
                        except:
                            campaign_count = 0
                            
                        try:
                            message_count = int(getattr(row, 'MessageCount', 0) or 0)
                        except:
                            message_count = 0
                            
                        # Skip if no campaigns or messages
                        if campaign_count == 0 and message_count == 0:
                            continue
                            
                        # Get sender name
                        sender_name = "Unknown"
                        if sender_id > 0:
                            sender_name = DatabaseHelper.get_person_name(sender_id)
                                
                        # Calculate delivery rate if SMSItems table exists
                        delivery_rate = 0
                        delivered_count = 0
                        
                        items_exists = DatabaseHelper.table_exists('SMSItems')
                        if items_exists and message_count > 0:
                            try:
                                # Get delivered message count for this sender
                                delivered_sql = """
                                SELECT COUNT(*) FROM SMSItems si
                                JOIN SMSList sl ON si.ListID = sl.ID
                                WHERE sl.SenderID = {0}
                                AND sl.Created BETWEEN '{1} 00:00:00' AND '{2} 23:59:59.999'
                                AND si.Sent = 1
                                """.format(sender_id, sDate, eDate)
                                
                                delivered_result = q.QuerySqlScalar(delivered_sql)
                                if delivered_result is not None:
                                    try:
                                        delivered_count = int(delivered_result)
                                    except:
                                        delivered_count = 0
                            except:
                                # If checking for SMSItems fails, assume all messages delivered
                                delivered_count = message_count
                            
                        # Calculate delivery rate
                        if message_count > 0:
                            if delivered_count == 0:
                                # If we couldn't get delivered count, assume all delivered
                                delivered_count = message_count
                                
                            delivery_rate = (float(delivered_count) / float(message_count)) * 100
                                
                        # Add to data
                        data.append({
                            'SenderName': sender_name,
                            'CampaignCount': campaign_count,
                            'MessageCount': FormatHelper.format_number(message_count),
                            'DeliveredCount': FormatHelper.format_number(delivered_count),
                            'DeliveryRate': FormatHelper.format_percent(delivery_rate)
                        })
                    except:
                        # Skip problematic rows
                        continue
                        
            return data
        except Exception as e:
            # Log error but return empty array
            print("<div class='alert alert-danger'>Error retrieving SMS top senders: {0}</div>".format(str(e)))
            return []

#####################################################################
#### UI RENDERING FUNCTIONS
#####################################################################

class UIRenderer:
    """Class for rendering UI components"""
    
    @staticmethod
    def render_filter_form():
        """Render the filter form"""
        
        # Get programs for dropdown
        try:
            sql_programs = """SELECT Id, Name AS ProgramName FROM Program ORDER BY Name"""
            programs = q.QuerySql(sql_programs)
            
            # Get failure classifications for dropdown
            sql_failure_classifications = """SELECT DISTINCT fe.Fail FROM FailedEmails fe WHERE fe.Fail IS NOT NULL ORDER BY fe.Fail"""
            failure_classifications = q.QuerySql(sql_failure_classifications)
            
            # Build program options HTML
            program_options = '<option value="999999">All Programs</option>'
            for program in programs:
                selected = 'selected="selected"' if str(model.Data.program) == str(program.Id) else ''
                program_options += '<option value="{}" {}>{}</option>'.format(
                    program.Id, selected, program.ProgramName
                )
            
            # Build failure classification options HTML
            failure_options = '<option value="999999">All Failure Types</option>'
            for failure in failure_classifications:
                fail_value = getattr(failure, 'Fail', '')
                selected = 'selected="selected"' if str(model.Data.failclassification) == str(fail_value) else ''
                failure_options += '<option value="{}" {}>{}</option>'.format(
                    fail_value, selected, fail_value
                )
                
            # Determine if checkbox should be checked
            success_checked = 'checked' if model.Data.HideSuccess != 'yes' else ''
            
            # Build the filter form HTML
            html = """
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Filter Options</h3>
                </div>
                <div class="panel-body">
                    <form action="" method="GET" class="form-horizontal" id="filter-form">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="form-group">
                                    <label class="col-sm-4 control-label">From Date:</label>
                                    <div class="col-sm-8">
                                        <input type="date" name="sDate" required class="form-control" value="{}">
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="col-sm-4 control-label">To Date:</label>
                                    <div class="col-sm-8">
                                        <input type="date" name="eDate" required class="form-control" value="{}">
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="form-group">
                                    <label class="col-sm-4 control-label">Program:</label>
                                    <div class="col-sm-8">
                                        <select name="program" class="form-control">
                                            {}
                                        </select>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="col-sm-4 control-label">Failure Type:</label>
                                    <div class="col-sm-8">
                                        <select name="failclassification" class="form-control">
                                            {}
                                        </select>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <div class="col-sm-offset-4 col-sm-8">
                                        <div class="checkbox">
                                            <label>
                                                <input type="checkbox" name="HideSuccess" value="no" {}>
                                                Show Successfully Sent Emails <i>(May be slower for large datasets)</i>
                                            </label>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <input type="hidden" name="activeTab" value="{}">
                        
                        <div class="form-group">
                            <div class="col-sm-offset-2 col-sm-10">
                                <button type="submit" class="btn btn-primary">
                                    <i class="glyphicon glyphicon-filter"></i> Apply Filters
                                </button>
                                <a href="?" class="btn btn-default">
                                    <i class="glyphicon glyphicon-refresh"></i> Reset Filters
                                </a>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
            """.format(
                model.Data.sDate, 
                model.Data.eDate,
                program_options,
                failure_options,
                success_checked,
                model.Data.activeTab
            )
            
            return html
            
        except Exception as e:
            return ErrorHandler.handle_error(e, "filter form")

    @staticmethod
    def render_nav_tabs():
        """Render navigation tabs for different sections"""
        
        # Define tabs
        tabs = [
            {'id': 'dashboard', 'label': 'Dashboard Overview', 'icon': 'dashboard'},
            {'id': 'email', 'label': 'Email Stats', 'icon': 'envelope'},
            {'id': 'sms', 'label': 'SMS Stats', 'icon': 'phone'},
            {'id': 'senders', 'label': 'Top Senders', 'icon': 'user'},
            {'id': 'programs', 'label': 'Program Stats', 'icon': 'list-alt'}
        ]
        
        html = """
        <ul class="nav nav-tabs" role="tablist">
        """
        
        for tab in tabs:
            active = 'active' if model.Data.activeTab == tab['id'] else ''
            html += """
            <li role="presentation" class="{}">
                <a href="?activeTab={}&sDate={}&eDate={}&program={}&failclassification={}&HideSuccess={}">
                    <i class="glyphicon glyphicon-{}"></i> {}
                </a>
            </li>
            """.format(
                active,
                tab['id'],
                model.Data.sDate,
                model.Data.eDate,
                model.Data.program,
                model.Data.failclassification,
                'no' if model.Data.HideSuccess != 'yes' else 'yes',
                tab['icon'],
                tab['label']
            )
        
        html += """
        </ul>
        <div class="tab-content" style="padding-top: 20px;">
        """
        
        return html

    @staticmethod
    def render_dashboard_overview(data):
        """Render the main dashboard overview"""
        
        if not data:
            return '<div class="alert alert-warning">No data available for the selected time period.</div>'
        
        # Extract values
        email_summary = data.get('email_summary', {})
        sms_stats = data.get('sms_stats', {})
        email_campaign_count = data.get('email_campaign_count', 0)
        sms_campaign_count = data.get('sms_campaign_count', 0)
        date_range = data.get('date_range', {})
        
        # Format dates for display
        start_date = FormatHelper.format_date(date_range.get('start', ''))
        end_date = FormatHelper.format_date(date_range.get('end', ''))
        
        # Email metrics
        total_emails = email_summary.get('total_emails', 0)
        email_delivery_rate = email_summary.get('delivery_rate', 0)
        
        # SMS metrics
        total_sms = sms_stats.get('total_count', 0)
        sms_delivery_rate = sms_stats.get('delivery_rate', 0)
        
        # Calculate overall stats
        combined_delivery_items = total_emails + total_sms
        combined_delivery_rate = 0
        
        if combined_delivery_items > 0:
            combined_delivery_rate = (
                (email_summary.get('sent_emails', 0) + sms_stats.get('sent_count', 0)) / 
                float(combined_delivery_items)
            ) * 100
        
        # Create HTML
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h3 class="panel-title">Communication Summary ({0} to {1})</h3>
            </div>
            <div class="panel-body">
                <!-- Overall metrics -->
                <div class="row text-center">
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{2}</h2>
                            <p>Total Communications</p>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{3}</h2>
                            <p>Email Campaigns</p>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{4}</h2>
                            <p>SMS Campaigns</p>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{5}</h2>
                            <p>Overall Delivery Rate</p>
                        </div>
                    </div>
                </div>
                
                <!-- Detailed sections in columns -->
                <div class="row">
                    <!-- Email summary section -->
                    <div class="col-md-6">
                        <div class="panel panel-info">
                            <div class="panel-heading">
                                <h3 class="panel-title">Email Summary</h3>
                            </div>
                            <div class="panel-body">
                                <div class="row text-center">
                                    <div class="col-xs-6">
                                        <div class="well well-sm">
                                            <h3>{6}</h3>
                                            <p>Total Emails</p>
                                        </div>
                                    </div>
                                    <div class="col-xs-6">
                                        <div class="well well-sm">
                                            <h3>{7}</h3>
                                            <p>Delivery Rate</p>
                                        </div>
                                    </div>
                                </div>
                                <div class="progress">
                                    <div class="progress-bar progress-bar-info" role="progressbar" 
                                        style="width: {8}%;" aria-valuenow="{9}" aria-valuemin="0" aria-valuemax="100">
                                        {10}%
                                    </div>
                                </div>
                                <div class="text-center">
                                    <a href="?activeTab=email&sDate={11}&eDate={12}" class="btn btn-info btn-sm">
                                        <i class="glyphicon glyphicon-stats"></i> Detailed Email Stats
                                    </a>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- SMS summary section -->
                    <div class="col-md-6">
                        <div class="panel panel-success">
                            <div class="panel-heading">
                                <h3 class="panel-title">SMS Summary</h3>
                            </div>
                            <div class="panel-body">
                                <div class="row text-center">
                                    <div class="col-xs-6">
                                        <div class="well well-sm">
                                            <h3>{13}</h3>
                                            <p>Total SMS</p>
                                        </div>
                                    </div>
                                    <div class="col-xs-6">
                                        <div class="well well-sm">
                                            <h3>{14}</h3>
                                            <p>Delivery Rate</p>
                                        </div>
                                    </div>
                                </div>
                                <div class="progress">
                                    <div class="progress-bar progress-bar-success" role="progressbar" 
                                        style="width: {15}%;" aria-valuenow="{16}" aria-valuemin="0" aria-valuemax="100">
                                        {17}%
                                    </div>
                                </div>
                                <div class="text-center">
                                    <a href="?activeTab=sms&sDate={18}&eDate={19}" class="btn btn-success btn-sm">
                                        <i class="glyphicon glyphicon-stats"></i> Detailed SMS Stats
                                    </a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Additional shortcuts -->
                <div class="row">
                    <div class="col-md-12">
                        <div class="panel panel-default">
                            <div class="panel-heading">
                                <h3 class="panel-title">Quick Links</h3>
                            </div>
                            <div class="panel-body">
                                <div class="row text-center">
                                    <div class="col-xs-4">
                                        <a href="?activeTab=senders&sDate={20}&eDate={21}" class="btn btn-primary btn-block">
                                            <i class="glyphicon glyphicon-user"></i> Top Senders
                                        </a>
                                    </div>
                                    <div class="col-xs-4">
                                        <a href="?activeTab=programs&sDate={22}&eDate={23}" class="btn btn-primary btn-block">
                                            <i class="glyphicon glyphicon-list-alt"></i> Program Stats
                                        </a>
                                    </div>
                                    <div class="col-xs-4">
                                        <a href="https://www.twilio.com/docs/sendgrid/ui/analytics-and-reporting/bounce-and-block-classifications" 
                                           target="_blank" class="btn btn-default btn-block">
                                            <i class="glyphicon glyphicon-question-sign"></i> Failure Documentation
                                        </a>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """.format(
            start_date, end_date,
            FormatHelper.format_number(combined_delivery_items),
            FormatHelper.format_number(email_campaign_count),
            FormatHelper.format_number(sms_campaign_count),
            FormatHelper.format_percent(combined_delivery_rate),
            
            FormatHelper.format_number(total_emails),
            FormatHelper.format_percent(email_delivery_rate),
            min(100, max(0, email_delivery_rate)),
            int(email_delivery_rate),
            int(email_delivery_rate),
            model.Data.sDate, model.Data.eDate,
            
            FormatHelper.format_number(total_sms),
            FormatHelper.format_percent(sms_delivery_rate),
            min(100, max(0, sms_delivery_rate)),
            int(sms_delivery_rate),
            int(sms_delivery_rate),
            model.Data.sDate, model.Data.eDate,
            
            model.Data.sDate, model.Data.eDate,
            model.Data.sDate, model.Data.eDate
        )
        
        return html

    @staticmethod
    def render_email_summary(summary_data):
        """Render email summary metrics section with Mailchimp-style metrics"""
        
        # Extract values from summary data with defaults
        total_emails = summary_data.get('total_emails', 0)
        sent_emails = summary_data.get('sent_emails', 0)
        failed_emails = summary_data.get('failed_emails', 0)
        delivery_rate = summary_data.get('delivery_rate', 0)
        opens = summary_data.get('opens', 0)
        clicks = summary_data.get('clicks', 0)
        open_rate = summary_data.get('open_rate', 0)
        click_rate = summary_data.get('click_rate', 0)
        bounces = summary_data.get('bounces', 0)
        bounce_rate = summary_data.get('bounce_rate', 0)
        unsubscribes = summary_data.get('unsubscribes', 0)
        unsubscribe_rate = summary_data.get('unsubscribe_rate', 0)
        
        # Create metrics HTML with Mailchimp-style layout
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h3 class="panel-title">Email Campaign Performance Metrics</h3>
            </div>
            <div class="panel-body">
                <!-- Primary Metrics Row -->
                <div class="row text-center" style="margin-bottom: 30px;">
                    <div class="col-md-3 col-sm-6">
                        <div class="metric-card" style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                            <div style="font-size: 36px; font-weight: bold; color: #2c3e50;">{0}</div>
                            <div style="color: #7f8c8d; margin-top: 10px;">Total Emails Sent</div>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="metric-card" style="background: #e8f5e9; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                            <div style="font-size: 36px; font-weight: bold; color: #27ae60;">{1}</div>
                            <div style="color: #7f8c8d; margin-top: 10px;">Delivery Rate</div>
                            <div style="font-size: 14px; color: #95a5a6;">({2} delivered)</div>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="metric-card" style="background: #e3f2fd; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                            <div style="font-size: 36px; font-weight: bold; color: #2196f3;">{3}</div>
                            <div style="color: #7f8c8d; margin-top: 10px;">Open Rate</div>
                            <div style="font-size: 14px; color: #95a5a6;">({4} opens)</div>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="metric-card" style="background: #f3e5f5; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                            <div style="font-size: 36px; font-weight: bold; color: #9c27b0;">{5}</div>
                            <div style="color: #7f8c8d; margin-top: 10px;">Click Rate</div>
                            <div style="font-size: 14px; color: #95a5a6;">({6} total clicks)</div>
                        </div>
                    </div>
                </div>
                
                <!-- Secondary Metrics Row -->
                <div class="row text-center">
                    <div class="col-md-4">
                        <div class="metric-card" style="background: #fff3e0; padding: 15px; border-radius: 8px;">
                            <div style="font-size: 24px; font-weight: bold; color: #ff9800;">{7}</div>
                            <div style="color: #7f8c8d; margin-top: 5px;">
                                Bounce Rate 
                                <a href="#" onclick="$('#bounceHelpModal').modal('show'); return false;" 
                                   style="color: #95a5a6; font-size: 12px;">
                                    <i class="glyphicon glyphicon-question-sign"></i>
                                </a>
                            </div>
                            <div style="font-size: 12px; color: #95a5a6;">({8} bounces)</div>
                            <div class="progress" style="height: 6px; margin-top: 10px;">
                                <div class="progress-bar" style="background-color: #ff9800; width: {9}%;"></div>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="metric-card" style="background: #ffebee; padding: 15px; border-radius: 8px;">
                            <div style="font-size: 24px; font-weight: bold; color: #f44336;">{10}</div>
                            <div style="color: #7f8c8d; margin-top: 5px;">Unsubscribe Rate</div>
                            <div style="font-size: 12px; color: #95a5a6;">({11} unsubscribes)</div>
                            <div class="progress" style="height: 6px; margin-top: 10px;">
                                <div class="progress-bar" style="background-color: #f44336; width: {12}%;"></div>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="metric-card" style="background: #fce4ec; padding: 15px; border-radius: 8px;">
                            <div style="font-size: 24px; font-weight: bold; color: #e91e63;">{13}</div>
                            <div style="color: #7f8c8d; margin-top: 5px;">Failed Emails</div>
                            <div style="font-size: 12px; color: #95a5a6;">({14} emails)</div>
                            <div class="progress" style="height: 6px; margin-top: 10px;">
                                <div class="progress-bar" style="background-color: #e91e63; width: {15}%;"></div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Engagement Funnel Visualization -->
                <div style="margin-top: 30px; padding: 20px; background: #f8f9fa; border-radius: 8px;">
                    <h4 style="margin-bottom: 20px; color: #2c3e50;">Engagement Funnel</h4>
                    <div class="funnel-chart">
                        <div class="funnel-stage" style="background: #3498db; width: 100%; padding: 10px; color: white; margin-bottom: 2px; position: relative; min-height: 40px;">
                            <span style="position: relative; z-index: 1;">Sent: {16}</span>
                        </div>
                        <div class="funnel-stage-wrapper" style="width: 100%; background: #e8e8e8; margin-bottom: 2px; position: relative;">
                            <div class="funnel-stage" style="background: #2ecc71; width: {17}%; padding: 10px; color: white; min-width: 1px; min-height: 40px;">
                                <span style="position: absolute; left: 10px; color: #2c3e50; white-space: nowrap;">Delivered: {18} ({21})</span>
                            </div>
                        </div>
                        <div class="funnel-stage-wrapper" style="width: 100%; background: #e8e8e8; margin-bottom: 2px; position: relative;">
                            <div class="funnel-stage" style="background: #f39c12; width: {19}%; padding: 10px; color: white; min-width: 1px; min-height: 40px;">
                                <span style="position: absolute; left: 10px; color: #2c3e50; white-space: nowrap;">Opened: {22} ({23})</span>
                            </div>
                        </div>
                        <div class="funnel-stage-wrapper" style="width: 100%; background: #e8e8e8; position: relative;">
                            <div class="funnel-stage" style="background: #e74c3c; width: {20}%; padding: 10px; color: white; min-width: 1px; min-height: 40px;">
                                <span style="position: absolute; left: 10px; color: #2c3e50; white-space: nowrap;">Clicked: {24} ({25})</span>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Industry Benchmarks Comparison -->
                <div style="margin-top: 20px; padding: 15px; background: #ecf0f1; border-radius: 8px;">
                    <p style="margin: 0; font-size: 12px; color: #7f8c8d;">
                        <strong>Industry Benchmarks:</strong> 
                        Average Open Rate: 20-25% | Average Click Rate: 2-3% | Average Bounce Rate: <2% | Average Unsubscribe Rate: <0.5%
                    </p>
                </div>
            </div>
        </div>
        
        <style>
        .metric-card {
            transition: transform 0.2s;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        .funnel-stage {
            text-align: center;
            transition: all 0.3s;
        }
        </style>
        """
        
        # Calculate some values for the funnel
        funnel_delivery_width = delivery_rate
        funnel_open_width = open_rate if sent_emails > 0 else 0
        funnel_click_width = (float(clicks) / float(total_emails) * 100) if total_emails > 0 else 0
        
        # Format all values
        formatted_values = {
            'total_emails_fmt': FormatHelper.format_number(total_emails),
            'delivery_rate_fmt': FormatHelper.format_percent(delivery_rate),
            'sent_emails_fmt': FormatHelper.format_number(sent_emails),
            'open_rate_fmt': FormatHelper.format_percent(open_rate),
            'opens_fmt': FormatHelper.format_number(opens),
            'click_rate_fmt': FormatHelper.format_percent(click_rate),
            'clicks_fmt': FormatHelper.format_number(clicks),
            'bounce_rate_fmt': FormatHelper.format_percent(bounce_rate),
            'bounces_fmt': FormatHelper.format_number(bounces),
            'bounce_progress': min(100, max(0, bounce_rate * 10)),
            'unsub_rate_fmt': FormatHelper.format_percent(unsubscribe_rate),
            'unsubs_fmt': FormatHelper.format_number(unsubscribes),
            'unsub_progress': min(100, max(0, unsubscribe_rate * 100)),
            'failed_rate_fmt': FormatHelper.format_percent((float(failed_emails) / float(total_emails) * 100) if total_emails > 0 else 0),
            'failed_emails_fmt': FormatHelper.format_number(failed_emails),
            'failed_progress': min(100, max(0, (float(failed_emails) / float(total_emails) * 100) if total_emails > 0 else 0)),
            'funnel_delivery_width': funnel_delivery_width,
            'funnel_open_width': funnel_open_width,
            'funnel_click_width': funnel_click_width
        }
        
        # Now substitute each value individually
        html = html.replace('{0}', formatted_values['total_emails_fmt'])
        html = html.replace('{1}', formatted_values['delivery_rate_fmt'])
        html = html.replace('{2}', formatted_values['sent_emails_fmt'])
        html = html.replace('{3}', formatted_values['open_rate_fmt'])
        html = html.replace('{4}', formatted_values['opens_fmt'])
        html = html.replace('{5}', formatted_values['click_rate_fmt'])
        html = html.replace('{6}', formatted_values['clicks_fmt'])
        html = html.replace('{7}', formatted_values['bounce_rate_fmt'])
        html = html.replace('{8}', formatted_values['bounces_fmt'])
        html = html.replace('{9}', str(formatted_values['bounce_progress']))
        html = html.replace('{10}', formatted_values['unsub_rate_fmt'])
        html = html.replace('{11}', formatted_values['unsubs_fmt'])
        html = html.replace('{12}', str(formatted_values['unsub_progress']))
        html = html.replace('{13}', formatted_values['failed_rate_fmt'])
        html = html.replace('{14}', formatted_values['failed_emails_fmt'])
        html = html.replace('{15}', str(formatted_values['failed_progress']))
        html = html.replace('{16}', formatted_values['total_emails_fmt'])
        html = html.replace('{17}', str(formatted_values['funnel_delivery_width']))
        html = html.replace('{18}', formatted_values['sent_emails_fmt'])
        html = html.replace('{19}', str(formatted_values['funnel_open_width']))
        html = html.replace('{20}', str(formatted_values['funnel_click_width']))
        html = html.replace('{21}', formatted_values['delivery_rate_fmt'])
        html = html.replace('{22}', formatted_values['opens_fmt'])
        html = html.replace('{23}', formatted_values['open_rate_fmt'])
        html = html.replace('{24}', formatted_values['clicks_fmt'])
        html = html.replace('{25}', formatted_values['click_rate_fmt'])
        
        # Add bounce help modal
        bounce_modal = """
        <!-- Bounce Help Modal -->
        <div class="modal fade" id="bounceHelpModal" tabindex="-1" role="dialog">
            <div class="modal-dialog" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                            <span aria-hidden="true">&times;</span>
                        </button>
                        <h4 class="modal-title">Understanding Email Bounce Rate</h4>
                    </div>
                    <div class="modal-body">
                        <p><strong>What is included in the bounce rate?</strong></p>
                        <p>The bounce rate includes emails that failed to deliver due to permanent or temporary delivery issues. Specifically, we count the following failure types as bounces:</p>
                        
                        <ul>
                            <li><strong>bouncedaddress</strong> - Email address has explicitly bounced</li>
                            <li><strong>Mailbox Unavailable</strong> - The recipient's mailbox doesn't exist or is full</li>
                            <li><strong>Invalid Address</strong> - The email address format is incorrect or doesn't exist</li>
                            <li><strong>invalid</strong> - General invalid address indicator</li>
                            <li><strong>bounce/hardbounce</strong> - Traditional bounce classifications</li>
                            <li><strong>blocked</strong> - Email was blocked by the recipient's server</li>
                        </ul>
                        
                        <p><strong>What is NOT included in bounce rate?</strong></p>
                        <ul>
                            <li>Spam reports or spam-related blocks</li>
                            <li>Reputation issues</li>
                            <li>Technical failures on our end</li>
                            <li>Unclassified failures</li>
                        </ul>
                        
                        <p><strong>Industry Benchmark:</strong> A healthy bounce rate should be below 2%. Higher rates may indicate issues with your email list quality.</p>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                    </div>
                </div>
            </div>
        </div>
        """
        
        return html + bounce_modal

    @staticmethod
    def render_failure_breakdown(failed_types):
        """Render failure type breakdown section"""
        
        if not failed_types:
            return '<div class="alert alert-info">No failures detected in the selected time period.</div>'
        
        # Create HTML for failure breakdown
        html = """
        <div class="panel panel-warning">
            <div class="panel-heading">
                <h3 class="panel-title">Failure Type Breakdown</h3>
            </div>
            <div class="panel-body">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Failure Type</th>
                            <th class="text-right">Count</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        # Add each failure type
        for failure in failed_types:
            html += """
                        <tr>
                            <td>{0}</td>
                            <td class="text-right">{1}</td>
                        </tr>
            """.format(failure.get('Status', ''), FormatHelper.format_number(failure.get('Count', 0)))
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """
        
        return html

    @staticmethod
    def render_recent_campaigns(campaigns, page=1):
        """Render recent campaigns table"""
        
        if not campaigns:
            return '<div class="alert alert-info">No campaigns found in the selected time period.</div>'
        
        # Check if we're showing single-recipient emails
        show_single = getattr(model.Data, 'show_single_recipient', 'false') == 'true'
        
        # No pagination needed
        
        # Create HTML for campaigns table with server-side pagination
        html = """
        <div class="panel panel-info">
            <div class="panel-heading">
                <h3 class="panel-title">
                    Recent Email Campaigns with Engagement Metrics
                    <span class="badge" id="campaign-count">{0}</span>
                    <small class="pull-right" id="single-recipient-info"></small>
                </h3>
            </div>
            <div class="panel-body">
                <form id="campaigns-form" method="GET" action="">
                    <!-- Hidden fields to maintain state -->
                    <input type="hidden" name="sDate" value="{1}">
                    <input type="hidden" name="eDate" value="{2}">
                    <input type="hidden" name="program" value="{3}">
                    <input type="hidden" name="failclassification" value="{4}">
                    <input type="hidden" name="HideSuccess" value="{5}">
                    <input type="hidden" name="activeTab" value="email">
                    <input type="hidden" id="show_single_recipient" name="show_single_recipient" value="{6}">
                    
                    <div style="margin-bottom: 15px;">
                        <div class="checkbox">
                            <label>
                                <input type="checkbox" id="showSingleRecipient" {8} onchange="toggleSingleRecipientFilter()">
                                Show single-recipient emails
                            </label>
                            <span class="text-muted" style="margin-left: 10px;">
                                <small>(Showing up to {7} campaigns{9})</small>
                            </span>
                        </div>
                    </div>
                </form>
                
                <div class="table-responsive">
                    <table class="table table-striped table-bordered table-hover table-condensed" id="recent-campaigns-table">
                        <thead>
                            <tr>
                                <th>Date</th>
                                <th>Campaign</th>
                                <th>Sender</th>
                                <th>Recipients</th>
                                <th>Opens</th>
                                <th>Open Rate</th>
                                <th>Clicks</th>
                                <th>Click Rate</th>
                                <th>Failures</th>
                                <th>Fail Rate</th>
                                <th>Unsubs</th>
                            </tr>
                        </thead>
                        <tbody id="campaigns-tbody">""".format(
            len(campaigns),  # Show count
            model.Data.sDate,
            model.Data.eDate,
            model.Data.program,
            model.Data.failclassification,
            'no' if model.Data.HideSuccess != 'yes' else 'yes',
            'true' if show_single else 'false',
            Config.MAX_ROWS_PER_TABLE,
            'checked' if show_single else '',
            ' - Performance mode' if Config.PERFORMANCE_MODE else ''  # Performance mode indicator
        )
        
        # Add all campaigns from this page
        for campaign in campaigns:
            row_style = 'background-color: #f9f9f9;' if campaign.get('IsSingleRecipient', False) else ''
            
            html += """
                            <tr style="{0}">
                                <td>{1}</td>
                                <td><a href="/Manage/Emails/Details/{2}">{3}</a></td>
                                <td><a href="/Manage/Emails/SentBy/{4}">{5}</a></td>
                                <td>{6}</td>
                                <td>{7}</td>
                                <td>{8}</td>
                                <td>{9}</td>
                                <td>{10}</td>
                                <td>{11}</td>
                                <td>{12}</td>
                                <td>{13}</td>
                            </tr>""".format(
                row_style,
                campaign['SentDate'],
                campaign['EmailQueueId'],
                campaign['CampaignName'],
                campaign['SenderPeopleId'] or '0',
                campaign['Sender'],
                campaign['RecipientCountFormatted'],
                campaign['OpenCount'],
                campaign['OpenRate'],
                campaign['ClickCount'],
                campaign['ClickRate'],
                campaign['FailureCount'],
                campaign['FailureRate'],
                campaign['UnsubscribeCount']
            )
        
        html += """
                        </tbody>
                    </table>
                </div>
                
                <div class="text-muted" style="margin-top: 10px;">
                    <small>
                        <i class="glyphicon glyphicon-info-sign"></i> 
                        Showing up to {0} campaigns. Single-recipient emails are highlighted in gray.
                    </small>
                </div>
            </div>
        </div>
        
        <script>
        function toggleSingleRecipientFilter() {{
            // Update the hidden field value
            var checkbox = document.getElementById('showSingleRecipient');
            document.getElementById('show_single_recipient').value = checkbox.checked ? 'true' : 'false';
            
            // Submit form to reload with filter applied
            document.getElementById('campaigns-form').submit();
        }}
        </script>
        """.format(
            Config.MAX_ROWS_PER_TABLE
        )
        
        return html

    @staticmethod
    def render_active_senders(senders):
        """Render most active senders table"""
        
        if not senders:
            return '<div class="alert alert-info">No sender data available for the selected time period.</div>'
        
        # Define column headers and labels
        header_labels = {
            'SenderName': 'Sender Name',
            'CampaignCount': 'Campaigns Sent',
            'TotalRecipients': 'Total Recipients',
            'SuccessfulDeliveries': 'Successful Deliveries',
            'DeliveryRate': 'Delivery Rate',
            'Unsubscribes': 'Unsubscribes'
        }
        
        # Create HTML for senders table
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h3 class="panel-title">Most Active Email Senders</h3>
            </div>
            <div class="panel-body">
        """
        
        html += TableGenerator.generate_html_table(
            senders,
            column_order=['SenderName', 'CampaignCount', 'TotalRecipients', 'SuccessfulDeliveries', 'DeliveryRate', 'Unsubscribes'],
            header_labels=header_labels,
            id="active-senders-table"
        )
        
        html += """
            </div>
        </div>
        """
        
        return html

    @staticmethod
    def render_program_stats(program_stats):
        """Render program stats table"""
        
        if not program_stats:
            return '<div class="alert alert-info">No program data available for the selected time period.</div>'
        
        # Define column headers and labels
        header_labels = {
            'ProgramName': 'Program',
            'CampaignCount': 'Campaigns',
            'TotalRecipients': 'Recipients',
            'SuccessfulDeliveries': 'Deliveries',
            'FailureCount': 'Failures',
            'DeliveryRate': 'Delivery Rate'
        }
        
        # Create HTML for program stats table
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h3 class="panel-title">Email Stats by Program</h3>
            </div>
            <div class="panel-body">
        """
        
        html += TableGenerator.generate_html_table(
            program_stats,
            column_order=['ProgramName', 'CampaignCount', 'TotalRecipients', 'SuccessfulDeliveries', 'FailureCount', 'DeliveryRate'],
            header_labels=header_labels,
            id="program-stats-table"
        )
        
        html += """
            </div>
        </div>
        """
        
        return html

    @staticmethod
    def render_failed_recipients(recipients):
        """Render recipients with failed emails with expandable rows"""
        
        if not recipients:
            return '<div class="alert alert-info">No recipients with failed emails found in the selected time period.</div>'
        
        # Determine if we need to limit display
        total_recipients = len(recipients)
        initial_display_limit = 10
        show_limited = total_recipients > initial_display_limit
        
        # Create HTML for recipients table
        html = """
        <div class="panel panel-danger">
            <div class="panel-heading">
                <h3 class="panel-title">
                    Recipients with Failed Emails 
                    <span class="badge">{0}</span>
                </h3>
            </div>
            <div class="panel-body">
        """.format(total_recipients)
        
        if show_limited:
            # Show limited rows initially
            html += """
                <div id="failed-recipients-limited">
                    <table class="table table-striped table-bordered table-hover" id="failed-recipients-table-limited">
                        <thead>
                            <tr>
                                <th>Recipient Name</th>
                                <th>Email Address</th>
                                <th>Failure Type</th>
                                <th>Failure Count</th>
                            </tr>
                        </thead>
                        <tbody>
            """
            
            # Add limited data rows
            for recipient in recipients[:initial_display_limit]:
                html += '<tr>'
                html += '<td><a href="/Person2/{0}">{1}</a></td>'.format(
                    recipient.get('PeopleId', 0), recipient.get('Name', '')
                )
                html += '<td>{0}</td>'.format(recipient.get('EmailAddress', ''))
                html += '<td>{0}</td>'.format(recipient.get('FailureType', ''))
                html += '<td>{0}</td>'.format(recipient.get('FailureCount', 0))
                html += '</tr>'
            
            html += """
                        </tbody>
                    </table>
                    <div class="text-center" style="margin-top: 15px;">
                        <button class="btn btn-default" onclick="$('#failed-recipients-limited').hide(); $('#failed-recipients-full').show(); return false;">
                            <i class="glyphicon glyphicon-chevron-down"></i> 
                            Show All {0} Recipients
                        </button>
                    </div>
                </div>
                
                <div id="failed-recipients-full" style="display: none;">
            """.format(total_recipients)
        
        # Full table (or only table if under limit)
        html += """
                    <table class="table table-striped table-bordered table-hover" id="failed-recipients-table" {0}>
                        <thead>
                            <tr>
                                <th>Recipient Name</th>
                                <th>Email Address</th>
                                <th>Failure Type</th>
                                <th>Failure Count</th>
                            </tr>
                        </thead>
                        <tbody>
        """.format('style="display: none;"' if show_limited else '')
        
        # Add all data rows
        for recipient in recipients:
            html += '<tr>'
            html += '<td><a href="/Person2/{0}">{1}</a></td>'.format(
                recipient.get('PeopleId', 0), recipient.get('Name', '')
            )
            html += '<td>{0}</td>'.format(recipient.get('EmailAddress', ''))
            html += '<td>{0}</td>'.format(recipient.get('FailureType', ''))
            html += '<td>{0}</td>'.format(recipient.get('FailureCount', 0))
            html += '</tr>'
        
        html += """
                        </tbody>
                    </table>
        """
        
        if show_limited:
            html += """
                    <div class="text-center" style="margin-top: 15px;">
                        <button class="btn btn-default" onclick="$('#failed-recipients-full').hide(); $('#failed-recipients-limited').show(); return false;">
                            <i class="glyphicon glyphicon-chevron-up"></i> 
                            Show Less
                        </button>
                    </div>
                </div>
            """
        
        html += """
            </div>
        </div>
        """
        
        return html
    
    @staticmethod
    def render_subscriber_growth_chart(growth_data):
        """Render email recipient growth chart with first-time recipients vs unsubscribes"""
        
        if not growth_data:
            return '<div class="alert alert-info">No email recipient growth data available for the selected time period.</div>'
        
        # Prepare data for Chart.js
        labels = []
        new_subscribers = []
        unsubscribes = []
        net_growth = []
        
        for row in growth_data:
            labels.append(str(row.get('DateGroup', '')))
            new_subscribers.append(row.get('NewSubscribers', 0))
            unsubscribes.append(row.get('Unsubscribes', 0))
            net_growth.append(row.get('NetGrowth', 0))
        
        # Convert to JSON strings for JavaScript
        labels_json = str(labels).replace("'", '"')
        new_subs_json = str(new_subscribers)
        unsubs_json = str(unsubscribes)
        net_growth_json = str(net_growth)
        
        html = """
        <div class="panel panel-info">
            <div class="panel-heading">
                <h3 class="panel-title">
                    Email Recipient Growth
                    <span class="pull-right">
                        <a href="#" onclick="$('#growthHelp').toggle(); return false;" style="color: white; text-decoration: none;">
                            <i class="glyphicon glyphicon-question-sign"></i>
                        </a>
                    </span>
                </h3>
            </div>
            <div class="panel-body">
                <div id="growthHelp" style="display: none; background: #f0f0f0; padding: 15px; margin-bottom: 20px; border-radius: 5px;">
                    <h4><i class="glyphicon glyphicon-info-sign"></i> Understanding Email Recipient Growth</h4>
                    <p><strong>What is a "period"?</strong> The time period depends on your date range selection:</p>
                    <ul>
                        <li><strong>Daily:</strong> Each day is a separate period</li>
                        <li><strong>Weekly:</strong> Each week (Sunday-Saturday) is a period</li>
                        <li><strong>Monthly:</strong> Each calendar month is a period</li>
                    </ul>
                    <p><strong>How are email recipients tracked?</strong></p>
                    <ul>
                        <li><strong>First-Time Recipients:</strong> People who received their first email from the church during this period (not necessarily subscribers who signed up)</li>
                        <li><strong>Unsubscribes:</strong> People who opted out (DoNotMail flag set) during this period</li>
                        <li><strong>Net Growth:</strong> First-time recipients minus unsubscribes for each period</li>
                    </ul>
                    <p class="text-muted"><small>Note: This tracks actual email activity. Recipients may have been added through organization membership, event registration, or staff selection - not necessarily through voluntary subscription.</small></p>
                </div>
                <canvas id="subscriberGrowthChart" height="100" style="max-height: 400px;"></canvas>
                
                <div class="row text-center" style="margin-top: 20px;">
                    <div class="col-md-4">
                        <div class="metric-card" style="background: #e8f5e9; padding: 15px; border-radius: 8px;">
                            <div style="font-size: 24px; font-weight: bold; color: #4caf50;">{0}</div>
                            <div style="color: #7f8c8d;">Total First-Time Recipients</div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="metric-card" style="background: #ffebee; padding: 15px; border-radius: 8px;">
                            <div style="font-size: 24px; font-weight: bold; color: #f44336;">{1}</div>
                            <div style="color: #7f8c8d;">Total Unsubscribes</div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="metric-card" style="background: #e3f2fd; padding: 15px; border-radius: 8px;">
                            <div style="font-size: 24px; font-weight: bold; color: #2196f3;">{2}</div>
                            <div style="color: #7f8c8d;">Net Growth</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
        <script>
        (function() {{
            // Destroy any existing chart instance
            var existingChart = Chart.getChart('subscriberGrowthChart');
            if (existingChart) {{
                existingChart.destroy();
            }}
            
            var ctx = document.getElementById('subscriberGrowthChart').getContext('2d');
            window.subscriberChart = new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: {3},
                    datasets: [{{
                        label: 'First-Time Recipients',
                        data: {4},
                        borderColor: '#4caf50',
                        backgroundColor: 'rgba(76, 175, 80, 0.1)',
                        borderWidth: 2,
                        tension: 0.1
                    }}, {{
                        label: 'Unsubscribes',
                        data: {5},
                        borderColor: '#f44336',
                        backgroundColor: 'rgba(244, 67, 54, 0.1)',
                        borderWidth: 2,
                        tension: 0.1
                    }}, {{
                        label: 'Net Growth',
                        data: {6},
                        borderColor: '#2196f3',
                        backgroundColor: 'rgba(33, 150, 243, 0.1)',
                        borderWidth: 2,
                        tension: 0.1,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    animation: {{
                        duration: 0  // Disable animation to prevent performance issues
                    }},
                    plugins: {{
                        legend: {{
                            position: 'top',
                        }},
                        title: {{
                            display: false
                        }},
                        tooltip: {{
                            mode: 'index',
                            intersect: false
                        }}
                    }},
                    scales: {{
                        x: {{
                            display: true,
                            title: {{
                                display: true,
                                text: 'Date'
                            }}
                        }},
                        y: {{
                            display: true,
                            title: {{
                                display: true,
                                text: 'Count'
                            }},
                            beginAtZero: true
                        }}
                    }}
                }}
            }});
        }})();
        </script>
        """.format(
            sum(new_subscribers),
            sum(unsubscribes), 
            sum(net_growth),
            labels_json,
            new_subs_json,
            unsubs_json,
            net_growth_json
        )
        
        return html

    @staticmethod
    def render_sms_summary(sms_stats):
        """Render SMS summary with fixed display"""
        # Extract values with defaults
        total_count = sms_stats.get('total_count', 0)
        sent_count = sms_stats.get('sent_count', 0)
        delivery_rate = sms_stats.get('delivery_rate', 0)
        
        # Create metrics HTML
        html = """
        <div class="panel panel-success">
            <div class="panel-heading">
                <h3 class="panel-title">SMS Delivery Summary</h3>
            </div>
            <div class="panel-body">
                <div class="row text-center">
                    <div class="col-md-4 col-sm-4">
                        <div class="well well-sm">
                            <h2>{0}</h2>
                            <p>Total SMS Sent</p>
                        </div>
                    </div>
                    <div class="col-md-4 col-sm-4">
                        <div class="well well-sm">
                            <h2>{1}</h2>
                            <p>Successfully Delivered</p>
                        </div>
                    </div>
                    <div class="col-md-4 col-sm-4">
                        <div class="well well-sm">
                            <h2>{2}</h2>
                            <p>Delivery Rate</p>
                        </div>
                    </div>
                </div>
                
                <div class="progress">
                    <div class="progress-bar progress-bar-success" role="progressbar" style="width: {3}%;" 
                        aria-valuenow="{4}" aria-valuemin="0" aria-valuemax="100">
                        {5}%
                    </div>
                </div>
                <p class="text-center">SMS Delivery Performance</p>
            </div>
        </div>
        """.format(
            FormatHelper.format_number(total_count),
            FormatHelper.format_number(sent_count),
            FormatHelper.format_percent(delivery_rate),
            min(100, max(0, delivery_rate)),
            int(delivery_rate),
            int(delivery_rate)
        )
        
        return html

    @staticmethod
    def render_sms_campaigns(campaigns):
        """Render SMS campaigns table with correct column references"""
        
        if not campaigns:
            return '<div class="alert alert-info">No SMS campaigns found in the selected time period.</div>'
        
        # Define column headers and labels
        header_labels = {
            'Title': 'Campaign Title',
            'Message': 'Message Content',
            'SentDate': 'Date Sent',
            'SentCount': 'Delivered',
            'FailedCount': 'Failed',
            'DeliveryRate': 'Delivery Rate',
            'Sender': 'Sender'
        }
        
        # Create HTML for SMS campaigns table
        html = """
        <div class="panel panel-success">
            <div class="panel-heading">
                <h3 class="panel-title">Recent SMS Campaigns</h3>
            </div>
            <div class="panel-body">
        """
        
        html += TableGenerator.generate_html_table(
            campaigns,
            column_order=['SentDate', 'Title', 'Message', 'SentCount', 'FailedCount', 'DeliveryRate', 'Sender'],
            header_labels=header_labels,
            id="sms-campaigns-table"
        )
        
        html += """
            </div>
        </div>
        """
        
        return html
        
    @staticmethod
    def render_sms_failure_breakdown(failed_types):
        """Render SMS failure type breakdown section"""
        
        if not failed_types:
            return '<div class="alert alert-info">No SMS failures detected in the selected time period.</div>'
        
        # Create HTML for failure breakdown
        html = """
        <div class="panel panel-warning">
            <div class="panel-heading">
                <h3 class="panel-title">SMS Failure Type Breakdown</h3>
            </div>
            <div class="panel-body">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Failure Reason</th>
                            <th class="text-right">Count</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        # Add each failure type
        for failure in failed_types:
            html += """
                        <tr>
                            <td>{0}</td>
                            <td class="text-right">{1}</td>
                        </tr>
            """.format(failure.get('Status', ''), FormatHelper.format_number(failure.get('Count', 0)))
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """
        
        return html

    @staticmethod
    def render_sms_top_senders(senders):
        """Render SMS top senders table"""
        
        if not senders:
            return '<div class="alert alert-info">No SMS sender data available for the selected time period.</div>'
        
        # Define column headers and labels
        header_labels = {
            'SenderName': 'Sender Name',
            'CampaignCount': 'Campaigns Sent',
            'MessageCount': 'Total Messages',
            'DeliveredCount': 'Delivered Messages',
            'DeliveryRate': 'Delivery Rate'
        }
        
        # Create HTML for senders table
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h3 class="panel-title">Top SMS Senders</h3>
            </div>
            <div class="panel-body">
        """
        
        html += TableGenerator.generate_html_table(
            senders,
            column_order=['SenderName', 'CampaignCount', 'MessageCount', 'DeliveredCount', 'DeliveryRate'],
            header_labels=header_labels,
            id="sms-top-senders-table"
        )
        
        html += """
            </div>
        </div>
        """
        
        return html

    @staticmethod
    def render_dashboard_resources():
        """Add required CSS and JavaScript resources"""
        
        html = """
        <!-- jQuery (required for Bootstrap and our AJAX) -->
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
        
        <!-- Bootstrap CSS -->
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.4.1/css/bootstrap.min.css">
        
        <!-- Bootstrap JS -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.4.1/js/bootstrap.min.js"></script>
        
        <style>
        .dashboard-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 15px;
        }
        
        .well {
            margin-bottom: 0;
        }
        
        .well h2, .well h3 {
            margin-top: 10px;
            margin-bottom: 5px;
        }
        
        .progress {
            height: 25px;
            margin-top: 20px;
            margin-bottom: 10px;
        }
        
        .progress-bar {
            line-height: 25px;
            font-size: 14px;
        }
        
        .panel {
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .panel-heading {
            font-weight: bold;
        }
        
        .table th {
            background-color: #f8f8f8;
        }
        
        /* Loading indicator */
        .loading-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.7);
            z-index: 9999;
        }
        
        .loading-spinner {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
        }
        
        .spinner {
            border: 5px solid #f3f3f3;
            border-top: 5px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 2s linear infinite;
            margin: 0 auto 15px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Responsive table handling */
        @media (max-width: 767px) {
            .table-responsive {
                border: none;
            }
        }
        </style>
        
        <script>
        // Show loading overlay when form is submitted
        document.addEventListener('DOMContentLoaded', function() {
            var form = document.getElementById('filter-form');
            if (form) {
                form.addEventListener('submit', function() {
                    var overlay = document.getElementById('loading-overlay');
                    if (overlay) {
                        overlay.style.display = 'block';
                    }
                });
            }
            
            // Handle tab links to show loading indicator
            var tabLinks = document.querySelectorAll('.nav-tabs a');
            for (var i = 0; i < tabLinks.length; i++) {
                tabLinks[i].addEventListener('click', function() {
                    var overlay = document.getElementById('loading-overlay');
                    if (overlay) {
                        overlay.style.display = 'block';
                    }
                });
            }
        });
        
        // Clean up charts when navigating away
        window.addEventListener('beforeunload', function() {
            if (window.subscriberChart) {
                window.subscriberChart.destroy();
                window.subscriberChart = null;
            }
        });
        </script>
        
        <!-- Loading overlay -->
        <div id="loading-overlay" class="loading-overlay">
            <div class="loading-spinner">
                <div class="spinner"></div>
                <p>Loading data, please wait...</p>
            </div>
        </div>
        """
        
        return html

#####################################################################
#### MAIN EXECUTION CODE
#####################################################################

def main():
    """Main function to render the dashboard with extra error handling"""
    try:
        # No AJAX handling needed - pagination removed
        
        # Surround everything with try/except to ensure we always return something
        try:
            # Get filter parameters
            sDate = model.Data.sDate
            eDate = model.Data.eDate
            hide_success = model.Data.HideSuccess
            program_filter = model.Data.program
            failure_filter = model.Data.failclassification
            active_tab = model.Data.activeTab
            
            # Build dashboard content
            output = UIRenderer.render_dashboard_resources()
            output += '<div class="dashboard-container">'
            
            # Header
            output += """
            <div class="page-header">
                <h1>{0}                  <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                  </svg> <small>Data as of {1}</small></h1>
            </div>
            """.format(Config.DASHBOARD_TITLE, current_date)
            
            # Filter form
            output += UIRenderer.render_filter_form()
            
            # Tab navigation
            output += UIRenderer.render_nav_tabs()
            
            # Wrap each tab's content in a separate try/except to ensure partial functionality
            try:
                # Dashboard Overview Tab
                if active_tab == 'dashboard':
                    try:
                        dashboard_data = DataRetrieval.get_dashboard_summary(sDate, eDate)
                        output += UIRenderer.render_dashboard_overview(dashboard_data)
                    except Exception as e:
                        output += ErrorHandler.handle_error(e, "dashboard overview")
                
                # Email Stats Tab
                elif active_tab == 'email':
                    try:
                        # Get and display email summary data
                        email_summary = DataRetrieval.get_email_summary(sDate, eDate, hide_success, program_filter, failure_filter)
                        output += UIRenderer.render_email_summary(email_summary)
                        
                        # Add subscriber growth chart
                        try:
                            # Determine interval based on date range
                            date_range_days = (datetime.strptime(eDate, '%Y-%m-%d') - datetime.strptime(sDate, '%Y-%m-%d')).days
                            if date_range_days <= 31:
                                interval = 'daily'
                            elif date_range_days <= 90:
                                interval = 'weekly'
                            else:
                                interval = 'monthly'
                            
                            growth_data = DataRetrieval.get_subscriber_growth(sDate, eDate, interval)
                            output += UIRenderer.render_subscriber_growth_chart(growth_data)
                        except Exception as eg:
                            output += ErrorHandler.handle_error(eg, "subscriber growth chart")
                        
                        # Recent campaigns (full width for the enhanced table with metrics)
                        try:
                            # Check if we should exclude single-recipient emails
                            show_single_initial = getattr(model.Data, 'show_single_recipient', 'false') == 'true'
                            
                            # Get campaigns with appropriate filtering (no pagination)
                            campaigns = DataRetrieval.get_recent_campaigns(sDate, eDate, program_filter, exclude_single_recipient=not show_single_initial)
                            output += UIRenderer.render_recent_campaigns(campaigns)
                        except Exception as e2:
                            output += ErrorHandler.handle_error(e2, "recent campaigns")
                        
                        # Failure breakdown (full width below campaigns)
                        try:
                            output += UIRenderer.render_failure_breakdown(email_summary.get('failed_types', []))
                        except Exception as e1:
                            output += ErrorHandler.handle_error(e1, "failure breakdown")
                        
                        # Failed recipients (full width)
                        try:
                            failed_recipients = DataRetrieval.get_failure_recipients(sDate, eDate, hide_success, program_filter, failure_filter)
                            output += UIRenderer.render_failed_recipients(failed_recipients)
                        except Exception as e3:
                            output += ErrorHandler.handle_error(e3, "failed recipients")
                    except Exception as e:
                        output += ErrorHandler.handle_error(e, "email statistics")
                
                # SMS Stats Tab
                elif active_tab == 'sms':
                    try:
                        # Get and display SMS summary data
                        sms_stats = DataRetrieval.get_sms_stats(sDate, eDate)
                        output += UIRenderer.render_sms_summary(sms_stats)
                        
                        # Layout in two columns for smaller sections
                        output += '<div class="row">'
                        
                        # Left column - SMS failure breakdown
                        output += '<div class="col-md-6">'
                        try:
                            # Add SMS failure breakdown - FIXED
                            sms_failures = DataRetrieval.get_sms_failure_breakdown(sDate, eDate)
                            output += UIRenderer.render_sms_failure_breakdown(sms_failures)
                        except Exception as e1:
                            output += ErrorHandler.handle_error(e1, "SMS failure breakdown")
                        output += '</div>'
                        
                        # Right column - SMS top senders
                        output += '<div class="col-md-6">'
                        try:
                            # Add SMS top senders - FIXED
                            sms_top_senders = DataRetrieval.get_sms_top_senders(sDate, eDate)
                            output += UIRenderer.render_sms_top_senders(sms_top_senders)
                        except Exception as e2:
                            output += ErrorHandler.handle_error(e2, "SMS top senders")
                        output += '</div>'
                        
                        output += '</div>'  # End row
                        
                        # Get and display SMS campaigns
                        try:
                            sms_campaigns = DataRetrieval.get_sms_campaigns(sDate, eDate)
                            output += UIRenderer.render_sms_campaigns(sms_campaigns)
                        except Exception as e3:
                            output += ErrorHandler.handle_error(e3, "SMS campaigns")
                    except Exception as e:
                        output += ErrorHandler.handle_error(e, "SMS statistics")
                
                # Top Senders Tab
                elif active_tab == 'senders':
                    try:
                        # Get and display top senders
                        active_senders = DataRetrieval.get_active_senders(sDate, eDate, program_filter)
                        output += UIRenderer.render_active_senders(active_senders)
                        
                        # Add SMS top senders as well
                        try:
                            sms_top_senders = DataRetrieval.get_sms_top_senders(sDate, eDate)
                            output += UIRenderer.render_sms_top_senders(sms_top_senders)
                        except Exception as e1:
                            output += ErrorHandler.handle_error(e1, "SMS top senders")
                    except Exception as e:
                        output += ErrorHandler.handle_error(e, "top senders")
                    
                # Program Stats Tab
                elif active_tab == 'programs':
                    try:
                        # Get and display program stats
                        program_stats = DataRetrieval.get_program_email_stats(sDate, eDate)
                        output += UIRenderer.render_program_stats(program_stats)
                    except Exception as e:
                        output += ErrorHandler.handle_error(e, "program statistics")
            except Exception as tab_error:
                # If a tab completely fails, show a user-friendly error
                output += """
                <div class="alert alert-danger">
                    <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Tab Error</h4>
                    <p>An error occurred while displaying this tab. Please try another tab or refresh the page.</p>
                    <p>Error details: {0}</p>
                </div>
                """.format(str(tab_error))
            
            output += '</div>'  # End tab content
            output += '</div>'  # End dashboard container
            
            return output
        except Exception as general_error:
            # Error handling for general issues
            return ErrorHandler.handle_error(general_error, "dashboard")
            
    except Exception as critical_error:
        # Last resort error handling if something goes wrong at the top level
        return """
        <div class="alert alert-danger">
            <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Critical Error</h4>
            <p>A critical error occurred while rendering the dashboard. Please contact support.</p>
            <pre style="max-height: 200px; overflow-y: auto;">{0}</pre>
        </div>
        """.format(traceback.format_exc())

# Run the main function and print the output
try:
    print main()
except Exception as e:
    # Last resort error handling if something goes wrong at the top level
    print """
    <div class="alert alert-danger">
        <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Critical Error</h4>
        <p>A critical error occurred while rendering the dashboard. Please contact support.</p>
        <pre style="max-height: 200px; overflow-y: auto;">{0}</pre>
    </div>
    """.format(traceback.format_exc())
# FormatHelper.format_number(total_sms),
# FormatHelper.format_percent(sms_delivery_rate),
# min(100, max(0, sms_delivery_rate)),
# int(sms_delivery_rate),
# int(sms_delivery_rate),
# model.Data.sDate, model.Data.eDate
