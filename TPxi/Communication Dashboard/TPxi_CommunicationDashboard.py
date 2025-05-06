#####################################################################
#### COMMUNICATION DASHBOARD
#####################################################################
# This dashboard provides comprehensive communication analytics including:
# - Email delivery metrics with success/failure breakdown
# - SMS/text message statistics with delivery performance
# - Top senders analytics with delivery rates
# - Program-specific communication metrics
#
# Features:
# - Modular design for easier maintenance
# - Better error handling with user-friendly messages
# - Loading indicators for long-running operations
# - Configurable options for customization
# - Fixed SQL queries with proper column references
# - Improved UI with responsive design
# - Complete SMS analytics with sender information and failure breakdowns
#
#####################################################################
#### UPLOAD INSTRUCTIONS
#####################################################################
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python script "CommunicationDashboard" and paste all this code
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
    MAX_ROWS_PER_TABLE = 20
    
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
            
            # Return final stats with guaranteed values
            return {
                'total_emails': total_emails,
                'sent_emails': sent_emails,
                'failed_emails': failed_emails,
                'delivery_rate': delivery_rate,
                'failed_types': failed_types
            }
        except Exception as e:
            # Log error but return a valid structure with defaults
            print("<div class='alert alert-danger'>Error in email summary: {0}</div>".format(str(e)))
            return {
                'total_emails': 0,
                'sent_emails': 0,
                'failed_emails': 0,
                'delivery_rate': 0,
                'failed_types': []
            }
            
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
        GROUP BY p.Name, p.EmailAddress, fe.Fail
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
                    'Name': getattr(row, 'Name', ''),
                    'EmailAddress': getattr(row, 'EmailAddress', ''),
                    'FailureType': getattr(row, 'FailureType', ''),
                    'FailureCount': getattr(row, 'FailureCount', 0)
                })
            
            return data
        except Exception as e:
            raise Exception("Error retrieving failure recipients: {}".format(str(e)))

    @staticmethod
    def get_recent_campaigns(sDate, eDate, program_filter):
        """Get recent email campaigns"""
        # Format filter conditions
        filter_program = '' if program_filter == str(999999) else ' AND pro.Id = {}'.format(program_filter)
        
        # SQL for basic campaign data - using 'Sent' column instead of 'SendingDate'
        sql = """
        SELECT TOP {3}
            eq.Subject AS CampaignName,
            COUNT(eqt.PeopleId) AS RecipientCount,
            MAX(CAST(eqt.Sent AS DATE)) AS SentDate,
            eq.FromName AS Sender
        FROM EmailQueue eq
        JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
        LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
        LEFT JOIN Division d ON d.Id = o.DivisionId
        LEFT JOIN Program pro ON pro.Id = d.ProgId
        WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2}
        GROUP BY eq.Subject, eq.FromName
        ORDER BY MAX(eqt.Sent) DESC
        """
        
        try:
            # Format SQL with parameters
            formatted_sql = sql.format(sDate, eDate, filter_program, Config.MAX_ROWS_PER_TABLE)
            results = q.QuerySql(formatted_sql)
            
            # Convert results to list of dictionaries
            data = []
            for row in results:
                data.append({
                    'CampaignName': getattr(row, 'CampaignName', ''),
                    'RecipientCount': FormatHelper.format_number(getattr(row, 'RecipientCount', 0)),
                    'SentDate': getattr(row, 'SentDate', ''),
                    'Sender': getattr(row, 'Sender', '')
                })
            
            return data
        except Exception as e:
            raise Exception("Error retrieving campaign data: {}".format(str(e)))

    @staticmethod
    def get_active_senders(sDate, eDate, program_filter):
        """Get most active email senders"""
        # Format filter conditions
        filter_program = '' if program_filter == str(999999) else ' AND pro.Id = {}'.format(program_filter)
        
        # SQL for top senders
        sql = """
        SELECT TOP {3}
            eq.FromName AS SenderName,
            COUNT(DISTINCT eq.Id) AS CampaignCount,
            COUNT(eqt.PeopleId) AS TotalRecipients,
            SUM(CASE WHEN fe.Id IS NULL THEN 1 ELSE 0 END) AS SuccessfulDeliveries
        FROM EmailQueue eq
        JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
        LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
        LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
        LEFT JOIN Division d ON d.Id = o.DivisionId
        LEFT JOIN Program pro ON pro.Id = d.ProgId
        WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2}
        GROUP BY eq.FromName
        ORDER BY COUNT(DISTINCT eq.Id) DESC
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
                delivery_rate = FormatHelper.safe_div(successful_deliveries, total_recipients) * 100
                
                data.append({
                    'SenderName': getattr(row, 'SenderName', ''),
                    'CampaignCount': getattr(row, 'CampaignCount', 0),
                    'TotalRecipients': FormatHelper.format_number(total_recipients),
                    'SuccessfulDeliveries': FormatHelper.format_number(successful_deliveries),
                    'DeliveryRate': FormatHelper.format_percent(delivery_rate)
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
        """Render email summary metrics section"""
        
        # Extract values from summary data with defaults
        total_emails = summary_data.get('total_emails', 0)
        sent_emails = summary_data.get('sent_emails', 0)
        failed_emails = summary_data.get('failed_emails', 0)
        delivery_rate = summary_data.get('delivery_rate', 0)
        
        # Create metrics HTML
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h3 class="panel-title">Email Delivery Summary</h3>
            </div>
            <div class="panel-body">
                <div class="row text-center">
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{0}</h2>
                            <p>Total Emails Sent</p>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{1}</h2>
                            <p>Successfully Delivered</p>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{2}</h2>
                            <p>Failed Emails</p>
                        </div>
                    </div>
                    <div class="col-md-3 col-sm-6">
                        <div class="well well-sm">
                            <h2>{3}</h2>
                            <p>Delivery Rate</p>
                        </div>
                    </div>
                </div>
                
                <div class="progress">
                    <div class="progress-bar progress-bar-success" role="progressbar" style="width: {4}%;" 
                        aria-valuenow="{5}" aria-valuemin="0" aria-valuemax="100">
                        {6}%
                    </div>
                </div>
                <p class="text-center">Email Delivery Performance</p>
            </div>
        </div>
        """.format(
            FormatHelper.format_number(total_emails),
            FormatHelper.format_number(sent_emails),
            FormatHelper.format_number(failed_emails),
            FormatHelper.format_percent(delivery_rate),
            min(100, max(0, delivery_rate)),
            int(delivery_rate),
            int(delivery_rate)
        )
        
        return html

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
    def render_recent_campaigns(campaigns):
        """Render recent campaigns table"""
        
        if not campaigns:
            return '<div class="alert alert-info">No campaigns found in the selected time period.</div>'
        
        # Define column headers and labels
        header_labels = {
            'CampaignName': 'Campaign Name',
            'Sender': 'Sender',
            'SentDate': 'Date Sent',
            'RecipientCount': 'Recipients'
        }
        
        # Create HTML for campaigns table
        html = """
        <div class="panel panel-info">
            <div class="panel-heading">
                <h3 class="panel-title">Recent Email Campaigns</h3>
            </div>
            <div class="panel-body">
        """
        
        html += TableGenerator.generate_html_table(
            campaigns,
            column_order=['SentDate', 'CampaignName', 'Sender', 'RecipientCount'],
            header_labels=header_labels,
            id="recent-campaigns-table"
        )
        
        html += """
            </div>
        </div>
        """
        
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
            'DeliveryRate': 'Delivery Rate'
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
            column_order=['SenderName', 'CampaignCount', 'TotalRecipients', 'SuccessfulDeliveries', 'DeliveryRate'],
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
        """Render recipients with failed emails"""
        
        if not recipients:
            return '<div class="alert alert-info">No recipients with failed emails found in the selected time period.</div>'
        
        # Define column headers and labels
        header_labels = {
            'Name': 'Recipient Name',
            'EmailAddress': 'Email Address',
            'FailureType': 'Failure Type',
            'FailureCount': 'Failure Count'
        }
        
        # Create HTML for recipients table
        html = """
        <div class="panel panel-danger">
            <div class="panel-heading">
                <h3 class="panel-title">Recipients with Failed Emails</h3>
            </div>
            <div class="panel-body">
        """
        
        html += TableGenerator.generate_html_table(
            recipients,
            column_order=['Name', 'EmailAddress', 'FailureType', 'FailureCount'],
            header_labels=header_labels,
            id="failed-recipients-table"
        )
        
        html += """
            </div>
        </div>
        """
        
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
        <!-- Bootstrap CSS -->
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.4.1/css/bootstrap.min.css">
        
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
                        
                        # Layout in two columns for smaller sections
                        output += '<div class="row">'
                        
                        # Left column
                        output += '<div class="col-md-6">'
                        try:
                            output += UIRenderer.render_failure_breakdown(email_summary.get('failed_types', []))
                        except Exception as e1:
                            output += ErrorHandler.handle_error(e1, "failure breakdown")
                        output += '</div>'
                        
                        # Right column
                        output += '<div class="col-md-6">'
                        try:
                            campaigns = DataRetrieval.get_recent_campaigns(sDate, eDate, program_filter)
                            output += UIRenderer.render_recent_campaigns(campaigns)
                        except Exception as e2:
                            output += ErrorHandler.handle_error(e2, "recent campaigns")
                        output += '</div>'
                        
                        output += '</div>'  # End row
                        
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
        <pre style="max-height: 200px; overflow-y: auto;">{}</pre>
    </div>
    """.format(traceback.format_exc())
# FormatHelper.format_number(total_sms),
# FormatHelper.format_percent(sms_delivery_rate),
# min(100, max(0, sms_delivery_rate)),
# int(sms_delivery_rate),
# int(sms_delivery_rate),
# model.Data.sDate, model.Data.eDate
