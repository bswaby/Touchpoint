# ::START:: Staff Anniversaries Widget with Email Support
# This widget displays upcoming staff birthdays, anniversaries, and work anniversaries 
# with navigation and email functionality. Designed to work as a homepage widget.
#
# --Upload Instructions Start--
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python Scripts
# 2. Click New Python Script File
# 3. Choose your script name (example: "TPxi_AnniversaryWidget")
# 4. ⭐ IMPORTANT: Update the script_name in the AnniversaryConfig class below to match exactly
# 5. Paste all this code
# 6. Add the word "widget" to the content keywords by the script name
# 7. Test the script standalone first: /PyScript/YourScriptName
# 8. Add as homepage widget: Admin ~ Advanced ~ Homepage Widget
# --Upload Instructions End--
#
# Written by: Ben Swaby (modified for improved widget compatibility)
# Email: bswaby@fbchtn.org

import hashlib
import time

# ::START:: Configuration Parameters
class AnniversaryConfig:
    """Configuration class for anniversary widget settings"""
    
    def __init__(self):
        # IMPORTANT: Set this to match your TouchPoint script name exactly
        self.script_name = 'TPxi_AnniversaryWidget'  # ⭐ CHANGE THIS TO YOUR SCRIPT NAME
        
        # Display settings
        self.title = 'Staff Anniversaries'
        self.days_to_look_out = 30  # Days to show initially
        self.data_range = 180  # Days to preload (6 months)
        self.saved_query = 'Dashboard_Birthday-Wedding'  # Name of saved query
        
        # Feature toggles
        self.show_birthdays = True
        self.show_weddings = True
        self.show_extra_value = True
        
        # Extra value field configuration
        self.extra_value_field = 'WorkAnniversary'
        self.extra_value_friendly_name = 'Work'
        
        # Email settings (fallback if user email unavailable)
        self.system_from_email = 'system@yourchurch.org'
        self.system_from_name = 'Church Family'
        
        # Widget appearance
        self.widget_border_color = '#4CAF50'
        self.widget_hover_color = '#45a049'

config = AnniversaryConfig()

# ::START:: Configuration Validation
def validate_configuration():
    """Validate that the configuration is set up correctly"""
    errors = []
    warnings = []
    
    # Check script name configuration
    if config.script_name == 'TPxi_AnniversaryWidget':
        warnings.append("Script name is still set to the example value. Consider updating it to match your actual script name.")
    
    if not config.script_name or len(config.script_name.strip()) == 0:
        errors.append("Script name cannot be empty. Please set config.script_name to your TouchPoint script name.")
    
    # Check query configuration
    if not config.saved_query:
        errors.append("Saved query name is required. Please set config.saved_query to your TouchPoint saved query name.")
    
    # Check email configuration
    if config.system_from_email == 'system@yourchurch.org':
        warnings.append("System email is still set to example value. Consider updating config.system_from_email.")
    
    return errors, warnings

# ::START:: User Authentication and Info Retrieval
class UserManager:
    """Handles user authentication and information retrieval"""
    
    @staticmethod
    def get_user_info():
        """Get logged-in user's name and email with fallback"""
        try:
            user_info = q.QuerySqlTop1(
                "SELECT p.Name, p.EmailAddress FROM People p WHERE p.PeopleId = {}".format(
                    model.UserPeopleId
                )
            )
            
            if user_info and hasattr(user_info, 'Name') and hasattr(user_info, 'EmailAddress'):
                return {
                    'name': user_info.Name or config.system_from_name,
                    'email': user_info.EmailAddress or config.system_from_email
                }
            else:
                return {
                    'name': config.system_from_name,
                    'email': config.system_from_email
                }
        except Exception:
            return {
                'name': config.system_from_name,
                'email': config.system_from_email
            }

# ::START:: Anniversary Data Management
class AnniversaryDataManager:
    """Manages anniversary data queries and processing"""
    
    @staticmethod
    def build_sql_query():
        """Build the SQL query based on configuration"""
        base_sql = '''
WITH weddingDate AS (
    SELECT DISTINCT 
        p.PeopleId,
        Name,
        'Wedding'  + ' (' + CAST(DATEDIFF(year, p.WeddingDate, GETDATE()) AS VARCHAR) + ')' AS [Type],
        FORMAT(p.WeddingDate, 'MM/dd') AS [dDate],
        p.WeddingDate AS bDate,
        DATEADD(year, DATEPART(year, GETDATE()) - DATEPART(year, p.WeddingDate), p.WeddingDate) AS [ThisYearDate],
        p.EmailAddress AS Email,
        DATEDIFF(year, p.WeddingDate, GETDATE()) AS YearsCount
    FROM People p
    WHERE p.PeopleId IN ({0}) 
    AND DATEADD(year, DATEPART(year, GETDATE()) - DATEPART(year, p.WeddingDate), p.WeddingDate) 
        BETWEEN CONVERT(datetime, DATEADD(day, -{1}, GETDATE()), 101) 
        AND CONVERT(datetime, DATEADD(day, {1}, GETDATE()), 101)
),
bDay AS (
    SELECT DISTINCT
        p.PeopleId,
        Name,
        FORMAT(p.BDate, 'MM/dd') AS [dDate],
        'Birthday' AS [Type],
        p.BDate AS bDate,
        DATEADD(year, DATEPART(year, GETDATE()) - DATEPART(year, p.BDate), p.BDate) AS [ThisYearDate],
        p.EmailAddress AS Email,
        0 AS YearsCount
    FROM People p
    WHERE p.PeopleId IN ({0})
    AND DATEADD(year, DATEPART(year, GETDATE()) - DATEPART(year, p.BDate), p.BDate) 
        BETWEEN CONVERT(datetime, DATEADD(day, -{1}, GETDATE()), 101) 
        AND CONVERT(datetime, DATEADD(day, {1}, GETDATE()), 101)
)

SELECT * FROM (
'''
        
        # Build union parts based on configuration
        union_parts = []
        
        # Add extra value field query if enabled
        if config.show_extra_value and config.extra_value_field:
            union_part = '''
    SELECT DISTINCT
        p.PeopleId,
        Name,
        FORMAT(pe.DateValue, 'MM/dd') AS [dDate],
        '{2}' + ' (' + CAST(DATEDIFF(year, pe.DateValue, GETDATE()) AS VARCHAR) + ')' AS [Type],
        pe.DateValue AS bDate,
        DATEADD(year, DATEPART(year, GETDATE()) - DATEPART(year, pe.DateValue), pe.DateValue) AS [ThisYearDate],
        p.EmailAddress AS Email,
        DATEDIFF(year, pe.DateValue, GETDATE()) AS YearsCount
    FROM People p
    INNER JOIN PeopleExtra pe ON pe.PeopleId = p.PeopleId AND pe.Field = '{3}'
    WHERE p.PeopleId IN ({0})
    AND DATEADD(year, DATEPART(year, GETDATE()) - DATEPART(year, pe.DateValue), pe.DateValue) 
        BETWEEN CONVERT(datetime, DATEADD(day, -{1}, GETDATE()), 101)
        AND CONVERT(datetime, DATEADD(day, {1}, GETDATE()), 101)
'''
            union_parts.append(union_part)
        
        # Add wedding date query if enabled
        if config.show_weddings:
            union_parts.append('''
    SELECT
        wd.PeopleId,
        wd.Name,
        wd.dDate,
        wd.Type,
        wd.bDate,
        wd.ThisYearDate,
        wd.Email,
        wd.YearsCount
    FROM WeddingDate wd
''')
        
        # Add birthday query if enabled
        if config.show_birthdays:
            union_parts.append('''
    SELECT
        bd.PeopleId,
        bd.Name,
        bd.dDate,
        bd.Type,
        bd.bDate,
        bd.ThisYearDate,
        bd.Email,
        bd.YearsCount
    FROM bDay bd
''')
        
        # Join all enabled parts with UNION ALL
        if union_parts:
            base_sql += "\n    UNION ALL\n    ".join(union_parts)
        else:
            # Default empty query if nothing is enabled
            base_sql += "SELECT 1 AS PeopleId, '' AS Name, '' AS dDate, '' AS Type, GETDATE() AS bDate, GETDATE() AS ThisYearDate, '' AS Email, 0 AS YearsCount WHERE 1=0"
        
        # Complete the SQL query
        base_sql += '''
) AS CombinedResults
ORDER BY MONTH(bDate), DAY(bDate)
'''
        
        return base_sql
    
    @staticmethod
    def get_people_list():
        """Get list of people from saved query"""
        try:
            people_list = []
            for p in q.QueryList(config.saved_query, "PeopleId"):
                people_list.append(str(p.PeopleId))
            return ','.join(people_list) if people_list else '0'
        except Exception as e:
            raise Exception("Error getting people list: " + str(e))

# ::START:: Email Handler
class EmailHandler:
    """Handles email sending functionality with improved widget compatibility"""
    
    @staticmethod
    def process_email_request():
        """Process email request with improved error handling"""
        try:
            # Get form data safely
            logged_in_user = model.UserPeopleId
            
            # Access form data with error checking
            try:
                pid = int(str(getattr(model.Data, 'email_pid', '')))
            except (ValueError, AttributeError):
                raise Exception("Invalid person ID provided")
            
            try:
                subject = str(getattr(model.Data, 'email_subject', ''))
                message = str(getattr(model.Data, 'email_message', ''))
                from_email = str(getattr(model.Data, 'email_from', ''))
                from_name = str(getattr(model.Data, 'email_fromname', ''))
            except AttributeError as e:
                raise Exception("Missing required email field: " + str(e))
            
            if not all([subject, message, from_email, from_name]):
                raise Exception("All email fields are required")
            
            # Send the email using TouchPoint's email function
            model.Email(pid, logged_in_user, from_email, from_name, subject, message)
            
            # Return success response
            return {
                'success': True,
                'message': 'Email sent successfully to ' + str(pid)
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': 'Error sending email: ' + str(e)
            }

# ::START:: Widget Renderer
class WidgetRenderer:
    """Renders the widget HTML and JavaScript"""
    
    def __init__(self, widget_id, user_info):
        self.widget_id = widget_id
        self.user_info = user_info
    
    def render_styles(self):
        """Render CSS styles for the widget"""
        return '''
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js"></script>

<style>
.anniversary-widget {
  background-color: White;
  border: 1px solid ''' + config.widget_border_color + ''';
  border-radius: 5px;
  padding: 3px;
  margin: 0px;
  margin-top: 4px;
  box-shadow: 0 2px 5px rgba(0,0,0,0.1);
}
.smallpadding {
  padding: 3px;
  margin: 1px;
}
.email-button {
  color: ''' + config.widget_border_color + ''';
  cursor: pointer;
  border: none;
  background: none;
  font-size: 16px;
  transition: all 0.3s;
  margin-left: 8px;
}
.email-button:hover {
  color: ''' + config.widget_hover_color + ''';
  transform: scale(1.2);
}
.anniversary-row {
  padding: 2px 10px;
  border-bottom: 1px solid #f0f0f0;
  display: flex;
  align-items: center;
  transition: background-color 0.2s;
}
.anniversary-row:hover {
  background-color: #f9f9f9;
}
.today-highlight {
  background-color: #e8f5e9;
  border-left: 3px solid ''' + config.widget_border_color + ''';
}
.anniversary-date {
  font-weight: bold;
  margin-right: 8px;
  min-width: 45px;
}
.anniversary-name {
  flex-grow: 1;
}
.anniversary-type {
  margin: 0 10px;
  color: #666;
  min-width: 100px;
  text-align: right;
}
.navigation-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.nav-arrow {
  cursor: pointer;
  font-size: 24px;
  color: ''' + config.widget_border_color + ''';
  padding: 0 10px;
  transition: transform 0.2s;
}
.nav-arrow:hover {
  transform: scale(1.2);
}
.date-range {
  font-size: 14px;
  color: #666;
  font-style: italic;
}
.center-title {
  flex-grow: 1;
  text-align: center;
}
.loading-indicator {
  text-align: center;
  padding: 20px;
  color: #666;
}
</style>
'''
    
    def render_widget_html(self):
        """Render the main widget HTML structure"""
        return '''
<div id="''' + self.widget_id + '''" class="anniversary-widget">
  <div class="navigation-header smallpadding">
    <span class="nav-arrow" onclick="''' + self.widget_id + '''_navigate(-1)" title="Previous ''' + str(config.days_to_look_out) + ''' days">&lt;</span>
    <h3 class="center-title">''' + config.title + '''
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
        <text x="100" y="120" font-family="Arial, sans-serif" font-weight="bold" font-size="60" fill="#333333">TP</text>
        <g transform="translate(190, 107)">
          <circle cx="0" cy="0" r="13.5" fill="#0099FF"/>
          <circle cx="0" cy="0" r="10.5" fill="white"/>
          <circle cx="0" cy="0" r="7.5" fill="#0099FF"/>
          <path d="M-9 -9 L9 9 M-9 9 L9 -9" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
        </g>
        <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">i</text>
      </svg>
    </h3>
    <span class="nav-arrow" onclick="''' + self.widget_id + '''_navigate(1)" title="Next ''' + str(config.days_to_look_out) + ''' days">&gt;</span>
  </div>
  <div id="''' + self.widget_id + '''_range" class="date-range" style="text-align: center; display:none;"></div>
  <hr class="smallpadding">
  
  <div id="''' + self.widget_id + '''_container">
    <div class="loading-indicator">Loading anniversaries...</div>
  </div>
</div>
'''
    
    def render_javascript_start(self):
        """Render the beginning of JavaScript code"""
        return '''
<script>
(function() {
    window.''' + self.widget_id + ''' = {
        daysToShow: ''' + str(config.days_to_look_out) + ''',
        currentOffset: 0,
        allAnniversaries: [],
        
        init: function() {
            // Detect execution context and use configured script name
            this.isWidget = window.location.pathname === '/' || window.location.pathname.includes('Default');
            this.configuredScriptName = ''' + repr(config.script_name) + ''';
            this.detectedScriptName = window.location.pathname.split('/').pop();
            
            this.loadData();
            this.updateOffsetDisplay();
            this.displayVisibleAnniversaries();
            document.getElementById(''' + repr(self.widget_id) + ''').setAttribute("data-initialized", "true");
        },
        
        loadData: function() {
'''
    
    def render_anniversary_data(self, data):
        """Render anniversary data as JavaScript objects"""
        js_data = ""
        if data:
            for item in data:
                # Safely escape strings for JavaScript
                name_safe = item.Name.replace("'", "\\'").replace('"', '\\"').replace('\n', ' ').replace('\r', ' ')
                type_safe = item.Type.replace("'", "\\'").replace('"', '\\"').replace('\n', ' ').replace('\r', ' ')
                date_safe = item.dDate
                email_safe = ""
                if hasattr(item, 'Email') and item.Email:
                    email_safe = item.Email.replace("'", "\\'").replace('"', '\\"')
                
                js_data += '''            this.allAnniversaries.push({
                pid: ''' + str(item.PeopleId) + ''',
                name: "''' + name_safe + '''",
                date: "''' + date_safe + '''",
                type: "''' + type_safe + '''",
                yearsCount: ''' + str(item.YearsCount) + ''',
                email: "''' + email_safe + '''"
            });
'''
        return js_data
    
    def render_javascript_functions(self):
        """Render JavaScript functions for widget functionality"""
        # Pre-build the widget ID strings to avoid JavaScript concatenation issues
        widget_container_id = self.widget_id + '_container'
        widget_range_id = self.widget_id + '_range'
        
        return '''        },
        
        parseAnniversaryDate: function(dateStr) {
            const parts = dateStr.split('/');
            if (parts.length !== 2) return null;
            
            const month = parseInt(parts[0]) - 1;
            const day = parseInt(parts[1]);
            
            const date = new Date();
            date.setMonth(month);
            date.setDate(day);
            
            return date;
        },
        
        navigate: function(direction) {
            this.currentOffset += direction * this.daysToShow;
            this.updateOffsetDisplay();
            this.displayVisibleAnniversaries();
        },
        
        sendEmail: function(pid, celebrationType, personName, yearsCount, personEmail) {
            let subject = "";
            let emailBody = "";
            
            if (celebrationType.includes("Birthday")) {
                subject = "Happy Birthday " + personName + "!";
                emailBody = "Happy Birthday " + personName + "!\\n\\nWishing you a wonderful birthday celebration!\\n\\n";
            } else if (celebrationType.includes("Wedding")) {
                subject = "Happy " + yearsCount + (yearsCount === 1 ? " Year" : " Years") + " Wedding Anniversary " + personName + "!";
                emailBody = "Congratulations on " + yearsCount + (yearsCount === 1 ? " year" : " years") + " of marriage, " + personName + "!\\n\\nMay your love continue to grow stronger with each passing year.\\n\\n";
            } else if (celebrationType.includes("Work")) {
                subject = "Happy " + yearsCount + (yearsCount === 1 ? " Year" : " Years") + " Work Anniversary " + personName + "!";
                emailBody = "Congratulations on " + yearsCount + (yearsCount === 1 ? " year" : " years") + " with us, " + personName + "!\\n\\nThank you for your dedication and all your contributions.\\n\\n";
            }
            
            emailBody += "Best wishes,\\n" + ''' + repr(self.user_info['name']) + ''';
            
            Swal.fire({
                title: "Send Celebration Email",
                html: '<div style="text-align: left; margin-bottom: 15px;"><strong>To:</strong> ' + personName + (personEmail ? ' (' + personEmail + ')' : '') + '<br><strong>Subject:</strong> ' + subject + '</div><textarea id="swal-email-message" class="swal2-textarea" style="height: 200px;">' + emailBody + '</textarea>',
                showCancelButton: true,
                confirmButtonText: 'Send Email',
                confirmButtonColor: ''' + repr(config.widget_border_color) + ''',
                cancelButtonColor: '#d33',
                showLoaderOnConfirm: true,
                preConfirm: () => {
                    const message = document.getElementById('swal-email-message').value;
                    if (!message.trim()) {
                        Swal.showValidationMessage('Please enter a message');
                        return false;
                    }
                    
                    // Create a more reliable form submission approach
                    const formData = new FormData();
                    formData.append('a', 'email');
                    formData.append('email_pid', pid);
                    formData.append('email_subject', subject);
                    formData.append('email_message', message);
                    formData.append('email_from', ''' + repr(self.user_info['email']) + ''');
                    formData.append('email_fromname', ''' + repr(self.user_info['name']) + ''');
                    
                    // Try multiple submission strategies
                    return this.submitEmailForm(formData);
                }
            }).then((result) => {
                if (result.isConfirmed) {
                    Swal.fire({
                        title: 'Email Sent!',
                        text: 'Your celebration email has been sent successfully.',
                        icon: 'success',
                        timer: 2000
                    });
                }
            });
        },
        
        submitEmailForm: function(formData) {
            // Different strategies based on execution context
            const urlParams = new URLSearchParams(formData);
            
            let submissionUrls = [];
            
            if (this.isWidget) {
                // Widget context: homepage or dashboard
                submissionUrls = [
                    '/PyScriptForm/' + this.configuredScriptName,  // Primary: use configured name
                    '/PyScriptForm/WidgetStaffAnniversaries',      // Fallback: legacy name
                    '/PyScriptForm/' + this.detectedScriptName     // Last resort: detected name
                ];
            } else {
                // Standalone context: direct script URL
                let currentPath = window.location.pathname;
                submissionUrls = [
                    currentPath.replace('/PyScript/', '/PyScriptForm/'),  // Primary: current URL
                    '/PyScriptForm/' + this.configuredScriptName,         // Fallback: configured name
                    '/PyScriptForm/' + this.detectedScriptName,           // Fallback: detected name
                    '/PyScriptForm/WidgetStaffAnniversaries'              // Last resort: legacy name
                ];
            }
            
            // Try each URL in sequence
            return this.trySubmissionSequence(submissionUrls, urlParams);
        },
        
        trySubmissionSequence: function(urls, urlParams) {
            if (urls.length === 0) {
                return Promise.reject(new Error('No more URLs to try'));
            }
            
            const currentUrl = urls[0];
            const remainingUrls = urls.slice(1);
            
            return this.trySubmission(currentUrl, urlParams)
                .catch(error => {
                    console.warn('Failed to submit to:', currentUrl, error.message);
                    if (remainingUrls.length > 0) {
                        return this.trySubmissionSequence(remainingUrls, urlParams);
                    } else {
                        throw new Error('All submission attempts failed. Last error: ' + error.message);
                    }
                });
        },
        
        trySubmission: function(url, urlParams) {
            return fetch(url, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: urlParams
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error('HTTP ' + response.status + ': ' + response.statusText);
                }
                return response.text();
            })
            .then(text => {
                // Check if response indicates success
                if (text.toLowerCase().includes('error') && !text.toLowerCase().includes('success')) {
                    throw new Error('Server returned error response');
                }
                return text;
            });
        },
        
                        updateOffsetDisplay: function() {
            const rangeDisplay = document.getElementById(''' + repr(widget_range_id) + ''');
            if (!rangeDisplay) return;
            
            if (this.currentOffset === 0) {
                rangeDisplay.style.display = 'none';
            } else {
                rangeDisplay.style.display = 'block';
                if (this.currentOffset < 0) {
                    rangeDisplay.textContent = '(Viewing ' + Math.abs(this.currentOffset) + ' days ago)';
                } else {
                    rangeDisplay.textContent = '(Viewing ' + this.currentOffset + ' days ahead)';
                }
            }
        },
        
        isDateInRange: function(dateObj) {
            const today = new Date();
            const startDate = new Date(today);
            const endDate = new Date(today);
            
            startDate.setDate(startDate.getDate() + this.currentOffset);
            endDate.setDate(endDate.getDate() + this.currentOffset + this.daysToShow);
            
            return dateObj >= startDate && dateObj <= endDate;
        },
        
        displayVisibleAnniversaries: function() {
            const container = document.getElementById(''' + repr(widget_container_id) + ''');
            if (!container) return;
            
            container.innerHTML = '';
            let visibleCount = 0;
            
            const today = new Date();
            const todayFormatted = (today.getMonth() + 1).toString().padStart(2, '0') + '/' + 
                                today.getDate().toString().padStart(2, '0');
            
            for (const anniv of this.allAnniversaries) {
                const anniversaryDate = this.parseAnniversaryDate(anniv.date);
                if (!anniversaryDate) continue;
                
                if (this.isDateInRange(anniversaryDate)) {
                    visibleCount++;
                    
                    const row = document.createElement('div');
                    row.className = 'anniversary-row';
                    
                    if (anniv.date === todayFormatted) {
                        row.className += ' today-highlight';
                    }
                    
                    const safeType = anniv.type.replace(/'/g, "\\'").replace(/"/g, '\\"');
                    const safeName = anniv.name.replace(/'/g, "\\'").replace(/"/g, '\\"');
                    
                    row.innerHTML = 
                        '<span class="anniversary-date">' + anniv.date + '</span>' +
                        '<span class="anniversary-name"><a href="/Person2/' + anniv.pid + '#tab-communications">' + anniv.name + '</a></span>' +
                        '<span class="anniversary-type">' + anniv.type + '</span>' +
                        '<button class="email-button" title="Send celebration email" data-pid="' + anniv.pid + '" data-type="' + safeType + '" data-name="' + safeName + '" data-years="' + anniv.yearsCount + '" data-email="' + anniv.email + '">✉️</button>';
                    
                    // Add click event listener to the button
                    const emailButton = row.querySelector('.email-button');
                    emailButton.addEventListener('click', function(e) {
                        e.preventDefault();
                        ''' + self.widget_id + '''.sendEmail(
                            parseInt(this.getAttribute('data-pid')),
                            this.getAttribute('data-type'),
                            this.getAttribute('data-name'),
                            parseInt(this.getAttribute('data-years')),
                            this.getAttribute('data-email')
                        );
                    });
                    
                    container.appendChild(row);
                }
            }
            
            if (visibleCount === 0) {
                const emptyMsg = document.createElement('p');
                emptyMsg.style.textAlign = 'center';
                emptyMsg.style.padding = '10px';
                emptyMsg.textContent = 'No anniversaries for this date range';
                container.appendChild(emptyMsg);
            }
        }
    };
    
    window.''' + self.widget_id + '''_navigate = function(direction) {
        ''' + self.widget_id + '''.navigate(direction);
    };
    
    function initWidget() {
        if (document.getElementById(''' + repr(self.widget_id) + ''')) {
            if (!document.getElementById(''' + repr(self.widget_id) + ''').getAttribute('data-initialized')) {
                ''' + self.widget_id + '''.init();
            }
        } else {
            setTimeout(initWidget, 100);
        }
    }
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initWidget);
    } else {
        initWidget();
    }
})();
</script>
'''

# ::START:: Main Controller
def main_controller():
    """Main controller that orchestrates the widget functionality"""
    
    # ::STEP:: Validate Configuration
    config_errors, config_warnings = validate_configuration()
    if config_errors:
        print '''
<div class="anniversary-widget" style="border-color: #d32f2f;">
    <h3 style="color: #d32f2f; text-align: center;">Configuration Error</h3>
    <div style="padding: 10px; background-color: #ffebee; border-radius: 3px; margin: 5px;">
        <p><strong>Please fix the following configuration issues:</strong></p>
        <ul>'''
        for error in config_errors:
            print '<li style="color: #d32f2f;">' + error + '</li>'
        print '''</ul>
        <p><em>Update the AnniversaryConfig class at the top of this script.</em></p>
    </div>
</div>
'''
        return
    
    # Show warnings if any (but continue execution)
    if config_warnings:
        print '''
<div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 3px; padding: 8px; margin: 5px 0;">
    <strong>Configuration Notice:</strong>
    <ul style="margin: 5px 0;">'''
        for warning in config_warnings:
            print '<li style="color: #856404;">' + warning + '</li>'
        print '''</ul>
</div>
'''
    
    # ::STEP:: Set widget header
    model.Header = config.title
    
    # ::STEP:: Generate unique widget ID
    widget_id = "sa_" + hashlib.md5(str(time.time())).hexdigest()[:8]
    
    try:
        # ::STEP:: Handle email POST requests
        if model.HttpMethod == 'post' and hasattr(model.Data, 'a') and str(getattr(model.Data, 'a', '')) == 'email':
            print ""  # Important: print blank line first for POST responses
            
            email_result = EmailHandler.process_email_request()
            if email_result['success']:
                print "Email sent successfully"
            else:
                print "<h4 style='color:red;'>Error</h4>"
                print "<p>" + email_result['message'] + "</p>"
            return
        
        # ::STEP:: Get user information
        user_info = UserManager.get_user_info()
        
        # ::STEP:: Initialize widget renderer
        renderer = WidgetRenderer(widget_id, user_info)
        
        # ::STEP:: Render styles and widget HTML
        print renderer.render_styles()
        print renderer.render_widget_html()
        print renderer.render_javascript_start()
        
        # ::STEP:: Get anniversary data
        data_manager = AnniversaryDataManager()
        people_list = data_manager.get_people_list()
        
        if people_list and people_list != '0':
            sql_query = data_manager.build_sql_query()
            anniversary_data = q.QuerySql(
                sql_query.format(
                    people_list, 
                    config.data_range, 
                    config.extra_value_friendly_name, 
                    config.extra_value_field
                )
            )
            
            # ::STEP:: Render anniversary data as JavaScript
            print renderer.render_anniversary_data(anniversary_data)
        
        # ::STEP:: Complete JavaScript rendering
        print renderer.render_javascript_functions()
        
    except Exception as e:
        # ::STEP:: Error handling with user-friendly display
        import traceback
        print '''
<div class="anniversary-widget" style="border-color: #d32f2f;">
    <h3 style="color: #d32f2f; text-align: center;">Error Loading Anniversaries</h3>
    <div style="padding: 10px; background-color: #ffebee; border-radius: 3px; margin: 5px;">
        <p><strong>Error:</strong> ''' + str(e) + '''</p>
        <details style="margin-top: 10px;">
            <summary style="cursor: pointer; color: #666;">Technical Details</summary>
            <pre style="font-size: 10px; color: #666; margin-top: 5px; white-space: pre-wrap;">'''
        traceback.print_exc()
        print '''</pre>
        </details>
    </div>
</div>
'''

# ::START:: Script Execution
# Execute main controller directly (Python 2.7.3 in TouchPoint doesn't support __name__ == "__main__")
main_controller()
