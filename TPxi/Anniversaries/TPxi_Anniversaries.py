#This is to show upcoming staff birthdays and anniversaries with navigation.  
#-- Steps --
#1. Copy this code to a python script (Admin ~ Advanced ~ Special Content ~ Python Scripts) 
#   and call it something like WidgetStaffBirthdays and make sure to add the word widget to 
#   the content keywords by the script name
#2. Update config parameters below
#3. Add as a home page widget (Admin ~ Advanced ~ Homepage Widget) by selecting the script, 
#   adding name, and setting permissions that can see it

# written by: Ben Swaby
# email: bswaby@fbchtn.org


#------------------------------------------------------
#config parameters
#------------------------------------------------------
title = '''Staff Anniversaries'''
daysToLookOut = '30' #set to how many days you want to look out initially
dataRange = '180' #set to how many days to load in each direction (pre-load 6 months)
savedQuery = 'Dashboard_Birthday-Wedding' #Name of saved query
extraValueField = 'WorkAnniversary' #add extra value field name if you want to pull in another date
extraValueFieldFriendlyName = 'Work' # just a friendly name of extra value

# System email will only be used if user email can't be retrieved
system_fromEmail = 'system@yourchurch.org' 
system_fromName = 'Church Family'

#------------------------------------------------------
# start of code
#------------------------------------------------------

model.Header = title

# Get logged-in user's name and email
try:
    userInfo = q.QuerySqlTop1("SELECT p.Name, p.EmailAddress FROM People p WHERE p.PeopleId = {}".format(model.UserPeopleId))
    fromName = userInfo.Name if userInfo and hasattr(userInfo, 'Name') else system_fromName
    fromEmail = userInfo.EmailAddress if userInfo and hasattr(userInfo, 'EmailAddress') else system_fromEmail
except Exception as e:
    # Fallback to system values if query fails
    fromName = system_fromName
    fromEmail = system_fromEmail

sql = '''
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
    
    UNION ALL
    
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
    
    UNION ALL
    
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
) AS CombinedResults
ORDER BY MONTH(bDate), DAY(bDate)
'''

# Generate a unique ID for this widget instance to avoid conflicts on the homepage
import hashlib
import time
widget_id = "sa_" + hashlib.md5(str(time.time())).hexdigest()[:8]

print '''
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js"></script>

<style>
.anniversary-widget {
  background-color: White;
  border: 1px solid #4CAF50;
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
  color: #4CAF50;
  cursor: pointer;
  border: none;
  background: none;
  font-size: 16px;
  transition: all 0.3s;
  margin-left: 8px;
}
.email-button:hover {
  color: #45a049;
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
  color: #4CAF50;
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
</style>

<div id="''' + widget_id + '''" class="anniversary-widget">
  <div class="navigation-header smallpadding">
    <span class="nav-arrow" onclick="''' + widget_id + '''_navigate(-1)" title="Previous 30 days">&lt;</span>
    <h3 class="center-title">''' + title + '''
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
      </svg>
    </h3>
    <span class="nav-arrow" onclick="''' + widget_id + '''_navigate(1)" title="Next 30 days">&gt;</span>
  </div>
  <div id="''' + widget_id + '''_range" class="date-range" style="text-align: center; display:none;"></div>
  <hr class="smallpadding">
  
  <!-- Container for anniversaries - populated by JavaScript -->
  <div id="''' + widget_id + '''_container"></div>
</div>

<script>
(function() {
    // Namespace all functions and variables to avoid conflicts
    window.''' + widget_id + ''' = {
        // Constants
        daysToShow: ''' + daysToLookOut + ''',
        
        // State
        currentOffset: 0,
        allAnniversaries: [],
        
        // Initialize the widget
        init: function() {
            this.loadData();
            this.updateOffsetDisplay();
            this.displayVisibleAnniversaries();
            
            // Add initialization complete flag
            document.getElementById("''' + widget_id + '''").setAttribute("data-initialized", "true");
        },
        
        // Load all anniversary data
        loadData: function() {
'''

# Handle POST request for sending emails
if model.HttpMethod == 'post' and model.Data.a == "email":  
    print " "  # Print a blank line first - IMPORTANT!

    try:
        # Get data from request
        LoggedInUser = model.UserPeopleId
        pid = int(model.Data.email_pid)
        subject = model.Data.email_subject
        message = model.Data.email_message
        from_email = model.Data.email_from
        from_name = model.Data.email_fromname
        
        # Send the email
        model.Email(pid, LoggedInUser, from_email, from_name, subject, message)
        
        # Success! Note: we must return SOMETHING or the request will 404
        print "Email sent successfully"
    except Exception as e:
        # More detailed error handling
        import traceback
        print "<h4 style='color:red;'>Error</h4>"
        print "<p>An error occurred while sending email: " + str(e) + "</p>"
        print "<pre style='font-size:10px;color:#666;'>"
        traceback.print_exc()
        print "</pre>"
        print " "

#--------------------------
#Add people from savedQuery
#--------------------------
try:
    people = ''
    for p in q.QueryList(savedQuery, "PeopleId"):
        if people:
            people += ',' + str(p.PeopleId)
        else:
            people += str(p.PeopleId)
    
    # Get anniversaries from a wider date range (6 months before/after)
    data = q.QuerySql(sql.format(people, dataRange, extraValueFieldFriendlyName, extraValueField))
    
    # For each anniversary, add a JavaScript object to the allAnniversaries array
    if data:
        for i in data:
            name_safe = i.Name.replace("'", "\\'").replace('"', '\\"')
            type_safe = i.Type.replace("'", "\\'").replace('"', '\\"')
            # Convert ThisYearDate from SQL DateTime to JavaScript format
            # Format: MM/DD
            date_safe = i.dDate
            email_safe = ""
            if hasattr(i, 'Email') and i.Email:
                email_safe = i.Email.replace("'", "\\'").replace('"', '\\"')
            
            print "            this.allAnniversaries.push({"
            print "                pid: " + str(i.PeopleId) + ","
            print "                name: \"" + name_safe + "\","
            print "                date: \"" + date_safe + "\","
            print "                type: \"" + type_safe + "\","
            print "                yearsCount: " + str(i.YearsCount) + ","
            print "                email: \"" + email_safe + "\""
            print "            });"
    
    print """
        },
        
        // Parse a date string in MM/dd format into a JS Date object for the current year
        parseAnniversaryDate: function(dateStr) {
            const parts = dateStr.split('/');
            if (parts.length !== 2) return null;
            
            const month = parseInt(parts[0]) - 1; // JS months are 0-based
            const day = parseInt(parts[1]);
            
            const date = new Date();
            date.setMonth(month);
            date.setDate(day);
            
            return date;
        },
        
        // Handle navigation - changes the visible anniversaries without reloading
        navigate: function(direction) {
            // Update the offset
            this.currentOffset += direction * this.daysToShow;
            
            // Update the header to show the current offset
            this.updateOffsetDisplay();
            
            // Filter and display the appropriate anniversaries
            this.displayVisibleAnniversaries();
        },
        
        // Send email function
        sendEmail: function(pid, celebrationType, personName, yearsCount, personEmail) {
            // Create appropriate subject and email body based on celebration type
            let subject = "";
            let emailBody = "";
            
            if (celebrationType.includes("Birthday")) {
                subject = "Happy Birthday " + personName + "!";
                emailBody = "Happy Birthday " + personName + "!\\n\\n";
                emailBody += "Wishing you a wonderful birthday celebration!\\n\\n";
            } else if (celebrationType.includes("Wedding")) {
                subject = "Happy " + yearsCount + (yearsCount === 1 ? " Year" : " Years") + " Wedding Anniversary " + personName + "!";
                emailBody = "Congratulations on " + yearsCount + (yearsCount === 1 ? " year" : " years") + " of marriage, " + personName + "!\\n\\n";
                emailBody += "May your love continue to grow stronger with each passing year.\\n\\n";
            } else if (celebrationType.includes("Work")) {
                subject = "Happy " + yearsCount + (yearsCount === 1 ? " Year" : " Years") + " Work Anniversary " + personName + "!";
                emailBody = "Congratulations on " + yearsCount + (yearsCount === 1 ? " year" : " years") + " with us, " + personName + "!\\n\\n";
                emailBody += "Thank you for your dedication and all your contributions.\\n\\n";
            }
            
            emailBody += "Best wishes,\\n""" + fromName + """";
            
            // Open email dialog with SweetAlert
            Swal.fire({
                title: "Send Celebration Email",
                html: `
                    <div style="text-align: left; margin-bottom: 15px;">
                        <strong>To:</strong> ${personName}${personEmail ? ' (' + personEmail + ')' : ''}<br>
                        <strong>Subject:</strong> ${subject}
                    </div>
                    <textarea id="swal-email-message" class="swal2-textarea" style="height: 200px;">${emailBody}</textarea>
                `,
                showCancelButton: true,
                confirmButtonText: 'Send Email',
                confirmButtonColor: '#4CAF50',
                cancelButtonColor: '#d33',
                showLoaderOnConfirm: true,
                preConfirm: () => {
                    // Get the message from the textarea
                    const message = document.getElementById('swal-email-message').value;
                    if (!message.trim()) {
                        Swal.showValidationMessage('Please enter a message');
                        return false;
                    }
                    
                    // Get the current page URL path
                    let currentPath = window.location.pathname;
                    let formPath = currentPath.replace('/PyScript/', '/PyScriptForm/');
                    
                    // Build form data
                    const formData = new FormData();
                    formData.append('a', 'email');
                    formData.append('email_pid', pid);
                    formData.append('email_subject', subject);
                    formData.append('email_message', message);
                    formData.append('email_from', '""" + fromEmail + """');
                    formData.append('email_fromname', '""" + fromName + """');
                    
                    // Send request
                    return fetch(formPath, {
                        method: 'POST',
                        body: new URLSearchParams(formData)
                    })
                    .then(response => {
                        if (!response.ok) {
                            throw new Error(`Request failed with status ${response.status}`);
                        }
                        return response.text();
                    })
                    .catch(error => {
                        Swal.showValidationMessage(`Request failed: ${error}`);
                        return false;
                    });
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
        
        // Update the header to show the current offset
        updateOffsetDisplay: function() {
            const rangeDisplay = document.getElementById('""" + widget_id + """_range');
            if (!rangeDisplay) return;
            
            if (this.currentOffset === 0) {
                rangeDisplay.style.display = 'none';
            } else {
                rangeDisplay.style.display = 'block';
                if (this.currentOffset < 0) {
                    rangeDisplay.textContent = `(Viewing ${Math.abs(this.currentOffset)} days ago)`;
                } else {
                    rangeDisplay.textContent = `(Viewing ${this.currentOffset} days ahead)`;
                }
            }
        },
        
        // Check if a date is within the current view range
        isDateInRange: function(dateObj) {
            const today = new Date();
            const startDate = new Date(today);
            const endDate = new Date(today);
            
            // Add the offset to our date range
            startDate.setDate(startDate.getDate() + this.currentOffset);
            endDate.setDate(endDate.getDate() + this.currentOffset + this.daysToShow);
            
            // Check if the anniversary date is in this range
            return dateObj >= startDate && dateObj <= endDate;
        },
        
        // Filter and display only the anniversaries in the current view range
        displayVisibleAnniversaries: function() {
            const container = document.getElementById('""" + widget_id + """_container');
            if (!container) return;
            
            container.innerHTML = ''; // Clear current display
            
            let visibleCount = 0;
            
            for (const anniv of this.allAnniversaries) {
                const anniversaryDate = this.parseAnniversaryDate(anniv.date);
                if (!anniversaryDate) continue; // Skip if date is invalid
                
                if (this.isDateInRange(anniversaryDate)) {
                    visibleCount++;
                    
                    // Create a row for this anniversary
                    const row = document.createElement('div');
                    row.className = 'anniversary-row';
                    
                    // Create safe versions of strings for the onclick handler
                    const safeType = anniv.type.replace(/'/g, "\\'").replace(/"/g, '\\"');
                    const safeName = anniv.name.replace(/'/g, "\\'").replace(/"/g, '\\"');
                    
                    row.innerHTML = `
                        <span class="anniversary-date">${anniv.date}</span>
                        <span class="anniversary-name"><a href="/Person2/${anniv.pid}#tab-communications">${anniv.name}</a></span>
                        <span class="anniversary-type">${anniv.type}</span>
                        <button class="email-button" title="Send celebration email" 
                            onclick="event.preventDefault(); """ + widget_id + """.sendEmail(${anniv.pid}, '${safeType}', '${safeName}', ${anniv.yearsCount}, '${anniv.email}')">✉️</button>
                    `;
                    
                    container.appendChild(row);
                }
            }
            
            // Show a message if no anniversaries in this range
            if (visibleCount === 0) {
                const emptyMsg = document.createElement('p');
                emptyMsg.style.textAlign = 'center';
                emptyMsg.style.padding = '10px';
                emptyMsg.textContent = 'No anniversaries for this date range';
                container.appendChild(emptyMsg);
            }
        }
    };
    
    // Define global navigation function for the arrow buttons
    window.""" + widget_id + """_navigate = function(direction) {
        """ + widget_id + """.navigate(direction);
    };
    
    // Wait for document to be ready, then initialize
    function initWidget() {
        if (document.getElementById('""" + widget_id + """')) {
            if (!document.getElementById('""" + widget_id + """').getAttribute('data-initialized')) {
                """ + widget_id + """.init();
            }
        } else {
            // If the widget element doesn't exist yet, try again in 100ms
            setTimeout(initWidget, 100);
        }
    }
    
    // Start the initialization process
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initWidget);
    } else {
        // Document already loaded
        initWidget();
    }
})();
</script>
"""

except Exception as e:
    # Improved error handling
    import traceback
    print "<h3>Error</h3>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
