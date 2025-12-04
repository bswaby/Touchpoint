#####################################################################
####TECH EMAIL REPORT INFORMATION - ENHANCED VERSION
#####################################################################
#This comprehensive email troubleshooting tool shows email success/failures 
#from multiple perspectives with advanced diagnostics.
#
#Written By: Ben Swaby
#Email: bswaby@fbchtn.org
#
#Features:
# - Email status overview with failure classifications
# - Organization-level email statistics
# - User-level failure analysis
# - Bounce rate trends
# - Email queue monitoring
# - Domain reputation analysis
# - Quick links to email logs and individual records
# - Detailed failure explanations and remediation steps
# - Campaign send performance analysis
# - Email body size and image analysis
#
#Update Log:
# 2025-12-04:
#   - Added Campaign Send Performance section with configurable minimum recipients filter
#   - Added recipient size bucket analysis with visual bar chart (how recipient count affects send time)
#   - Added email body size analysis with size buckets (Small to Huge 500KB+)
#   - Added image detection: counts <img> tags and detects base64 embedded images
#   - Fixed scheduled email timing calculation (uses SendWhen instead of Queued for scheduled emails)
#   - Enhanced Recent Campaign Details table with Size (KB), Images, Embedded?, and Type columns
#   - Added info box explaining linked vs embedded images (linked ~100 bytes, embedded 1-4MB each)
#
#Installation:
#  Installation is easy, but it does require rights to Special Content.
#  1. Copy all this code
#  2. In TP, go to Admin~Advanced~Special Content~Python Scripts Tab
#  3. Select New Python Script File, Name the File, and click submit
#  4. Paste in the code and run
#  5. Add to menu if needed


#####################################################################
####USER CONFIG FIELDS
#####################################################################
#These are defined variables that are required for the report to run.

model.Header = 'Email Technical Diagnostics Dashboard' #Page Name

#######################################################################
####START OF CODE.  No configuration should be needed beyond this point
#######################################################################
#######################################################################
from datetime import datetime  # Import datetime at the beginning of the function
import datetime
from decimal import Decimal
import re
import json

current_date = datetime.date.today().strftime("%B %d, %Y")
sDate = model.Data.sDate
eDate = model.Data.eDate


if model.Data.HideSuccess == 'yes':
    optionHideSuccess = 'checked'
    sqlHideSuccess = ''
else:
    optionHideSuccess = ''
    sqlHideSuccess = ' AND fe.Fail IS NOT NULL '

from datetime import datetime

# Email failure classifications and remediation steps based on SendGrid documentation and TouchPoint specifics
FAILURE_CLASSIFICATIONS = {
    'Technical': {
        'description': 'A permanent failure due to a technical issue (e.g., DNS problems, invalid domain)',
        'causes': ['Invalid or non-existent domain', 'DNS configuration issues', 'MX records not found'],
        'remediation': 'Verify the email domain exists and has proper MX records. Check for typos in email addresses.',
        'severity': 'high',
        'touchpoint_note': 'These failures typically require fixing the underlying technical issue with the domain or email server configuration.'
    },
    'Invalid Address': {
        'description': 'The email address format is invalid or the mailbox does not exist. Address is added to SendGrid\'s Invalid bucket.',
        'causes': ['Malformed email address', 'Misspelled email address', 'Non-existent mailbox', 'Deleted email account'],
        'remediation': 'Correct misspellings in email addresses. For persistent invalid addresses, remove from your lists. Check the Failed tab in TouchPoint for details.',
        'severity': 'high',
        'touchpoint_note': 'Invalid addresses are stored in SendGrid\'s Invalid bucket. After 3 bounces, emails are dropped. After 30+ days, the cycle may restart.'
    },
    'invalid': {
        'description': 'Email address bounced as invalid but was theoretically valid format. Gets "dropped" after repeated failures.',
        'causes': ['Mailbox no longer exists', 'Account closed', 'Temporary but persistent server rejection'],
        'remediation': 'Check the Failed tab for bounce history. After 30 days, TouchPoint may retry. Consider verifying the email address with the recipient.',
        'severity': 'high',
        'touchpoint_note': 'Status shows as "Not Delivered" after being dropped. The retry cycle restarts after 30+ days.'
    },
    'Mailbox Unavailable': {
        'description': 'The mailbox exists but is temporarily unavailable',
        'causes': ['Mailbox full', 'Account suspended', 'Temporary server issues', 'Out of office auto-reply'],
        'remediation': 'TouchPoint will automatically retry. If persistent, contact recipient through alternative means.',
        'severity': 'medium',
        'touchpoint_note': 'These are often temporary issues that resolve themselves.'
    },
    'Reputation': {
        'description': 'Email blocked due to sender reputation issues',
        'causes': ['IP/domain on blocklist', 'Poor sending history', 'High spam complaint rate', 'Authentication failures'],
        'remediation': 'Contact TouchPoint support - may require SendGrid IP reputation management. Ensure SPF/DKIM/DMARC are properly configured.',
        'severity': 'critical',
        'touchpoint_note': 'This affects all emails from your organization. Requires immediate attention from TouchPoint support.'
    },
    'Content': {
        'description': 'Email blocked due to content that triggered spam filters',
        'causes': ['Spam-like content', 'Suspicious links', 'Blocked attachments', 'Trigger words'],
        'remediation': 'Review and modify email content. Avoid spam trigger words, excessive formatting, and suspicious links.',
        'severity': 'medium',
        'touchpoint_note': 'Consider using TouchPoint\'s email templates which are pre-tested for deliverability.'
    },
    'Unclassified': {
        'description': 'Bounce could not be categorized into other classifications',
        'causes': ['Unique server responses', 'Custom rejection messages', 'Unknown errors'],
        'remediation': 'Review the specific bounce message in the Failed tab for details. May require TouchPoint support assistance.',
        'severity': 'low',
        'touchpoint_note': 'Often requires manual investigation. Check the Failed tab for specific error messages.'
    },
    'spamreport': {
        'description': 'Recipient marked the email as spam. Email address is blocked in SendGrid\'s Spam Report bucket.',
        'causes': ['Recipient clicked spam/junk button', 'Unwanted emails', 'No clear unsubscribe option', 'Frequency too high'],
        'remediation': 'IMPORTANT: Only TouchPoint staff can remove spam blocks. Contact support to remove from SendGrid\'s Spam Report bucket or use the Failed tab if available.',
        'severity': 'critical',
        'touchpoint_note': 'Once marked as spam, the person won\'t receive ANY emails through TouchPoint until the block is removed by TouchPoint staff.'
    },
    'Spam Report': {
        'description': 'Recipient marked the email as spam. Email address is blocked in SendGrid\'s Spam Report bucket.',
        'causes': ['Recipient clicked spam/junk button', 'Unwanted emails', 'No clear unsubscribe option', 'Frequency too high'],
        'remediation': 'IMPORTANT: Only TouchPoint staff can remove spam blocks. Contact support to remove from SendGrid\'s Spam Report bucket or use the Failed tab if available.',
        'severity': 'critical',
        'touchpoint_note': 'Once marked as spam, the person won\'t receive ANY emails through TouchPoint until the block is removed by TouchPoint staff.'
    },
    'bouncedaddress': {
        'description': 'Address has consistently bounced and is now suppressed',
        'causes': ['Persistent delivery failures', 'Abandoned mailbox', 'Multiple bounce types'],
        'remediation': 'Address is likely invalid. Verify with recipient through other means. May need removal from SendGrid suppression list.',
        'severity': 'high',
        'touchpoint_note': 'These addresses are on SendGrid\'s suppression list due to repeated failures.'
    },
    'dropped': {
        'description': 'Email was dropped and not attempted for delivery',
        'causes': ['Previous hard bounces', 'On suppression list', 'Invalid address in SendGrid bucket'],
        'remediation': 'Check Failed tab for history. Address may be in SendGrid suppression. After 30+ days, retry cycle may begin again.',
        'severity': 'high',
        'touchpoint_note': 'Dropped emails indicate the address is on a suppression list. Status shows as "Not Delivered" in TouchPoint.'
    },
    'Missing Subject': {
        'description': 'Email queued without a subject line and will not send',
        'causes': ['No subject provided when email was created', 'Draft/test email accidentally queued', 'API call missing subject parameter'],
        'remediation': 'Edit the email to add a subject line or delete if it was a test. TouchPoint requires a subject line for all emails.',
        'severity': 'high',
        'touchpoint_note': 'Emails without subjects appear as "sent" in UI but never actually process. This is a safety feature to prevent malformed emails.'
    },
    'No Schedule Date': {
        'description': 'Email has no SendWhen date set for immediate sending but has not processed',
        'causes': ['Missing subject line', 'Empty email body', 'System processing issue', 'Email in draft state'],
        'remediation': 'Check if email has subject and body content. For immediate send, SendWhen should be NULL. May need to recreate the email.',
        'severity': 'medium',
        'touchpoint_note': 'Immediate emails (SendWhen=NULL) should send right away. If stuck, usually missing required fields like subject or body.'
    },
    'Empty Body': {
        'description': 'Email queued without any body content and will not send',
        'causes': ['No body content provided', 'Test email with only subject', 'API call missing body parameter', 'Template loading failure'],
        'remediation': 'Add body content to the email or delete if it was a test. TouchPoint requires both subject AND body content for emails to send.',
        'severity': 'high',
        'touchpoint_note': 'Emails without body content are blocked from sending even if they have valid subjects and recipients. This prevents empty emails.'
    },
    'Processing Issue': {
        'description': 'Email has all required fields but has not been processed by the email system',
        'causes': ['Email service temporarily down', 'Queue processing paused', 'System maintenance', 'SendGrid API issues'],
        'remediation': 'Contact TouchPoint support. Email appears valid but may be stuck in processing queue. May need to be requeued.',
        'severity': 'critical',
        'touchpoint_note': 'These emails should have sent immediately but are stuck. Usually indicates a system-level issue requiring support intervention.'
    },
    'Past Due': {
        'description': 'Scheduled email that should have sent but failed',
        'causes': ['SendGrid rejection', 'API failures', 'Service interruption at scheduled time', 'Invalid recipients discovered at send time'],
        'remediation': 'Review email details and error logs. May need to reschedule or recreate the email.',
        'severity': 'high',
        'touchpoint_note': 'These emails were scheduled for a past date/time but never sent. Check the Failed tab for specific error details.'
    },
    'No Recipients': {
        'description': 'Email has no recipients in the EmailQueueTo table',
        'causes': ['Query returned no results', 'Recipients removed after queuing', 'Test email without recipients', 'Template email not personalized'],
        'remediation': 'Add recipients or delete if this is a template. Emails without recipients cannot be sent.',
        'severity': 'medium',
        'touchpoint_note': 'Common for email templates and test emails. These will never send without recipients being added.'
    }
}

def generate_html_table(data, expanded_list=None, match_column=None, title=None, hide_columns=None, url_columns=None, 
                        bold_columns=None, column_order=None, divider_after_column=None, sum_columns=None, 
                        date_columns=None, column_widths=None, table_width="100%", remove_borders=False,
                        remove_header_borders=False, header_padding="5px", content_padding="5px", header_font_size="14px", 
                        content_font_size="12px", header_bg_color="#f4f4f4", slanted_headers=None, slant_angle=45, 
                        row_colors=None, header_rename=None, column_alignment=None):  
    
    if not data:
        return "<p>No data available.</p>"
    
    hide_columns = hide_columns or []
    url_columns = url_columns or {}
    bold_columns = bold_columns or []
    column_order = column_order or []
    sum_columns = sum_columns or []
    date_columns = date_columns or []
    column_widths = column_widths or {}
    slanted_headers = slanted_headers or []
    row_colors = row_colors or ["#ffffff", "#f9f9f9"]
    header_rename = header_rename or {}
    column_alignment = column_alignment or {}
    
    from datetime import datetime
    
    html = ""
    html += """
        <style>
            th.rotate {
              height: 140px;
              white-space: nowrap;
              position: relative;
            }
            
            th.rotate > div {
              transform: translate(15px, 35px) rotate(315deg); 
              width: 100%;
              position: absolute;
              top: 50%;
              left: 0;
            }
            
            th.rotate > div > span {
              border-bottom: 1px solid #ccc;
              padding: 5px 10px;
            }
        </style>
        """
    
    if title:
        html += "<h2>{}</h2>\n".format(title)
    
    headers = data[0].keys()
    
    if column_order:
        headers = [h for h in column_order if h in headers] + [h for h in headers if h not in column_order]
    
    table_style = "width: {}; border-collapse: collapse;".format(table_width)
    if remove_borders:
        table_style += " border: none;"
    
    html += '<table style="{}">\n'.format(table_style)
    
    html += "    <tr>\n"
    
    for h in headers:
        if h not in hide_columns:
            th_style = "text-align: {}; vertical-align: bottom; white-space: nowrap;"
            th_style = th_style.format(column_alignment.get(h, "center"))  # Align header as per column_alignment
            th_style += " padding: {}; font-size: {}; background-color: {};".format(header_padding, header_font_size, header_bg_color)
            
            if h in slanted_headers:
                th_style = "text-align: {}; vertical-align: bottom; white-space: nowrap; padding: {}; font-size: {};".format(column_alignment.get(h, "center"), header_padding, header_font_size)
            
            header_display_name = header_rename.get(h, h)
            
            if remove_header_borders and h not in slanted_headers:
                th_style += " border: none;"
            elif not remove_borders and h not in slanted_headers:
                th_style += " border: 1px solid #ddd;"
            
            if h in column_widths:
                th_style += " width: {};".format(column_widths[h])
    
            if h in slanted_headers:
                html += "<th class='rotate' style='{}'><div><span>{}</span></div></th>\n".format(th_style, header_display_name)
            else:
                html += "<th style='{}'>{}</th>\n".format(th_style, header_display_name)

    html += "    </tr>\n"
    
    expanded_rows = {}
    if expanded_list:
        expanded_rows = {row[match_column]: row for row in expanded_list}
    
    for idx, row in enumerate(data):
        row_color = row_colors[idx % len(row_colors)]  # Alternate row colors
        
        html += "    <tr style='background-color: {};'>\n".format(row_color)
        for h in headers:
            if h not in hide_columns:
                cell_value = row.get(h, "")
                
                # Apply header renaming to date_columns
                if h in header_rename:
                    h = header_rename[h]  # Rename the header
                
                for column in date_columns:
                    if h == column:
                        if cell_value:
                            cell_value = str(cell_value)  # Ensure it's a string
                            cell_value = cell_value.split(" ")[0]  # Take only the date part
                        else:
                            cell_value = ""  # Handle None values

                if cell_value is None:
                    cell_value = ""
                if h in bold_columns:
                    cell_value = "<b>{}</b>".format(cell_value)
                if h in url_columns and cell_value:
                    url_format = url_columns[h]
                    url_value = url_format.format(**row)
                    cell_value = '<a href="{}" target="_blank" style="color: #1e90ff; text-decoration: none;">{}</a>'.format(url_value, cell_value)
                
                alignment = column_alignment.get(h, "left")
                td_style = "text-align: {}; padding: {}; font-size: {};".format(alignment, content_padding, content_font_size)
                if h in column_widths:
                    td_style += " width: {};".format(column_widths[h])
                if not remove_borders:
                    td_style += " border: 1px solid #ddd;"
                html += "<td style='{}'>{}</td>\n".format(td_style, cell_value)
        html += "    </tr>\n"
        
        if expanded_list and row[match_column] in expanded_rows:
            expanded_row = expanded_rows[row[match_column]]
            html += "    <tr style='background-color: #e0e0e0;'>\n"
            for h in headers:
                if h not in hide_columns:
                    cell_value = expanded_row.get(h, "")
                    td_style = "padding-left: 20px; text-align: {}; font-size: {};".format(column_alignment.get(h, "left"), content_font_size)
                    html += "<td style='{}'>{}</td>\n".format(td_style, cell_value)
            html += "    </tr>\n"
        
    if sum_columns:
        sums = {col: 0 for col in sum_columns}
        for row in data:
            for col in sum_columns:
                if isinstance(row.get(col), (int, float)):
                    sums[col] += row[col]

        html += "    <tr>\n"
        for h in headers:
            if h not in hide_columns:
                if h in sum_columns:
                    html += "<td style='padding: 3px 4px; text-align: left; font-weight: bold; border: none;'>{}</td>\n".format(sums[h])
                else:
                    html += "<td style='padding: 3px 4px; text-align: left; border: none;'>&nbsp;</td>\n"
        html += "    </tr>\n"
    
    html += "</table>\n"
    return html




sql = '''
SELECT 
    COALESCE(fe.Fail, 'Sent') AS Status,  
    COUNT(*) AS TotalCount,
    CAST(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER() AS DECIMAL(5,2)) AS Percentage
FROM EmailQueueTo eqt
LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id aND fe.PeopleId = eqt.PeopleId  -- Adjust join column if needed
LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
LEFT JOIN Division d ON d.Id = o.DivisionId
LEFT JOIN Program pro ON pro.Id = d.ProgId
WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2} {3} {4} {5}
GROUP BY COALESCE(fe.Fail, 'Sent')
ORDER BY TotalCount DESC;
'''



sqlOrgStat = '''
   DECLARE @DynamicColumns NVARCHAR(MAX);
    DECLARE @SQLQuery NVARCHAR(MAX);
    
    -- Step 1: Retrieve unique fe.Fail values dynamically
    SELECT @DynamicColumns = STRING_AGG(QUOTENAME(Fail), ', ')
    FROM (SELECT DISTINCT Fail FROM FailedEmails WHERE Fail IS NOT NULL) AS Failures;
    
    -- Ensure @DynamicColumns is not NULL
    IF @DynamicColumns IS NULL 
        SET @DynamicColumns = '[Sent]';
    
    -- Step 2: Construct the Dynamic SQL Query
    SET @SQLQuery = '
    SELECT *
    FROM (
        SELECT 
            eqt.OrgId,
			eqt.id,
            pro.Name AS Program,
			eq.Subject,
			CAST(eqt.Sent AS DATE) AS SentDate,
            o.OrganizationName,
            COALESCE(fe.Fail, ''Sent'') AS Status,  
            COUNT(*) AS TotalCount
        FROM EmailQueueTo eqt
        LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId  
		LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
        LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
        LEFT JOIN Division d ON d.Id = o.DivisionId
        LEFT JOIN Program pro ON pro.Id = d.ProgId
        WHERE eqt.Sent BETWEEN ''{0} 00:00:00'' AND ''{1} 23:59:59.999'' {2} {3} {4} {5}
        GROUP BY COALESCE(fe.Fail, ''Sent''), eqt.OrgId, o.OrganizationName, pro.Name,eq.Subject,CAST(eqt.Sent AS DATE),eqt.id
    ) SourceTable
    PIVOT (
        SUM(TotalCount) 
        FOR Status IN (' + @DynamicColumns + ')
    ) PivotTable
    ORDER BY Program, OrganizationName;';
    
    -- Step 3: Execute the Dynamic SQL
    EXEC(@SQLQuery);
'''


sqlOrgDetails = '''
SELECT eqt.id, eq.Subject, CAST(eqt.Sent AS DATE) AS SentDate, eqt.OrgId
FROM EmailQueueTo eqt
LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id aND fe.PeopleId = eqt.PeopleId
WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2}
GROUP BY eqt.id, eq.Subject, CAST(eqt.Sent AS DATE), eqt.OrgId
ORDER BY OrgId, SentDate;
'''


sqlUserFailedStat = '''
SELECT 
	eqt.PeopleId,
	p.Name,
    fe.Fail AS Status,  
    COUNT(*) AS TotalCount,
    p.SendEmailAddress1,
	p.EmailAddress,
	p.SendEmailAddress2,
	p.EmailAddress2
FROM EmailQueueTo eqt
LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id aND fe.PeopleId = eqt.PeopleId  -- Adjust join column if needed
LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
LEFT JOIN Division d ON d.Id = o.DivisionId
LEFT JOIN Program pro ON pro.Id = d.ProgId
LEFT JOIN People p ON p.PeopleId = eqt.PeopleId
WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2} {3} {4} {5}
GROUP BY fe.Fail, eqt.PeopleId, p.Name,	p.EmailAddress,p.EmailAddress2,p.SendEmailAddress1,p.SendEmailAddress2
ORDER BY TotalCount DESC;
'''

sqlPrograms = """select Id, Name AS ProgramName From Program Order by Name"""
sqlFailClassifications = """select distinct fe.Fail from FailedEmails fe Order by fe.Fail"""

# Additional SQL for enhanced diagnostics
sqlBounceRateTrend = '''
SELECT 
    CAST(eqt.Sent AS DATE) AS Date,
    COUNT(*) AS TotalSent,
    SUM(CASE WHEN fe.Fail IS NOT NULL THEN 1 ELSE 0 END) AS TotalFailed,
    CAST(SUM(CASE WHEN fe.Fail IS NOT NULL THEN 1 ELSE 0 END) * 100.0 / COUNT(*) AS DECIMAL(5,2)) AS BounceRate
FROM EmailQueueTo eqt
LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
LEFT JOIN Division d ON d.Id = o.DivisionId
LEFT JOIN Program pro ON pro.Id = d.ProgId
WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2} {3} {4} {5}
GROUP BY CAST(eqt.Sent AS DATE)
ORDER BY Date DESC;
'''

sqlDomainAnalysis = '''
SELECT 
    SUBSTRING(p.EmailAddress, CHARINDEX('@', p.EmailAddress) + 1, LEN(p.EmailAddress)) AS Domain,
    COUNT(DISTINCT p.PeopleId) AS TotalRecipients,
    SUM(CASE WHEN fe.Fail IS NOT NULL THEN 1 ELSE 0 END) AS FailedEmails,
    CAST(SUM(CASE WHEN fe.Fail IS NOT NULL THEN 1 ELSE 0 END) * 100.0 / COUNT(*) AS DECIMAL(5,2)) AS FailureRate
FROM EmailQueueTo eqt
LEFT JOIN EmailQueue eq ON eq.Id = eqt.Id
LEFT JOIN FailedEmails fe ON fe.Id = eqt.Id AND fe.PeopleId = eqt.PeopleId
LEFT JOIN People p ON p.PeopleId = eqt.PeopleId
LEFT JOIN Organizations o ON o.OrganizationId = eqt.OrgId
LEFT JOIN Division d ON d.Id = o.DivisionId
LEFT JOIN Program pro ON pro.Id = d.ProgId
WHERE eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999' {2} {3} {4} {5}
    AND p.EmailAddress IS NOT NULL
GROUP BY SUBSTRING(p.EmailAddress, CHARINDEX('@', p.EmailAddress) + 1, LEN(p.EmailAddress))
HAVING COUNT(*) > 5
ORDER BY FailureRate DESC, TotalRecipients DESC;
'''

sqlEmailQueueStatus = '''
SELECT 
    eq.QueuedBy,
    p.Name AS QueuedByName,
    COUNT(*) AS PendingEmails,
    MIN(eq.Queued) AS OldestQueued,
    MAX(eq.Queued) AS NewestQueued,
    DATEDIFF(hour, MIN(eq.Queued), GETDATE()) AS HoursOld,
    STRING_AGG(CAST(eq.Id AS VARCHAR), ',') AS EmailIds
FROM EmailQueue eq
LEFT JOIN People p ON p.PeopleId = eq.QueuedBy
WHERE eq.Sent IS NULL
GROUP BY eq.QueuedBy, p.Name
ORDER BY MIN(eq.Queued);
'''

sqlEmailQueueDetails = '''
SELECT TOP 10
    eq.Id,
    eq.Subject,
    eq.Queued,
    eq.SendWhen,
    eq.QueuedBy,
    p.Name AS QueuedByName,
    DATEDIFF(hour, eq.Queued, GETDATE()) AS HoursOld,
    (SELECT COUNT(*) FROM EmailQueueTo WHERE Id = eq.Id) AS RecipientCount,
    CASE
        WHEN eq.Subject IS NULL OR eq.Subject = '' THEN 'Missing Subject'
        WHEN eq.Body IS NULL OR LEN(eq.Body) = 0 THEN 'Empty Body'
        WHEN eq.SendWhen IS NULL AND DATEDIFF(hour, eq.Queued, GETDATE()) > 1
             AND (eq.Subject IS NOT NULL AND eq.Subject != '')
             AND (eq.Body IS NOT NULL AND LEN(eq.Body) > 0) THEN 'Processing Issue'
        WHEN eq.SendWhen IS NOT NULL AND eq.SendWhen <= GETDATE() AND (SELECT COUNT(*) FROM EmailQueueTo WHERE Id = eq.Id) > 0 THEN 'Past Due'
        WHEN (SELECT COUNT(*) FROM EmailQueueTo WHERE Id = eq.Id) = 0 THEN 'No Recipients'
        ELSE 'Unknown'
    END AS FailureReason
FROM EmailQueue eq
LEFT JOIN People p ON p.PeopleId = eq.QueuedBy
WHERE eq.Sent IS NULL
ORDER BY eq.Queued
'''

# Campaign Send Performance SQL - measures time from scheduled/queue time to completion
# For scheduled emails (SendWhen IS NOT NULL), measure from SendWhen to Sent
# For immediate emails (SendWhen IS NULL), measure from Queued to Sent
sqlCampaignPerformance = """
SELECT
    eq.Id,
    eq.Subject,
    eq.Queued,
    eq.SendWhen,
    eq.Sent AS CompletedAt,
    p.Name AS SentBy,
    COUNT(eqt.PeopleId) AS Recipients,
    DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS SendDurationSeconds,
    CASE
        WHEN DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) < 60 THEN CAST(DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS VARCHAR) + ' sec'
        WHEN DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) < 3600 THEN CAST(DATEDIFF(minute, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS VARCHAR) + ' min'
        ELSE CAST(CAST(DATEDIFF(minute, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) / 60.0 AS DECIMAL(10,1)) AS VARCHAR) + ' hrs'
    END AS SendDuration,
    CASE WHEN eq.SendWhen IS NOT NULL THEN 'Scheduled' ELSE 'Immediate' END AS SendType,
    LEN(eq.Body) AS BodySizeBytes,
    CAST(LEN(eq.Body) / 1024.0 AS DECIMAL(10,1)) AS BodySizeKB,
    (LEN(eq.Body) - LEN(REPLACE(eq.Body, '<img', ''))) / 4 AS ImageCount,
    CASE WHEN eq.Body LIKE '%data:image%' THEN 'Yes' ELSE 'No' END AS HasEmbeddedImages
FROM EmailQueue eq
JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
LEFT JOIN People p ON p.PeopleId = eq.QueuedBy
WHERE eq.Sent IS NOT NULL
    AND eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
    {2} {3} {4}
GROUP BY eq.Id, eq.Subject, eq.Queued, eq.SendWhen, eq.Sent, p.Name, eq.Body
HAVING COUNT(eqt.PeopleId) >= {5}
ORDER BY eq.Sent DESC
"""

# Campaign Performance Summary SQL - overall stats
# Uses COALESCE to measure from SendWhen (for scheduled) or Queued (for immediate)
sqlCampaignPerformanceSummary = '''
WITH CampaignStats AS (
    SELECT
        eq.Id,
        eq.Queued,
        eq.SendWhen,
        eq.Sent,
        COUNT(eqt.PeopleId) AS Recipients,
        DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS SendDurationSeconds
    FROM EmailQueue eq
    JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
    WHERE eq.Sent IS NOT NULL
        AND eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
        {2} {3} {4}
    GROUP BY eq.Id, eq.Queued, eq.SendWhen, eq.Sent
    HAVING COUNT(eqt.PeopleId) >= {5}
)
SELECT
    COUNT(*) AS TotalCampaigns,
    SUM(Recipients) AS TotalRecipients,
    AVG(Recipients) AS AvgRecipients,
    AVG(SendDurationSeconds) AS AvgSendDurationSeconds,
    MIN(SendDurationSeconds) AS FastestSendSeconds,
    MAX(SendDurationSeconds) AS SlowestSendSeconds
FROM CampaignStats
'''

# Monthly Campaign Performance Trend SQL
# Uses COALESCE to measure from SendWhen (for scheduled) or Queued (for immediate)
sqlCampaignMonthlyTrend = '''
WITH MonthlyCampaigns AS (
    SELECT
        FORMAT(eq.Sent, 'yyyy-MM') AS Month,
        eq.Id,
        DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS SendDurationSeconds,
        COUNT(eqt.PeopleId) AS Recipients
    FROM EmailQueue eq
    JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
    WHERE eq.Sent IS NOT NULL
        AND eq.Sent >= DATEADD(month, -12, GETDATE())
        {0} {1} {2}
    GROUP BY FORMAT(eq.Sent, 'yyyy-MM'), eq.Id, eq.Queued, eq.SendWhen, eq.Sent
    HAVING COUNT(eqt.PeopleId) >= {3}
)
SELECT
    Month,
    COUNT(*) AS Campaigns,
    SUM(Recipients) AS TotalRecipients,
    AVG(Recipients) AS AvgRecipients,
    AVG(SendDurationSeconds) AS AvgSendSeconds,
    MIN(SendDurationSeconds) AS FastestSendSeconds,
    MAX(SendDurationSeconds) AS SlowestSendSeconds
FROM MonthlyCampaigns
GROUP BY Month
ORDER BY Month DESC
'''

# Recipient Size Bucket Analysis SQL - groups campaigns by recipient count ranges
sqlRecipientBucketAnalysis = '''
WITH CampaignStats AS (
    SELECT
        eq.Id,
        COUNT(eqt.PeopleId) AS Recipients,
        DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS SendDurationSeconds,
        LEN(eq.Body) AS BodySize
    FROM EmailQueue eq
    JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
    WHERE eq.Sent IS NOT NULL
        AND eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
        {2} {3} {4}
    GROUP BY eq.Id, eq.Queued, eq.SendWhen, eq.Sent, eq.Body
    HAVING COUNT(eqt.PeopleId) >= {5}
)
SELECT
    CASE
        WHEN Recipients BETWEEN {5} AND 250 THEN '1. {5}-250'
        WHEN Recipients BETWEEN 251 AND 500 THEN '2. 251-500'
        WHEN Recipients BETWEEN 501 AND 1000 THEN '3. 501-1000'
        WHEN Recipients BETWEEN 1001 AND 2500 THEN '4. 1001-2500'
        WHEN Recipients BETWEEN 2501 AND 5000 THEN '5. 2501-5000'
        ELSE '6. 5000+'
    END AS RecipientBucket,
    COUNT(*) AS CampaignCount,
    AVG(Recipients) AS AvgRecipients,
    AVG(SendDurationSeconds) AS AvgSendSeconds,
    MIN(SendDurationSeconds) AS FastestSendSeconds,
    MAX(SendDurationSeconds) AS SlowestSendSeconds,
    AVG(BodySize) AS AvgBodySize
FROM CampaignStats
GROUP BY
    CASE
        WHEN Recipients BETWEEN {5} AND 250 THEN '1. {5}-250'
        WHEN Recipients BETWEEN 251 AND 500 THEN '2. 251-500'
        WHEN Recipients BETWEEN 501 AND 1000 THEN '3. 501-1000'
        WHEN Recipients BETWEEN 1001 AND 2500 THEN '4. 1001-2500'
        WHEN Recipients BETWEEN 2501 AND 5000 THEN '5. 2501-5000'
        ELSE '6. 5000+'
    END
ORDER BY RecipientBucket
'''

# Email Body Size Analysis SQL - groups by body size ranges
# Note: LEN(Body) includes base64 embedded images which can be HUGE (4MB+)
# Linked/external images only add a small URL string to the body
sqlBodySizeAnalysis = """
WITH CampaignStats AS (
    SELECT
        eq.Id,
        COUNT(eqt.PeopleId) AS Recipients,
        DATEDIFF(second, COALESCE(eq.SendWhen, eq.Queued), eq.Sent) AS SendDurationSeconds,
        LEN(eq.Body) AS BodySize,
        (LEN(eq.Body) - LEN(REPLACE(eq.Body, '<img', ''))) / 4 AS ImageCount,
        CASE WHEN eq.Body LIKE '%data:image%' THEN 1 ELSE 0 END AS HasBase64Images
    FROM EmailQueue eq
    JOIN EmailQueueTo eqt ON eq.Id = eqt.Id
    WHERE eq.Sent IS NOT NULL
        AND eqt.Sent BETWEEN '{0} 00:00:00' AND '{1} 23:59:59.999'
        AND eq.Body IS NOT NULL
        {2} {3} {4}
    GROUP BY eq.Id, eq.Queued, eq.SendWhen, eq.Sent, eq.Body
    HAVING COUNT(eqt.PeopleId) >= {5}
)
SELECT
    CASE
        WHEN BodySize < 5000 THEN '1. Small (<5KB)'
        WHEN BodySize BETWEEN 5000 AND 15000 THEN '2. Medium (5-15KB)'
        WHEN BodySize BETWEEN 15001 AND 50000 THEN '3. Large (15-50KB)'
        WHEN BodySize BETWEEN 50001 AND 100000 THEN '4. XL (50-100KB)'
        WHEN BodySize BETWEEN 100001 AND 500000 THEN '5. XXL (100-500KB)'
        ELSE '6. Huge (500KB+, likely embedded images)'
    END AS SizeBucket,
    COUNT(*) AS CampaignCount,
    AVG(Recipients) AS AvgRecipients,
    AVG(SendDurationSeconds) AS AvgSendSeconds,
    MIN(SendDurationSeconds) AS FastestSendSeconds,
    MAX(SendDurationSeconds) AS SlowestSendSeconds,
    AVG(BodySize) AS AvgBodySizeBytes,
    MIN(BodySize) AS MinBodySize,
    MAX(BodySize) AS MaxBodySize,
    AVG(ImageCount) AS AvgImageCount,
    SUM(HasBase64Images) AS EmailsWithBase64Images
FROM CampaignStats
GROUP BY
    CASE
        WHEN BodySize < 5000 THEN '1. Small (<5KB)'
        WHEN BodySize BETWEEN 5000 AND 15000 THEN '2. Medium (5-15KB)'
        WHEN BodySize BETWEEN 15001 AND 50000 THEN '3. Large (15-50KB)'
        WHEN BodySize BETWEEN 50001 AND 100000 THEN '4. XL (50-100KB)'
        WHEN BodySize BETWEEN 100001 AND 500000 THEN '5. XXL (100-500KB)'
        ELSE '6. Huge (500KB+, likely embedded images)'
    END
ORDER BY SizeBucket
"""

if sDate is not None:
    optionsDate = ' value="' + sDate + '"'
else:
    optionsDate = ''

if eDate is not None:
    optioneDate = ' value="' + eDate + '"'
else:
    optioneDate = ''


frmProgramOption = """
    <label for="program" id="program">Program:</label>
    <select name="program" id="program">
        <option value="999999">All</option>
    """

rsqlPrograms = q.QuerySql(sqlPrograms)

for sp in rsqlPrograms:
    if str(model.Data.program) == str(sp.Id):
        frmProgramOption +=  """<option value="{0}" selected="selected">{1}</option>""".format(sp.Id, sp.ProgramName)
    else:
        frmProgramOption +=  """<option value="{0}">{1}</option>""".format(sp.Id, sp.ProgramName)

frmProgramOption += """</select>"""

frmFailClassificationsOption = """
    <label for="failclassification" id="program">Failure Type:</label>
    <select name="failclassification" id="failclassification">
        <option value="999999">All</option>
    """

rsqlFailClassifications = q.QuerySql(sqlFailClassifications)

for sp in rsqlFailClassifications:
    if str(model.Data.failclassification) == str(sp.Fail):
        frmFailClassificationsOption +=  """<option value="{0}" selected="selected">{0}</option>""".format(sp.Fail)
    else:
        frmFailClassificationsOption +=  """<option value="{0}">{0}</option>""".format(sp.Fail)

frmFailClassificationsOption += """</select>"""

# Create text input for Sent By PeopleIDs
sentByValue = model.Data.sentby if model.Data.sentby else ''
frmSentByOption = """
    <label for="sentby" id="sentby">Sent By (PeopleIDs):</label>
    <input type="text" name="sentby" id="sentby" value="{0}" placeholder="e.g. 123 or 123,456,789" style="width: 200px;">
    <span style="font-size: 11px; color: #666; margin-left: 5px;">Enter one or more PeopleIDs separated by commas</span>
""".format(sentByValue)

# Create text input for Minimum Recipients filter (for campaign performance)
minRecipientsValue = model.Data.minrecipients if model.Data.minrecipients else '100'
try:
    minRecipients = int(minRecipientsValue)
except:
    minRecipients = 100
frmMinRecipientsOption = """
    <label for="minrecipients">Min Recipients:</label>
    <input type="number" name="minrecipients" id="minrecipients" value="{0}" min="1" style="width: 80px;">
    <span style="font-size: 11px; color: #666; margin-left: 5px;">For campaign performance (show emails with at least this many recipients)</span>
""".format(minRecipients)


filterProgram = '' if model.Data.program == str(999999) else ' AND pro.Id = {}'.format(model.Data.program) if model.Data.program else ''
filterFailClassfication = '' if model.Data.failclassification == str(999999) else """ AND fe.Fail = '{}'""".format(model.Data.failclassification) if model.Data.failclassification else ''
filterFailClassficationOrgStat = '' if model.Data.failclassification == str(999999) else """ AND fe.Fail = ''{}''""".format(model.Data.failclassification) if model.Data.failclassification else ''

# Process Sent By filter - handle comma-separated list of PeopleIDs
filterSentBy = ''
filterSentByOrgStat = ''
if model.Data.sentby:
    # Clean and split the input
    sentByIds = [id.strip() for id in str(model.Data.sentby).split(',') if id.strip().isdigit()]
    if sentByIds:
        # Create IN clause for multiple IDs
        filterSentBy = ' AND eq.QueuedBy IN ({})'.format(','.join(sentByIds))
        filterSentByOrgStat = ' AND eq.QueuedBy IN ({})'.format(','.join(sentByIds))

        
#<p style="margin: 0; padding: 0;">

headerTemplate = '''
    <style>
        .dashboard-header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }}
        .dashboard-header h1 {{
            margin: 0;
            font-size: 24px;
        }}
        .dashboard-header p {{
            margin: 5px 0;
            opacity: 0.9;
        }}
        .filter-form {{
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }}
        .filter-form label {{
            margin-right: 10px;
            font-weight: bold;
        }}
        .filter-form input, .filter-form select {{
            margin-right: 15px;
        }}
        .stat-card {{
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .stat-card h3 {{
            margin-top: 0;
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }}
        .warning-box {{
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 10px;
            margin: 10px 0;
        }}
        .success-box {{
            background: #d4edda;
            border-left: 4px solid #28a745;
            padding: 10px;
            margin: 10px 0;
        }}
        .error-box {{
            background: #f8d7da;
            border-left: 4px solid #dc3545;
            padding: 10px;
            margin: 10px 0;
        }}
        .info-icon {{
            display: inline-block;
            width: 16px;
            height: 16px;
            background: #007bff;
            color: white;
            border-radius: 50%;
            text-align: center;
            line-height: 16px;
            font-size: 12px;
            cursor: help;
            margin-left: 5px;
        }}
        .failure-info {{
            background: #f0f8ff;
            border: 1px solid #007bff;
            border-radius: 8px;
            padding: 15px;
            margin: 20px 0;
        }}
        .failure-info h4 {{
            margin: 0 0 10px 0;
            color: #007bff;
        }}
        .failure-info ul {{
            margin: 5px 0;
            padding-left: 20px;
        }}
        .severity-critical {{
            color: #dc3545;
            font-weight: bold;
        }}
        .severity-high {{
            color: #fd7e14;
            font-weight: bold;
        }}
        .severity-medium {{
            color: #ffc107;
        }}
        .severity-low {{
            color: #28a745;
        }}
    </style>

    <div class="dashboard-header">
        <h1>Email Technical Diagnostics Dashboard</h1>
        <p>Comprehensive email delivery analysis and troubleshooting tool</p>
        <p style="font-size: 12px;">Report generated: {5}</p>
    </div>
    
    <div class="container" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
        <div style="font-size: 12px; color: #666;">
            <span class="severity-critical">●</span> Critical &nbsp;
            <span class="severity-high">●</span> High &nbsp;
            <span class="severity-medium">●</span> Medium &nbsp;
            <span class="severity-low">●</span> Low
        </div>
        <div>
            <a href="https://www.twilio.com/docs/sendgrid/ui/analytics-and-reporting/bounce-and-block-classifications" target="_blank" style="margin-right: 10px;">SendGrid Documentation</a>
            <a href="/Emails" target="_blank">Email Queue Management</a>
        </div>
    </div>
    
    <div class="filter-form">
        <form action="" method="GET">
            <div style="margin-bottom: 10px;">
                {4} {3}
            </div>
            <div style="margin-bottom: 10px;">
                {6}
            </div>
            <div style="margin-bottom: 10px;">
                {7}
            </div>
            <div style="margin-bottom: 10px;">
                <label for="sDate">Start Date:</label>
                <input type="date" id="sDate" name="sDate" required {0}>
                <label for="eDate">End Date:</label>
                <input type="date" id="eDate" name="eDate" required {1}>
            </div>
            <div style="margin-bottom: 10px;">
                <input type="checkbox" id="HideSuccess" name="HideSuccess" value="yes" {2}>
                <label for="HideSuccess" class="no-print">Show Successfully Sent <i>(This may impact performance)</i></label>
            </div>
            <input type="submit" value="Generate Report" style="background: #667eea; color: white; border: none; padding: 8px 20px; border-radius: 4px; cursor: pointer;">
        </form>
    </div>
'''.format(optionsDate,optioneDate,optionHideSuccess,frmProgramOption,frmFailClassificationsOption,current_date,frmSentByOption,frmMinRecipientsOption)



rsql = q.QuerySql(sql.format(sDate,eDate,sqlHideSuccess,filterProgram,filterFailClassfication,filterSentBy))

TotalEmails = 0
bodyTemplate = ''


####### Email Health Summary #######

# Calculate overall health metrics
total_emails = sum(row.TotalCount for row in rsql)
total_failed = sum(row.TotalCount for row in rsql if row.Status != 'Sent')
total_sent = sum(row.TotalCount for row in rsql if row.Status == 'Sent') if any(row.Status == 'Sent' for row in rsql) else 0
delivery_rate = (float(total_sent) / total_emails * 100) if total_emails > 0 else 0
failure_rate = (float(total_failed) / total_emails * 100) if total_emails > 0 else 0

# Identify failure types present in the data
failure_types_present = [row.Status for row in rsql if row.Status != 'Sent' and row.TotalCount > 0]

# Create health status summary
health_status = "success-box" if failure_rate < 5 else "warning-box" if failure_rate < 10 else "error-box"
health_message = "Excellent" if failure_rate < 5 else "Needs Attention" if failure_rate < 10 else "Critical"

# Format numbers for display
import locale
try:
    locale.setlocale(locale.LC_ALL, '')
except:
    pass

# Simple number formatting for Python 2
def format_number(n):
    """Format number with thousands separator"""
    if n >= 1000:
        return '{:,}'.format(int(n)) if hasattr('{:,}', 'format') else str(int(n))
    return str(int(n))

total_emails_str = format_number(total_emails)
total_sent_str = format_number(total_sent)
total_failed_str = format_number(total_failed)
delivery_rate_str = "%.1f" % delivery_rate
failure_rate_str = "%.1f" % failure_rate

bodyTemplate += '''
<div class="stat-card">
    <h3>Email Delivery Health Summary</h3>
    <div class="%s">
        <strong>Overall Status: %s</strong><br>
        Total Emails: %s<br>
        Successfully Delivered: %s (%s%%)<br>
        Failed/Bounced: %s (%s%%)
    </div>
</div>
''' % (health_status, health_message, total_emails_str, total_sent_str, 
       delivery_rate_str, total_failed_str, failure_rate_str)

# Add failure classification information if there are failures
if failure_types_present:
    bodyTemplate += '''
    <div class="failure-info">
        <h4>Failure Classifications Detected <span class="info-icon" title="Based on SendGrid bounce classifications">?</span></h4>
    '''
    
    for failure_type in failure_types_present:
        if failure_type in FAILURE_CLASSIFICATIONS:
            info = FAILURE_CLASSIFICATIONS[failure_type]
            severity_class = 'severity-' + info['severity']
            
            # Count for this failure type
            failure_count = sum(row.TotalCount for row in rsql if row.Status == failure_type)
            
            bodyTemplate += '''
        <details style="margin: 10px 0;">
            <summary style="cursor: pointer; font-weight: bold;">
                <span class="%s">%s</span> (%d occurrences)
            </summary>
            <div style="margin-left: 20px; margin-top: 10px;">
                <p><strong>Description:</strong> %s</p>
                <p><strong>Common Causes:</strong></p>
                <ul>
    ''' % (severity_class, failure_type, failure_count, info['description'])
            
            for cause in info['causes']:
                bodyTemplate += '<li>%s</li>' % cause
            
            bodyTemplate += '''
                </ul>
                <p><strong>Recommended Action:</strong> %s</p>
    ''' % info['remediation']
            
            # Add TouchPoint-specific note if available
            if 'touchpoint_note' in info:
                bodyTemplate += '''
                <div style="background: #fff3cd; border-left: 3px solid #ffc107; padding: 8px; margin-top: 10px;">
                    <strong>TouchPoint Note:</strong> %s
                </div>
    ''' % info['touchpoint_note']
            
            bodyTemplate += '''
            </div>
        </details>
    '''
    
    bodyTemplate += '</div>'

####### Total Stats #######

# Extract relevant attributes (ignoring anything that starts with "_")
data_list = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsql]


####advanced sub columns
#expanded_list=expanded_list 
#match_column="OrgId"

# Define all dictionaries separately
table_config = {
    "title": "<div class='stat-card'><h3>Email Status Breakdown</h3>",
    "table_width": "100%",
    "remove_borders": False,
}

column_settings = {
    "hide_columns": ['PeopleId'],
    "column_order": ['Status'],
    "divider_after_column": "", #Adds a vertical line the table to seperate
}

formatting = {
    "bold_columns": [''], 
    "date_columns": [''],  #supposed to set to date only w/o time, but needs work
    "sum_columns": ['TotalCount'],
    "row_colors": ["#f9f9f9", "#ffffff"],  # Alternating row colors
}

column_widths = {
    'some column': '20%',
}

url_columns = {
    '': '', #add's url link to columns. Example: 'Name': '/Person2/{PeopleId}'  Typically using in conjuction with hide_columns
}

header_rename = {
    "TotalCount": "Count",
    "Percentage": "% of Total",
}

header_style = {
    "header_padding": "6px 10px",
    "header_font_size": "14px",
    "header_bg_color": "#f0f0f0",
    "remove_header_borders": True,
}

content_style = {
    "content_padding": "3px 8px",
    "content_font_size": "12px",
}

column_alignment = {
    "Name": "left", 
    "Status": "left", 
}

slanted_headers_config = {
    "slanted_headers": [''], #kinda works, but not aligning 100%.  Still some css work to do
    "slant_angle": 45, #not working
}


# Merge all settings into a single dictionary
merged_settings = {}
merged_settings.update(table_config)
merged_settings.update(column_settings)
merged_settings.update(formatting)
merged_settings.update(header_style)
merged_settings.update(content_style)
merged_settings.update(slanted_headers_config)
merged_settings.update({"header_rename": header_rename})  
merged_settings.update({"column_alignment": column_alignment}) 

# Now pass everything at once (Python 2 allows only ONE **kwargs)
bodyTemplate += generate_html_table(
    data_list, 
    column_widths=column_widths, 
    url_columns=url_columns, 
    **merged_settings
)
bodyTemplate += "</div>"  # Close stat-card div

####### Bounce Rate Trend #######
if sDate and eDate:
    rsqlBounceRateTrend = q.QuerySql(sqlBounceRateTrend.format(sDate,eDate,sqlHideSuccess,filterProgram,filterFailClassfication,filterSentBy))
    if rsqlBounceRateTrend:
        trend_data = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlBounceRateTrend]
        
        # Create trend analysis
        trend_config = {
            "title": "<div class='stat-card'><h3>Daily Bounce Rate Trend</h3>",
            "table_width": "100%",
            "remove_borders": False,
        }
        
        trend_columns = {
            "hide_columns": [],
            "column_order": ['Date', 'TotalSent', 'TotalFailed', 'BounceRate'],
        }
        
        trend_formatting = {
            "bold_columns": ['BounceRate'],
            "date_columns": ['Date'],
            "row_colors": ["#f9f9f9", "#ffffff"],
        }
        
        trend_rename = {
            "TotalSent": "Sent",
            "TotalFailed": "Failed",
            "BounceRate": "Bounce %",
        }
        
        trend_alignment = {
            "Date": "left",
            "TotalSent": "right",
            "TotalFailed": "right",
            "BounceRate": "right",
        }
        
        merged_trend = {}
        merged_trend.update(trend_config)
        merged_trend.update(trend_columns)
        merged_trend.update(trend_formatting)
        merged_trend.update({"header_rename": trend_rename})
        merged_trend.update({"column_alignment": trend_alignment})
        
        bodyTemplate += generate_html_table(trend_data, **merged_trend)
        bodyTemplate += "</div>"

####### Domain Analysis #######
rsqlDomainAnalysis = q.QuerySql(sqlDomainAnalysis.format(sDate,eDate,sqlHideSuccess,filterProgram,filterFailClassfication,filterSentBy))
if rsqlDomainAnalysis:
    domain_data = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlDomainAnalysis]
    
    # Find problematic domains
    problematic_domains = [row for row in domain_data if float(row.get('FailureRate', 0)) > 10]
    
    if problematic_domains:
        bodyTemplate += '''
        <div class="stat-card">
            <h3>Domain Reputation Analysis</h3>
            <div class="warning-box">
                <strong>Warning:</strong> {0} domain(s) showing high failure rates (>10%)
            </div>
        '''.format(len(problematic_domains))
        
        domain_config = {
            "table_width": "100%",
            "remove_borders": False,
        }
        
        domain_columns = {
            "column_order": ['Domain', 'TotalRecipients', 'FailedEmails', 'FailureRate'],
        }
        
        domain_rename = {
            "TotalRecipients": "Recipients",
            "FailedEmails": "Failed",
            "FailureRate": "Failure %",
        }
        
        domain_alignment = {
            "Domain": "left",
            "TotalRecipients": "right",
            "FailedEmails": "right",
            "FailureRate": "right",
        }
        
        merged_domain = {}
        merged_domain.update(domain_config)
        merged_domain.update(domain_columns)
        merged_domain.update({"header_rename": domain_rename})
        merged_domain.update({"column_alignment": domain_alignment})
        merged_domain.update({"row_colors": ["#fff3cd", "#ffffff"]})
        
        bodyTemplate += generate_html_table(problematic_domains[:10], **merged_domain)  # Show top 10
        bodyTemplate += "</div>"

####### Email Queue Status #######
rsqlEmailQueueStatus = q.QuerySql(sqlEmailQueueStatus)
rsqlEmailQueueDetails = q.QuerySql(sqlEmailQueueDetails) if rsqlEmailQueueStatus else None

if rsqlEmailQueueStatus:
    queue_data = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlEmailQueueStatus]
    
    if queue_data:
        total_pending = sum(row.get('PendingEmails', 0) for row in queue_data)
        
        if total_pending > 0:
            # Find old emails (>24 hours)
            old_emails = [row for row in queue_data if row.get('HoursOld', 0) > 24]
            very_old_emails = [row for row in queue_data if row.get('HoursOld', 0) > 72]
            
            bodyTemplate += '''
            <div class="stat-card">
                <h3>Email Queue Status</h3>
            '''
            
            # Add alerts for old emails
            if very_old_emails:
                bodyTemplate += '''
                <div class="error-box">
                    <strong>Critical:</strong> {0} sender(s) have emails older than 72 hours in queue!
                </div>
                '''.format(len(very_old_emails))
            elif old_emails:
                bodyTemplate += '''
                <div class="warning-box">
                    <strong>Warning:</strong> {0} sender(s) have emails older than 24 hours in queue
                </div>
                '''.format(len(old_emails))
            else:
                bodyTemplate += '''
                <div class="success-box">
                    <strong>Queue Status:</strong> {0} emails pending (all recent)
                </div>
                '''.format(total_pending)
            
            # Summary by sender with age information
            bodyTemplate += '<h4>Queue Summary by Sender</h4>'
            
            queue_config = {
                "table_width": "100%",
                "remove_borders": False,
            }
            
            queue_columns = {
                "hide_columns": ['QueuedBy', 'EmailIds'],
                "column_order": ['QueuedByName', 'PendingEmails', 'HoursOld', 'OldestQueued', 'NewestQueued'],
            }
            
            queue_rename = {
                "QueuedByName": "Queued By",
                "PendingEmails": "Count",
                "HoursOld": "Hours Waiting",
                "OldestQueued": "Oldest",
                "NewestQueued": "Newest",
            }
            
            # Add URL for viewing emails
            queue_url_columns = {
                "QueuedByName": '/Person2/{QueuedBy}'
            }
            
            merged_queue = {}
            merged_queue.update(queue_config)
            merged_queue.update(queue_columns)
            merged_queue.update({"header_rename": queue_rename})
            merged_queue.update({"date_columns": ['OldestQueued', 'NewestQueued']})
            merged_queue.update({"bold_columns": ['HoursOld']})
            
            bodyTemplate += generate_html_table(queue_data, url_columns=queue_url_columns, **merged_queue)
            
            # Show detailed list of oldest emails if there are old ones
            if rsqlEmailQueueDetails:
                detail_data = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlEmailQueueDetails]
                
                if detail_data:
                    bodyTemplate += '<h4>Oldest Queued Emails (Top 10)</h4>'
                    
                    detail_config = {
                        "table_width": "100%",
                        "remove_borders": False,
                    }
                    
                    detail_columns = {
                        "hide_columns": ['QueuedBy', 'SendWhen'],
                        "column_order": ['Id', 'Subject', 'FailureReason', 'QueuedByName', 'RecipientCount', 'HoursOld', 'Queued'],
                    }
                    
                    detail_rename = {
                        "Id": "Email ID",
                        "Subject": "Subject",
                        "FailureReason": "Issue",
                        "QueuedByName": "Sender",
                        "RecipientCount": "Recipients",
                        "HoursOld": "Hours Old",
                        "Queued": "Queued Date",
                    }
                    
                    detail_url_columns = {
                        "Id": '/Emails/Details/{Id}',
                        "QueuedByName": '/Person2/{QueuedBy}'
                    }
                    
                    merged_detail = {}
                    merged_detail.update(detail_config)
                    merged_detail.update(detail_columns)
                    merged_detail.update({"header_rename": detail_rename})
                    merged_detail.update({"date_columns": ['Queued']})
                    merged_detail.update({"bold_columns": ['HoursOld']})
                    
                    bodyTemplate += generate_html_table(detail_data, url_columns=detail_url_columns, **merged_detail)
            
            bodyTemplate += '''
                <div style="margin-top: 15px; padding: 10px; background: #f0f8ff; border-radius: 5px;">
                    <strong>Troubleshooting Tips:</strong>
                    <ul style="margin: 5px 0;">
                        <li>Emails older than 24 hours may indicate a sending issue</li>
                        <li>Check if the sender's account is properly configured</li>
                        <li>Very old emails (72+ hours) should be investigated or cancelled</li>
                        <li>Click on Email IDs to view details and manage individual emails</li>
                        <li>Visit <a href="/Emails" target="_blank">Email Queue</a> to manage all pending emails</li>
                    </ul>
                </div>
                
                <div style="margin-top: 10px; padding: 10px; background: #fff3cd; border-left: 3px solid #ffc107; border-radius: 5px;">
                    <strong>⚠️ Known Issue:</strong> Some emails in the queue may not have a sent date and appear as "not sent" even though they may have been delivered. 
                    This is a known issue that has been reported to TouchPoint support. We are currently awaiting their response and resolution. 
                    In the meantime, check the recipient's email history or SendGrid logs for actual delivery status.
                </div>
            '''
            
            bodyTemplate += "</div>"

####### Campaign Send Performance #######
# Helper function to format seconds into human-readable duration
def format_duration(seconds):
    """Convert seconds to a human-readable format"""
    if seconds is None:
        return "N/A"
    seconds = int(seconds)
    if seconds < 60:
        return "{0} sec".format(seconds)
    elif seconds < 3600:
        minutes = seconds / 60
        return "{0} min".format(int(minutes))
    else:
        hours = seconds / 3600.0
        return "{0:.1f} hrs".format(hours)

if sDate and eDate:
    # Get campaign performance summary stats
    rsqlCampaignSummary = q.QuerySql(sqlCampaignPerformanceSummary.format(sDate, eDate, filterProgram, filterFailClassfication, filterSentBy, minRecipients))

    # Safe check for results - TotalCampaigns may be None
    hasCampaignData = False
    totalCampaigns = 0
    summary = None
    summaryList = []

    # Convert to list and safely extract data - wrap in try/except for .NET errors
    try:
        if rsqlCampaignSummary:
            for row in rsqlCampaignSummary:
                rowDict = {}
                for attr in dir(row):
                    if not attr.startswith("_"):
                        try:
                            rowDict[attr] = getattr(row, attr, None)
                        except BaseException:
                            rowDict[attr] = None
                summaryList.append(rowDict)
    except BaseException:
        summaryList = []

    if len(summaryList) > 0:
        summary = summaryList[0]
        totalCampaigns = summary.get('TotalCampaigns', 0) or 0
        if totalCampaigns > 0:
            hasCampaignData = True

    if hasCampaignData:
        # summary is already a dictionary from summaryList[0]

        bodyTemplate += '''
        <div class="stat-card">
            <h3>Email Campaign Send Performance</h3>
            <p style="font-size: 12px; color: #666; margin-bottom: 15px;">
                Showing campaigns with {0}+ recipients | Measures time from queue to send completion
            </p>

            <div style="display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 20px;">
                <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">
                    <div style="font-size: 28px; font-weight: bold; color: #667eea;">{1}</div>
                    <div style="font-size: 12px; color: #666;">Total Campaigns</div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">
                    <div style="font-size: 28px; font-weight: bold; color: #667eea;">{2}</div>
                    <div style="font-size: 12px; color: #666;">Total Recipients</div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center;">
                    <div style="font-size: 28px; font-weight: bold; color: #667eea;">{3}</div>
                    <div style="font-size: 12px; color: #666;">Avg Recipients/Email</div>
                </div>
            </div>

            <div style="display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 20px;">
                <div style="flex: 1; min-width: 150px; background: #d4edda; padding: 15px; border-radius: 8px; text-align: center;">
                    <div style="font-size: 24px; font-weight: bold; color: #28a745;">{4}</div>
                    <div style="font-size: 12px; color: #666;">Avg Time to Complete</div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #d1ecf1; padding: 15px; border-radius: 8px; text-align: center;">
                    <div style="font-size: 24px; font-weight: bold; color: #17a2b8;">{5}</div>
                    <div style="font-size: 12px; color: #666;">Fastest Send</div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #fff3cd; padding: 15px; border-radius: 8px; text-align: center;">
                    <div style="font-size: 24px; font-weight: bold; color: #856404;">{6}</div>
                    <div style="font-size: 12px; color: #666;">Slowest Send</div>
                </div>
            </div>
        '''.format(
            minRecipients,
            summary.get('TotalCampaigns', 0),
            int(summary.get('TotalRecipients', 0) or 0),
            int(summary.get('AvgRecipients', 0) or 0),
            format_duration(summary.get('AvgSendDurationSeconds')),
            format_duration(summary.get('FastestSendSeconds')),
            format_duration(summary.get('SlowestSendSeconds'))
        )

        # Get monthly trend
        rsqlMonthlyTrend = q.QuerySql(sqlCampaignMonthlyTrend.format(filterProgram, filterFailClassfication, filterSentBy, minRecipients))

        if rsqlMonthlyTrend:
            trend_data = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlMonthlyTrend]

            # Convert seconds to readable format for display
            for row in trend_data:
                row['AvgSendTime'] = format_duration(row.get('AvgSendSeconds'))
                row['FastestSend'] = format_duration(row.get('FastestSendSeconds'))
                row['SlowestSend'] = format_duration(row.get('SlowestSendSeconds'))

            bodyTemplate += '<h4>Monthly Campaign Performance Trend (Last 12 Months)</h4>'

            trend_config = {
                "table_width": "100%",
                "remove_borders": False,
            }

            trend_columns = {
                "hide_columns": ['AvgSendSeconds', 'FastestSendSeconds', 'SlowestSendSeconds'],
                "column_order": ['Month', 'Campaigns', 'TotalRecipients', 'AvgRecipients', 'AvgSendTime', 'FastestSend', 'SlowestSend'],
            }

            trend_rename = {
                "TotalRecipients": "Total Recipients",
                "AvgRecipients": "Avg Recipients",
                "AvgSendTime": "Avg Send Time",
                "FastestSend": "Fastest",
                "SlowestSend": "Slowest",
            }

            trend_alignment = {
                "Month": "left",
                "Campaigns": "right",
                "TotalRecipients": "right",
                "AvgRecipients": "right",
                "AvgSendTime": "right",
                "FastestSend": "right",
                "SlowestSend": "right",
            }

            merged_trend = {}
            merged_trend.update(trend_config)
            merged_trend.update(trend_columns)
            merged_trend.update({"header_rename": trend_rename})
            merged_trend.update({"column_alignment": trend_alignment})
            merged_trend.update({"row_colors": ["#f9f9f9", "#ffffff"]})

            bodyTemplate += generate_html_table(trend_data, **merged_trend)

        # Get recipient bucket analysis
        try:
            rsqlRecipientBucket = q.QuerySql(sqlRecipientBucketAnalysis.format(sDate, eDate, filterProgram, filterFailClassfication, filterSentBy, minRecipients))
            if rsqlRecipientBucket:
                bucket_data = []
                for row in rsqlRecipientBucket:
                    rowDict = {}
                    for attr in dir(row):
                        if not attr.startswith("_"):
                            try:
                                rowDict[attr] = getattr(row, attr, None)
                            except BaseException:
                                rowDict[attr] = None
                    bucket_data.append(rowDict)

                if len(bucket_data) > 0:
                    # Find max avg send time for scaling the bar chart
                    max_avg_time = max([row.get('AvgSendSeconds', 0) or 0 for row in bucket_data])
                    if max_avg_time == 0:
                        max_avg_time = 1

                    bodyTemplate += '''
                    <h4>Send Time by Recipient Count (Size Analysis)</h4>
                    <p style="font-size: 11px; color: #666;">Does email size (by recipients) affect send time?</p>
                    <div style="margin: 15px 0;">
                    '''

                    for row in bucket_data:
                        bucket_name = row.get('RecipientBucket', 'Unknown')
                        # Remove the sorting prefix (e.g., "1. " from "1. 100-250")
                        display_name = bucket_name[3:] if len(bucket_name) > 3 else bucket_name
                        campaign_count = row.get('CampaignCount', 0) or 0
                        avg_recipients = int(row.get('AvgRecipients', 0) or 0)
                        avg_seconds = row.get('AvgSendSeconds', 0) or 0
                        fastest = row.get('FastestSendSeconds', 0) or 0
                        slowest = row.get('SlowestSendSeconds', 0) or 0
                        bar_width = int((avg_seconds / max_avg_time) * 100) if max_avg_time > 0 else 0

                        bodyTemplate += '''
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 120px; font-weight: bold; font-size: 12px;">{0} recipients</div>
                            <div style="flex: 1; background: #e9ecef; border-radius: 4px; height: 28px; position: relative;">
                                <div style="width: {1}%; background: linear-gradient(90deg, #667eea, #764ba2); height: 100%; border-radius: 4px; min-width: 2px;"></div>
                                <span style="position: absolute; left: 10px; top: 50%; transform: translateY(-50%); font-size: 11px; color: #333; font-weight: bold;">{2}</span>
                            </div>
                            <div style="width: 180px; text-align: right; font-size: 11px; color: #666; margin-left: 10px;">
                                {3} campaigns | Fastest: {4} | Slowest: {5}
                            </div>
                        </div>
                        '''.format(
                            display_name,
                            bar_width,
                            format_duration(avg_seconds),
                            campaign_count,
                            format_duration(fastest),
                            format_duration(slowest)
                        )

                    bodyTemplate += '</div>'

                    # Also show as a table for detailed view
                    for row in bucket_data:
                        row['AvgSendTime'] = format_duration(row.get('AvgSendSeconds'))
                        row['FastestSend'] = format_duration(row.get('FastestSendSeconds'))
                        row['SlowestSend'] = format_duration(row.get('SlowestSendSeconds'))
                        row['AvgBodySizeKB'] = '{:.1f} KB'.format((row.get('AvgBodySize', 0) or 0) / 1024.0)

                    bucket_config = {
                        "table_width": "100%",
                        "remove_borders": False,
                    }
                    bucket_columns = {
                        "hide_columns": ['AvgSendSeconds', 'FastestSendSeconds', 'SlowestSendSeconds', 'AvgBodySize'],
                        "column_order": ['RecipientBucket', 'CampaignCount', 'AvgRecipients', 'AvgSendTime', 'FastestSend', 'SlowestSend', 'AvgBodySizeKB'],
                    }
                    bucket_rename = {
                        "RecipientBucket": "Recipient Range",
                        "CampaignCount": "Campaigns",
                        "AvgRecipients": "Avg Recipients",
                        "AvgSendTime": "Avg Send Time",
                        "FastestSend": "Fastest",
                        "SlowestSend": "Slowest",
                        "AvgBodySizeKB": "Avg Email Size",
                    }
                    bucket_alignment = {
                        "RecipientBucket": "left",
                        "CampaignCount": "right",
                        "AvgRecipients": "right",
                        "AvgSendTime": "right",
                        "FastestSend": "right",
                        "SlowestSend": "right",
                        "AvgBodySizeKB": "right",
                    }
                    merged_bucket = {}
                    merged_bucket.update(bucket_config)
                    merged_bucket.update(bucket_columns)
                    merged_bucket.update({"header_rename": bucket_rename})
                    merged_bucket.update({"column_alignment": bucket_alignment})
                    merged_bucket.update({"row_colors": ["#f9f9f9", "#ffffff"]})

                    bodyTemplate += generate_html_table(bucket_data, **merged_bucket)
        except BaseException:
            pass  # Skip bucket analysis if there's an error

        # Get body size analysis
        try:
            rsqlBodySize = q.QuerySql(sqlBodySizeAnalysis.format(sDate, eDate, filterProgram, filterFailClassfication, filterSentBy, minRecipients))
            if rsqlBodySize:
                body_data = []
                for row in rsqlBodySize:
                    rowDict = {}
                    for attr in dir(row):
                        if not attr.startswith("_"):
                            try:
                                rowDict[attr] = getattr(row, attr, None)
                            except BaseException:
                                rowDict[attr] = None
                    body_data.append(rowDict)

                if len(body_data) > 0:
                    # Find max avg send time for scaling the bar chart
                    max_avg_time = max([row.get('AvgSendSeconds', 0) or 0 for row in body_data])
                    if max_avg_time == 0:
                        max_avg_time = 1

                    bodyTemplate += '''
                    <h4>Send Time by Email Body Size</h4>
                    <p style="font-size: 11px; color: #666;">Does email content size affect send time? (Body size includes base64 embedded images)</p>
                    <div style="margin: 15px 0;">
                    '''

                    for row in body_data:
                        bucket_name = row.get('SizeBucket', 'Unknown')
                        display_name = bucket_name[3:] if len(bucket_name) > 3 else bucket_name
                        campaign_count = row.get('CampaignCount', 0) or 0
                        avg_recipients = int(row.get('AvgRecipients', 0) or 0)
                        avg_seconds = row.get('AvgSendSeconds', 0) or 0
                        avg_body_kb = (row.get('AvgBodySizeBytes', 0) or 0) / 1024.0
                        avg_images = int(row.get('AvgImageCount', 0) or 0)
                        base64_count = int(row.get('EmailsWithBase64Images', 0) or 0)
                        bar_width = int((avg_seconds / max_avg_time) * 100) if max_avg_time > 0 else 0

                        bodyTemplate += '''
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 180px; font-weight: bold; font-size: 12px;">{0}</div>
                            <div style="flex: 1; background: #e9ecef; border-radius: 4px; height: 28px; position: relative;">
                                <div style="width: {1}%; background: linear-gradient(90deg, #28a745, #20c997); height: 100%; border-radius: 4px; min-width: 2px;"></div>
                                <span style="position: absolute; left: 10px; top: 50%; transform: translateY(-50%); font-size: 11px; color: #333; font-weight: bold;">{2}</span>
                            </div>
                            <div style="width: 280px; text-align: right; font-size: 11px; color: #666; margin-left: 10px;">
                                {3} emails | {4:.0f} KB | ~{5} imgs | {6} embedded
                            </div>
                        </div>
                        '''.format(
                            display_name,
                            bar_width,
                            format_duration(avg_seconds),
                            campaign_count,
                            avg_body_kb,
                            avg_images,
                            base64_count
                        )

                    bodyTemplate += '</div>'

                    # Info box about image types
                    bodyTemplate += '''
                    <div style="margin: 10px 0; padding: 10px; background: #fff3cd; border-left: 3px solid #ffc107; font-size: 11px;">
                        <strong>Note:</strong> Body size includes base64 embedded images (can be 1-4MB each!).
                        Linked/external images only add ~100 bytes per URL. "Huge" emails likely have embedded images.
                    </div>
                    '''

                    # Also show as a table
                    for row in body_data:
                        row['AvgSendTime'] = format_duration(row.get('AvgSendSeconds'))
                        row['FastestSend'] = format_duration(row.get('FastestSendSeconds'))
                        row['SlowestSend'] = format_duration(row.get('SlowestSendSeconds'))
                        row['AvgSizeKB'] = '{:.1f} KB'.format((row.get('AvgBodySizeBytes', 0) or 0) / 1024.0)
                        row['AvgImages'] = int(row.get('AvgImageCount', 0) or 0)
                        row['Base64Count'] = int(row.get('EmailsWithBase64Images', 0) or 0)

                    body_config = {
                        "table_width": "100%",
                        "remove_borders": False,
                    }
                    body_columns = {
                        "hide_columns": ['AvgSendSeconds', 'FastestSendSeconds', 'SlowestSendSeconds', 'AvgBodySizeBytes', 'MinBodySize', 'MaxBodySize', 'AvgImageCount', 'EmailsWithBase64Images'],
                        "column_order": ['SizeBucket', 'CampaignCount', 'AvgRecipients', 'AvgSizeKB', 'AvgImages', 'Base64Count', 'AvgSendTime', 'FastestSend', 'SlowestSend'],
                    }
                    body_rename = {
                        "SizeBucket": "Email Size",
                        "CampaignCount": "Campaigns",
                        "AvgRecipients": "Avg Recipients",
                        "AvgSizeKB": "Avg Size",
                        "AvgImages": "Avg Images",
                        "Base64Count": "Has Embedded",
                        "AvgSendTime": "Avg Send Time",
                        "FastestSend": "Fastest",
                        "SlowestSend": "Slowest",
                    }
                    body_alignment = {
                        "SizeBucket": "left",
                        "CampaignCount": "right",
                        "AvgRecipients": "right",
                        "AvgSizeKB": "right",
                        "AvgImages": "right",
                        "Base64Count": "right",
                        "AvgSendTime": "right",
                        "FastestSend": "right",
                        "SlowestSend": "right",
                    }
                    merged_body = {}
                    merged_body.update(body_config)
                    merged_body.update(body_columns)
                    merged_body.update({"header_rename": body_rename})
                    merged_body.update({"column_alignment": body_alignment})
                    merged_body.update({"row_colors": ["#f9f9f9", "#ffffff"]})

                    bodyTemplate += generate_html_table(body_data, **merged_body)
        except BaseException:
            pass  # Skip body size analysis if there's an error

        # Get individual campaign details
        rsqlCampaignPerf = q.QuerySql(sqlCampaignPerformance.format(sDate, eDate, filterProgram, filterFailClassfication, filterSentBy, minRecipients))

        if rsqlCampaignPerf:
            campaign_data = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlCampaignPerf]

            # Limit to top 50 campaigns for performance
            campaign_data = campaign_data[:50]

            bodyTemplate += '<h4>Recent Campaign Details (Top 50)</h4>'

            campaign_config = {
                "table_width": "100%",
                "remove_borders": False,
            }

            campaign_columns = {
                "hide_columns": ['SendDurationSeconds', 'Queued', 'BodySizeBytes', 'SendWhen', 'Id'],
                "column_order": ['Subject', 'SentBy', 'Recipients', 'BodySizeKB', 'ImageCount', 'HasEmbeddedImages', 'SendDuration', 'SendType', 'CompletedAt'],
            }

            campaign_rename = {
                "Subject": "Email Subject",
                "SentBy": "Sent By",
                "Recipients": "Recipients",
                "BodySizeKB": "Size (KB)",
                "ImageCount": "Images",
                "HasEmbeddedImages": "Embedded?",
                "SendDuration": "Send Time",
                "SendType": "Type",
                "CompletedAt": "Completed",
            }

            campaign_url_columns = {
                "Subject": '/Manage/Emails/Details/{Id}'
            }

            campaign_alignment = {
                "Subject": "left",
                "SentBy": "left",
                "Recipients": "right",
                "BodySizeKB": "right",
                "ImageCount": "right",
                "HasEmbeddedImages": "center",
                "SendDuration": "right",
                "SendType": "center",
                "CompletedAt": "left",
            }

            merged_campaign = {}
            merged_campaign.update(campaign_config)
            merged_campaign.update(campaign_columns)
            merged_campaign.update({"header_rename": campaign_rename})
            merged_campaign.update({"column_alignment": campaign_alignment})
            merged_campaign.update({"date_columns": ['CompletedAt']})
            merged_campaign.update({"row_colors": ["#f9f9f9", "#ffffff"]})

            bodyTemplate += generate_html_table(campaign_data, url_columns=campaign_url_columns, **merged_campaign)

        bodyTemplate += '''
            <div style="margin-top: 15px; padding: 10px; background: #f0f8ff; border-radius: 5px;">
                <strong>Understanding Campaign Performance:</strong>
                <ul style="margin: 5px 0;">
                    <li><strong>Send Time</strong> = Time from when email was queued/scheduled to when sending completed</li>
                    <li><strong>Size (KB)</strong> = Email body size including HTML and any base64 embedded images</li>
                    <li><strong>Images</strong> = Count of &lt;img&gt; tags (both linked and embedded)</li>
                    <li><strong>Embedded?</strong> = "Yes" if email contains base64 embedded images (can be 1-4MB each!)</li>
                    <li><strong>Type</strong> = Scheduled (set to send later) vs Immediate (sent right away)</li>
                    <li>Linked images (URLs) only add ~100 bytes each. Embedded images add the full image size.</li>
                </ul>
            </div>
        '''

        bodyTemplate += "</div>"
    else:
        bodyTemplate += '''
        <div class="stat-card">
            <h3>Email Campaign Send Performance</h3>
            <p style="color: #666;">No campaigns found with {0}+ recipients in the selected date range. Try lowering the minimum recipients filter.</p>
        </div>
        '''.format(minRecipients)

####### Involvement Stats #######
rsqlOrgStat = q.QuerySql(sqlOrgStat.format(sDate,eDate,sqlHideSuccess,filterProgram,filterFailClassficationOrgStat,filterSentByOrgStat))
#rsqlOrgDetails = q.QuerySql(sqlOrgDetails.format(sDate,eDate,sqlHideSuccess))


# Extract relevant attributes (ignoring anything that starts with "_")
data_list = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlOrgStat]

# Define all dictionaries separately
table_config = {
    "title": "<div class='stat-card'><h3>Organization Email Statistics</h3>",
    "table_width": "100%",
    "remove_borders": False,
}

column_settings = {
    "hide_columns": ['OrgId', 'id'],
    "column_order": ['Program', 'OrganizationName', 'Subject', 'SentDate'],
    "divider_after_column": "OrganizationName",
}

formatting = {
    "bold_columns": ['Program'],
    "date_columns": ['Sent'], #Removes the time
    "sum_columns": ['Technical', 'Invalid Address', 'bouncedaddress', 'spamreport', 
                    'Unclassified', 'Mailbox Unavailable', 'Reputation', 'Content', 
                    'invalid', 'spamreporting'],
    "row_colors": ["#f9f9f9", "#ffffff"],  # Alternating row colors
}

column_widths = {
    'OrgId': '20%',
    'Program': '150px',
    'SentDate': 'auto',
    'TotalCount': '100px',
}

url_columns = {
    'Involvement': '/Org/{OrgId}',
    'Email Subject': '/Manage/Emails/Details/{id}',
}

header_rename = {
    "OrganizationName": "Involvement",
    "Subject": "Email Subject",
    "SentDate": "Sent",
}

header_style = {
    "header_padding": "8px 12px",
    "header_font_size": "14px",
    "header_bg_color": "#FFFFFF",
    "remove_header_borders": True,
}

content_style = {
    "content_padding": "5px 10px",
    "content_font_size": "12px",
}

column_alignment = {
    "Program": "left", 
    "OrganizationName": "left", 
    "Subject": "left", 
    "SentDate": "left", 
}

slanted_headers_config = {
    "slanted_headers": ['Technical', 'Invalid Address', 'bouncedaddress', 'spamreport', 
                        'Unclassified', 'Mailbox Unavailable', 'Reputation', 'Content', 
                        'invalid', 'spamreporting'],
    "slant_angle": 45,
}

# Merge all settings into a single dictionary
merged_settings = {}
merged_settings.update(table_config)
merged_settings.update(column_settings)
merged_settings.update(formatting)
merged_settings.update(header_style)
merged_settings.update(content_style)
merged_settings.update(slanted_headers_config)
merged_settings.update({"header_rename": header_rename})  
merged_settings.update({"column_alignment": column_alignment}) 

# Now pass everything at once (Python 2 allows only ONE **kwargs)
bodyTemplate += generate_html_table(
    data_list, 
    column_widths=column_widths, 
    url_columns=url_columns, 
    **merged_settings
)
bodyTemplate += "</div>"  # Close stat-card div


###### User Failed Stats  ######
rsqlUserFailedStat = q.QuerySql(sqlUserFailedStat.format(sDate,eDate,sqlHideSuccess,filterProgram,filterFailClassfication,filterSentBy))

# Extract relevant attributes (ignoring anything that starts with "_")
data_list = [{attr: getattr(row, attr) for attr in dir(row) if not attr.startswith("_")} for row in rsqlUserFailedStat]

# Define all dictionaries separately
table_config = {
    "title": "<div class='stat-card'><h3>User Email Failure Details</h3>",
    "table_width": "100%",
    "remove_borders": False,
}

column_settings = {
    "hide_columns": ['PeopleId'],
    "column_order": ['Name', 'TotalCount','Status'],
    "divider_after_column": "",  # Adds a vertical line the table to separate
}

formatting = {
    "bold_columns": [''],
    "date_columns": [''],  # Supposed to set to date only w/o time, but needs work
    "sum_columns": ['Total Count'],
    "row_colors": ["#f9f9f9", "#ffffff"],  # Alternating row colors
}

column_widths = {
    'some column': '20%',
}

url_columns = {
    'Name': '/Person2/{PeopleId}',  # Adds url link to columns, typically used in conjunction with hide_columns
}

header_rename = {
    "Organization": "Involvement",
}

header_style = {
    "header_padding": "6px 10px",
    "header_font_size": "14px",
    "header_bg_color": "#f0f0f0",
    "remove_header_borders": True,
}

content_style = {
    "content_padding": "3px 8px",
    "content_font_size": "12px",
}

column_alignment = {
    "Name": "left", 
    "Status": "left", 
}

slanted_headers_config = {
    "slanted_headers": [''],  # Kinda works, but not aligning 100%. Still needs some CSS work
    "slant_angle": 45,  # Not working yet
}



# Merge all settings into a single dictionary
merged_settings = {}
merged_settings.update(table_config)
merged_settings.update(column_settings)
merged_settings.update(formatting)
merged_settings.update(header_style)
merged_settings.update(content_style)
merged_settings.update(slanted_headers_config)
merged_settings.update({"header_rename": header_rename})  
merged_settings.update({"column_alignment": column_alignment}) 


# Now pass everything at once (Python 2 allows only ONE **kwargs)
bodyTemplate += generate_html_table(
    data_list, 
    column_widths=column_widths, 
    url_columns=url_columns, 
    **merged_settings
)
bodyTemplate += "</div>"  # Close stat-card div

# Add footer with additional troubleshooting resources
footerTemplate = '''
<div class="stat-card" style="margin-top: 30px; background: #f8f9fa;">
    <h3>Troubleshooting Resources</h3>
    <ul style="margin: 10px 0;">
        <li><a href="/Emails" target="_blank">Email Queue Management</a></li>
        <li><a href="/Reports" target="_blank">Additional Reports</a></li>
        <li><a href="https://www.twilio.com/docs/sendgrid/ui/analytics-and-reporting" target="_blank">SendGrid Analytics Documentation</a></li>
        <li><a href="https://docs.touchpointsoftware.com/EmailTroubleshooting" target="_blank">TouchPoint Email Troubleshooting Guide</a></li>
    </ul>
    <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #dee2e6; font-size: 11px; color: #6c757d;">
        Report Period: {0} to {1}<br>
        Generated: {2}<br>
        Filters Applied: Program={3}, Classification={4}, Sent By={5}, Min Recipients={7}, Hide Success={6}
    </div>
</div>
'''.format(
    sDate if sDate else 'Not specified',
    eDate if eDate else 'Not specified',
    current_date,
    'All' if model.Data.program == str(999999) or not model.Data.program else model.Data.program,
    'All' if model.Data.failclassification == str(999999) or not model.Data.failclassification else model.Data.failclassification,
    model.Data.sentby if model.Data.sentby else 'All',
    'No' if model.Data.HideSuccess == 'yes' else 'Yes',
    minRecipients
)

Report = model.RenderTemplate(headerTemplate)
Report += model.RenderTemplate(bodyTemplate)
Report += model.RenderTemplate(footerTemplate)
print Report
