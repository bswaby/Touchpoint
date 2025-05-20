# -*- coding: utf-8 -*-
"""
--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File  
3. Name the Python "ModernPaymentManager" and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Modern Payment Manager System
============================
A complete, responsive payment management system with:
- Program/Division/Payer drill-down navigation
- Integrated payment processing (no external redirects)
- Email history and preview functionality
- Transaction history with filtering
- AJAX-based user experience
- TouchPoint-compatible design

Written By: Ben Swaby
Email: bswaby@fbchtn.org
Version: 2.0
Date: May 2025
"""

from datetime import datetime
import sys
import ast
import re 
import json
from collections import defaultdict

# General Settings
DEFAULT_PROGRAM_ID = 0
PAGE_TITLE = "Modern Payment Manager"

# Payment Settings
PAYMENT_LINK_DESCRIPTION = "CreditCardLinkSent"
CREDIT_CHARGE_PAYMENT_TYPE = "CreditCharge"
CHECK_PAYMENT_TYPE = "CHK|"
CASH_PAYMENT_TYPE = "CSH|" 

# Email Settings
EMAIL_SUBJECT_PREFIX = "Payment Notification - "
DEFAULT_EMAIL_SENDER = {
    'sender_id': '3134',
    'sender_email': 'noreply@church.com',
    'sender_alias': 'Payment System',
    'email_title': 'Payment Notification',
    'sender_phone': '(555) 123-4567'
}

# Template Settings >>> See "EMAIL TEMPLATE SETUP INSTRUCTIONS" below for template usage <<<
PAYMENT_CONFIRMATION_TEMPLATE_NAME = "PM_PaymentMade_Email"
PAYMENT_NOTIFICATION_TEMPLATE_NAME = "PM_Notification_Email"

# Email Templates (default templates if not found in HTML Content)
DEFAULT_PAYMENT_EMAIL_TEMPLATE = """
<p>Dear {name},</p>
<p>You have an outstanding balance of {totalDue}.</p>
<p>Previous Balance: {previousDue}<br>
   New Charges: {chargeNotes}<br>
   <strong>Total Due: {totalDue}</strong></p>
<p>Please use the following link to make your payment:</p>
<p><a href="{paylink}">Pay Now</a></p>
<p>Thank you,<br>{sender_alias}<br>{sender_phone}<br>{sender_email}</p>
"""

DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE = """
<p>Dear {name},</p>
<p>We have received your payment. Thank you!</p>
<p>Payment Details: {chargeNotes}<br>
   Previous Balance: {previousDue}<br>
   <strong>New Balance: {newTotalDue}</strong></p>
<p>Thank you,<br>{sender_alias}<br>{sender_phone}<br>{sender_email}</p>
"""

"""
-----------------------------------
EMAIL TEMPLATE SETUP INSTRUCTIONS
-----------------------------------

This payment manager uses two customizable email templates:

1. PAYMENT_CONFIRMATION_TEMPLATE_NAME (default: "PM_PaymentMade_Email")
   - Used when recording a payment
   - Contains variables:
     - {name} - Payer's name
     - {chargeNotes} - Payment details (type, amount, description)
     - {previousDue} - Previous balance amount with formatting
     - {newTotalDue} - New balance after payment with formatting
     - {sender_alias} - Name of the email sender
     - {sender_phone} - Phone number for contact
     - {sender_email} - Email address for contact

2. PAYMENT_NOTIFICATION_TEMPLATE_NAME (default: "PM_Notification_Email")
   - Used when sending payment links
   - Contains variables:
     - {name} - Payer's name
     - {chargeNotes} - Charge details with formatting
     - {previousDue} - Previous balance amount with formatting
     - {totalDue} - Total amount due with formatting
     - {paylink} - Generated payment link URL
     - {sender_alias} - Name of the email sender
     - {sender_phone} - Phone number for contact
     - {sender_email} - Email address for contact

To set up these templates:
1. Go to Special Content > HTML in TouchPoint
2. Create a new HTML file with the name specified in the configuration
3. Design your email using the variables listed above
4. Save the template

Default templates will be used if the custom templates are not found.
"""

try:
    class ModernPaymentManager:
        """Complete Payment Manager with integrated functionality"""
        
        def __init__(self):
            self.program_id = self.get_program_id()
            self.current_action = self.get_current_action()
            self.division_filter = getattr(model.Data, 'divFilter', '')
            self.search_filters = self.get_search_filters()
            
        def get_program_id(self):
            """Get the program ID from form data"""
            try:
                return str(getattr(model.Data, 'ProgramID', DEFAULT_PROGRAM_ID))
            except:
                return str(DEFAULT_PROGRAM_ID)
                
        def get_current_action(self):
            """Determine current action from URL parameters"""
            try:
                return str(getattr(model.Data, 'action', 'programs'))
            except:
                return 'programs'
                
        def get_search_filters(self):
            """Extract search filters from form data"""
            return {
                'first_name': str(getattr(model.Data, 'FirstNameSearch', '')),
                'last_name': str(getattr(model.Data, 'LastNameSearch', '')),
                'show_all': str(getattr(model.Data, 'ShowAll', ''))
            }
            
        def format_phone(self, phone_number):
            """Format phone number for display"""
            if not phone_number:
                return ""
            phone_str = str(phone_number).strip()
            if len(phone_str) == 10:
                return '(' + phone_str[:3] + ') ' + phone_str[3:6] + '-' + phone_str[6:]
            elif len(phone_str) == 7:
                return '(615) ' + phone_str[:3] + '-' + phone_str[3:]
            return phone_str
            
        def format_currency(self, amount):
            """Format amount as currency"""
            try:
                return '${:,.2f}'.format(float(amount) + 0.00)
            except:
                return '$0.00'
                
        def safe_get_attr(self, obj, attr, default=''):
            """Safely get attribute from object"""
            try:
                return getattr(obj, attr, default)
            except:
                return default
        
        # Data retrieval methods
        def get_programs_with_dues(self):
            """Get all programs that have outstanding dues"""
            sql = """
            SELECT 
                pro.Name AS ProgramName,
                pro.Id AS ProgramId,
                SUM(ts.IndDue) AS Outstanding,
                COUNT(DISTINCT ts.PeopleId) AS PayerCount
            FROM [TransactionSummary] ts
            INNER JOIN [People] p ON ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE 
                ts.IndDue <> 0
                AND ts.IsLatestTransaction = 1
                --AND pro.Id <> 1152
            GROUP BY 
                pro.Name, pro.Id
            ORDER BY pro.Name
            """
            return q.QuerySql(sql)
            
        def get_divisions_with_dues(self, program_id):
            """Get divisions within a program that have outstanding dues"""
            sql = """
            SELECT 
                d.Name AS DivisionName,
                d.Id AS DivisionId,
                o.OrganizationName,
                o.OrganizationId,
                SUM(ts.IndDue) AS Outstanding,
                COUNT(DISTINCT ts.PeopleId) AS PayerCount
            FROM [TransactionSummary] ts
            INNER JOIN [People] p ON ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE 
                ts.IndDue <> 0
                AND ts.IsLatestTransaction = 1
                AND pro.Id = {0}
                {1}
            GROUP BY 
                d.Name, d.Id, o.OrganizationName, o.OrganizationId
            ORDER BY d.Name, o.OrganizationName
            """.format(program_id, self.get_division_filter_sql())
            return q.QuerySql(sql)
            
        def get_payers_with_dues(self, org_id=None, program_id=None):
            """Get individual payers with outstanding dues"""
            where_clause = "ts.IndDue <> 0 AND ts.IsLatestTransaction = 1"
            
            if org_id:
                where_clause += " AND o.OrganizationId = {0}".format(org_id)
            elif program_id:
                where_clause += " AND pro.Id = {0}".format(program_id)
                
            # Add name search filters
            if self.search_filters['first_name']:
                where_clause += " AND p.FirstName LIKE '{0}%'".format(self.search_filters['first_name'])
            if self.search_filters['last_name']:
                where_clause += " AND p.LastName LIKE '{0}%'".format(self.search_filters['last_name'])
                
            sql = """
            SELECT 
                pro.Name AS Program,
                pro.Id AS ProgramId,
                d.Name AS Division,
                d.Id AS DivisionId,
                o.OrganizationName,
                o.OrganizationId,
                ts.PeopleId,
                p.Name2,
                p.Age,
                p.FirstName,
                p.LastName,
                p.EmailAddress,
                p.CellPhone,
                p.HomePhone,
                p.FamilyId,
                SUM(ts.TotPaid) AS Paid,
                SUM(ts.TotCoupon) AS Coupons,
                SUM(ts.IndDue) AS Outstanding,
                ts.TranDate
            FROM [TransactionSummary] ts
            INNER JOIN [People] p ON ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE {0}
            GROUP BY 
                d.Name, o.OrganizationName, o.OrganizationId, pro.Name, pro.Id,
                d.Name, d.Id, ts.PeopleId, p.Name2, p.Age, p.FirstName, p.LastName,
                p.EmailAddress, p.CellPhone, p.HomePhone, p.FamilyId, ts.TranDate
            ORDER BY o.OrganizationName, p.Name2
            """.format(where_clause)
            return q.QuerySql(sql)
            
        def get_division_filter_sql(self):
            """Generate SQL filter for division"""
            if self.division_filter and self.division_filter != 'All':
                try:
                    div_id = int(self.division_filter)
                    return " AND d.Id = {0}".format(div_id)
                except:
                    return " AND d.Name = '{0}'".format(self.division_filter)
            return ""
            
        def get_parent_emails(self, family_id):
            """Get parent email addresses for CC purposes"""
            try:
                sql = """
                SELECT DISTINCT p.PeopleId, p.EmailAddress, p.FirstName, p.LastName, 
                       p.CellPhone, p.HomePhone 
                FROM dbo.People AS p 
                INNER JOIN dbo.Families AS t1 ON t1.FamilyId = p.FamilyId 
                WHERE (p.FamilyId = {0}) 
                    AND (p.PositionInFamilyId = 10) 
                    AND (NOT (p.IsDeceased = 1)) 
                    AND (NOT (p.ArchivedFlag = 1)) 
                    AND p.EmailAddress <> '' 
                    AND (t1.HeadOfHouseholdId = p.PeopleId OR t1.HeadOfHouseholdSpouseId = p.PeopleId)
                """.format(family_id)
                
                parents = q.QuerySql(sql)
                emails = []
                parent_info = {}
                
                for parent in parents:
                    email = self.safe_get_attr(parent, 'EmailAddress')
                    if email:
                        emails.append(email)
                        parent_info = {
                            'id': self.safe_get_attr(parent, 'PeopleId'),
                            'email': email,
                            'first_name': self.safe_get_attr(parent, 'FirstName'),
                            'last_name': self.safe_get_attr(parent, 'LastName'),
                            'phone': self.format_phone(self.safe_get_attr(parent, 'CellPhone')) or 
                                   self.format_phone(self.safe_get_attr(parent, 'HomePhone'))
                        }
                        
                    # Ensure user account exists for payment links
                    people_id = self.safe_get_attr(parent, 'PeopleId')
                    if people_id:
                        user_count = q.QuerySqlInt("SELECT COUNT(UserId) FROM Users WHERE PeopleId = {0}".format(people_id))
                        if user_count == 0:
                            model.AddRole(people_id, "Access")
                            model.RemoveRole(people_id, "Access")
                
                return ','.join(emails), parent_info
            except Exception as e:
                return '', {}

        def get_email_history(self, people_id):
            """Get email history for a person"""
            sql = '''
            SELECT Top 50
                p.Name, eq.Subject, eq.Body, p.PeopleId, eq.Sent, eq.FromName, 
                eq.Id AS messageId, 
                Count(er.Id) AS Opened
            FROM dbo.EmailQueueTo eqt
            INNER JOIN dbo.EmailQueue eq ON (eqt.Id = eq.Id) 
            INNER JOIN dbo.People p ON (eqt.PeopleId = p.PeopleId)
            LEFT JOIN EmailResponses er ON er.PeopleId = p.PeopleId AND eq.Id = er.EmailQueueId
            WHERE eqt.PeopleId = {0}
            GROUP BY p.Name, eq.Subject, eq.Body, p.PeopleId, eq.Sent, eq.FromName, eq.Id
            ORDER BY eq.Sent DESC
            '''.format(people_id)
            return q.QuerySql(sql)

        def get_transaction_history(self, people_id):
            """Get transaction history for a person"""
            sql = '''
            SELECT 
                p.Name Person, pro.Name as Program, d.Name as Division, o.OrganizationName OrgName,
                FORMAT(ts.TranDate, 'yyyy-MM-dd') as TranDate, ts.TotDue as TSAmount,
                FORMAT(t.TransactionDate, 'yyyy-MM-dd') as TransactionDate, t.amtdue as Amount,
                CASE WHEN ts.TotDue > 0 THEN 'Church Received'
                     WHEN ts.TotDue < 0 THEN 'Person Received'
                     ELSE 'Zero Amount' END TransactionDirection,
                t.TransactionId, t.Message, t.Description,
                CASE WHEN [message] like 'CHK%' THEN 'Check'
                     WHEN [message] like 'CSH%' THEN 'Cash'
                     WHEN [message] like 'Response%' THEN 'Credit Card'
                     WHEN [message] like 'FEE%' THEN 'Church Adjustment'
                     WHEN [transactionid] like 'Coupon%' THEN 'Coupon'
                     ELSE 'Unknown' END TransactionType
            FROM TransactionSummary ts
            LEFT JOIN [Transaction] t on ts.RegId = t.OriginalId
            INNER JOIN [People] p on ts.PeopleId = p.PeopleId
            LEFT JOIN Organizations o ON o.OrganizationId = ts.OrganizationId
            LEFT JOIN Division d ON d.Id = o.DivisionId
            LEFT JOIN Program pro ON pro.Id = d.ProgId
            WHERE ts.PeopleId = {0} AND t.amtdue <> 0
            ORDER BY t.TransactionDate DESC
            '''.format(people_id)
            return q.QuerySql(sql)

        # Payment processing methods
        def process_payment_link(self):
            """Process payment link request"""
            try:
                payer_id = str(getattr(model.Data, 'pid', ''))
                org_id = str(getattr(model.Data, 'PaymentOrg', ''))
                amount = str(getattr(model.Data, 'PayFee', ''))
                payer_name = str(getattr(model.Data, 'payerName', ''))
                cc_emails = str(getattr(model.Data, 'cc_emails', ''))
                
                if not payer_id or not org_id or not amount:
                    return self.create_json_response(False, "Missing required payment information")
                
                # Get email sender details
                program_id = str(getattr(model.Data, 'ProgramID', self.program_id))
                email_details = self.get_email_details(program_id)
                
                if not email_details:
                    return self.create_json_response(False, "Email configuration not found for program")
                
                # Get current balance
                previous_due = self.get_current_balance(payer_id, org_id)
                
                # Generate payment link
                paylink = model.GetPayLink(int(payer_id), int(org_id))
                paylinkauth = model.GetAuthenticatedUrl(int(org_id), paylink, True)
                
                if str(paylinkauth).split('/')[-1].lower() == 'none':
                    return self.create_json_response(False, "Unable to generate payment link")
                
                # Calculate total due
                total_due = float(previous_due) + float(amount)
                
                # Send email
                message = self.build_payment_email(
                    payer_name, amount, previous_due, total_due, 
                    paylinkauth, email_details
                )
                
                try:
                    model.Email(
                        int(payer_id), 
                        int(email_details['sender_id']), 
                        email_details['sender_email'],
                        email_details['sender_alias'], 
                        email_details['email_title'], 
                        message, 
                        cc_emails
                    )
                    return self.create_json_response(True, "Payment link sent successfully")
                except Exception as e:
                    return self.create_json_response(False, "Failed to send email: " + str(e))
                    
            except Exception as e:
                return self.create_json_response(False, "Error processing payment link: " + str(e))

        def process_payment_record(self):
            """
            Process payment recording
            
            Uses template: PAYMENT_CONFIRMATION_TEMPLATE_NAME
            Variables: name, chargeNotes, previousDue, newTotalDue, sender_alias, sender_phone, sender_email
            """
            try:
                # Extract form data with safe fallbacks
                payer_id = str(getattr(model.Data, 'pid', ''))
                org_id = str(getattr(model.Data, 'PaymentOrg', ''))
                payment_amount = str(getattr(model.Data, 'PaidAmount', ''))
                payment_type = str(getattr(model.Data, 'PaymentType', CHECK_PAYMENT_TYPE))
                payment_desc = str(getattr(model.Data, 'PaymentDescription', ''))
                payer_name = str(getattr(model.Data, 'payerName', ''))
                cc_emails = str(getattr(model.Data, 'cc_emails', ''))
                
                # Validate required fields
                if not all([payer_id, org_id, payment_amount, payment_type, payment_desc]):
                    return self.create_json_response(False, "Missing required payment information")
                
                # Validate amount format
                if payment_amount != "" and str(payment_amount).replace(".", "").isdigit():
                    # Payment comes in as a positive number but needs to be negative in the database
                    amt = -float(payment_amount)
                    PayType = 'Payment'
                    payment_amount_display = '${0:.2f}'.format(abs(float(amt))).replace("$-", "-$")
                    chargeNotes = payment_type + ' ' + payment_amount_display + ': ' + payment_desc
                else:
                    return self.create_json_response(False, "Invalid payment amount")
                
                # Get email configuration
                program_id = str(getattr(model.Data, 'ProgramID', self.program_id))
                email_details = self.get_email_details(program_id)
                
                # Prepare transaction description
                messageDescription = payment_type + payment_desc
                
                # Get previous balance
                previousDue = "<br />" + '$0.00' + "........Previous Credit Amount"
                IndDue = 0
                
                # Query current balance
                try:
                    inddue_list = q.QuerySql(
                        "SELECT Sum(IndDue) as IndDue, Sum(TotDue) as TotDue " +
                        "FROM dbo.TransactionSummary " +
                        "WHERE PeopleId = {0} AND OrganizationId = {1}".format(payer_id, org_id)
                    )
                    
                    if inddue_list:
                        for pc in inddue_list:
                            # Handle different balance scenarios
                            if pc.IndDue is None or pc.IndDue == 0:
                                IndDue = float(0)
                            elif pc.IndDue > 0 or pc.IndDue < 0:
                                IndDue = float(pc.IndDue)
                            
                            # Create initial transaction if needed
                            if pc.TotDue is None:
                                model.AddTransaction(int(payer_id), int(org_id), 0, "Initial charge of $0")
                            
                            # Format previous balance for display
                            previousDue = "<br />" + '${:,.2f}'.format(IndDue + 0.00).replace("$-","-$") + "........Previous Balance"
                except Exception as query_error:
                    print("<!-- Error querying balance: " + str(query_error) + " -->")
                    # Continue with default values
                
                # Check for overpayment
                if IndDue + float(amt) < 0:
                    PayType = 'Payment'
                
                # Process payment
                if PayType == 'Payment':
                    # Calculate new total due
                    newTotalDue = float(IndDue) + amt
                    newTotalDue_display = '${:,.2f}'.format(float(newTotalDue)*1.0 + 0.00).replace("$-","-$")
                    
                    # Add transaction to database
                    try:
                        # In DB: positive = charge, negative = payment
                        transamount = -amt  # Negate the already negative amount to make it positive
                        transaction_id = model.AddTransaction(int(payer_id), int(org_id), transamount, messageDescription)
                        
                        if not transaction_id:
                            return self.create_json_response(False, "Transaction not recorded - please check credentials")
                    except Exception as e:
                        return self.create_json_response(False, "Failed to record transaction: " + str(e))
                    
                    # Send confirmation email
                    try:
                        # Try to load configured template
                        try:
                            message = model.HtmlContent(PAYMENT_CONFIRMATION_TEMPLATE_NAME)
                        except Exception as template_error:
                            # Fall back to default template if configured one not found
                            print("<!-- Template not found: " + str(template_error) + " -->")
                            message = DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE
                        
                        # Format template with payment details
                        formatted_message = message.format(
                            name=payer_name, 
                            chargeNotes=chargeNotes, 
                            previousDue=previousDue,
                            newTotalDue=newTotalDue_display, 
                            sender_alias=email_details.get('sender_alias', ''), 
                            sender_phone=email_details.get('sender_phone', ''), 
                            sender_email=email_details.get('sender_email', '')
                        )
                        
                        # Send the email
                        try:
                            model.Email(
                                int(payer_id), 
                                int(email_details.get('sender_id', 3134)), 
                                email_details.get('sender_email', 'noreply@church.com'),
                                email_details.get('sender_alias', 'Payment System'), 
                                EMAIL_SUBJECT_PREFIX + "Payment Received", 
                                formatted_message, 
                                cc_emails
                            )
                        except Exception as email_error:
                            # Log but don't fail the transaction if email fails
                            print("<!-- Email error: " + str(email_error) + " -->")
                    except Exception as template_error:
                        # Log but don't fail if template processing fails
                        print("<!-- Template error: " + str(template_error) + " -->")
                
                # Return success response
                return self.create_json_response(True, "Payment recorded successfully")
                
            except Exception as e:
                return self.create_json_response(False, "Error recording payment: " + str(e))

        def process_resend_email(self):
            """Process email resend request"""
            try:
                message_id = str(getattr(model.Data, 'messageId', ''))
                people_id = str(getattr(model.Data, 'PeopleId', ''))
                
                if not message_id or not people_id:
                    return self.create_json_response(False, "Missing message ID or people ID")
                
                # Use the same SQL as your original PM_EmailPreview
                sql = '''
                Select eq.Subject
                    ,eq.Body
                    ,eq.Sent
                    ,eq.FromName
                    ,eq.Id AS [messageId]
                    ,eqt.PeopleId
                    ,eq.QueuedBy
                    ,eq.FromAddr
                from EmailQueue eq
                left join dbo.EmailQueueTo eqt on (eqt.id = eq.id and eqt.PeopleId = {1})
                where eq.Id = {0}
                '''.format(message_id, people_id)
                
                email_data = q.QuerySql(sql)
                if not email_data or len(email_data) == 0:
                    return self.create_json_response(False, "Original email not found")
                
                # Use a for loop like your original code
                for a in email_data:
                    email_title = 'Email Copy - ' + str(self.safe_get_attr(a, 'Subject', 'No Subject'))
                    email_body = '<H2>Email Copy - Originally Sent on <i>' + str(self.safe_get_attr(a, 'Sent', '')) + '</i></H2><br />' + str(self.safe_get_attr(a, 'Body', 'No content'))
                    
                    try:
                        model.Email(
                            int(people_id), 
                            self.safe_get_attr(a, 'QueuedBy', 3134), 
                            str(self.safe_get_attr(a, 'FromAddr', 'noreply@church.com')),
                            str(self.safe_get_attr(a, 'FromName', 'Payment System')), 
                            email_title, 
                            email_body
                        )
                        return self.create_json_response(True, "Email copy sent successfully")
                    except Exception as e:
                        return self.create_json_response(False, "Failed to send email: " + str(e))
                        
            except Exception as e:
                return self.create_json_response(False, "Error processing email resend: " + str(e))

        def get_email_details(self, program_id):
            """Get email configuration for program"""
            try:
                email_details_text = model.TextContent("PM_EmailSenders")
                email_details = ast.literal_eval(email_details_text)
                
                if program_id in email_details:
                    return email_details[program_id]
                return DEFAULT_EMAIL_SENDER
            except:
                # Return default email configuration
                return DEFAULT_EMAIL_SENDER

        def get_current_balance(self, payer_id, org_id):
            """Get current balance for payer/org"""
            try:
                sql = "SELECT Sum(IndDue) as IndDue FROM dbo.TransactionSummary WHERE PeopleId = {0} AND OrganizationId = {1}".format(payer_id, org_id)
                result = q.QuerySql(sql)
                if result and len(result) > 0:
                    balance = self.safe_get_attr(result[0], 'IndDue', 0)
                    return float(balance) if balance is not None else 0.0
                return 0.0
            except:
                return 0.0

        def build_payment_email(self, name, charge_amount, previous_due, total_due, paylink, email_details):
            """Build payment notification email"""
            try:
                message = model.HtmlContent(PAYMENT_NOTIFICATION_TEMPLATE_NAME)
            except:
                message = DEFAULT_PAYMENT_EMAIL_TEMPLATE
            
            charge_notes = '${:,.2f}........New Charge'.format(float(charge_amount))
            previous_due_text = '${:,.2f}........Previous Balance'.format(float(previous_due))
            total_due_text = '${:,.2f}'.format(float(total_due))
            
            return message.format(
                name=name,
                chargeNotes=charge_notes,
                previousDue=previous_due_text,
                totalDue=total_due_text,
                paylink=paylink,
                sender_alias=email_details['sender_alias'],
                sender_phone=email_details['sender_phone'],
                sender_email=email_details['sender_email']
            ) + '{track}{tracklinks}<br />'

        def build_payment_confirmation_email(self, name, payment_amount, previous_due, new_total, payment_type, payment_desc, email_details):
            """Build payment confirmation email"""
            try:
                # Try to load the template from TouchPoint's special content
                message = model.HtmlContent(PAYMENT_CONFIRMATION_TEMPLATE_NAME)
            except Exception as e:
                # If template not found, log error and use default template
                print("<!-- Warning: Email template not found. Using default. Error: {} -->".format(str(e)))
                message = DEFAULT_PAYMENT_CONFIRMATION_TEMPLATE
            
            # Format payment details
            payment_amount_abs = abs(float(payment_amount))
            charge_notes = '{0} ${1:.2f}: {2}'.format(
                payment_type.replace('|', ''), 
                payment_amount_abs, 
                payment_desc
            )
            
            # Format monetary values
            previous_due_text = '${:,.2f}'.format(float(previous_due))
            new_total_text = '${:,.2f}'.format(float(new_total))
            
            # Apply template replacements
            try:
                return message.format(
                    name=name,
                    chargeNotes=charge_notes,
                    previousDue=previous_due_text,
                    newTotalDue=new_total_text,
                    sender_alias=email_details.get('sender_alias', 'Payment System'),
                    sender_phone=email_details.get('sender_phone', ''),
                    sender_email=email_details.get('sender_email', '')
                )
            except Exception as e:
                # If template format fails, return simplified message
                print("<!-- Warning: Error formatting email template: {} -->".format(str(e)))
                return """
                <p>Dear {0},</p>
                <p>We have received your payment of {1}. Thank you!</p>
                <p>New Balance: {2}</p>
                <p>Thank you,<br>{3}</p>
                """.format(name, charge_notes, new_total_text, email_details.get('sender_alias', 'Payment System'))
            
            return message.format(
                name=name,
                chargeNotes=charge_notes,
                previousDue=previous_due_text,
                newTotalDue=new_total_text,
                sender_alias=email_details['sender_alias'],
                sender_phone=email_details['sender_phone'],
                sender_email=email_details['sender_email']
            )

        def create_json_response(self, success, message, data=None):
            """Create JSON response for AJAX calls"""
            response = {
                'success': success,
                'message': message
            }
            if data:
                response['data'] = data
            return json.dumps(response)

        # View rendering methods
        def render_programs_view(self):
            """Render the main programs overview"""
            programs = self.get_programs_with_dues()
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-credit-card"></i> Payment Manager - Programs Overview</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-content">
                    <table class="pm-table">
                        <thead>
                            <tr>
                                <th>Program Name</th>
                                <th>Outstanding Amount</th>
                                <th>Payers Count</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
            """
            
            total_outstanding = 0
            total_payers = 0
            
            try:
                for program in programs:
                    outstanding = float(self.safe_get_attr(program, 'Outstanding', 0) or 0)
                    payer_count = int(self.safe_get_attr(program, 'PayerCount', 0) or 0)
                    total_outstanding += outstanding
                    total_payers += payer_count
                    
                    program_name = self.safe_get_attr(program, 'ProgramName', 'Unknown')
                    program_id = self.safe_get_attr(program, 'ProgramId', 0)
                    
                    html += """
                                <tr>
                                    <td><strong>{0}</strong></td>
                                    <td class="pm-currency">{1}</td>
                                    <td class="pm-center">{2}</td>
                                    <td class="pm-center">
                                        <button class="btn btn-sm btn-outline-primary" 
                                                onclick="viewDivisions({3})">
                                            <i class="fa fa-list"></i> View Divisions
                                        </button>
                                        <button class="btn btn-sm btn-outline-success" 
                                                onclick="viewPayers(null, {3})">
                                            <i class="fa fa-users"></i> View All Payers
                                        </button>
                                    </td>
                                </tr>
                    """.format(
                        program_name,
                        self.format_currency(outstanding),
                        payer_count,
                        program_id
                    )
            except Exception as e:
                html += "<tr><td colspan='4'>Error loading programs: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                        <tfoot>
                            <tr class="pm-total">
                                <th style="text-align: left;">Total</th>
                                <th class="pm-currency" style="text-align: right;">{0}</th>
                                <th class="pm-center" style="text-align: center;">{1}</th>
                                <th style="text-align: center;"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            """.format(
                self.format_currency(total_outstanding),
                total_payers
            )
            
            return html
            
        def render_divisions_view(self, program_id):
            """Render divisions/involvements within a program"""
            divisions = self.get_divisions_with_dues(program_id)
            program_name = "Unknown Program"
            
            # Get program name
            try:
                program_data = q.QuerySql("SELECT Name FROM Program WHERE Id = {0}".format(program_id))
                if program_data and len(program_data) > 0:
                    program_name = self.safe_get_attr(program_data[0], 'Name', 'Unknown Program')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-sitemap"></i> {0} - Divisions & Involvements</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="viewPrograms()">
                            <i class="fa fa-arrow-left"></i> Back to Programs
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-content">
                    <table class="pm-table">
                        <thead>
                            <tr>
                                <th>Division</th>
                                <th>Involvement</th>
                                <th>Outstanding</th>
                                <th>Payers</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(program_name)
            
            current_division = ""
            total_outstanding = 0
            total_payers = 0
            
            try:
                for div in divisions:
                    outstanding = float(self.safe_get_attr(div, 'Outstanding', 0) or 0)
                    payer_count = int(self.safe_get_attr(div, 'PayerCount', 0) or 0)
                    total_outstanding += outstanding
                    total_payers += payer_count
                    
                    division_name = self.safe_get_attr(div, 'DivisionName', '')
                    org_name = self.safe_get_attr(div, 'OrganizationName', 'Unknown')
                    org_id = self.safe_get_attr(div, 'OrganizationId', 0)
                    
                    division_cell = ""
                    if division_name != current_division:
                        division_cell = division_name
                        current_division = division_name
                    
                    html += """
                                <tr>
                                    <td>{0}</td>
                                    <td>{1}</td>
                                    <td class="pm-currency">{2}</td>
                                    <td class="pm-center">{3}</td>
                                    <td class="pm-center">
                                        <button class="btn btn-sm btn-outline-success" 
                                                onclick="viewPayers({4}, {5})">
                                            <i class="fa fa-users"></i> View Payers
                                        </button>
                                    </td>
                                </tr>
                    """.format(
                        division_cell,
                        org_name,
                        self.format_currency(outstanding),
                        payer_count,
                        org_id,
                        program_id
                    )
            except Exception as e:
                html += "<tr><td colspan='5'>Error loading divisions: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                        <tfoot>
                            <tr class="pm-total">
                                <th colspan="2">Total</th>
                                <th class="pm-currency">{0}</th>
                                <th class="pm-center">{1}</th>
                                <th></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            """.format(
                self.format_currency(total_outstanding),
                total_payers
            )
            
            return html
            
        def render_payers_view(self, org_id=None, program_id=None):
            """Render individual payers with payment options"""
            payers = self.get_payers_with_dues(org_id, program_id)
            
            # Get context title
            context_title = "All Programs"
            try:
                if program_id:
                    program_data = q.QuerySql("SELECT Name FROM Program WHERE Id = {0}".format(program_id))
                    if program_data and len(program_data) > 0:
                        context_title = self.safe_get_attr(program_data[0], 'Name', 'Unknown Program')
                if org_id:
                    org_data = q.QuerySql("SELECT OrganizationName FROM Organizations WHERE OrganizationId = {0}".format(org_id))
                    if org_data and len(org_data) > 0:
                        context_title = self.safe_get_attr(org_data[0], 'OrganizationName', 'Unknown Organization')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-users"></i> {0} - Individual Payers</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-search">
                    <input type="text" id="searchInput" placeholder="Search by name, email, or phone..." 
                           class="form-control">
                </div>
                
                <div class="pm-content">
                    <table class="pm-table" id="payersTable">
                        <thead>
                            <tr>
                                <th>Name & Contact</th>
                                <th>Involvement</th>
                                <th>Outstanding</th>
                                <th>Payment Options</th>
                                <th>History</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(context_title)
            
            total_outstanding = 0
            
            try:
                for payer in payers:
                    outstanding = float(self.safe_get_attr(payer, 'Outstanding', 0) or 0)
                    total_outstanding += outstanding
                    
                    # Get payer details safely
                    payer_name = self.safe_get_attr(payer, 'Name2', 'Unknown')
                    payer_id = self.safe_get_attr(payer, 'PeopleId', 0)
                    payer_email = self.safe_get_attr(payer, 'EmailAddress', '')
                    org_name = self.safe_get_attr(payer, 'OrganizationName', 'Unknown')
                    org_id = self.safe_get_attr(payer, 'OrganizationId', 0)
                    division_name = self.safe_get_attr(payer, 'Division', '')
                    family_id = self.safe_get_attr(payer, 'FamilyId', 0)
                    
                    # Get parent contact info
                    cc_emails, parent_info = self.get_parent_emails(family_id)
                    
                    # Format contact information
                    phone_info = ""
                    cell_phone = self.safe_get_attr(payer, 'CellPhone', '')
                    home_phone = self.safe_get_attr(payer, 'HomePhone', '')
                    if cell_phone:
                        phone_info += " <small>C: {0}</small>".format(self.format_phone(cell_phone))
                    if home_phone:
                        phone_info += " <small>H: {0}</small>".format(self.format_phone(home_phone))
                    
                    email_display = payer_email if payer_email else "<em>No email</em>"
                    if parent_info and parent_info.get('email'):
                        email_display += "<br><small>Parent: {0}</small>".format(parent_info['email'])
                    
                    # Payment buttons
                    payment_buttons = ""
                    if outstanding > 0:
                        payment_buttons = """
                            <div class="pm-btn-group">
                                <button class="btn btn-sm btn-outline-primary" 
                                        onclick="sendPaymentLink({0}, {1}, '{2}', {3}, '{4}')">
                                    <i class="fa fa-credit-card"></i> Send Pay Link
                                </button>
                                <button class="btn btn-sm btn-outline-success" 
                                        onclick="recordPayment({0}, {1}, '{2}', {3})">
                                    <i class="fa fa-money"></i> Record Payment
                                </button>
                            </div>
                        """.format(
                            payer_id,
                            org_id,
                            payer_name.replace("'", "\\'"),
                            outstanding,
                            cc_emails
                        )
                    
                    html += """
                                <tr>
                                    <td>
                                        <div>
                                            <strong>{0}</strong>
                                            <a href="{1}/Person2/{2}" target="_blank">
                                                <i class="fa fa-external-link"></i>
                                            </a>
                                        </div>
                                        <div class="pm-contact">
                                            {3}<br>
                                            {4}
                                        </div>
                                    </td>
                                    <td>
                                        <div>{5}</div>
                                        <small class="pm-muted">{6}</small>
                                    </td>
                                    <td class="pm-currency">{7}</td>
                                    <td class="pm-center">
                                        {8}
                                    </td>
                                    <td class="pm-center">
                                        <div class="pm-btn-group">
                                            <button class="btn btn-sm btn-outline-info" 
                                                    onclick="viewTransactions({2})">
                                                <i class="fa fa-history"></i> Transactions
                                            </button>
                                            <button class="btn btn-sm btn-outline-secondary" 
                                                    onclick="viewEmails({2})">
                                                <i class="fa fa-envelope"></i> Email History
                                            </button>
                                        </div>
                                    </td>
                                </tr>
                    """.format(
                        payer_name,
                        model.CmsHost,
                        payer_id,
                        email_display,
                        phone_info,
                        org_name,
                        division_name,
                        self.format_currency(outstanding),
                        payment_buttons
                    )
            except Exception as e:
                html += "<tr><td colspan='5'>Error loading payers: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                        <tfoot>
                            <tr class="pm-total">
                                <th colspan="2">Total Outstanding</th>
                                <th class="pm-currency">{0}</th>
                                <th colspan="2"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            """.format(self.format_currency(total_outstanding))
            
            return html

        def render_email_history(self, people_id):
            """Render email history for a person"""
            emails = self.get_email_history(people_id)
            
            # Get person name
            person_name = "Unknown Person"
            try:
                person_data = q.QuerySql("SELECT Name FROM People WHERE PeopleId = {0}".format(people_id))
                if person_data and len(person_data) > 0:
                    person_name = self.safe_get_attr(person_data[0], 'Name', 'Unknown Person')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-envelope"></i> Email History - {0}</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-content">
                    <table class="pm-table">
                        <thead>
                            <tr>
                                <th>Sent Date</th>
                                <th>From</th>
                                <th>Subject</th>
                                <th>Opened</th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(person_name)
            
            try:
                if not emails:
                    html += "<tr><td colspan='5' class='pm-center'>No email history found</td></tr>"
                else:
                    for email in emails:
                        sent_date = self.safe_get_attr(email, 'Sent', '')
                        from_name = self.safe_get_attr(email, 'FromName', 'Unknown')
                        subject = self.safe_get_attr(email, 'Subject', 'No Subject')
                        opened = self.safe_get_attr(email, 'Opened', 0)
                        message_id = self.safe_get_attr(email, 'messageId', '')
                        
                        html += """
                                    <tr>
                                        <td>{0}</td>
                                        <td>{1}</td>
                                        <td>{2}</td>
                                        <td class="pm-center">{3}</td>
                                        <td class="pm-center">
                                            <button class="btn btn-sm btn-outline-primary" 
                                                    onclick="previewEmail({4}, {5})">
                                                <i class="fa fa-eye"></i> Preview
                                            </button>
                                            <button class="btn btn-sm btn-outline-success" 
                                                    onclick="resendEmail({4}, {5})">
                                                <i class="fa fa-paper-plane"></i> Resend
                                            </button>
                                        </td>
                                    </tr>
                        """.format(
                            sent_date,
                            from_name,
                            subject,
                            opened,
                            message_id,
                            people_id
                        )
            except Exception as e:
                html += "<tr><td colspan='5'>Error loading email history: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                    </table>
                </div>
            </div>
            """
            
            return html

        def render_email_preview(self, message_id, people_id):
            """Render email preview"""
            try:
                sql = '''
                Select eq.Subject
                    ,eq.Body
                    ,eq.Sent
                    ,eq.FromName
                    ,eq.Id AS [messageId]
                    ,eqt.PeopleId
                    ,eq.QueuedBy
                    ,eq.FromAddr
                from EmailQueue eq
                left join dbo.EmailQueueTo eqt on (eqt.id = eq.id and eqt.PeopleId = {1})
                where eq.Id = {0}
                '''.format(message_id, people_id)
                
                # Use the same pattern as your original PM_EmailPreview
                emailTitle = ""
                emailBody = ""
                
                email_data = q.QuerySql(sql)
                if not email_data:
                    return "<div class='alert alert-warning'>Email not found</div>"
                
                # Use for loop like your original code
                for a in email_data:
                    return """
                    <div class="pm-container">
                        <div class="pm-header">
                            <h3><i class="fa fa-envelope-open"></i> Email Preview</h3>
                            <div class="pm-actions">
                                <button class="btn btn-secondary" onclick="history.go(-1)">
                                    <i class="fa fa-arrow-left"></i> Go Back
                                </button>
                                <button class="btn btn-success" onclick="resendEmail({0}, {1})">
                                    <i class="fa fa-paper-plane"></i> Send Copy
                                </button>
                                <button class="btn btn-primary" onclick="window.print()">
                                    <i class="fa fa-print"></i> Print
                                </button>
                            </div>
                        </div>
                        
                        <div class="pm-content">
                            <div class="email-meta">
                                <h4>Email Details</h4>
                                <table class="pm-table">
                                    <tr><th width="120">Subject:</th><td>{2}</td></tr>
                                    <tr><th>From:</th><td>{3}</td></tr>
                                    <tr><th>Sent:</th><td>{4}</td></tr>
                                </table>
                            </div>
                            
                            <div class="email-body">
                                <h4>Email Content</h4>
                                <div class="email-content">
                                    {5}
                                </div>
                            </div>
                        </div>
                    </div>
                    """.format(
                        message_id,
                        people_id,
                        getattr(a, 'Subject', 'No Subject'),
                        getattr(a, 'FromName', 'Unknown'),
                        getattr(a, 'Sent', ''),
                        getattr(a, 'Body', 'No content')
                    )
                
                # If no data found
                return "<div class='alert alert-warning'>Email not found</div>"
                
            except Exception as e:
                return "<div class='alert alert-danger'>Error loading email preview: {0}</div>".format(str(e))

        def render_transaction_history(self, people_id):
            """Render transaction history for a person"""
            transactions = self.get_transaction_history(people_id)
            
            # Get person name
            person_name = "Unknown Person"
            try:
                person_data = q.QuerySql("SELECT Name FROM People WHERE PeopleId = {0}".format(people_id))
                if person_data and len(person_data) > 0:
                    person_name = self.safe_get_attr(person_data[0], 'Name', 'Unknown Person')
            except:
                pass
            
            html = """
            <div class="pm-container">
                <div class="pm-header">
                    <h3><i class="fa fa-history"></i> Transaction History - {0}</h3>
                    <div class="pm-actions">
                        <button class="btn btn-secondary" onclick="history.go(-1)">
                            <i class="fa fa-arrow-left"></i> Go Back
                        </button>
                        <button class="btn btn-primary" onclick="refreshData()">
                            <i class="fa fa-sync"></i> Refresh
                        </button>
                    </div>
                </div>
                
                <div class="pm-search">
                    <input type="text" id="transactionSearch" placeholder="Search transactions..." 
                           class="form-control">
                </div>
                
                <div class="pm-content">
                    <table class="pm-table" id="transactionsTable">
                        <thead>
                            <tr>
                                <th>Date</th>
                                <th>Organization</th>
                                <th>Amount</th>
                                <th>Type</th>
                                <th>Direction</th>
                                <th>Description</th>
                            </tr>
                        </thead>
                        <tbody>
            """.format(person_name)
            
            try:
                if not transactions:
                    html += "<tr><td colspan='6' class='pm-center'>No transaction history found</td></tr>"
                else:
                    for trans in transactions:
                        trans_date = self.safe_get_attr(trans, 'TransactionDate', '')
                        org_name = self.safe_get_attr(trans, 'OrgName', 'Unknown')
                        amount = float(self.safe_get_attr(trans, 'Amount', 0) or 0)
                        trans_type = self.safe_get_attr(trans, 'TransactionType', 'Unknown')
                        direction = self.safe_get_attr(trans, 'TransactionDirection', '')
                        description = self.safe_get_attr(trans, 'Description', '')
                        message = self.safe_get_attr(trans, 'Message', '')
                        
                        # Combine description and message
                        full_desc = description
                        if message and message != description:
                            full_desc = description + " (" + message + ")" if description else message
                        
                        # Color code amounts
                        amount_class = "pm-positive" if amount > 0 else "pm-negative"
                        amount_display = self.format_currency(abs(amount))
                        
                        html += """
                                    <tr>
                                        <td>{0}</td>
                                        <td>{1}</td>
                                        <td class="pm-currency {2}">{3}</td>
                                        <td>{4}</td>
                                        <td>{5}</td>
                                        <td>{6}</td>
                                    </tr>
                        """.format(
                            trans_date,
                            org_name,
                            amount_class,
                            amount_display,
                            trans_type,
                            direction,
                            full_desc
                        )
            except Exception as e:
                html += "<tr><td colspan='6'>Error loading transaction history: {0}</td></tr>".format(str(e))
            
            html += """
                        </tbody>
                    </table>
                </div>
            </div>
            """
            
            return html

        def render_page_structure(self, content):
            """Render the complete page with navigation and styling"""
            return """
            <script>
            // Helper function to get the correct form submission URL
            function getPyScriptAddress() {{
                let path = window.location.pathname;
                return path.replace("/PyScript/", "/PyScriptForm/");
            }}
            
            // Define all functions globally - before any HTML that uses them
            function showLoading() {{
                var loading = document.getElementById('pmLoading');
                if (loading) loading.style.display = 'flex';
            }}
            
            function hideLoading() {{
                var loading = document.getElementById('pmLoading');
                if (loading) loading.style.display = 'none';
            }}
            
            function showAlert(message, type) {{
                type = type || 'success';
                var alertContainer = document.getElementById('alertContainer');
                if (!alertContainer) return;
                
                var alertDiv = document.createElement('div');
                alertDiv.className = 'alert alert-' + type;
                alertDiv.innerHTML = '<button type="button" style="float: right; background: none; border: none; font-size: 18px; cursor: pointer;" onclick="this.parentNode.remove()">&times;</button>' + message;
                alertContainer.appendChild(alertDiv);
                
                setTimeout(function() {{
                    if (alertDiv.parentNode) {{
                        alertDiv.parentNode.removeChild(alertDiv);
                    }}
                }}, 5000);
            }}
            
            function viewPrograms() {{
                showLoading();
                window.location.href = window.location.pathname + '?action=programs';
            }}
            
            function viewDivisions(programId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=divisions&ProgramID=' + programId;
            }}
            
            function viewPayers(orgId, programId) {{
                showLoading();
                var url = window.location.pathname + '?action=payers';
                if (orgId && orgId !== 'null') url += '&OrganizationId=' + orgId;
                if (programId && programId !== 'null') url += '&ProgramID=' + programId;
                window.location.href = url;
            }}
            
            function refreshData() {{
                showLoading();
                window.location.reload();
            }}
            
            function sendPaymentLink(payerId, orgId, payerName, amount, ccEmails) {{
                showLoading();
                
                var formData = new FormData();
                formData.append('action', 'send_payment_link');
                formData.append('pid', payerId);
                formData.append('PaymentOrg', orgId);
                formData.append('PayFee', amount);
                formData.append('payerName', payerName);
                formData.append('cc_emails', ccEmails || '');
                formData.append('ProgramID', '{1}');
                formData.append('PaymentDescription', '{2}');
                formData.append('PaymentType', 'CreditCharge');
                
                fetch(getPyScriptAddress(), {{
                    method: 'POST',
                    body: formData
                }})
                .then(function(response) {{
                    return response.text().then(function(text) {{
                        try {{
                            return JSON.parse(text);
                        }} catch (e) {{
                            throw new Error('Invalid JSON response: ' + text);
                        }}
                    }});
                }})
                .then(function(data) {{
                    hideLoading();
                    if (data.success) {{
                        showAlert('Payment link sent successfully!', 'success');
                    }} else {{
                        showAlert('Error: ' + data.message, 'danger');
                    }}
                }})
                .catch(function(error) {{
                    hideLoading();
                    showAlert('Network error: ' + error.message, 'danger');
                }});
            }}
            
            function recordPayment(payerId, orgId, payerName, amount) {{
                var modal = document.getElementById('paymentModal');
                if (!modal) return;
                
                document.getElementById('modal-pid').value = payerId;
                document.getElementById('modal-org').value = orgId;
                document.getElementById('modal-name').value = payerName;
                document.getElementById('modal-emails').value = '';
                document.getElementById('modal-payer-info').textContent = payerName;
                document.getElementById('modal-amount-due').textContent = '$' + amount.toFixed(2);
                document.getElementById('PaidAmount').value = amount.toFixed(2);
                modal.style.display = 'flex';
            }}
            
            function closePaymentModal() {{
                var modal = document.getElementById('paymentModal');
                if (modal) modal.style.display = 'none';
            }}
            
            function submitPayment() {{
                var payerId = document.getElementById('modal-pid').value;
                var orgId = document.getElementById('modal-org').value;
                var payerName = document.getElementById('modal-name').value;
                var paymentTypeEl = document.querySelector('input[name="PaymentType"]:checked');
                var paymentType = paymentTypeEl ? paymentTypeEl.value : 'CHK|';
                var paymentDesc = document.getElementById('PaymentDescription').value;
                var paidAmount = document.getElementById('PaidAmount').value;
                
                if (!paymentDesc || !paidAmount) {{
                    showAlert('Please fill in all required fields', 'warning');
                    return;
                }}
                
                showLoading();
                closePaymentModal();
                
                var formData = new FormData();
                formData.append('action', 'record_payment');
                formData.append('pid', payerId);
                formData.append('PaymentOrg', orgId);
                formData.append('payerName', payerName);
                formData.append('PaymentType', paymentType);
                formData.append('PaymentDescription', paymentDesc);
                formData.append('PaidAmount', paidAmount);
                formData.append('ProgramID', '{1}');
                formData.append('cc_emails', '');
                
                fetch(getPyScriptAddress(), {{
                    method: 'POST',
                    body: formData
                }})
                .then(function(response) {{
                    return response.text().then(function(text) {{
                        try {{
                            return JSON.parse(text);
                        }} catch (e) {{
                            throw new Error('Invalid JSON response: ' + text);
                        }}
                    }});
                }})
                .then(function(data) {{
                    hideLoading();
                    if (data.success) {{
                        showAlert('Payment recorded successfully!', 'success');
                        setTimeout(function() {{ refreshData(); }}, 1500);
                    }} else {{
                        showAlert('Error: ' + data.message, 'danger');
                    }}
                }})
                .catch(function(error) {{
                    hideLoading();
                    showAlert('Network error: ' + error.message, 'danger');
                }});
            }}
            
            function viewTransactions(payerId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=transactions&PeopleId=' + payerId;
            }}
            
            function viewEmails(payerId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=emails&PeopleId=' + payerId;
            }}
            
            function previewEmail(messageId, peopleId) {{
                showLoading();
                window.location.href = window.location.pathname + '?action=email_preview&messageId=' + messageId + '&PeopleId=' + peopleId;
            }}
            
            function resendEmail(messageId, peopleId) {{
                if (!confirm('Are you sure you want to resend this email?')) {{
                    return;
                }}
                
                showLoading();
                
                var formData = new FormData();
                formData.append('action', 'resend_email');
                formData.append('messageId', messageId);
                formData.append('PeopleId', peopleId);
                
                fetch(getPyScriptAddress(), {{
                    method: 'POST',
                    body: formData
                }})
                .then(function(response) {{
                    return response.text().then(function(text) {{
                        try {{
                            return JSON.parse(text);
                        }} catch (e) {{
                            throw new Error('Invalid JSON response: ' + text);
                        }}
                    }});
                }})
                .then(function(data) {{
                    hideLoading();
                    if (data.success) {{
                        showAlert('Email resent successfully!', 'success');
                    }} else {{
                        showAlert('Error: ' + data.message, 'danger');
                    }}
                }})
                .catch(function(error) {{
                    hideLoading();
                    showAlert('Network error: ' + error.message, 'danger');
                }});
            }}
            </script>
            
            <style>
            .pm-container {{
                max-width: 1200px;
                margin: 20px auto;
                padding: 0 20px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }}
            .pm-header {{
                background: #f8f9fa;
                padding: 15px 20px;
                border-radius: 8px;
                margin-bottom: 20px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border: 1px solid #dee2e6;
            }}
            .pm-header h3 {{
                margin: 0;
                color: #495057;
                font-size: 1.5rem;
            }}
            .pm-actions {{
                display: flex;
                gap: 10px;
            }}
            .pm-search {{
                margin-bottom: 20px;
            }}
            .pm-search .form-control {{
                max-width: 400px;
                padding: 8px 12px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-size: 14px;
            }}
            .pm-content {{
                background: white;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                border: 1px solid #dee2e6;
            }}
            .pm-table {{
                width: 100%;
                border-collapse: collapse;
                margin: 0;
            }}
            .pm-table th {{
                background: #343a40;
                color: white;
                padding: 12px;
                text-align: left;
                font-weight: 600;
                border: none;
            }}
            .pm-table td {{
                padding: 12px;
                border-bottom: 1px solid #dee2e6;
                vertical-align: top;
            }}
            .pm-table tbody tr:hover {{
                background-color: #f8f9fa;
            }}
            .pm-table tfoot {{
                background: #f8f9fa;
                font-weight: 600;
            }}
            .pm-table tfoot th {{
                background: #e9ecef;
                color: #495057;
                text-align: left;
            }}
            .pm-table tfoot th.pm-currency {{
                text-align: right;
            }}
            .pm-table tfoot th.pm-center {{
                text-align: center;
            }}
            .pm-currency {{
                text-align: right;
                font-weight: 600;
                color: #28a745;
            }}
            .pm-positive {{
                color: #28a745;
            }}
            .pm-negative {{
                color: #dc3545;
            }}
            .pm-center {{
                text-align: center;
            }}
            .pm-contact {{
                color: #6c757d;
                font-size: 0.9em;
                margin-top: 4px;
            }}
            .pm-muted {{
                color: #6c757d;
            }}
            .pm-btn-group {{
                display: flex;
                flex-direction: column;
                gap: 4px;
                min-width: 120px;
            }}
            .btn {{
                padding: 6px 12px;
                border: 1px solid;
                border-radius: 4px;
                cursor: pointer;
                text-decoration: none;
                display: inline-block;
                text-align: center;
                font-size: 14px;
                line-height: 1.4;
                background: white;
                transition: all 0.2s;
            }}
            .btn:hover {{
                opacity: 0.8;
                transform: translateY(-1px);
            }}
            .btn-primary {{
                background: #007bff;
                border-color: #007bff;
                color: white;
            }}
            .btn-secondary {{
                background: #6c757d;
                border-color: #6c757d;
                color: white;
            }}
            .btn-success {{
                background: #28a745;
                border-color: #28a745;
                color: white;
            }}
            .btn-outline-primary {{
                border-color: #007bff;
                color: #007bff;
            }}
            .btn-outline-success {{
                border-color: #28a745;
                color: #28a745;
            }}
            .btn-outline-info {{
                border-color: #17a2b8;
                color: #17a2b8;
            }}
            .btn-outline-secondary {{
                border-color: #6c757d;
                color: #6c757d;
            }}
            .btn-sm {{
                padding: 4px 8px;
                font-size: 12px;
            }}
            .fa {{
                margin-right: 4px;
            }}
            .pm-loading {{
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(255,255,255,0.8);
                display: none;
                align-items: center;
                justify-content: center;
                z-index: 9999;
            }}
            .pm-loading-spinner {{
                text-align: center;
                padding: 20px;
                background: white;
                border-radius: 8px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            }}
            .pm-hidden {{
                display: none !important;
            }}
            .email-meta {{
                margin-bottom: 30px;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 8px;
            }}
            .email-meta h4 {{
                margin-top: 0;
                color: #495057;
            }}
            .email-body {{
                padding: 20px;
            }}
            .email-body h4 {{
                margin-top: 0;
                color: #495057;
            }}
            .email-content {{
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 20px;
                background: white;
                min-height: 300px;
            }}
            .alert {{
                padding: 15px;
                margin-bottom: 20px;
                border: 1px solid transparent;
                border-radius: 4px;
            }}
            .alert-warning {{
                color: #856404;
                background-color: #fff3cd;
                border-color: #ffeaa7;
            }}
            .alert-danger {{
                color: #721c24;
                background-color: #f8d7da;
                border-color: #f5c6cb;
            }}
            .alert-success {{
                color: #155724;
                background-color: #d4edda;
                border-color: #c3e6cb;
            }}
            @keyframes spin {{
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }}
            @media (max-width: 768px) {{
                .pm-container {{
                    padding: 0 10px;
                }}
                .pm-header {{
                    flex-direction: column;
                    gap: 10px;
                    text-align: center;
                }}
                .pm-table {{
                    font-size: 14px;
                }}
                .pm-table th, .pm-table td {{
                    padding: 8px;
                }}
                .pm-btn-group {{
                    flex-direction: row;
                    flex-wrap: wrap;
                }}
            }}
            </style>
            
            <div class="pm-loading" id="pmLoading">
                <div class="pm-loading-spinner">
                    <div style="width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #007bff; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto;"></div>
                    <div style="margin-top: 10px;">Processing...</div>
                </div>
            </div>
            
            <!-- Alert container for messages -->
            <div id="alertContainer" style="position: fixed; top: 20px; right: 20px; z-index: 10001;"></div>
            
            {0}
            
            <!-- Payment Modal -->
            <div id="paymentModal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 10000; align-items: center; justify-content: center;">
                <div style="background: white; padding: 20px; border-radius: 8px; max-width: 500px; width: 90%;">
                    <h4>Record Payment</h4>
                    <div id="paymentForm">
                        <input type="hidden" id="modal-pid">
                        <input type="hidden" id="modal-org">
                        <input type="hidden" id="modal-name">
                        <input type="hidden" id="modal-emails">
                        
                        <div style="margin-bottom: 15px;">
                            <label>Payer:</label>
                            <div id="modal-payer-info" style="font-weight: bold;"></div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Amount Due:</label>
                            <div id="modal-amount-due" style="color: green; font-weight: bold;"></div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Payment Type:</label><br>
                            <label><input type="radio" name="PaymentType" value="CHK|" checked> Check</label>
                            <label style="margin-left: 15px;"><input type="radio" name="PaymentType" value="CSH|"> Cash</label>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Description:</label>
                            <input type="text" id="PaymentDescription" style="width: 100%; padding: 8px; border: 1px solid #ccc;" required>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <label>Payment Amount:</label>
                            <input type="number" id="PaidAmount" step="0.01" min="0" style="width: 100%; padding: 8px; border: 1px solid #ccc;" required>
                        </div>
                        
                        <div style="text-align: right;">
                            <button type="button" onclick="closePaymentModal()" class="btn btn-secondary">Cancel</button>
                            <button type="button" onclick="submitPayment()" class="btn btn-primary" style="margin-left: 10px;">Record Payment</button>
                        </div>
                    </div>
                </div>
            </div>
            
            <script>
            // Setup page functionality after DOM is loaded
            document.addEventListener('DOMContentLoaded', function() {{
                // Setup search functionality
                var searchInput = document.getElementById('searchInput');
                if (searchInput) {{
                    searchInput.addEventListener('input', function() {{
                        var filter = this.value.toLowerCase();
                        var table = document.getElementById('payersTable');
                        if (!table) return;
                        
                        var rows = table.querySelectorAll('tbody tr');
                        for (var i = 0; i < rows.length; i++) {{
                            var text = rows[i].textContent.toLowerCase();
                            rows[i].style.display = text.indexOf(filter) > -1 ? '' : 'none';
                        }}
                    }});
                }}
                
                // Setup transaction search
                var transactionSearch = document.getElementById('transactionSearch');
                if (transactionSearch) {{
                    transactionSearch.addEventListener('input', function() {{
                        var filter = this.value.toLowerCase();
                        var table = document.getElementById('transactionsTable');
                        if (!table) return;
                        
                        var rows = table.querySelectorAll('tbody tr');
                        for (var i = 0; i < rows.length; i++) {{
                            var text = rows[i].textContent.toLowerCase();
                            rows[i].style.display = text.indexOf(filter) > -1 ? '' : 'none';
                        }}
                    }});
                }}
                
                // Setup modal functionality
                var paymentModal = document.getElementById('paymentModal');
                if (paymentModal) {{
                    paymentModal.addEventListener('click', function(e) {{
                        if (e.target === this) {{
                            closePaymentModal();
                        }}
                    }});
                }}
                
                // Setup form submission
                var paymentForm = document.getElementById('paymentForm');
                if (paymentForm) {{
                    paymentForm.addEventListener('keypress', function(e) {{
                        if (e.key === 'Enter') {{
                            e.preventDefault();
                            submitPayment();
                        }}
                    }});
                }}
                
                hideLoading();
            }});
            </script>
            """.format(content, self.program_id, PAYMENT_LINK_DESCRIPTION)
            
        def render(self):
            """Main render method - determines which view to show and handles POST requests"""
            try:
                # Handle POST requests (AJAX actions)
                if hasattr(model.Data, 'action'):
                    action = str(model.Data.action)
                    
                    # Check if this is a POST request for AJAX actions
                    if action in ['send_payment_link', 'record_payment', 'resend_email']:
                        if action == 'send_payment_link':
                            return self.process_payment_link()
                        elif action == 'record_payment':
                            return self.process_payment_record()
                        elif action == 'resend_email':
                            return self.process_resend_email()
                
                # Handle GET requests (page views)
                content = ""
                
                if self.current_action == 'divisions':
                    program_id = str(getattr(model.Data, 'ProgramID', self.program_id))
                    content = self.render_divisions_view(program_id)
                elif self.current_action == 'payers':
                    org_id = getattr(model.Data, 'OrganizationId', None)
                    program_id = getattr(model.Data, 'ProgramID', None)
                    if org_id:
                        org_id = str(org_id)
                    if program_id:
                        program_id = str(program_id)
                    content = self.render_payers_view(org_id, program_id)
                elif self.current_action == 'emails':
                    people_id = str(getattr(model.Data, 'PeopleId', ''))
                    if people_id:
                        content = self.render_email_history(people_id)
                    else:
                        content = "<div class='alert alert-warning'>People ID is required for email history</div>"
                elif self.current_action == 'email_preview':
                    message_id = str(getattr(model.Data, 'messageId', ''))
                    people_id = str(getattr(model.Data, 'PeopleId', ''))
                    if message_id and people_id:
                        content = self.render_email_preview(message_id, people_id)
                    else:
                        content = "<div class='alert alert-warning'>Message ID and People ID are required for email preview</div>"
                elif self.current_action == 'transactions':
                    people_id = str(getattr(model.Data, 'PeopleId', ''))
                    if people_id:
                        content = self.render_transaction_history(people_id)
                    else:
                        content = "<div class='alert alert-warning'>People ID is required for transaction history</div>"
                else:
                    content = self.render_programs_view()
                
                return self.render_page_structure(content)
                
            except Exception as e:
                return self.render_error_page(str(e))
            
        def render_error_page(self, error_message):
            """Render an error page with details"""
            import traceback
            error_details = traceback.format_exc()
            
            return """
            <div style="max-width: 800px; margin: 50px auto; padding: 20px;">
                <div style="background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; padding: 20px; color: #721c24;">
                    <h3><i class="fa fa-exclamation-triangle"></i> Error</h3>
                    <p>An error occurred while processing your request:</p>
                    <p><strong>{0}</strong></p>
                    <hr>
                    <details>
                        <summary>Technical Details</summary>
                        <pre style="background: white; padding: 10px; border-radius: 4px; overflow: auto;">{1}</pre>
                    </details>
                    <div style="margin-top: 20px;">
                        <button onclick="history.go(-1)" class="btn btn-primary">Go Back</button>
                        <button onclick="location.reload()" class="btn btn-secondary" style="margin-left: 10px;">Retry</button>
                    </div>
                </div>
            </div>
            """.format(error_message, error_details)

    # Main execution
    payment_manager = ModernPaymentManager()
    
    # Set page title
    model.Title = PAGE_TITLE
    
    # Print the rendered page
    print(payment_manager.render())

except Exception as e:
    # Print any errors
    import traceback
    print("<div style='max-width: 800px; margin: 50px auto; padding: 20px;'>")
    print("<div style='background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; padding: 20px; color: #721c24;'>")
    print("<h2><i class='fa fa-exclamation-triangle'></i> System Error</h2>")
    print("<p>An unexpected error occurred: <strong>{0}</strong></p>".format(str(e)))
    print("<hr>")
    print("<details><summary>Technical Details</summary>")
    print("<pre style='background: white; padding: 10px; border-radius: 4px;'>")
    traceback.print_exc()
    print("</pre></details>")
    print("<div style='margin-top: 20px;'>")
    print("<button onclick='history.go(-1)' style='padding: 8px 16px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer;'>Go Back</button>")
    print("<button onclick='location.reload()' style='padding: 8px 16px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px;'>Retry</button>")
    print("</div></div></div>")
