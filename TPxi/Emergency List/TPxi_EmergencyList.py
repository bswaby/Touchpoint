'''
Enhanced Emergency List for TouchPoint
---------------------------------------
This report generates a comprehensive emergency contact and medical information list for selected individuals.
Perfect for events, trips, or emergency preparedness.

Written By: Ben Swaby
Email: bswaby@fbchtn.org

Features:
- Title page with counts and church information
- Missing information summary (helps identify incomplete records)
- Quick reference allergy page
- Comprehensive medical information page
- Detailed individual cards with photos, emergency contacts, and medical data

Update 2025113:
- Added variables to increase photo size for screen and print

Installation:
1. Go to Admin > Advanced > Special Content > Python Scripts
2. Click "New Python Script File"
3. Name it "EmergencyListEnhanced"
4. Paste this entire code and save

To add to Blue Toolbar:
1. Go to Admin > Advanced > Special Content > Text > CustomReports.xml
2. Add: <Report name="Enhanced Emergency List" type="PyScript" role="Access">/PyScript/EmergencyListEnhanced</Report>

Configuration:
- Adjust settings in the Config class below to customize pages and features
- Modify EXCLUDE_ANSWERS list to filter out non-meaningful responses


'''

# ===== CONFIGURATION SECTION =====
class Config:
    # Display settings
    PAGE_TITLE = "Enhanced Emergency List"
    ENTRIES_PER_PAGE = 6  # Max number of people per page (reduced due to larger photos)

    # Photo size settings (in pixels)
    # Adjust these values to make photos larger or smaller
    # Recommended ratios: Print size should be ~70% of screen size for best results
    # Larger photos may reduce ENTRIES_PER_PAGE to maintain layout
    PHOTO_SIZE_SCREEN = 140  # Profile photo size for screen display (default: 120)
    PHOTO_SIZE_PRINT = 140    # Profile photo size for print (default: 83)
    PHOTO_SIZE_FLOAT_SCREEN = 90  # Floated photo size for screen - if used (default: 70)
    PHOTO_SIZE_FLOAT_PRINT = 90   # Floated photo size for print - if used (default: 55)
    
    # Page Enable/Disable Settings
    SHOW_TITLE_PAGE = True  # Show title page with counts and church info
    SHOW_MISSING_INFO_PAGE = True  # Show missing information page
    SHOW_ALLERGY_PAGE = True  # Show quick reference allergy page
    SHOW_MEDICAL_INFO_PAGE = True  # Show comprehensive medical information page
    
    # Missing Information Flags - Set to True to flag as missing, False to ignore
    FLAG_MISSING_EMERGENCY_CONTACT = True  # Flag if emergency contact is missing
    FLAG_MISSING_AGE = True  # Flag if age/birthdate is missing
    FLAG_MISSING_PHOTO = True  # Flag if photo is missing
    
    # Extra Values to display 
    EXTRA_VALUE_FIELDS = [
        'MedicalCondition'  # Keep this in case you have extra values too
    ]
    
    # Exclude these values from being displayed as allergies or medical info
    # These are common "non-answers" that don't provide useful information
    EXCLUDE_ANSWERS = [
        'Ma', 'Mon', 'Mone', 'No ', 'Non', 'Non ', 'Nine',
        'None', 'None.', 'None ', 'None. ', 'None know', 'None know ', 'no allergies', 'No alleriges',
        'No concerns', 'None Known', 'None Known ', 'None \\nNone', 'Nonr', 'Nome',
        'No food allergies ', 'null', 'N/A', 'N/', 'N_A',
        'NKDA', 'KNA', 'NA', 'NA ', 'NKA', 'N-A', 'NS', 'N/S', 'No',
        'no food allergies', '5', ''
    ]
    
    # Note: OTC medications are dynamically pulled from the database
    # They come from either RegAnswer table (for recent registrations) or 
    # RecReg table fields (Tylenol, Advil, Maalox, Robitussin columns)
    
    # Security settings
    REQUIRED_ROLE = "Access"

# Set page header
model.Header = ''

from datetime import datetime
import sys
import re
import json

# ===== UTILITY FUNCTIONS =====

def check_permissions():
    """Check if user has required permissions"""
    if not model.UserIsInRole(Config.REQUIRED_ROLE):
        print """
        <div class="alert alert-danger">
            <h4><i class="fa fa-lock"></i> Access Denied</h4>
            <p>You need the "{0}" role to access this tool.</p>
        </div>
        """.format(Config.REQUIRED_ROLE)
        return False
    return True

def format_phone(phone):
    """Format phone number"""
    if not phone:
        return ""
    return model.FmtPhone(phone)

def get_extra_value_list():
    """Build the list of extra values for SQL query"""
    return ", ".join(["'{0}'".format(field) for field in Config.EXTRA_VALUE_FIELDS])

def print_styles():
    """Print enhanced CSS styles"""
    print """
    <style>
        /* General Styles */
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333;
        }}

        h2 {{
            color: #2c5282;
            border-bottom: 3px solid #2c5282;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}

        /* Container and layout styles */
        .emergency-list-container {{
            width: 100%;
        }}

        .person-container {{
            border-bottom: 2px solid #e2e8f0;
            page-break-inside: avoid;
            margin-bottom: 15px;
        }}

        .person-content {{
            display: flex;
            flex-direction: row;
            align-items: flex-start;
            gap: 10px;
            padding: 5px 0;
        }}
        
        /* Column specific styles */
        .photo-column {{
            flex: 0 0 {0}px;
            text-align: center;
        }}

        .medical-column {{
            flex: 1 1 45%;
            padding: 0 5px;
        }}

        .contacts-column {{
            flex: 1 1 45%;
            padding: 0 5px;
        }}

        /* Person Header */
        .person-header {{
            background-color: #f7fafc;
            border-left: 4px solid #2c5282;
            padding-left: 12px;
            margin-bottom: 10px;
        }}

        .person-number {{
            display: inline-block;
            background-color: #2c5282;
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-weight: bold;
            margin-right: 8px;
        }}

        .person-name {{
            font-size: 18px;
            font-weight: bold;
            color: #2c5282;
        }}

        .person-age {{
            font-size: 16px;
            font-weight: normal;
            color: #718096;
            margin-left: 10px;
        }}

        .member-type {{
            color: #4299e1;
            font-weight: bold;
            font-size: 14px;
            margin-top: 4px;
        }}
        
        /* Profile Image */
        .profile-img {{
            width: {0}px;
            height: {0}px;
            object-fit: cover;
            border-radius: 6px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
        }}

        /* Floated profile image */
        .profile-img-float {{
            width: {1}px;
            height: {1}px;
            object-fit: cover;
            border-radius: 6px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
            float: left;
            margin-right: 8px;
            margin-bottom: 5px;
        }}
        
        /* Contact Information */
        .contact-section {{
            background-color: #f7fafc;
            padding: 10px;
            border-radius: 6px;
            margin: 10px 0;
        }}

        .contact-label {{
            font-weight: bold;
            color: #4a5568;
            display: inline-block;
            width: 100px;
        }}

        /* Medical Information */
        .medical-section {{
            background-color: #fff5f5;
            border: 1px solid #feb2b2;
            padding: 8px;
            border-radius: 6px;
            margin: 0;
            font-size: 12px;
        }}

        .medical-header {{
            color: #c53030;
            font-weight: bold;
            font-size: 14px;
            margin-bottom: 6px;
            border-bottom: 1px solid #feb2b2;
            padding-bottom: 3px;
        }}

        .medical-field {{
            margin: 4px 0;
            padding-left: 8px;
            line-height: 1.3;
        }}

        .medical-label {{
            font-weight: bold;
            color: #742a2a;
            display: inline-block;
            width: 110px;
            font-size: 11px;
        }}

        .medical-value {{
            color: #4a5568;
        }}

        /* Highlight medical notes */
        .medical-notes {{
            background-color: #fef3c7;
            border-left: 3px solid #f59e0b;
            padding: 4px 6px;
            margin: 4px 0;
            font-weight: bold;
        }}

        /* Medication pills style */
        .medication-pills {{
            display: inline-flex;
            gap: 6px;
            flex-wrap: wrap;
        }}

        .med-pill {{
            background-color: #4299e1;
            color: white;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 500;
        }}

        /* Emergency Contact */
        .emergency-contact {{
            padding: 6px 10px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
            margin: 0 0 10px 0;
            font-size: 13px;
        }}

        .emergency-present {{
            background-color: #c6f6d5;
            color: #22543d;
            border: 1px solid #9ae6b4;
        }}

        .emergency-missing {{
            background-color: #fed7d7;
            color: #742a2a;
            border: 1px solid #fc8181;
        }}

        /* Family Information */
        .family-section {{
            background-color: #e6fffa;
            border-left: 4px solid #319795;
            padding: 6px;
            margin: 0;
            font-size: 11px;
        }}

        .family-header {{
            color: #234e52;
            font-weight: bold;
            margin-bottom: 4px;
            font-size: 12px;
        }}

        .family-member {{
            margin: 3px 0;
            padding-left: 6px;
            line-height: 1.3;
        }}

        .family-name {{
            font-weight: bold;
            color: #234e52;
        }}

        .family-email {{
            color: #2c5282;
            text-decoration: none;
        }}

        .family-email:hover {{
            text-decoration: underline;
        }}
        
        /* Print Styles */
        @media print {{
            * {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}

            /* Hide the TouchPoint page header when printing */
            #page-header {{
                display: none !important;
            }}

            body {{
                font-size: 11pt;
            }}
            
            /* Container adjustments for print */
            .emergency-list-container {{
                width: 100% !important;
            }}

            .person-container {{
                page-break-inside: avoid !important;
                margin-bottom: 10px !important;
                border-bottom: 1px solid #e2e8f0 !important;
            }}

            /* Flexbox layout for print with minimal spacing */
            .person-content {{
                display: flex !important;
                flex-direction: row !important;
                align-items: flex-start !important;
                gap: 5px !important;
                padding: 2px 0 !important;
            }}
            
            /* Photo column - fixed width, no grow */
            .photo-column {{
                flex: 0 0 {2}px !important;
                width: {2}px !important;
                padding: 0 !important;
                margin: 0 !important;
                text-align: center !important;
            }}
            
            /* Medical column - takes available space */
            .medical-column {{
                flex: 1 1 48% !important;
                padding: 0 2px !important;
                margin: 0 !important;
            }}

            /* Contacts column */
            .contacts-column {{
                flex: 1 1 48% !important;
                padding: 0 2px !important;
                margin: 0 !important;
            }}

            .person-header {{
                padding: 6px !important;
                margin-bottom: 6px !important;
                background-color: #f7fafc !important;
                border-left: 4px solid #2c5282 !important;
            }}

            .person-number {{
                background-color: #2c5282 !important;
                color: white !important;
            }}

            .person-age {{
                font-size: 14px !important;
                color: #718096 !important;
                margin-left: 8px !important;
            }}

            .medical-section {{
                background-color: #fff5f5 !important;
                border: 1px solid #feb2b2 !important;
                padding: 3px !important;
                margin: 0 0 3px 0 !important;
            }}

            .medical-header {{
                color: #c53030 !important;
                border-bottom: 1px solid #feb2b2 !important;
                font-weight: bold !important;
            }}

            .family-section {{
                background-color: #e6fffa !important;
                border-left: 3px solid #319795 !important;
                padding: 3px !important;
                margin: 0 0 3px 0 !important;
            }}

            .emergency-present {{
                background-color: #c6f6d5 !important;
                border: 1px solid #22543d !important;
                color: #22543d !important;
                padding: 4px 6px !important;
                font-size: 11px !important;
                font-weight: bold !important;
            }}

            .emergency-missing {{
                background-color: #fed7d7 !important;
                border: 1px solid #742a2a !important;
                color: #742a2a !important;
                padding: 4px 6px !important;
                font-size: 11px !important;
                font-weight: bold !important;
                font-style: italic;
            }}

            tr {{
                page-break-inside: avoid !important;
            }}
            
            .profile-img {{
                width: {2}px !important;
                height: {2}px !important;
                margin: 0 !important;
                display: block !important;
            }}

            /* Floated image in print */
            .profile-img-float {{
                width: {3}px !important;
                height: {3}px !important;
                float: left !important;
                margin-right: 5px !important;
                margin-bottom: 3px !important;
            }}
            
            .medical-label {{
                width: 95px !important;
                font-size: 10px !important;
            }}

            .medical-field {{
                margin: 2px 0 !important;
                padding-left: 5px !important;
            }}

            .medical-header,
            .family-header {{
                font-size: 12px !important;
                margin-bottom: 3px !important;
            }}

            .med-pill {{
                background-color: #4299e1 !important;
                color: white !important;
                padding: 2px 6px !important;
                font-size: 10px !important;
                border-radius: 10px !important;
                font-weight: bold !important;
            }}

            .medical-notes {{
                background-color: #fef3c7 !important;
                border-left: 3px solid #f59e0b !important;
                padding: 4px 6px !important;
                margin: 4px 0 !important;
                font-weight: bold !important;
            }}
        }}
        
        /* Responsive adjustments - only for screen, not print */
        @media screen and (max-width: 768px) {{
            .person-content {{
                flex-direction: column;
                gap: 10px;
            }}

            .photo-column {{
                flex: 0 0 auto;
                width: 100%;
                text-align: center;
            }}

            .medical-column,
            .contacts-column {{
                flex: 1 1 100%;
                width: 100%;
            }}

            .profile-img {{
                margin: 0 auto;
                display: block;
            }}
        }}
    </style>
    """.format(
        Config.PHOTO_SIZE_SCREEN,
        Config.PHOTO_SIZE_FLOAT_SCREEN,
        Config.PHOTO_SIZE_PRINT,
        Config.PHOTO_SIZE_FLOAT_PRINT
    )

def get_member_type_description(people_id, org_id):
    """Get organization member type description"""
    if not org_id:
        return ""
        
    sql = """
    SELECT mt.Description AS OrgMemType 
    FROM OrganizationMembers om
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    WHERE om.PeopleId = {0} AND om.OrganizationId = {1}
    """.format(people_id, org_id)
    
    result = q.QuerySqlTop1(sql)
    if result and result.OrgMemType and result.OrgMemType != 'Member':
        return result.OrgMemType
    return ""

def format_medical_value(value, label=""):
    """Format medical field value"""
    if not value:
        return ""
    
    # Handle boolean values (medication fields)
    if isinstance(value, bool):
        return ""
    
    # Convert to string and strip whitespace
    value_str = str(value).strip()
    
    # Check if value is in the exclusion list (case-insensitive)
    if value_str in Config.EXCLUDE_ANSWERS:
        return ""
    
    # Also check lowercase version
    if value_str.lower() in [x.lower() for x in Config.EXCLUDE_ANSWERS]:
        return ""
    
    # Check for empty/invalid values
    if value_str.upper() in ['UNKNOWN TYPE:', '']:
        return ""
    
    return value_str

def print_title_page(people_data, org_name=None):
    """Print title page with counts and statistics"""
    if not Config.SHOW_TITLE_PAGE:
        return  # Skip this page if disabled
    male_count = 0
    female_count = 0
    unknown_count = 0
    
    for person in people_data:
        # Get gender from database
        gender_sql = """
        SELECT GenderId FROM People WHERE PeopleId = {0}
        """.format(person.PeopleId)
        gender_result = q.QuerySqlTop1(gender_sql)
        
        if gender_result:
            if gender_result.GenderId == 1:  # Male
                male_count += 1
            elif gender_result.GenderId == 2:  # Female
                female_count += 1
            else:
                unknown_count += 1
        else:
            unknown_count += 1
    
    # Get church information from Setting table
    church_info_sql = """
    SELECT TOP 1
        Setting
    FROM dbo.Setting
    WHERE Id = 'NameOfChurch'
    """
    church_name_result = q.QuerySqlTop1(church_info_sql)
    church_name = church_name_result.Setting if church_name_result else "Church"
    
    # Get church contact information
    contact_sql = """
    SELECT TOP 1
        Setting
    FROM dbo.Setting
    WHERE Id IN ('ChurchPhone', 'ChurchEmail')
    ORDER BY Id
    """
    contact_info = q.QuerySql(contact_sql)
    church_phone = ""
    church_email = ""
    for info in contact_info:
        if info.Setting:
            if '@' in info.Setting:
                church_email = info.Setting
            else:
                church_phone = format_phone(info.Setting)
    
    total_count = len(people_data)
    current_time = datetime.now().strftime("%B %d, %Y at %I:%M %p")
    
    print """
    <div style="page-break-after: always; padding: 50px; text-align: center; position: relative;">
        <div style="position: absolute; top: 20px; right: 20px; color: #c53030; font-weight: bold; font-size: 16px; border: 2px solid #c53030; padding: 5px 15px; background-color: #fed7d7;">
            CONFIDENTIAL
        </div>
        <h1 style="font-size: 36px; color: #2c5282; margin-bottom: 10px;">Emergency List</h1>
        <h2 style="font-size: 24px; color: #4a5568; margin-bottom: 30px;">{0}</h2>
        {1}
        <div style="margin: 40px auto; padding: 30px; background-color: #f7fafc; border-radius: 10px; max-width: 600px;">
            <h2 style="color: #4a5568; margin-bottom: 20px;">Report Summary</h2>
            <div style="font-size: 18px; line-height: 2;">
                <div><strong>Total People:</strong> {2}</div>
                <div><strong>Male:</strong> {3}</div>
                <div><strong>Female:</strong> {4}</div>
                {5}
                <div style="margin-top: 20px; padding-top: 20px; border-top: 2px solid #e2e8f0;">
                    <strong>Report Generated:</strong><br>{6}
                </div>
            </div>
        </div>
        <div style="margin-top: 30px; padding: 20px; background-color: #e6fffa; border-radius: 10px; max-width: 600px; margin: 30px auto;">
            <h3 style="color: #234e52; margin-bottom: 15px;">Church Contact Information</h3>
            <div style="font-size: 16px; line-height: 1.8;">
                {7}
                {8}
            </div>
        </div>
    </div>
    """.format(
        church_name,
        '<h3 style="color: #718096; margin-bottom: 10px;">{0}</h3>'.format(org_name) if org_name else '',
        total_count,
        male_count,
        female_count,
        '<div><strong>Unknown Gender:</strong> {0}</div>'.format(unknown_count) if unknown_count > 0 else '',
        current_time,
        '<div><strong>Phone:</strong> {0}</div>'.format(church_phone) if church_phone else '',
        '<div><strong>Email:</strong> {0}</div>'.format(church_email) if church_email else ''
    )

def print_missing_items_page(people_data):
    """Print page with missing information that needs to be resolved"""
    if not Config.SHOW_MISSING_INFO_PAGE:
        return  # Skip this page if disabled
        
    missing_items = []
    
    for person in people_data:
        missing_fields = []
        
        # Check for missing emergency contact (if flagged)
        if Config.FLAG_MISSING_EMERGENCY_CONTACT:
            if not person.emcontact or not person.emphone:
                missing_fields.append('Emergency Contact')
        
        # Check for missing age (if flagged)
        if Config.FLAG_MISSING_AGE:
            if not person.Age:
                missing_fields.append('Age/Birthdate')
        
        # Check for missing picture (if flagged)
        if Config.FLAG_MISSING_PHOTO:
            if not person.pic:
                missing_fields.append('Photo')
        
        # If any fields are missing, add to the list
        if missing_fields:
            missing_items.append({
                'name': person.Name2,
                'age': person.Age if person.Age else 'Missing',
                'missing': ', '.join(missing_fields),
                'people_id': person.PeopleId
            })
    
    if missing_items:
        print """
        <div style="page-break-after: always; padding: 20px; position: relative;">
            <div style="position: absolute; top: 10px; right: 20px; color: #c53030; font-weight: bold; font-size: 14px; border: 1px solid #c53030; padding: 3px 10px; background-color: #fed7d7;">
                CONFIDENTIAL
            </div>
            <h2 style="color: #f59e0b; border-bottom: 3px solid #f59e0b; padding-bottom: 10px;">Missing Information - Action Required</h2>
            <p style="color: #92400e; font-weight: bold; margin-bottom: 20px;">
                The following people have incomplete emergency information that needs to be updated:
            </p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
                <thead>
                    <tr style="background-color: #fef3c7;">
                        <th style="padding: 8px; text-align: left; border: 1px solid #f59e0b;">Name</th>
                        <th style="padding: 8px; text-align: center; border: 1px solid #f59e0b; width: 80px;">Age</th>
                        <th style="padding: 8px; text-align: left; border: 1px solid #f59e0b;">Missing Information</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for item in sorted(missing_items, key=lambda x: x['name']):
            # Highlight row based on severity of missing info
            row_style = ""
            if 'Emergency Contact' in item['missing']:
                row_style = "background-color: #fee2e2;"  # Light red for critical
            elif 'Age/Birthdate' in item['missing']:
                row_style = "background-color: #fef3c7;"  # Light yellow for important
            
            print """
                    <tr style="{0}">
                        <td style="padding: 6px; border: 1px solid #e2e8f0; font-weight: bold;">{1}</td>
                        <td style="padding: 6px; border: 1px solid #e2e8f0; text-align: center;">{2}</td>
                        <td style="padding: 6px; border: 1px solid #e2e8f0; color: #92400e; font-weight: 500;">{3}</td>
                    </tr>
            """.format(row_style, item['name'], item['age'], item['missing'])
        
        print """
                </tbody>
            </table>
            <div style="margin-top: 20px; padding: 15px; background-color: #fef3c7; border-left: 4px solid #f59e0b;">
                <strong style="color: #92400e;">Priority Legend:</strong>
                <ul style="margin-top: 10px; color: #92400e;">
                    <li><span style="background-color: #fee2e2; padding: 2px 8px;">Red Background</span> - Missing Emergency Contact (Critical)</li>
                    <li><span style="background-color: #fef3c7; padding: 2px 8px;">Yellow Background</span> - Missing Age/Birthdate (Important)</li>
                    <li>White Background - Missing Photo Only (Low Priority)</li>
                </ul>
            </div>
        </div>
        """
    else:
        # If no missing items, print a success page
        print """
        <div style="page-break-after: always; padding: 20px; position: relative;">
            <div style="position: absolute; top: 10px; right: 20px; color: #c53030; font-weight: bold; font-size: 14px; border: 1px solid #c53030; padding: 3px 10px; background-color: #fed7d7;">
                CONFIDENTIAL
            </div>
            <h2 style="color: #059669; border-bottom: 3px solid #059669; padding-bottom: 10px;">Data Completeness Status</h2>
            <div style="margin-top: 30px; padding: 30px; background-color: #d1fae5; border-radius: 10px; text-align: center;">
                <i class="fa fa-check-circle" style="font-size: 48px; color: #059669; margin-bottom: 20px;"></i>
                <h3 style="color: #065f46; margin-bottom: 10px;">All Information Complete!</h3>
                <p style="color: #047857; font-size: 18px;">
                    Every person in this report has complete emergency information including contact details, age, and photos.
                </p>
            </div>
        </div>
        """

def print_allergy_page(people_data):
    """Print allergy page with name, age, and allergies"""
    if not Config.SHOW_ALLERGY_PAGE:
        return  # Skip this page if disabled
    allergies_list = []
    
    for person in people_data:
        if person.MedAllergy and format_medical_value(person.MedAllergy):
            allergies_list.append({
                'name': person.Name2,
                'age': person.Age if person.Age else 'N/A',
                'allergy': format_medical_value(person.MedAllergy)
            })
    
    if allergies_list:
        print """
        <div style="page-break-after: always; padding: 20px; position: relative;">
            <div style="position: absolute; top: 10px; right: 20px; color: #c53030; font-weight: bold; font-size: 14px; border: 1px solid #c53030; padding: 3px 10px; background-color: #fed7d7;">
                CONFIDENTIAL
            </div>
            <h2 style="color: #c53030; border-bottom: 3px solid #c53030; padding-bottom: 10px;">Allergy Information</h2>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
                <thead>
                    <tr style="background-color: #fed7d7;">
                        <th style="padding: 10px; text-align: left; border: 1px solid #fc8181;">Name</th>
                        <th style="padding: 10px; text-align: center; border: 1px solid #fc8181; width: 80px;">Age</th>
                        <th style="padding: 10px; text-align: left; border: 1px solid #fc8181;">Allergies</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for allergy_info in sorted(allergies_list, key=lambda x: x['name']):
            print """
                    <tr>
                        <td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">{0}</td>
                        <td style="padding: 8px; border: 1px solid #e2e8f0; text-align: center;">{1}</td>
                        <td style="padding: 8px; border: 1px solid #e2e8f0; color: #c53030;">{2}</td>
                    </tr>
            """.format(allergy_info['name'], allergy_info['age'], allergy_info['allergy'])
        
        print """
                </tbody>
            </table>
        </div>
        """

def print_adhoc_values_page(people_data):
    """Print medical information page with all medical data"""
    if not Config.SHOW_MEDICAL_INFO_PAGE:
        return  # Skip this page if disabled
    # Get all people with medical conditions, allergies, or other medical info
    adhoc_dict = {}  # Use dict to group by person
    
    for person in people_data:
        person_key = person.Name2
        person_age = person.Age if person.Age else 'N/A'
        
        # Check for allergies from MedAllergy field
        if person.MedAllergy and format_medical_value(person.MedAllergy):
            if person_key not in adhoc_dict:
                adhoc_dict[person_key] = {
                    'age': person_age,
                    'items': []
                }
            adhoc_dict[person_key]['items'].append({
                'type': 'Allergies',
                'value': format_medical_value(person.MedAllergy)
            })
        
        # Check for allergies from MedicalDescription field (also used for allergies)
        if person.MedicalDescription and format_medical_value(person.MedicalDescription):
            if person_key not in adhoc_dict:
                adhoc_dict[person_key] = {
                    'age': person_age,
                    'items': []
                }
            adhoc_dict[person_key]['items'].append({
                'type': 'Allergies',
                'value': format_medical_value(person.MedicalDescription)
            })
        
        # Check for medical conditions from extra values
        if person.MedicalCondition and format_medical_value(person.MedicalCondition):
            if person_key not in adhoc_dict:
                adhoc_dict[person_key] = {
                    'age': person_age,
                    'items': []
                }
            adhoc_dict[person_key]['items'].append({
                'type': 'Medical Condition',
                'value': format_medical_value(person.MedicalCondition)
            })
    
    if adhoc_dict:
        print """
        <div style="page-break-after: always; padding: 20px; position: relative;">
            <div style="position: absolute; top: 10px; right: 20px; color: #c53030; font-weight: bold; font-size: 14px; border: 1px solid #c53030; padding: 3px 10px; background-color: #fed7d7;">
                CONFIDENTIAL
            </div>
            <h2 style="color: #319795; border-bottom: 3px solid #319795; padding-bottom: 10px;">Medical Information</h2>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
                <thead>
                    <tr style="background-color: #e6fffa;">
                        <th style="padding: 6px; text-align: left; border: 1px solid #4fd1c5;">Name</th>
                        <th style="padding: 6px; text-align: center; border: 1px solid #4fd1c5; width: 60px;">Age</th>
                        <th style="padding: 6px; text-align: left; border: 1px solid #4fd1c5; width: 150px;">Type</th>
                        <th style="padding: 6px; text-align: left; border: 1px solid #4fd1c5;">Information</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for person_name in sorted(adhoc_dict.keys()):
            person_data = adhoc_dict[person_name]
            first_row = True
            
            for item in person_data['items']:
                if first_row:
                    print """
                    <tr>
                        <td style="padding: 4px 6px; border: 1px solid #e2e8f0; font-weight: bold; vertical-align: top;" rowspan="{0}">{1}</td>
                        <td style="padding: 4px 6px; border: 1px solid #e2e8f0; text-align: center; vertical-align: top;" rowspan="{0}">{2}</td>
                        <td style="padding: 4px 6px; border: 1px solid #e2e8f0; color: #234e52; font-weight: bold;">{3}</td>
                        <td style="padding: 4px 6px; border: 1px solid #e2e8f0;">{4}</td>
                    </tr>
                    """.format(len(person_data['items']), person_name, person_data['age'], item['type'], item['value'])
                    first_row = False
                else:
                    print """
                    <tr>
                        <td style="padding: 4px 6px; border: 1px solid #e2e8f0; color: #234e52; font-weight: bold;">{0}</td>
                        <td style="padding: 4px 6px; border: 1px solid #e2e8f0;">{1}</td>
                    </tr>
                    """.format(item['type'], item['value'])
        
        print """
                </tbody>
            </table>
        </div>
        """

def main():
    """Main function"""
    try:
        # Check permissions
        if not check_permissions():
            return
            
        # Get organization ID if running from organization context
        org_id = getattr(model.Data, 'CurrentOrgId', None)
        org_name = None
        
        # Configuration for pagination
        count_loop = 94
        first_page = Config.ENTRIES_PER_PAGE + count_loop
        
        # Print styles
        print_styles()
        
        # Get organization name if applicable
        if org_id:
            sql_header = """
            SELECT TOP 1 os.Organization, os.Program, os.Division 
            FROM OrganizationStructure os 
            WHERE OrgId = {0}
            """.format(org_id)
            
            header_data = q.QuerySqlTop1(sql_header)
            if header_data:
                org_name = header_data.Organization
        
        # Build main query with medical fields from RecReg table
        sql_main = """
        SELECT DISTINCT 
            p.Name2,
            p.PeopleId, 
            p.FamilyId,
            p.Age,
            p.PrimaryAddress,
            p.PrimaryCity,
            p.PrimaryState,
            p.PrimaryZip,
            p.CellPhone,
            -- Medical fields from RecReg
            rr.MedicalDescription,
            rr.MedAllergy,
            rr.emcontact,
            rr.emphone,
            rr.doctor,
            rr.docphone,
            rr.insurance,
            rr.policy,
            rr.Comments,
            -- OTC Medications
            rr.Tylenol,
            rr.Advil,
            rr.Maalox,
            rr.Robitussin,
            -- Picture
            pic1.SmallUrl AS pic,
            -- Extra values if needed
            pe1.Data AS MedicalCondition
        FROM People p
            LEFT JOIN Picture pic1 ON pic1.PictureId = p.PictureId
            LEFT JOIN RecReg rr ON rr.PeopleId = p.PeopleId
            LEFT JOIN PeopleExtra pe1 ON pe1.PeopleId = p.PeopleId AND pe1.Field = 'MedicalCondition'
        WHERE 
            p.PeopleId IN (SELECT p.PeopleId 
                          FROM dbo.People p 
                          JOIN dbo.TagPerson tp ON tp.PeopleId = p.PeopleId 
                          WHERE tp.Id = @BlueToolbarTagId)
        ORDER BY p.Name2
        """
        
        # Get data
        people_data = q.QuerySql(sql_main)
        
        # Convert to list if needed for multiple iterations
        people_list = list(people_data)
        
        # Print title page first
        print_title_page(people_list, org_name)
        
        # Print missing items page (for data quality)
        print_missing_items_page(people_list)
        
        # Print allergy page (quick reference for allergies only)
        print_allergy_page(people_list)
        
        # Print medical information page (comprehensive medical data)
        print_adhoc_values_page(people_list)
        
        # Now print the main emergency list
        person_count = 0
        
        # Print organization header if applicable with CONFIDENTIAL marking
        if org_name:
            print """
            <div style="position: relative;">
                <div style="position: absolute; top: 10px; right: 20px; color: #c53030; font-weight: bold; font-size: 14px; border: 1px solid #c53030; padding: 3px 10px; background-color: #fed7d7;">
                    CONFIDENTIAL
                </div>
                <h2>{0}</h2>
            </div>
            """.format(org_name)
            first_page = first_page - 1
        else:
            print """
            <div style="position: absolute; top: 10px; right: 20px; color: #c53030; font-weight: bold; font-size: 14px; border: 1px solid #c53030; padding: 3px 10px; background-color: #fed7d7;">
                CONFIDENTIAL
            </div>
            """
        
        print '<div class="emergency-list-container">'
        
        # Process each person
        for person in people_list:
            person_count += 1
            
            # Get member type if from organization
            member_type = ""
            if org_id:
                member_type = get_member_type_description(person.PeopleId, org_id)
            
            # Start person container
            print '<div class="person-container">'
            
            # Person header
            print """
            <div class="person-header">
                <span class="person-number">{0}</span>
                <span class="person-name">{1}</span>
                <span class="person-age">(Age {2})</span>
                {3}
            </div>
            """.format(
                person_count,
                person.Name2,
                person.Age if person.Age else 'N/A',
                '<div class="member-type">{0}</div>'.format(member_type) if member_type else ''
            )
            
            # Main content row with flexbox
            print '<div class="person-content">'
            
            # Column 1 - Photo
            print """
            <div class="photo-column">
                <img class="profile-img" src="{0}" 
                     onerror="this.onerror=null; this.src='https://c4265878.ssl.cf2.rackcdn.com/fbchville.2502091552.Hey__I_am_beautiful._Consider_adding_a_photo_-1-.png';" 
                     alt="Profile Photo">
            </div>
            """.format(person.pic or '')
            
            # Column 2 - Medical Information
            print '<div class="medical-column">'
            
            # Medical Information
            has_medical_info = any([
                format_medical_value(person.MedicalDescription),
                format_medical_value(person.MedAllergy),
                format_medical_value(person.MedicalCondition),
                format_medical_value(person.doctor),
                format_medical_value(person.insurance),
                person.Tylenol or person.Advil or person.Maalox or person.Robitussin
            ])
            
            if has_medical_info:
                print '<div class="medical-section">'
                print '<div class="medical-header">Medical Information</div>'
                
                # Allergies from MedAllergy field
                allergies = format_medical_value(person.MedAllergy)
                if allergies:
                    print '<div class="medical-field"><span class="medical-label">Allergies:</span> <span class="medical-value">{0}</span></div>'.format(allergies)
                
                # Additional Allergies from MedicalDescription field (often used for allergies)
                med_desc = format_medical_value(person.MedicalDescription)
                if med_desc:
                    print '<div class="medical-notes"><span class="medical-label" style="color: #92400e;">Allergies:</span> {0}</div>'.format(med_desc)
                
                # Medical Conditions from Extra Values
                conditions = format_medical_value(person.MedicalCondition)
                if conditions:
                    print '<div class="medical-field"><span class="medical-label">Medical Conditions:</span> <span class="medical-value">{0}</span></div>'.format(conditions)
                
                # Doctor Information from RecReg
                doctor = format_medical_value(person.doctor)
                if doctor:
                    doctor_info = doctor
                    if person.docphone:
                        doctor_info += ' - ' + format_phone(person.docphone)
                    print '<div class="medical-field"><span class="medical-label">Primary Doctor:</span> <span class="medical-value">{0}</span></div>'.format(doctor_info)
                
                # Insurance Information from RecReg
                insurance = format_medical_value(person.insurance)
                if insurance:
                    insurance_info = insurance
                    if person.policy:
                        insurance_info += ' (Policy: {0})'.format(person.policy)
                    print '<div class="medical-field"><span class="medical-label">Health Insurance:</span> <span class="medical-value">{0}</span></div>'.format(insurance_info)
                
                # OTC Medications allowed
                # First check if there's a registration answer override for this specific question
                # The RegAnswer table stores answers in JSON array format like ["Tylenol","Advil"]
                reg_answer_sql = """
                SELECT ra.AnswerValue
                FROM RegAnswer ra
                INNER JOIN RegPeople rp ON rp.RegPeopleId = ra.RegPeopleId
                WHERE rp.PeopleId = {0} 
                  AND ra.RegQuestionId = '8A9F1199-6F1C-480C-8A2D-146EEEAE55B8'
                """.format(person.PeopleId)
                
                reg_answer = q.QuerySqlTop1(reg_answer_sql)
                allowed_meds = []
                
                if reg_answer and reg_answer.AnswerValue:
                    # Parse the JSON-like format ["meda","medb"]
                    answer_value = reg_answer.AnswerValue
                    # Remove brackets and quotes, then split by comma
                    if answer_value.startswith('[') and answer_value.endswith(']'):
                        answer_value = answer_value[1:-1]  # Remove brackets
                    # Split by comma and clean up each medication
                    meds_list = answer_value.split(',')
                    for med in meds_list:
                        # Remove quotes and whitespace
                        cleaned_med = med.strip().strip('"').strip("'")
                        if cleaned_med:
                            allowed_meds.append(cleaned_med)
                else:
                    # Fall back to RecReg table values
                    if person.Tylenol:
                        allowed_meds.append('Tylenol')
                    if person.Advil:
                        allowed_meds.append('Advil')
                    if person.Maalox:
                        allowed_meds.append('Maalox')
                    if person.Robitussin:
                        allowed_meds.append('Robitussin')
                
                if allowed_meds:
                    print '<div class="medical-field"><span class="medical-label">Allowed OTC Meds:</span> <span class="medication-pills">'
                    for med in allowed_meds:
                        print '<span class="med-pill">{0}</span>'.format(med)
                    print '</span></div>'
                
                
                print '</div>'
            
            print '</div>'  # End medical column
            
            # Column 3 - Contacts
            print '<div class="contacts-column">'
            
            # Emergency Contact
            if person.emcontact:
                print '<div class="emergency-contact emergency-present">Emergency: {0} {1}</div>'.format(
                    person.emcontact,
                    format_phone(person.emphone)
                )
            else:
                print '<div class="emergency-contact emergency-missing">Emergency Contact Missing</div>'
            
            # Family Information
            sql_family = """
            SELECT p.Name, p.CellPhone, p.HomePhone, p.WorkPhone, p.EmailAddress
            FROM People p 
            WHERE FamilyId = {0} AND PositionInFamilyId IN (10, 20)
            ORDER BY p.Name
            """.format(person.FamilyId)
            
            family_data = q.QuerySql(sql_family)
            if family_data:
                print '<div class="family-section">'
                print '<div class="family-header">Family Contacts</div>'
                
                for family in family_data:
                    print '<div class="family-member">'
                    print '<span class="family-name">{0}</span><br>'.format(family.Name)
                    
                    # Only show cell phone to save space
                    if family.CellPhone:
                        print 'Cell: {0}<br>'.format(format_phone(family.CellPhone))
                    elif family.HomePhone:
                        print 'Home: {0}<br>'.format(format_phone(family.HomePhone))
                    
                    print '</div>'
                
                print '</div>'
            
            print '</div>'  # End contacts column
            print '</div>'  # End person-content
            print '</div>'  # End person-container
            
            # Check for page break
            count_loop += 1
            if count_loop == Config.ENTRIES_PER_PAGE or count_loop == first_page:
                print '</div><p style="page-break-after: always;">&nbsp;</p><div class="emergency-list-container">'
                count_loop = 0
        
        print '</div>'
        
        # Only show a message if no people were found
        if person_count == 0:
            print """
            <div class="alert alert-info">
                <i class="fa fa-info-circle"></i> No people found in the current selection.
            </div>
            """
            
    except Exception as e:
        print """
        <div class="alert alert-danger">
            <h4><i class="fa fa-exclamation-circle"></i> Error</h4>
            <p>{0}</p>
        </div>
        """.format(str(e))

# Execute main function
main()
