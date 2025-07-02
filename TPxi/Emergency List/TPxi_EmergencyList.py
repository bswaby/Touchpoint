'''
Purpose: Enhanced Emergency List - Medical and emergency contact information with improved visual design

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python script "EmergencyList" and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Add to Blue Toolbar by:
1. Open Admin ~ Advanced ~ Special Content ~ Text > CustomReports.xml
2. Add: <Report name="Enhanced Emergency List" type="PyScript" role="Access">/PyScript/EmergencyList</Report>

Note:  CustomReports.xml take up to 24hrs to show

Written By: Ben Swaby
Email: bswaby@fbchtn.org
'''

# ===== CONFIGURATION SECTION =====
class Config:
    # Display settings
    PAGE_TITLE = "Enhanced Emergency List"
    ENTRIES_PER_PAGE = 5  # Max number of people per page (reduced due to larger photos)
    
    # Extra Values to display (if any - currently using RecReg table fields)
    EXTRA_VALUE_FIELDS = [
        'MedicalCondition'  # Keep this in case you have extra values too
    ]
    
    # Over-the-counter medications from RecReg table
    OTC_MEDICATIONS = ['Tylenol', 'Advil', 'Maalox', 'Robitussin']
    
    # Security settings
    REQUIRED_ROLE = "Edit"

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
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333;
        }
        
        h2 {
            color: #2c5282;
            border-bottom: 3px solid #2c5282;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        
        /* Container and layout styles */
        .emergency-list-container {
            width: 100%;
        }
        
        .person-container {
            border-bottom: 2px solid #e2e8f0;
            page-break-inside: avoid;
            margin-bottom: 15px;
        }
        
        .person-content {
            display: flex;
            flex-direction: row;
            align-items: flex-start;
            gap: 10px;
            padding: 5px 0;
        }
        
        /* Column specific styles */
        .photo-column {
            flex: 0 0 113px;
            text-align: center;
        }
        
        .medical-column {
            flex: 1 1 45%;
            padding: 0 5px;
        }
        
        .contacts-column {
            flex: 1 1 45%;
            padding: 0 5px;
        }
        
        /* Person Header */
        .person-header {
            background-color: #f7fafc;
            border-left: 4px solid #2c5282;
            padding-left: 12px;
            margin-bottom: 10px;
        }
        
        .person-number {
            display: inline-block;
            background-color: #2c5282;
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-weight: bold;
            margin-right: 8px;
        }
        
        .person-name {
            font-size: 18px;
            font-weight: bold;
            color: #2c5282;
        }
        
        .person-age {
            font-size: 16px;
            font-weight: normal;
            color: #718096;
            margin-left: 10px;
        }
        
        .member-type {
            color: #4299e1;
            font-weight: bold;
            font-size: 14px;
            margin-top: 4px;
        }
        
        /* Profile Image */
        .profile-img {
            width: 120px;
            height: 120px;
            object-fit: cover;
            border-radius: 6px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
        }
        
        /* Floated profile image */
        .profile-img-float {
            width: 70px;
            height: 70px;
            object-fit: cover;
            border-radius: 6px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
            float: left;
            margin-right: 8px;
            margin-bottom: 5px;
        }
        
        /* Contact Information */
        .contact-section {
            background-color: #f7fafc;
            padding: 10px;
            border-radius: 6px;
            margin: 10px 0;
        }
        
        .contact-label {
            font-weight: bold;
            color: #4a5568;
            display: inline-block;
            width: 100px;
        }
        
        /* Medical Information */
        .medical-section {
            background-color: #fff5f5;
            border: 1px solid #feb2b2;
            padding: 8px;
            border-radius: 6px;
            margin: 0;
            font-size: 12px;
        }
        
        .medical-header {
            color: #c53030;
            font-weight: bold;
            font-size: 14px;
            margin-bottom: 6px;
            border-bottom: 1px solid #feb2b2;
            padding-bottom: 3px;
        }
        
        .medical-field {
            margin: 4px 0;
            padding-left: 8px;
            line-height: 1.3;
        }
        
        .medical-label {
            font-weight: bold;
            color: #742a2a;
            display: inline-block;
            width: 110px;
            font-size: 11px;
        }
        
        .medical-value {
            color: #4a5568;
        }
        
        /* Highlight medical notes */
        .medical-notes {
            background-color: #fef3c7;
            border-left: 3px solid #f59e0b;
            padding: 4px 6px;
            margin: 4px 0;
            font-weight: bold;
        }
        
        /* Medication pills style */
        .medication-pills {
            display: inline-flex;
            gap: 6px;
            flex-wrap: wrap;
        }
        
        .med-pill {
            background-color: #4299e1;
            color: white;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 500;
        }
        
        /* Emergency Contact */
        .emergency-contact {
            padding: 6px 10px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
            margin: 0 0 10px 0;
            font-size: 13px;
        }
        
        .emergency-present {
            background-color: #c6f6d5;
            color: #22543d;
            border: 1px solid #9ae6b4;
        }
        
        .emergency-missing {
            background-color: #fed7d7;
            color: #742a2a;
            border: 1px solid #fc8181;
        }
        
        /* Family Information */
        .family-section {
            background-color: #e6fffa;
            border-left: 4px solid #319795;
            padding: 6px;
            margin: 0;
            font-size: 11px;
        }
        
        .family-header {
            color: #234e52;
            font-weight: bold;
            margin-bottom: 4px;
            font-size: 12px;
        }
        
        .family-member {
            margin: 3px 0;
            padding-left: 6px;
            line-height: 1.3;
        }
        
        .family-name {
            font-weight: bold;
            color: #234e52;
        }
        
        .family-email {
            color: #2c5282;
            text-decoration: none;
        }
        
        .family-email:hover {
            text-decoration: underline;
        }
        
        /* Print Styles */
        @media print {
            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            
            body {
                font-size: 11pt;
            }
            
            /* Container adjustments for print */
            .emergency-list-container {
                width: 100% !important;
            }
            
            .person-container {
                page-break-inside: avoid !important;
                margin-bottom: 10px !important;
                border-bottom: 1px solid #e2e8f0 !important;
            }
            
            /* Flexbox layout for print with minimal spacing */
            .person-content {
                display: flex !important;
                flex-direction: row !important;
                align-items: flex-start !important;
                gap: 5px !important;
                padding: 2px 0 !important;
            }
            
            /* Photo column - fixed width, no grow */
            .photo-column {
                flex: 0 0 90px !important;
                width: 90px !important;
                padding: 0 !important;
                margin: 0 !important;
                text-align: center !important;
            }
            
            /* Medical column - takes available space */
            .medical-column {
                flex: 1 1 48% !important;
                padding: 0 2px !important;
                margin: 0 !important;
            }
            
            /* Contacts column */
            .contacts-column {
                flex: 1 1 48% !important;
                padding: 0 2px !important;
                margin: 0 !important;
            }
            
            .person-header {
                padding: 6px !important;
                margin-bottom: 6px !important;
                background-color: #f7fafc !important;
                border-left: 4px solid #2c5282 !important;
            }
            
            .person-number {
                background-color: #2c5282 !important;
                color: white !important;
            }
            
            .person-age {
                font-size: 14px !important;
                color: #718096 !important;
                margin-left: 8px !important;
            }
            
            .medical-section {
                background-color: #fff5f5 !important;
                border: 1px solid #feb2b2 !important;
                padding: 3px !important;
                margin: 0 0 3px 0 !important;
            }
            
            .medical-header {
                color: #c53030 !important;
                border-bottom: 1px solid #feb2b2 !important;
                font-weight: bold !important;
            }
            
            .family-section {
                background-color: #e6fffa !important;
                border-left: 3px solid #319795 !important;
                padding: 3px !important;
                margin: 0 0 3px 0 !important;
            }
            
            .emergency-present {
                background-color: #c6f6d5 !important;
                border: 1px solid #22543d !important;
                color: #22543d !important;
                padding: 4px 6px !important;
                font-size: 11px !important;
                font-weight: bold !important;
            }
            
            .emergency-missing {
                background-color: #fed7d7 !important;
                border: 1px solid #742a2a !important;
                color: #742a2a !important;
                padding: 4px 6px !important;
                font-size: 11px !important;
                font-weight: bold !important;
                font-style: italic;
            }
            
            tr {
                page-break-inside: avoid !important;
            }
            
            .profile-img {
                width: 83px !important;
                height: 83px !important;
                margin: 0 !important;
                display: block !important;
            }
            
            /* Floated image in print */
            .profile-img-float {
                width: 55px !important;
                height: 55px !important;
                float: left !important;
                margin-right: 5px !important;
                margin-bottom: 3px !important;
            }
            
            .medical-label {
                width: 95px !important;
                font-size: 10px !important;
            }
            
            .medical-field {
                margin: 2px 0 !important;
                padding-left: 5px !important;
            }
            
            .medical-header,
            .family-header {
                font-size: 12px !important;
                margin-bottom: 3px !important;
            }
            
            .med-pill {
                background-color: #4299e1 !important;
                color: white !important;
                padding: 2px 6px !important;
                font-size: 10px !important;
                border-radius: 10px !important;
                font-weight: bold !important;
            }
            
            .medical-notes {
                background-color: #fef3c7 !important;
                border-left: 3px solid #f59e0b !important;
                padding: 4px 6px !important;
                margin: 4px 0 !important;
                font-weight: bold !important;
            }
        }
        
        /* Responsive adjustments - only for screen, not print */
        @media screen and (max-width: 768px) {
            .person-content {
                flex-direction: column;
                gap: 10px;
            }
            
            .photo-column {
                flex: 0 0 auto;
                width: 100%;
                text-align: center;
            }
            
            .medical-column,
            .contacts-column {
                flex: 1 1 100%;
                width: 100%;
            }
            
            .profile-img {
                margin: 0 auto;
                display: block;
            }
        }
    </style>
    """

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
    
    # Convert to string and check for empty/invalid values
    value_str = str(value)
    if value_str.upper() in ['UNKNOWN TYPE:', 'N/A', 'NONE', '']:
        return ""
    
    return value_str

def main():
    """Main function"""
    try:
        # Check permissions
        if not check_permissions():
            return
            
        # Get organization ID if running from organization context
        org_id = getattr(model.Data, 'CurrentOrgId', None)
        
        # Configuration for pagination
        count_loop = 94
        first_page = Config.ENTRIES_PER_PAGE + count_loop
        
        # Print styles
        print_styles()
        
        # Print organization header if applicable
        if org_id:
            sql_header = """
            SELECT TOP 1 os.Organization, os.Program, os.Division 
            FROM OrganizationStructure os 
            WHERE OrgId = {0}
            """.format(org_id)
            
            header_data = q.QuerySqlTop1(sql_header)
            if header_data:
                print '<h2>{0}</h2>'.format(header_data.Organization)
                first_page = first_page - 1
        
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
        person_count = 0
        
        print '<div class="emergency-list-container">'
        
        # Process each person
        for person in people_data:
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
                
                # Allergies from RecReg
                allergies = format_medical_value(person.MedAllergy)
                if allergies:
                    print '<div class="medical-field"><span class="medical-label">Allergies:</span> <span class="medical-value">{0}</span></div>'.format(allergies)
                
                # Medical Description from RecReg
                med_desc = format_medical_value(person.MedicalDescription)
                if med_desc:
                    print '<div class="medical-notes"><span class="medical-label" style="color: #92400e;">Medical Notes:</span> {0}</div>'.format(med_desc)
                
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
                allowed_meds = []
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
        
        # Print summary
        if person_count > 0:
            print """
            <div style="margin-top: 20px; padding: 10px; background-color: #f7fafc; border-radius: 6px;">
                <strong>Total People:</strong> {0}
            </div>
            """.format(person_count)
        else:
            print """
            <div class="alert alert-info">
                <i class="fa fa-info-circle"></i> No people found in the current selection.
            </div>
            """
        
        # Blue Toolbar report call - required for Blue Toolbar integration
        sql_bluetoolbar = """
        SELECT p.PeopleId 
        FROM dbo.People p 
        JOIN dbo.TagPerson tp ON tp.PeopleId = p.PeopleId 
        WHERE tp.Id = @BlueToolbarTagId
        """
        
        # This satisfies the Blue Toolbar requirement
        q.BlueToolbarReport("Enhanced Emergency List", sql_bluetoolbar)
            
    except Exception as e:
        print """
        <div class="alert alert-danger">
            <h4><i class="fa fa-exclamation-circle"></i> Error</h4>
            <p>{0}</p>
        </div>
        """.format(str(e))

# Execute main function
main()
