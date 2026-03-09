#roles=Edit

"""
TPxi Report Writer
==================================
A visual builder for creating polished, per-person registration reports.
Staff can select an involvement, configure report layout with drag-and-drop
sections and fields, preview the output, and print beautiful forms.

Templates are saved per-org so they persist across sessions.

Features:
- 4-step wizard: Select Org > Configure Layout > Preview > Generate/Print
- Standard templates: Basic Registration, Detailed Form (WEE-style), Custom
- Drag-sortable sections and fields
- Two-column and single-column layouts via CSS Grid
- Live preview with person selector
- Print-optimized CSS with per-person page breaks
- Saves template to org extra value for reuse
- Handles both new (RegQuestion/RegAnswer) and old (RegistrationData XML) formats

Written By: Ben Swaby
Email: bswaby@fbchtn.org
Version: 1.0
Date: February 2026

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python "ReportWriter" and paste all this code
4. Access via /PyScript/ReportWriter
--Upload Instructions End--
"""

import json
import re

model.Header = 'Registration Report Builder'

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def html_escape(text):
    """Escape HTML special characters"""
    if not text:
        return ''
    s = str(text)
    s = s.replace('&', '&amp;')
    s = s.replace('<', '&lt;')
    s = s.replace('>', '&gt;')
    s = s.replace('"', '&quot;')
    s = s.replace("'", '&#39;')
    return s

def safe_int(val, default=0):
    """Safely convert to int"""
    try:
        return int(val)
    except:
        return default

def parse_filter_people(val):
    """Parse comma-separated people IDs from filter parameter"""
    if not val:
        return None
    try:
        pids = [int(p.strip()) for p in str(val).split(',') if p.strip()]
        return pids if pids else None
    except:
        return None

def fmt_date(dt):
    """Format a date value"""
    if not dt:
        return ''
    try:
        s = str(dt)
        if len(s) >= 10:
            return s[:10]
        return s
    except:
        return ''

def fmt_phone(phone):
    """Format phone number"""
    if not phone:
        return ''
    try:
        return model.FmtPhone(phone)
    except:
        return str(phone)

# ============================================================================
# STANDARD TEMPLATE DEFINITIONS
# ============================================================================

BASIC_TEMPLATE = {
    "templateName": "Basic Registration",
    "baseTemplate": "basic",
    "printSettings": {
        "onePersonPerPage": True,
        "showPageNumbers": True,
        "showOrgHeader": True
    },
    "globalOptions": {
        "hideEmptyFields": True,
        "hideUnansweredQuestions": True,
        "showPersonPhoto": False,
        "headingColor": "#2c5282",
        "fontFamily": "Segoe UI, Tahoma, Geneva, Verdana, sans-serif"
    },
    "sections": [
        {
            "sectionId": "sec_1",
            "title": "Registrant Information",
            "order": 1,
            "layout": "two-column",
            "headerColor": "#2c5282",
            "visible": True,
            "fields": [
                {"fieldId": "fld_1", "fieldType": "person", "sourceField": "Name", "label": "Name", "displayFormat": "single-line", "order": 1, "visible": True, "colSpan": 1},
                {"fieldId": "fld_2", "fieldType": "person", "sourceField": "EmailAddress", "label": "Email", "displayFormat": "single-line", "order": 2, "visible": True, "colSpan": 1},
                {"fieldId": "fld_3", "fieldType": "person", "sourceField": "CellPhone", "label": "Cell Phone", "displayFormat": "single-line", "order": 3, "visible": True, "colSpan": 1},
                {"fieldId": "fld_4", "fieldType": "person", "sourceField": "Age", "label": "Age", "displayFormat": "single-line", "order": 4, "visible": True, "colSpan": 1}
            ]
        },
        {
            "sectionId": "sec_2",
            "title": "Registration Questions",
            "order": 2,
            "layout": "single-column",
            "headerColor": "#2c5282",
            "visible": True,
            "fields": []
        }
    ]
}

DETAILED_TEMPLATE = {
    "templateName": "Detailed Form",
    "baseTemplate": "detailed_form",
    "printSettings": {
        "onePersonPerPage": True,
        "showPageNumbers": True,
        "showOrgHeader": True
    },
    "globalOptions": {
        "hideEmptyFields": True,
        "hideUnansweredQuestions": True,
        "showPersonPhoto": True,
        "headingColor": "#2c5282",
        "fontFamily": "Segoe UI, Tahoma, Geneva, Verdana, sans-serif"
    },
    "sections": [
        {
            "sectionId": "sec_1",
            "title": "Personal Information",
            "order": 1,
            "layout": "two-column",
            "headerColor": "#2c5282",
            "visible": True,
            "fields": [
                {"fieldId": "fld_1", "fieldType": "person", "sourceField": "Name", "label": "Name", "displayFormat": "single-line", "order": 1, "visible": True, "colSpan": 1},
                {"fieldId": "fld_2", "fieldType": "person", "sourceField": "BDate", "label": "Date of Birth", "displayFormat": "single-line", "order": 2, "visible": True, "colSpan": 1},
                {"fieldId": "fld_3", "fieldType": "person", "sourceField": "Age", "label": "Age", "displayFormat": "single-line", "order": 3, "visible": True, "colSpan": 1},
                {"fieldId": "fld_4", "fieldType": "person", "sourceField": "Gender", "label": "Gender", "displayFormat": "single-line", "order": 4, "visible": True, "colSpan": 1}
            ]
        },
        {
            "sectionId": "sec_2",
            "title": "Contact Information",
            "order": 2,
            "layout": "two-column",
            "headerColor": "#2c5282",
            "visible": True,
            "fields": [
                {"fieldId": "fld_5", "fieldType": "family", "sourceField": "Parents", "label": "Parent(s)/Guardian(s)", "displayFormat": "full-width", "order": 1, "visible": True, "colSpan": 2},
                {"fieldId": "fld_6", "fieldType": "person", "sourceField": "PrimaryAddress", "label": "Address", "displayFormat": "full-width", "order": 2, "visible": True, "colSpan": 2},
                {"fieldId": "fld_7", "fieldType": "person", "sourceField": "EmailAddress", "label": "Email", "displayFormat": "single-line", "order": 3, "visible": True, "colSpan": 1},
                {"fieldId": "fld_8", "fieldType": "person", "sourceField": "CellPhone", "label": "Cell Phone", "displayFormat": "single-line", "order": 4, "visible": True, "colSpan": 1},
                {"fieldId": "fld_9", "fieldType": "person", "sourceField": "HomePhone", "label": "Home Phone", "displayFormat": "single-line", "order": 5, "visible": True, "colSpan": 1}
            ]
        },
        {
            "sectionId": "sec_3",
            "title": "Emergency Contact",
            "order": 3,
            "layout": "two-column",
            "headerColor": "#c53030",
            "visible": True,
            "fields": [
                {"fieldId": "fld_10", "fieldType": "medical", "sourceField": "emcontact", "label": "Emergency Contact", "displayFormat": "single-line", "order": 1, "visible": True, "colSpan": 1},
                {"fieldId": "fld_11", "fieldType": "medical", "sourceField": "emphone", "label": "Emergency Phone", "displayFormat": "single-line", "order": 2, "visible": True, "colSpan": 1},
                {"fieldId": "fld_12", "fieldType": "medical", "sourceField": "doctor", "label": "Doctor", "displayFormat": "single-line", "order": 3, "visible": True, "colSpan": 1},
                {"fieldId": "fld_13", "fieldType": "medical", "sourceField": "docphone", "label": "Doctor Phone", "displayFormat": "single-line", "order": 4, "visible": True, "colSpan": 1},
                {"fieldId": "fld_14", "fieldType": "medical", "sourceField": "insurance", "label": "Insurance", "displayFormat": "single-line", "order": 5, "visible": True, "colSpan": 1},
                {"fieldId": "fld_15", "fieldType": "medical", "sourceField": "policy", "label": "Policy #", "displayFormat": "single-line", "order": 6, "visible": True, "colSpan": 1}
            ]
        },
        {
            "sectionId": "sec_4",
            "title": "Medical Information",
            "order": 4,
            "layout": "single-column",
            "headerColor": "#c53030",
            "visible": True,
            "fields": [
                {"fieldId": "fld_16", "fieldType": "medical", "sourceField": "MedAllergy", "label": "Allergies", "displayFormat": "block", "order": 1, "visible": True, "colSpan": 1},
                {"fieldId": "fld_17", "fieldType": "medical", "sourceField": "MedicalDescription", "label": "Medical Conditions", "displayFormat": "block", "order": 2, "visible": True, "colSpan": 1}
            ]
        },
        {
            "sectionId": "sec_5",
            "title": "Registration Answers",
            "order": 5,
            "layout": "single-column",
            "headerColor": "#2c5282",
            "visible": True,
            "fields": []
        }
    ]
}

CUSTOM_TEMPLATE = {
    "templateName": "Custom",
    "baseTemplate": "custom",
    "printSettings": {
        "onePersonPerPage": True,
        "showPageNumbers": True,
        "showOrgHeader": True
    },
    "globalOptions": {
        "hideEmptyFields": False,
        "hideUnansweredQuestions": False,
        "showPersonPhoto": False,
        "headingColor": "#2c5282",
        "fontFamily": "Segoe UI, Tahoma, Geneva, Verdana, sans-serif"
    },
    "sections": [
        {
            "sectionId": "sec_1",
            "title": "Section 1",
            "order": 1,
            "layout": "single-column",
            "headerColor": "#2c5282",
            "visible": True,
            "fields": []
        }
    ]
}

# ============================================================================
# DATA QUERY FUNCTIONS
# ============================================================================

def get_registrant_data(org_id, filter_people_ids=None):
    """Get all registrant data for an org in batch queries.
    If filter_people_ids is set, only include those people (Blue Toolbar filter)."""
    result = {'people': [], 'questions': [], 'familyData': {}}

    filter_clause = ""
    if filter_people_ids:
        filter_clause = " AND p.PeopleId IN ({0})".format(','.join(str(int(pid)) for pid in filter_people_ids))

    # Query 1: Get all org members with person + medical data
    people_sql = """
        SELECT DISTINCT
            p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
            p.BDate, p.Age, p.GenderId, p.EmailAddress, p.CellPhone, p.HomePhone,
            p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
            p.FamilyId,
            rr.MedicalDescription, rr.MedAllergy, rr.emcontact, rr.emphone,
            rr.doctor, rr.docphone, rr.insurance, rr.policy,
            pic.SmallUrl as PhotoUrl
        FROM OrganizationMembers om
        JOIN People p ON om.PeopleId = p.PeopleId
        LEFT JOIN RecReg rr ON rr.PeopleId = p.PeopleId
        LEFT JOIN Picture pic ON pic.PictureId = p.PictureId
        WHERE om.OrganizationId = {0}
        {1}
        ORDER BY p.Name2
    """.format(int(org_id), filter_clause)

    people_list = []
    people_ids = []
    family_ids = set()

    try:
        for r in q.QuerySql(people_sql):
            person = {
                'PeopleId': r.PeopleId,
                'Name': r.Name2 or '',
                'FirstName': r.FirstName or '',
                'LastName': r.LastName or '',
                'NickName': r.NickName or '',
                'BDate': fmt_date(r.BDate),
                'Age': str(r.Age) if r.Age else '',
                'Gender': 'Male' if r.GenderId == 1 else 'Female' if r.GenderId == 2 else '',
                'EmailAddress': r.EmailAddress or '',
                'CellPhone': fmt_phone(r.CellPhone),
                'HomePhone': fmt_phone(r.HomePhone),
                'PrimaryAddress': r.PrimaryAddress or '',
                'PrimaryCity': r.PrimaryCity or '',
                'PrimaryState': r.PrimaryState or '',
                'PrimaryZip': r.PrimaryZip or '',
                'FullAddress': '',
                'FamilyId': r.FamilyId,
                'MedicalDescription': r.MedicalDescription or '',
                'MedAllergy': r.MedAllergy or '',
                'emcontact': r.emcontact or '',
                'emphone': fmt_phone(r.emphone),
                'doctor': r.doctor or '',
                'docphone': fmt_phone(r.docphone),
                'insurance': r.insurance or '',
                'policy': r.policy or '',
                'PhotoUrl': r.PhotoUrl or '',
                'answers': {}
            }
            # Build full address
            addr_parts = []
            if r.PrimaryAddress:
                addr_parts.append(r.PrimaryAddress)
            city_state = []
            if r.PrimaryCity:
                city_state.append(r.PrimaryCity)
            if r.PrimaryState:
                city_state.append(r.PrimaryState)
            if city_state:
                addr_parts.append(', '.join(city_state))
            if r.PrimaryZip:
                addr_parts.append(r.PrimaryZip)
            person['FullAddress'] = ', '.join(addr_parts) if addr_parts else ''

            people_list.append(person)
            people_ids.append(r.PeopleId)
            if r.FamilyId:
                family_ids.add(r.FamilyId)
    except Exception as e:
        result['error'] = 'Error loading people: ' + str(e)
        return result

    if not people_ids:
        result['people'] = []
        return result

    # Build a lookup map
    people_map = {}
    for p in people_list:
        people_map[p['PeopleId']] = p

    # Query 2: Registration answers (new system)
    questions = []
    question_set = set()

    answers_sql = """
        SELECT
            rp.PeopleId,
            rq.Label as Question,
            rq.[Order] as QuestionOrder,
            ra.AnswerValue as Answer
        FROM Registration r WITH (NOLOCK)
        JOIN RegPeople rp WITH (NOLOCK) ON r.RegistrationId = rp.RegistrationId
        LEFT JOIN RegAnswer ra WITH (NOLOCK) ON rp.RegPeopleId = ra.RegPeopleId
        LEFT JOIN RegQuestion rq WITH (NOLOCK) ON ra.RegQuestionId = rq.RegQuestionId
        WHERE r.OrganizationId = {0}
          AND rp.PeopleId IN ({1})
          AND rq.Label IS NOT NULL
        ORDER BY rq.[Order]
    """.format(int(org_id), ','.join(str(pid) for pid in people_ids))

    try:
        for r in q.QuerySql(answers_sql):
            if r.Question and r.Question not in question_set:
                questions.append(r.Question)
                question_set.add(r.Question)
            if r.PeopleId in people_map and r.Question:
                people_map[r.PeopleId]['answers'][r.Question] = r.Answer or ''
    except:
        pass

    # Query 2b: Fallback to RegistrationData XML for older registrations
    rd_sql = """
        SELECT CAST(rd.Data AS NVARCHAR(MAX)) as XmlData
        FROM RegistrationData rd WITH (NOLOCK)
        WHERE rd.OrganizationId = {0}
          AND rd.completed = 1
        ORDER BY rd.Stamp DESC
    """.format(int(org_id))

    try:
        for row in q.QuerySql(rd_sql):
            if not row.XmlData:
                continue
            xml_data = row.XmlData
            for p in people_list:
                pid = p['PeopleId']
                pid_tag = '<PeopleId>{0}</PeopleId>'.format(pid)
                if pid_tag not in xml_data:
                    continue
                person_pattern = r'<OnlineRegPersonModel[^>]*>.*?<PeopleId>{0}</PeopleId>.*?</OnlineRegPersonModel>'.format(pid)
                person_match = re.search(person_pattern, xml_data, re.DOTALL)
                if not person_match:
                    continue
                person_xml = person_match.group(0)

                # ExtraQuestion
                for match in re.finditer(r'<ExtraQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</ExtraQuestion>', person_xml):
                    qtext, ans = match.group(1), match.group(2).strip()
                    if qtext and qtext not in p['answers']:
                        p['answers'][qtext] = ans
                        if qtext not in question_set:
                            questions.append(qtext)
                            question_set.add(qtext)
                # Text
                for match in re.finditer(r'<Text[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</Text>', person_xml):
                    qtext, ans = match.group(1), match.group(2).strip()
                    if qtext and qtext not in p['answers']:
                        p['answers'][qtext] = ans
                        if qtext not in question_set:
                            questions.append(qtext)
                            question_set.add(qtext)
                # YesNoQuestion
                for match in re.finditer(r'<YesNoQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</YesNoQuestion>', person_xml):
                    qtext, ans_val = match.group(1), match.group(2).strip()
                    if qtext and qtext not in p['answers']:
                        ans = 'Yes' if ans_val == 'True' else 'No' if ans_val == 'False' else ans_val
                        p['answers'][qtext] = ans
                        if qtext not in question_set:
                            questions.append(qtext)
                            question_set.add(qtext)
    except:
        pass

    # Query 3: Family/parent data
    family_data = {}
    if family_ids:
        fam_sql = """
            SELECT p.FamilyId, p.FirstName, p.LastName, p.PositionInFamilyId,
                   p.CellPhone, p.EmailAddress
            FROM People p
            WHERE p.FamilyId IN ({0})
              AND p.PositionInFamilyId IN (10, 20)
            ORDER BY p.FamilyId, p.PositionInFamilyId, p.Name2
        """.format(','.join(str(fid) for fid in family_ids))

        try:
            for r in q.QuerySql(fam_sql):
                fid = r.FamilyId
                if fid not in family_data:
                    family_data[fid] = []
                family_data[fid].append({
                    'name': (r.FirstName or '') + ' ' + (r.LastName or ''),
                    'phone': fmt_phone(r.CellPhone),
                    'email': r.EmailAddress or ''
                })
        except:
            pass

    # Attach parent data to each person
    for p in people_list:
        fid = p.get('FamilyId')
        if fid and fid in family_data:
            parents = family_data[fid]
            parent_names = []
            parent_phones = []
            parent_emails = []
            for par in parents:
                pname = par['name'].strip()
                if pname != (p['FirstName'] + ' ' + p['LastName']).strip():
                    parent_names.append(pname)
                    if par['phone']:
                        parent_phones.append(pname + ': ' + par['phone'])
                    if par['email']:
                        parent_emails.append(par['email'])
            p['Parents'] = ', '.join(parent_names) if parent_names else ''
            p['ParentPhones'] = '; '.join(parent_phones) if parent_phones else ''
            p['ParentEmails'] = ', '.join(parent_emails) if parent_emails else ''
        else:
            p['Parents'] = ''
            p['ParentPhones'] = ''
            p['ParentEmails'] = ''

    result['people'] = people_list
    result['questions'] = questions
    result['familyData'] = family_data
    return result

def get_people_data_direct(people_ids):
    """Get person + family + medical data by people IDs (no org required).
    Used when Blue Toolbar is run outside an involvement context."""
    result = {'people': [], 'questions': [], 'familyData': {}}

    if not people_ids:
        return result

    pid_list = ','.join(str(int(pid)) for pid in people_ids)

    people_sql = """
        SELECT DISTINCT
            p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
            p.BDate, p.Age, p.GenderId, p.EmailAddress, p.CellPhone, p.HomePhone,
            p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
            p.FamilyId,
            rr.MedicalDescription, rr.MedAllergy, rr.emcontact, rr.emphone,
            rr.doctor, rr.docphone, rr.insurance, rr.policy,
            pic.SmallUrl as PhotoUrl
        FROM People p
        LEFT JOIN RecReg rr ON rr.PeopleId = p.PeopleId
        LEFT JOIN Picture pic ON pic.PictureId = p.PictureId
        WHERE p.PeopleId IN ({0})
        ORDER BY p.Name2
    """.format(pid_list)

    people_list = []
    family_ids = set()

    try:
        for r in q.QuerySql(people_sql):
            person = {
                'PeopleId': r.PeopleId,
                'Name': r.Name2 or '',
                'FirstName': r.FirstName or '',
                'LastName': r.LastName or '',
                'NickName': r.NickName or '',
                'BDate': fmt_date(r.BDate),
                'Age': str(r.Age) if r.Age else '',
                'Gender': 'Male' if r.GenderId == 1 else 'Female' if r.GenderId == 2 else '',
                'EmailAddress': r.EmailAddress or '',
                'CellPhone': fmt_phone(r.CellPhone),
                'HomePhone': fmt_phone(r.HomePhone),
                'PrimaryAddress': r.PrimaryAddress or '',
                'PrimaryCity': r.PrimaryCity or '',
                'PrimaryState': r.PrimaryState or '',
                'PrimaryZip': r.PrimaryZip or '',
                'FullAddress': '',
                'FamilyId': r.FamilyId,
                'MedicalDescription': r.MedicalDescription or '',
                'MedAllergy': r.MedAllergy or '',
                'emcontact': r.emcontact or '',
                'emphone': fmt_phone(r.emphone),
                'doctor': r.doctor or '',
                'docphone': fmt_phone(r.docphone),
                'insurance': r.insurance or '',
                'policy': r.policy or '',
                'PhotoUrl': r.PhotoUrl or '',
                'answers': {}
            }
            addr_parts = []
            if r.PrimaryAddress:
                addr_parts.append(r.PrimaryAddress)
            city_state = []
            if r.PrimaryCity:
                city_state.append(r.PrimaryCity)
            if r.PrimaryState:
                city_state.append(r.PrimaryState)
            if city_state:
                addr_parts.append(', '.join(city_state))
            if r.PrimaryZip:
                addr_parts.append(r.PrimaryZip)
            person['FullAddress'] = ', '.join(addr_parts) if addr_parts else ''

            people_list.append(person)
            if r.FamilyId:
                family_ids.add(r.FamilyId)
    except Exception as e:
        result['error'] = 'Error loading people: ' + str(e)
        return result

    # Family/parent data
    family_data = {}
    if family_ids:
        fam_sql = """
            SELECT p.FamilyId, p.FirstName, p.LastName, p.PositionInFamilyId,
                   p.CellPhone, p.EmailAddress
            FROM People p
            WHERE p.FamilyId IN ({0})
              AND p.PositionInFamilyId IN (10, 20)
            ORDER BY p.FamilyId, p.PositionInFamilyId, p.Name2
        """.format(','.join(str(fid) for fid in family_ids))
        try:
            for r in q.QuerySql(fam_sql):
                fid = r.FamilyId
                if fid not in family_data:
                    family_data[fid] = []
                family_data[fid].append({
                    'name': (r.FirstName or '') + ' ' + (r.LastName or ''),
                    'phone': fmt_phone(r.CellPhone),
                    'email': r.EmailAddress or ''
                })
        except:
            pass

    for p in people_list:
        fid = p.get('FamilyId')
        if fid and fid in family_data:
            parents = family_data[fid]
            parent_names = []
            parent_phones = []
            parent_emails = []
            for par in parents:
                pname = par['name'].strip()
                if pname != (p['FirstName'] + ' ' + p['LastName']).strip():
                    parent_names.append(pname)
                    if par['phone']:
                        parent_phones.append(pname + ': ' + par['phone'])
                    if par['email']:
                        parent_emails.append(par['email'])
            p['Parents'] = ', '.join(parent_names) if parent_names else ''
            p['ParentPhones'] = '; '.join(parent_phones) if parent_phones else ''
            p['ParentEmails'] = ', '.join(parent_emails) if parent_emails else ''
        else:
            p['Parents'] = ''
            p['ParentPhones'] = ''
            p['ParentEmails'] = ''

    result['people'] = people_list
    result['familyData'] = family_data
    return result

# ============================================================================
# REPORT GENERATION ENGINE
# ============================================================================

def get_field_value(person, field):
    """Get the value of a field from person data"""
    ft = field.get('fieldType', '')
    src = field.get('sourceField', '')

    if ft == 'static':
        if src == 'separator':
            return '---'
        return field.get('staticContent', '')

    if ft == 'person':
        if src == 'PrimaryAddress' or src == 'FullAddress':
            return html_escape(person.get('FullAddress', ''))
        return html_escape(person.get(src, ''))

    elif ft == 'family':
        return html_escape(person.get(src, ''))

    elif ft == 'medical':
        return html_escape(person.get(src, ''))

    elif ft == 'regquestion':
        return html_escape(person.get('answers', {}).get(src, ''))

    return ''

def render_report_html(people, template, org_name, questions, single_person_id=None):
    """Render the full report HTML from template + data"""
    opts = template.get('globalOptions', {})
    ps = template.get('printSettings', {})
    sections = template.get('sections', [])
    hide_empty = opts.get('hideEmptyFields', True)
    hide_unanswered = opts.get('hideUnansweredQuestions', True)
    show_photo = opts.get('showPersonPhoto', False)
    heading_color = opts.get('headingColor', '#2c5282')
    font_family = opts.get('fontFamily', 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif')
    one_per_page = ps.get('onePersonPerPage', True)
    show_org_header = ps.get('showOrgHeader', True)
    compact_rows = opts.get('compactRows', False)

    sorted_sections = sorted(sections, key=lambda s: s.get('order', 0))

    if single_person_id:
        pid = safe_int(single_person_id)
        people = [p for p in people if p['PeopleId'] == pid]

    parts = []
    report_class = 'rr-report' + (' rr-compact' if compact_rows else '')
    parts.append('<div class="{0}" style="font-family: {1};">'.format(report_class, html_escape(font_family)))

    for idx, person in enumerate(people):
        page_class = 'rr-person-page'
        if one_per_page and idx < len(people) - 1:
            page_class += ' rr-page-break'

        parts.append('<div class="{0}">'.format(page_class))

        if show_org_header and org_name:
            parts.append('<div class="rr-org-header" style="border-bottom-color: {0};">'.format(html_escape(heading_color)))
            parts.append('<h2 style="color: {0};">{1}</h2>'.format(html_escape(heading_color), html_escape(org_name)))
            parts.append('</div>')

        parts.append('<div class="rr-person-header">')
        if show_photo and person.get('PhotoUrl'):
            parts.append('<img class="rr-photo" src="{0}" alt="Photo" onerror="this.style.display=\'none\'">'.format(html_escape(person['PhotoUrl'])))
        parts.append('<div class="rr-person-name" style="color: {0};">{1}</div>'.format(
            html_escape(heading_color), html_escape(person.get('Name', ''))))
        parts.append('</div>')

        for sec in sorted_sections:
            if not sec.get('visible', True):
                continue

            sec_header_color = sec.get('headerColor', heading_color)
            layout = sec.get('layout', 'single-column')
            fields = sorted(sec.get('fields', []), key=lambda f: f.get('order', 0))

            # Auto-populate reg question sections when fields list is empty
            if not fields:
                title_lower = (sec.get('title', '') or '').lower()
                if 'question' in title_lower or 'answer' in title_lower or 'registration' in title_lower:
                    auto_fields = []
                    for qi, qtext in enumerate(questions):
                        auto_fields.append({
                            'fieldId': 'auto_q_' + str(qi),
                            'fieldType': 'regquestion',
                            'sourceField': qtext,
                            'label': qtext,
                            'displayFormat': 'block',
                            'order': qi + 1,
                            'visible': True,
                            'colSpan': 1
                        })
                    fields = auto_fields

            if not fields:
                continue

            # Check if section has any visible values
            if hide_empty:
                has_values = False
                for f in fields:
                    if not f.get('visible', True):
                        continue
                    val = get_field_value(person, f)
                    if val and val.strip():
                        has_values = True
                        break
                if not has_values:
                    continue

            parts.append('<div class="rr-section">')
            parts.append('<div class="rr-section-header" style="background-color: {0};">{1}</div>'.format(
                html_escape(sec_header_color), html_escape(sec.get('title', ''))))

            if layout == 'two-column':
                parts.append('<div class="rr-fields-grid rr-two-col">')
            else:
                parts.append('<div class="rr-fields-grid rr-one-col">')

            for f in fields:
                if not f.get('visible', True):
                    continue

                # Special handling for static fields (custom text, separators)
                if f.get('fieldType') == 'static':
                    if f.get('sourceField') == 'separator':
                        parts.append('<div class="rr-field rr-full-width"><hr style="border:none;border-top:1px solid #cbd5e0;margin:8px 0;"></div>')
                    else:
                        parts.append('<div class="rr-field rr-field-block rr-full-width">')
                        label_text = f.get('label', '')
                        if label_text:
                            parts.append('<div class="rr-field-label" style="font-size:13px;font-weight:700;margin-bottom:4px;">{0}</div>'.format(html_escape(label_text)))
                        content = html_escape(f.get('staticContent', ''))
                        if content:
                            parts.append('<div class="rr-field-value rr-block-value" style="font-style:italic;white-space:pre-wrap;">{0}</div>'.format(content))
                        parts.append('</div>')
                    continue

                val = get_field_value(person, f)
                display_fmt = f.get('displayFormat', 'single-line')
                col_span = f.get('colSpan', 1)

                if (not val or not val.strip()):
                    if f['fieldType'] == 'regquestion' and hide_unanswered:
                        continue
                    elif f['fieldType'] != 'regquestion' and hide_empty:
                        continue

                span_class = ''
                if col_span == 2 or display_fmt == 'full-width':
                    span_class = ' rr-full-width'

                if display_fmt == 'block':
                    parts.append('<div class="rr-field rr-field-block{0}">'.format(span_class))
                    parts.append('<div class="rr-field-label">{0}</div>'.format(html_escape(f.get('label', ''))))
                    parts.append('<div class="rr-field-value rr-block-value">{0}</div>'.format(val or '&nbsp;'))
                    parts.append('</div>')
                else:
                    parts.append('<div class="rr-field{0}">'.format(span_class))
                    parts.append('<span class="rr-field-label">{0}:</span> '.format(html_escape(f.get('label', ''))))
                    parts.append('<span class="rr-field-value">{0}</span>'.format(val or '&mdash;'))
                    parts.append('</div>')

            parts.append('</div>')
            parts.append('</div>')

        parts.append('</div>')

    parts.append('</div>')
    return ''.join(parts)


# ============================================================================
# AJAX HANDLER
# ============================================================================
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -------------------------------------------------------------------------
    # Search Orgs with Registration Questions
    # -------------------------------------------------------------------------
    if action == 'search_orgs':
        search_term = getattr(Data, 'search_term', '')
        program_id = getattr(Data, 'program_id', '')
        division_id = getattr(Data, 'division_id', '')

        where_clauses = ["o.OrganizationStatusId = 30"]

        if search_term:
            safe_term = str(search_term).replace("'", "''")
            where_clauses.append("""(o.OrganizationName LIKE '%{0}%'
                OR CAST(o.OrganizationId AS VARCHAR) = '{0}')""".format(safe_term))

        if program_id:
            where_clauses.append("d.ProgId = {0}".format(safe_int(program_id)))

        if division_id:
            where_clauses.append("o.DivisionId = {0}".format(safe_int(division_id)))

        sql = """
            SELECT TOP 50
                o.OrganizationId,
                o.OrganizationName,
                d.Name as DivisionName,
                p.Name as ProgramName,
                (SELECT COUNT(*) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId) as MemberCount,
                (SELECT COUNT(*) FROM RegQuestion rq WHERE rq.OrganizationId = o.OrganizationId) as QuestionCount
            FROM Organizations o
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program p ON d.ProgId = p.Id
            WHERE {0}
              AND (EXISTS (SELECT 1 FROM RegQuestion rq WHERE rq.OrganizationId = o.OrganizationId)
                   OR EXISTS (SELECT 1 FROM RegistrationData rd WHERE rd.OrganizationId = o.OrganizationId AND rd.completed = 1))
            ORDER BY o.OrganizationName
        """.format(" AND ".join(where_clauses))

        try:
            results = q.QuerySql(sql)
            orgs = []
            for r in results:
                orgs.append({
                    'id': r.OrganizationId,
                    'name': r.OrganizationName,
                    'division': r.DivisionName or '',
                    'program': r.ProgramName or '',
                    'memberCount': r.MemberCount or 0,
                    'questionCount': r.QuestionCount or 0
                })
            print json.dumps({'success': True, 'orgs': orgs})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Get Filter Dropdowns (Programs/Divisions)
    # -------------------------------------------------------------------------
    elif action == 'get_filters':
        try:
            prog_sql = """
                SELECT DISTINCT p.Id, p.Name
                FROM Program p
                JOIN Division d ON d.ProgId = p.Id
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE o.OrganizationStatusId = 30
                ORDER BY p.Name
            """
            programs = []
            for r in q.QuerySql(prog_sql):
                programs.append({'id': r.Id, 'name': r.Name})

            div_sql = """
                SELECT DISTINCT d.Id, d.Name, d.ProgId
                FROM Division d
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE o.OrganizationStatusId = 30
                ORDER BY d.Name
            """
            divisions = []
            for r in q.QuerySql(div_sql):
                divisions.append({'id': r.Id, 'name': r.Name, 'programId': r.ProgId})

            print json.dumps({'success': True, 'programs': programs, 'divisions': divisions})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Load Org Data (questions + people + person field list)
    # -------------------------------------------------------------------------
    elif action == 'load_org_data':
        org_id = getattr(Data, 'org_id', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))
        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                org_info = q.QuerySqlTop1("SELECT OrganizationName FROM Organizations WHERE OrganizationId = {0}".format(org_id))
                org_name = org_info.OrganizationName if org_info else 'Unknown'

                data = get_registrant_data(org_id, filter_ids)
                if 'error' in data:
                    print json.dumps({'success': False, 'message': data['error']})
                else:
                    person_list = []
                    for p in data['people']:
                        person_list.append({
                            'id': p['PeopleId'],
                            'name': p['Name']
                        })

                    person_fields = [
                        {'sourceField': 'Name', 'label': 'Full Name', 'fieldType': 'person'},
                        {'sourceField': 'FirstName', 'label': 'First Name', 'fieldType': 'person'},
                        {'sourceField': 'LastName', 'label': 'Last Name', 'fieldType': 'person'},
                        {'sourceField': 'NickName', 'label': 'Nickname', 'fieldType': 'person'},
                        {'sourceField': 'EmailAddress', 'label': 'Email', 'fieldType': 'person'},
                        {'sourceField': 'CellPhone', 'label': 'Cell Phone', 'fieldType': 'person'},
                        {'sourceField': 'HomePhone', 'label': 'Home Phone', 'fieldType': 'person'},
                        {'sourceField': 'Age', 'label': 'Age', 'fieldType': 'person'},
                        {'sourceField': 'BDate', 'label': 'Date of Birth', 'fieldType': 'person'},
                        {'sourceField': 'Gender', 'label': 'Gender', 'fieldType': 'person'},
                        {'sourceField': 'PrimaryAddress', 'label': 'Full Address', 'fieldType': 'person'}
                    ]

                    family_fields = [
                        {'sourceField': 'Parents', 'label': 'Parent(s)/Guardian(s)', 'fieldType': 'family'},
                        {'sourceField': 'ParentPhones', 'label': 'Parent Phone(s)', 'fieldType': 'family'},
                        {'sourceField': 'ParentEmails', 'label': 'Parent Email(s)', 'fieldType': 'family'}
                    ]

                    medical_fields = [
                        {'sourceField': 'emcontact', 'label': 'Emergency Contact', 'fieldType': 'medical'},
                        {'sourceField': 'emphone', 'label': 'Emergency Phone', 'fieldType': 'medical'},
                        {'sourceField': 'doctor', 'label': 'Doctor', 'fieldType': 'medical'},
                        {'sourceField': 'docphone', 'label': 'Doctor Phone', 'fieldType': 'medical'},
                        {'sourceField': 'insurance', 'label': 'Insurance', 'fieldType': 'medical'},
                        {'sourceField': 'policy', 'label': 'Policy #', 'fieldType': 'medical'},
                        {'sourceField': 'MedAllergy', 'label': 'Allergies', 'fieldType': 'medical'},
                        {'sourceField': 'MedicalDescription', 'label': 'Medical Conditions', 'fieldType': 'medical'}
                    ]

                    question_fields = []
                    for qtext in data['questions']:
                        question_fields.append({
                            'sourceField': qtext,
                            'label': qtext,
                            'fieldType': 'regquestion'
                        })

                    print json.dumps({
                        'success': True,
                        'orgName': org_name,
                        'personCount': len(data['people']),
                        'questionCount': len(data['questions']),
                        'questions': data['questions'],
                        'personList': person_list,
                        'availableFields': {
                            'person': person_fields,
                            'family': family_fields,
                            'medical': medical_fields,
                            'regquestion': question_fields
                        }
                    })
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Load BT People Data (no org - direct people IDs from Blue Toolbar)
    # -------------------------------------------------------------------------
    elif action == 'load_bt_data':
        bt_pids_str = getattr(Data, 'people_ids', '')
        if not bt_pids_str:
            print json.dumps({'success': False, 'message': 'No people IDs provided'})
        else:
            try:
                bt_pids = [int(x) for x in str(bt_pids_str).split(',') if x.strip()]
                data = get_people_data_direct(bt_pids)
                if 'error' in data:
                    print json.dumps({'success': False, 'message': data['error']})
                else:
                    person_list = []
                    for p in data['people']:
                        person_list.append({'id': p['PeopleId'], 'name': p['Name']})

                    person_fields = [
                        {'sourceField': 'Name', 'label': 'Full Name', 'fieldType': 'person'},
                        {'sourceField': 'FirstName', 'label': 'First Name', 'fieldType': 'person'},
                        {'sourceField': 'LastName', 'label': 'Last Name', 'fieldType': 'person'},
                        {'sourceField': 'NickName', 'label': 'Nickname', 'fieldType': 'person'},
                        {'sourceField': 'EmailAddress', 'label': 'Email', 'fieldType': 'person'},
                        {'sourceField': 'CellPhone', 'label': 'Cell Phone', 'fieldType': 'person'},
                        {'sourceField': 'HomePhone', 'label': 'Home Phone', 'fieldType': 'person'},
                        {'sourceField': 'Age', 'label': 'Age', 'fieldType': 'person'},
                        {'sourceField': 'BDate', 'label': 'Date of Birth', 'fieldType': 'person'},
                        {'sourceField': 'Gender', 'label': 'Gender', 'fieldType': 'person'},
                        {'sourceField': 'PrimaryAddress', 'label': 'Full Address', 'fieldType': 'person'}
                    ]
                    family_fields = [
                        {'sourceField': 'Parents', 'label': 'Parent(s)/Guardian(s)', 'fieldType': 'family'},
                        {'sourceField': 'ParentPhones', 'label': 'Parent Phone(s)', 'fieldType': 'family'},
                        {'sourceField': 'ParentEmails', 'label': 'Parent Email(s)', 'fieldType': 'family'}
                    ]
                    medical_fields = [
                        {'sourceField': 'emcontact', 'label': 'Emergency Contact', 'fieldType': 'medical'},
                        {'sourceField': 'emphone', 'label': 'Emergency Phone', 'fieldType': 'medical'},
                        {'sourceField': 'doctor', 'label': 'Doctor', 'fieldType': 'medical'},
                        {'sourceField': 'docphone', 'label': 'Doctor Phone', 'fieldType': 'medical'},
                        {'sourceField': 'insurance', 'label': 'Insurance', 'fieldType': 'medical'},
                        {'sourceField': 'policy', 'label': 'Policy #', 'fieldType': 'medical'},
                        {'sourceField': 'MedAllergy', 'label': 'Allergies', 'fieldType': 'medical'},
                        {'sourceField': 'MedicalDescription', 'label': 'Medical Conditions', 'fieldType': 'medical'}
                    ]
                    print json.dumps({
                        'success': True,
                        'orgName': 'Selected People',
                        'personCount': len(data['people']),
                        'questionCount': 0,
                        'questions': [],
                        'personList': person_list,
                        'availableFields': {
                            'person': person_fields,
                            'family': family_fields,
                            'medical': medical_fields,
                            'regquestion': []
                        }
                    })
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Load Saved Template
    # -------------------------------------------------------------------------
    elif action == 'load_template':
        org_id = getattr(Data, 'org_id', '')
        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                saved = model.ExtraValueTextOrg(org_id, 'RegReportTemplate')
                if saved:
                    template = json.loads(saved)
                    print json.dumps({'success': True, 'template': template, 'hasSaved': True})
                else:
                    print json.dumps({'success': True, 'template': None, 'hasSaved': False})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Save Template
    # -------------------------------------------------------------------------
    elif action == 'save_template':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        if not org_id or not template_json:
            print json.dumps({'success': False, 'message': 'Organization ID and template required'})
        else:
            try:
                org_id = int(org_id)
                parsed = json.loads(template_json)
                model.AddExtraValueTextOrg(org_id, 'RegReportTemplate', template_json)
                print json.dumps({'success': True, 'message': 'Template saved successfully'})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Preview Report (single person)
    # -------------------------------------------------------------------------
    elif action == 'preview_report':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        people_id = getattr(Data, 'people_id', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))

        if not org_id or not template_json:
            print json.dumps({'success': False, 'message': 'Organization and template required'})
        else:
            try:
                template = json.loads(template_json)
                if str(org_id) == 'bt_direct':
                    data = get_people_data_direct(filter_ids)
                    org_name = 'Selected People'
                else:
                    org_id = int(org_id)
                    org_info = q.QuerySqlTop1("SELECT OrganizationName FROM Organizations WHERE OrganizationId = {0}".format(org_id))
                    org_name = org_info.OrganizationName if org_info else ''
                    data = get_registrant_data(org_id, filter_ids)
                html_out = render_report_html(data['people'], template, org_name, data.get('questions', []), people_id)
                print json.dumps({'success': True, 'html': html_out})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Generate Full Report (all registrants)
    # -------------------------------------------------------------------------
    elif action == 'generate_report':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))

        if not org_id or not template_json:
            print json.dumps({'success': False, 'message': 'Organization and template required'})
        else:
            try:
                template = json.loads(template_json)
                if str(org_id) == 'bt_direct':
                    data = get_people_data_direct(filter_ids)
                    org_name = 'Selected People'
                else:
                    org_id = int(org_id)
                    org_info = q.QuerySqlTop1("SELECT OrganizationName FROM Organizations WHERE OrganizationId = {0}".format(org_id))
                    org_name = org_info.OrganizationName if org_info else ''
                    data = get_registrant_data(org_id, filter_ids)
                html_out = render_report_html(data['people'], template, org_name, data.get('questions', []))
                print json.dumps({'success': True, 'html': html_out, 'personCount': len(data['people'])})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Get Person List (for selector dropdown)
    # -------------------------------------------------------------------------
    elif action == 'get_person_list':
        org_id = getattr(Data, 'org_id', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))
        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                plist_filter = ""
                if filter_ids:
                    plist_filter = " AND p.PeopleId IN ({0})".format(','.join(str(int(pid)) for pid in filter_ids))
                sql = """
                    SELECT p.PeopleId, p.Name2
                    FROM OrganizationMembers om
                    JOIN People p ON om.PeopleId = p.PeopleId
                    WHERE om.OrganizationId = {0}
                    {1}
                    ORDER BY p.Name2
                """.format(org_id, plist_filter)
                persons = []
                for r in q.QuerySql(sql):
                    persons.append({'id': r.PeopleId, 'name': r.Name2})
                print json.dumps({'success': True, 'persons': persons})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + action})

# ============================================================================
# MAIN PAGE (GET request - normal mode or redirected from Blue Toolbar)
# ============================================================================
else:
    # Build template JSON for embedding in JavaScript
    tpl_basic_json = json.dumps(BASIC_TEMPLATE)
    tpl_detailed_json = json.dumps(DETAILED_TEMPLATE)
    tpl_custom_json = json.dumps(CUSTOM_TEMPLATE)

    # Blue Toolbar: gather people from @BlueToolbarTagId
    # Only useful in PyScript context (JS stores in sessionStorage then redirects).
    # In PyScriptForm context, JS ignores this data and reads sessionStorage instead.
    _bt_gathered_pids = []
    _bt_gathered_org = 0
    try:
        for r in q.QuerySql("SELECT tp.PeopleId FROM TagPerson tp WHERE tp.Id = @BlueToolbarTagId"):
            _bt_gathered_pids.append(int(r.PeopleId))
        if _bt_gathered_pids:
            try:
                if hasattr(Data, 'CurrentOrgId') and Data.CurrentOrgId:
                    _bt_gathered_org = int(Data.CurrentOrgId)
            except:
                pass
    except:
        pass

    bt_gathered_pids_json = json.dumps(_bt_gathered_pids)
    bt_gathered_org_json = str(_bt_gathered_org)

    html = '''<!DOCTYPE html>
<html>
<head>
<style>
/* ===== RR- Scoped Styles ===== */
.rr-root {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    color: #333;
    max-width: 1400px;
    margin: 0 auto;
}

/* Wizard Steps */
.rr-steps {
    display: flex;
    margin-bottom: 24px;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 0;
}
.rr-step {
    flex: 1;
    text-align: center;
    padding: 12px 8px;
    cursor: pointer;
    color: #a0aec0;
    font-weight: 600;
    border-bottom: 3px solid transparent;
    margin-bottom: -2px;
    transition: all 0.2s;
}
.rr-step.active { color: #2c5282; border-bottom-color: #2c5282; }
.rr-step.completed { color: #38a169; border-bottom-color: #38a169; }
.rr-step .rr-step-num {
    display: inline-block; width: 28px; height: 28px; line-height: 28px;
    border-radius: 50%; background: #e2e8f0; color: #718096; font-size: 14px; margin-right: 6px;
}
.rr-step.active .rr-step-num { background: #2c5282; color: #fff; }
.rr-step.completed .rr-step-num { background: #38a169; color: #fff; }

/* Panels */
.rr-panel { display: none; animation: rrFadeIn 0.3s ease; }
.rr-panel.active { display: block; }
@keyframes rrFadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }

/* Cards */
.rr-card {
    background: #fff; border: 1px solid #e2e8f0; border-radius: 8px;
    padding: 20px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.rr-card-title {
    font-size: 16px; font-weight: 700; color: #2d3748; margin-bottom: 12px;
    padding-bottom: 8px; border-bottom: 1px solid #e2e8f0;
}

/* Search */
.rr-search-box { display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }
.rr-search-box input, .rr-search-box select {
    padding: 8px 12px; border: 1px solid #cbd5e0; border-radius: 6px; font-size: 14px;
}
.rr-search-box input { flex: 1; min-width: 200px; }

/* Buttons */
.rr-btn {
    display: inline-block; padding: 8px 16px; border: none; border-radius: 6px;
    font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.15s;
}
.rr-btn-primary { background: #2c5282; color: #fff; }
.rr-btn-primary:hover { background: #2a4a7f; }
.rr-btn-success { background: #38a169; color: #fff; }
.rr-btn-success:hover { background: #2f855a; }
.rr-btn-secondary { background: #e2e8f0; color: #4a5568; }
.rr-btn-secondary:hover { background: #cbd5e0; }
.rr-btn-danger { background: #e53e3e; color: #fff; }
.rr-btn-sm { padding: 4px 10px; font-size: 12px; }
.rr-btn:disabled { opacity: 0.5; cursor: not-allowed; }

/* Org list */
.rr-org-item {
    display: flex; justify-content: space-between; align-items: center;
    padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 6px;
    margin-bottom: 6px; cursor: pointer; transition: all 0.15s;
}
.rr-org-item:hover { background: #ebf4ff; border-color: #90cdf4; }
.rr-org-item.selected { background: #ebf8ff; border-color: #2c5282; border-width: 2px; }
.rr-org-name { font-weight: 600; color: #2d3748; }
.rr-org-meta { font-size: 12px; color: #718096; }
.rr-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
.rr-badge-blue { background: #ebf8ff; color: #2c5282; }
.rr-badge-green { background: #f0fff4; color: #276749; }

/* Help guide */
.rr-help-toggle {
    display: inline-flex; align-items: center; gap: 6px; padding: 6px 14px;
    background: #ebf8ff; border: 1px solid #90cdf4; border-radius: 6px;
    color: #2c5282; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.15s;
}
.rr-help-toggle:hover { background: #bee3f8; }
.rr-help-toggle .fa { font-size: 14px; }
.rr-help-guide {
    display: none; background: #f7fafc; border: 1px solid #e2e8f0; border-radius: 8px;
    padding: 20px 24px; margin-bottom: 16px; font-size: 13px; color: #4a5568; line-height: 1.6;
}
.rr-help-guide.open { display: block; animation: rrFadeIn 0.3s ease; }
.rr-help-guide h4 {
    margin: 0 0 6px 0; font-size: 14px; color: #2d3748; display: flex; align-items: center; gap: 6px;
}
.rr-help-guide h4 .fa { color: #2c5282; font-size: 13px; }
.rr-help-cols { display: grid; grid-template-columns: 1fr 1fr; gap: 16px 24px; }
.rr-help-section { padding: 10px 14px; background: #fff; border-radius: 6px; border: 1px solid #e2e8f0; }
.rr-help-section ul { margin: 0; padding-left: 16px; }
.rr-help-section li { margin-bottom: 3px; }
.rr-help-section code { background: #edf2f7; padding: 1px 5px; border-radius: 3px; font-size: 12px; }
@media (max-width: 768px) { .rr-help-cols { grid-template-columns: 1fr; } }

/* Config layout */
.rr-config-layout { display: flex; gap: 20px; }
.rr-config-left { flex: 3; }
.rr-config-right { flex: 2; }

/* Section editor */
.rr-section-editor { border: 1px solid #e2e8f0; border-radius: 8px; margin-bottom: 12px; overflow: hidden; }
.rr-section-head {
    display: flex; justify-content: space-between; align-items: center;
    padding: 10px 14px; background: #f7fafc; cursor: pointer; border-bottom: 1px solid #e2e8f0;
}
.rr-section-head:hover { background: #edf2f7; }
.rr-section-body { padding: 12px; display: none; }
.rr-section-body.open { display: block; }
.rr-field-item {
    display: flex; align-items: center; gap: 8px; padding: 6px 8px;
    margin-bottom: 4px; background: #f7fafc; border-radius: 4px; font-size: 13px;
}
.rr-field-item .rr-drag-handle { cursor: grab; color: #a0aec0; user-select: none; }
.rr-field-item .rr-field-label-text { flex: 1; }
.rr-field-item select, .rr-field-item input[type="text"] {
    font-size: 12px; padding: 2px 4px; border: 1px solid #cbd5e0; border-radius: 4px;
}
.rr-add-field-row {
    display: flex; gap: 6px; margin-top: 8px; padding-top: 8px; border-top: 1px solid #e2e8f0;
}
.rr-add-field-row select { flex: 1; font-size: 13px; padding: 6px; border: 1px solid #cbd5e0; border-radius: 4px; }

/* Options */
.rr-option-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 8px 0; border-bottom: 1px solid #f0f0f0;
}
.rr-option-row:last-child { border-bottom: none; }
.rr-option-label { font-size: 13px; color: #4a5568; }

/* Toggle */
.rr-toggle { position: relative; width: 40px; height: 22px; display: inline-block; }
.rr-toggle input { display: none; }
.rr-toggle-slider {
    position: absolute; top: 0; left: 0; right: 0; bottom: 0;
    background: #cbd5e0; border-radius: 22px; cursor: pointer; transition: 0.2s;
}
.rr-toggle-slider:before {
    content: ""; position: absolute; width: 18px; height: 18px;
    left: 2px; bottom: 2px; background: #fff; border-radius: 50%; transition: 0.2s;
}
.rr-toggle input:checked + .rr-toggle-slider { background: #38a169; }
.rr-toggle input:checked + .rr-toggle-slider:before { transform: translateX(18px); }

/* Template cards */
.rr-template-cards { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
.rr-template-card {
    flex: 1; min-width: 140px; padding: 14px; border: 2px solid #e2e8f0; border-radius: 8px;
    cursor: pointer; text-align: center; transition: all 0.15s;
}
.rr-template-card:hover { border-color: #90cdf4; background: #ebf8ff; }
.rr-template-card.selected { border-color: #2c5282; background: #ebf4ff; }
.rr-template-card h4 { margin: 0 0 4px 0; font-size: 14px; color: #2d3748; }
.rr-template-card p { margin: 0; font-size: 12px; color: #718096; }

/* Preview */
.rr-preview-toolbar {
    display: flex; gap: 12px; align-items: center; margin-bottom: 12px;
    padding: 10px; background: #f7fafc; border-radius: 6px; flex-wrap: wrap;
}
.rr-preview-toolbar select { padding: 6px 10px; border: 1px solid #cbd5e0; border-radius: 4px; font-size: 13px; }
.rr-preview-frame {
    border: 1px solid #e2e8f0; border-radius: 8px; padding: 30px;
    background: #fff; min-height: 400px; box-shadow: inset 0 1px 4px rgba(0,0,0,0.05);
}

/* ===== Report Render Styles ===== */
.rr-report { color: #333; }
.rr-person-page { padding: 20px 0; }
.rr-page-break { border-bottom: 2px dashed #cbd5e0; margin-bottom: 20px; padding-bottom: 20px; }
.rr-org-header { text-align: center; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #2c5282; }
.rr-org-header h2 { margin: 0; font-size: 20px; }
.rr-person-header {
    display: flex; align-items: center; gap: 12px; margin-bottom: 16px;
    padding-bottom: 8px; border-bottom: 1px solid #e2e8f0;
}
.rr-photo { width: 60px; height: 60px; border-radius: 6px; object-fit: cover; }
.rr-person-name { font-size: 22px; font-weight: 700; }
.rr-section { margin-bottom: 16px; page-break-inside: avoid; }
.rr-section-header {
    padding: 6px 12px; color: #fff; font-weight: 700; font-size: 14px;
    border-radius: 4px 4px 0 0; margin-bottom: 0;
}
.rr-fields-grid { border: 1px solid #e2e8f0; border-top: none; border-radius: 0 0 4px 4px; padding: 10px 12px; }
.rr-two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 16px; }
.rr-one-col { display: grid; grid-template-columns: 1fr; gap: 6px; }
.rr-full-width { grid-column: 1 / -1; }
.rr-field { padding: 3px 0; }
.rr-field-label { font-weight: 600; color: #4a5568; font-size: 12px; }
.rr-field-value { color: #1a202c; font-size: 13px; }
.rr-field-block { grid-column: 1 / -1; padding: 4px 0; }
.rr-block-value {
    display: block; padding: 4px 8px; background: #f7fafc; border: 1px solid #e2e8f0;
    border-radius: 4px; margin-top: 2px; word-wrap: break-word; white-space: pre-wrap;
    min-height: 24px; font-size: 13px;
}

/* Compact row spacing */
.rr-compact .rr-person-page { padding: 8px 0; }
.rr-compact .rr-person-header { margin-bottom: 6px; padding-bottom: 4px; }
.rr-compact .rr-section { margin-bottom: 6px; }
.rr-compact .rr-section-header { padding: 3px 10px; font-size: 13px; }
.rr-compact .rr-fields-grid { padding: 4px 10px; }
.rr-compact .rr-two-col { gap: 1px 12px; }
.rr-compact .rr-one-col { gap: 1px; }
.rr-compact .rr-field { padding: 1px 0; }
.rr-compact .rr-field-block { padding: 1px 0; }
.rr-compact .rr-block-value { padding: 2px 6px; margin-top: 1px; min-height: 16px; }
.rr-compact .rr-org-header { margin-bottom: 6px; padding-bottom: 4px; }

/* Toast */
.rr-toast {
    position: fixed; top: 20px; right: 20px; padding: 12px 20px; border-radius: 6px;
    color: #fff; font-weight: 600; z-index: 9999; opacity: 0; transition: opacity 0.3s;
}
.rr-toast.show { opacity: 1; }
.rr-toast-success { background: #38a169; }
.rr-toast-danger { background: #e53e3e; }
.rr-toast-info { background: #3182ce; }

/* Loading */
.rr-loading { text-align: center; padding: 40px; color: #718096; }
.rr-spinner {
    display: inline-block; width: 32px; height: 32px;
    border: 3px solid #e2e8f0; border-top: 3px solid #2c5282;
    border-radius: 50%; animation: rrSpin 0.8s linear infinite;
}
@keyframes rrSpin { to { transform: rotate(360deg); } }

/* Generate panel */
.rr-generate-actions { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
.rr-stats { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.rr-stat-card { padding: 12px 20px; background: #f7fafc; border-radius: 6px; text-align: center; }
.rr-stat-num { font-size: 24px; font-weight: 700; color: #2c5282; }
.rr-stat-label { font-size: 12px; color: #718096; }

.rr-selected-org {
    display: flex; align-items: center; gap: 12px; padding: 10px 14px;
    background: #ebf8ff; border: 1px solid #90cdf4; border-radius: 6px; margin-bottom: 16px;
}
.rr-selected-org-name { font-weight: 700; color: #2c5282; flex: 1; }

.rr-empty { text-align: center; padding: 40px 20px; color: #a0aec0; }
.rr-empty i { font-size: 48px; margin-bottom: 12px; display: block; }

input[type="color"] { width: 36px; height: 28px; border: 1px solid #cbd5e0; border-radius: 4px; cursor: pointer; padding: 0; }

/* Print opens in a separate window - hide print output div on main page */
#rr-print-output { display: none; }

@media screen and (max-width: 768px) {
    .rr-config-layout { flex-direction: column; }
    .rr-template-cards { flex-direction: column; }
    .rr-steps { font-size: 12px; }
    .rr-step { padding: 8px 4px; }
}
</style>
</head>
<body>
<script>
(function() {
    var path = window.location.pathname;
    var isPyScript = path.indexOf('/PyScript/') > -1 && path.indexOf('/PyScriptForm/') === -1;
    if (!isPyScript) return;
    var parts = path.split('/');
    var scriptName = '';
    for (var i = 0; i < parts.length; i++) {
        if (parts[i] === 'PyScript') {
            if (i + 1 < parts.length) {
                scriptName = parts[i + 1].split('?')[0];
                break;
            }
        }
    }
    if (!scriptName) return;
    var btPeople = ''' + bt_gathered_pids_json + ''';
    var btOrg = ''' + bt_gathered_org_json + ''';
    if (btPeople.length > 0) {
        sessionStorage.setItem('rrBtPeople', JSON.stringify(btPeople));
        sessionStorage.setItem('rrBtOrg', btOrg.toString());
    }
    window.location.href = '/PyScriptForm/' + scriptName;
})();
</script>

<div class="rr-root">
    <div id="rrToast" class="rr-toast"></div>

    <div class="rr-steps">
        <div class="rr-step active" data-step="1" onclick="goToStep(1)">
            <span class="rr-step-num">1</span> Select Involvement
        </div>
        <div class="rr-step" data-step="2" onclick="goToStep(2)">
            <span class="rr-step-num">2</span> Configure Layout
        </div>
        <div class="rr-step" data-step="3" onclick="goToStep(3)">
            <span class="rr-step-num">3</span> Preview
        </div>
        <div class="rr-step" data-step="4" onclick="goToStep(4)">
            <span class="rr-step-num">4</span> Save &amp; Generate
        </div>
    </div>

    <!-- STEP 1: Select Involvement -->
    <div id="step1" class="rr-panel active">
        <div class="rr-card">
            <div class="rr-card-title">Search Involvements with Registration Questions</div>
            <div class="rr-search-box">
                <input type="text" id="rrSearchInput" placeholder="Search by name or ID..." onkeyup="if(event.keyCode===13) searchOrgs()">
                <select id="rrProgFilter" onchange="searchOrgs()"><option value="">All Programs</option></select>
                <select id="rrDivFilter" onchange="searchOrgs()"><option value="">All Divisions</option></select>
                <button class="rr-btn rr-btn-primary" onclick="searchOrgs()">Search</button>
            </div>
            <div id="rrOrgList">
                <div class="rr-empty"><i class="fa fa-search"></i>Search for an involvement to get started</div>
            </div>
        </div>
    </div>

    <!-- STEP 2: Configure Layout -->
    <div id="step2" class="rr-panel">
        <div id="rrSelectedOrgBar" class="rr-selected-org" style="display:none;">
            <span class="rr-selected-org-name" id="rrOrgNameBar"></span>
            <span class="rr-badge rr-badge-blue" id="rrOrgPeopleCount"></span>
            <span class="rr-badge rr-badge-green" id="rrOrgQuestionCount"></span>
            <button class="rr-btn rr-btn-secondary rr-btn-sm" id="rrChangeOrgBtn" onclick="goToStep(1)">Change</button>
        </div>

        <div class="rr-card">
            <div class="rr-card-title">Choose a Starting Template</div>
            <div class="rr-template-cards">
                <div class="rr-template-card selected" data-tpl="basic" onclick="selectTemplate('basic')">
                    <h4>Basic Registration</h4>
                    <p>Name, contact info, and all registration questions</p>
                </div>
                <div class="rr-template-card" data-tpl="detailed_form" onclick="selectTemplate('detailed_form')">
                    <h4>Detailed Form</h4>
                    <p>WEE-style with emergency contacts, medical info, and family data</p>
                </div>
                <div class="rr-template-card" data-tpl="custom" onclick="selectTemplate('custom')">
                    <h4>Custom</h4>
                    <p>Start with a blank canvas</p>
                </div>
                <div class="rr-template-card" data-tpl="saved" id="rrSavedTplCard" style="display:none;" onclick="selectTemplate('saved')">
                    <h4>Saved Template</h4>
                    <p>Load previously saved template for this org</p>
                </div>
            </div>
        </div>

        <div style="margin-bottom:12px;">
            <button class="rr-help-toggle" onclick="document.getElementById('rrHelpGuide').classList.toggle('open'); this.querySelector('.fa').classList.toggle('fa-question-circle'); this.querySelector('.fa').classList.toggle('fa-times-circle');">
                <i class="fa fa-question-circle"></i> How to use the layout builder
            </button>
        </div>

        <div id="rrHelpGuide" class="rr-help-guide">
            <div class="rr-help-cols">
                <div class="rr-help-section">
                    <h4><i class="fa fa-th-list"></i> Sections</h4>
                    <ul>
                        <li>Sections group related fields together (e.g. "Contact Info", "Medical")</li>
                        <li>Click a section header to <strong>expand/collapse</strong> it</li>
                        <li>Use the <strong>&#9650; &#9660; arrows</strong> to reorder sections</li>
                        <li>Edit the <strong>section title</strong> directly in the header</li>
                        <li>Pick <strong>Single Column</strong> or <strong>Two Column</strong> layout per section</li>
                        <li>Change the section <strong>header color</strong> with the color picker</li>
                    </ul>
                </div>
                <div class="rr-help-section">
                    <h4><i class="fa fa-plus-square"></i> Fields</h4>
                    <ul>
                        <li>Use the <strong>+ Add a field</strong> dropdown at the bottom of each section</li>
                        <li>Field categories: <strong>Person</strong> (name, phone, etc.), <strong>Family</strong> (parents), <strong>Medical/Emergency</strong>, <strong>Registration Questions</strong>, and <strong>Other</strong> (custom text, separator)</li>
                        <li>Use <strong>&#9650; &#9660; arrows</strong> to reorder fields within a section</li>
                        <li>Uncheck <strong>Show</strong> to temporarily hide a field without removing it</li>
                    </ul>
                </div>
                <div class="rr-help-section">
                    <h4><i class="fa fa-columns"></i> Display Formats</h4>
                    <ul>
                        <li><strong>Inline</strong> &ndash; label and value side by side, fits in one column cell</li>
                        <li><strong>Full Width</strong> &ndash; spans both columns in a two-column layout</li>
                        <li><strong>Block</strong> &ndash; label on top, value below (good for long text answers)</li>
                    </ul>
                </div>
                <div class="rr-help-section">
                    <h4><i class="fa fa-sliders"></i> Global Options</h4>
                    <ul>
                        <li><strong>Hide empty fields</strong> &ndash; skips fields with no data for a person</li>
                        <li><strong>Hide unanswered questions</strong> &ndash; skips questions with no answer</li>
                        <li><strong>Show person photo</strong> &ndash; includes profile photo in the report</li>
                        <li><strong>One person per page</strong> &ndash; adds page breaks when printing</li>
                        <li><strong>Compact row spacing</strong> &ndash; tighter spacing for dense reports</li>
                    </ul>
                </div>
            </div>
            <div style="margin-top:12px; padding:8px 12px; background:#ebf8ff; border-radius:6px; font-size:12px; color:#2c5282;">
                <i class="fa fa-lightbulb-o" style="margin-right:4px;"></i>
                <strong>Tip:</strong> Sections titled with "question", "answer", or "registration" will <strong>auto-populate</strong> with all registration questions if left empty. Add individual question fields to override this behavior.
            </div>
        </div>

        <div class="rr-config-layout">
            <div class="rr-config-left">
                <div class="rr-card">
                    <div class="rr-card-title" style="display:flex; justify-content:space-between; align-items:center;">
                        Sections &amp; Fields
                        <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addSection()">+ Add Section</button>
                    </div>
                    <div id="rrSectionsContainer"></div>
                </div>
            </div>

            <div class="rr-config-right">
                <div class="rr-card">
                    <div class="rr-card-title">Global Options</div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Hide empty fields</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptHideEmpty" checked onchange="updateGlobalOption('hideEmptyFields', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Hide unanswered questions</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptHideUnanswered" checked onchange="updateGlobalOption('hideUnansweredQuestions', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Show person photo</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptShowPhoto" onchange="updateGlobalOption('showPersonPhoto', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">One person per page</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptOnePerPage" checked onchange="updatePrintSetting('onePersonPerPage', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Show org header</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptShowOrgHeader" checked onchange="updatePrintSetting('showOrgHeader', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Compact row spacing</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptCompactRows" onchange="updateGlobalOption('compactRows', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Heading color</span>
                        <input type="color" id="rrOptHeadingColor" value="#2c5282" onchange="updateGlobalOption('headingColor', this.value)">
                    </div>
                </div>
                <div style="display:flex; gap:8px;">
                    <button class="rr-btn rr-btn-primary" onclick="goToStep(3)" style="flex:1;">Preview &rarr;</button>
                </div>
            </div>
        </div>
    </div>

    <!-- STEP 3: Preview -->
    <div id="step3" class="rr-panel">
        <div class="rr-preview-toolbar">
            <label style="font-weight:600; font-size:13px;">Preview Person:</label>
            <select id="rrPreviewPerson" onchange="refreshPreview()"></select>
            <button class="rr-btn rr-btn-secondary rr-btn-sm" onclick="refreshPreview()"><i class="fa fa-refresh"></i> Refresh</button>
            <div style="flex:1;"></div>
            <button class="rr-btn rr-btn-secondary rr-btn-sm" onclick="goToStep(2)"><i class="fa fa-cog"></i> Edit Layout</button>
            <button class="rr-btn rr-btn-success" onclick="goToStep(4)">Generate &rarr;</button>
        </div>
        <div id="rrPreviewFrame" class="rr-preview-frame">
            <div class="rr-loading"><div class="rr-spinner"></div><br>Loading preview...</div>
        </div>
    </div>

    <!-- STEP 4: Save & Generate -->
    <div id="step4" class="rr-panel">
        <div class="rr-card">
            <div class="rr-card-title">Save &amp; Generate Report</div>
            <div id="rrGenStats" class="rr-stats"></div>
            <div class="rr-generate-actions">
                <button class="rr-btn rr-btn-primary" onclick="saveTemplate()"><i class="fa fa-save"></i> Save Template</button>
                <button class="rr-btn rr-btn-success" onclick="generateFullReport()"><i class="fa fa-file-text-o"></i> Generate Report</button>
                <button class="rr-btn rr-btn-secondary" id="rrPrintBtn" style="display:none;" onclick="printReport()"><i class="fa fa-print"></i> Print Report</button>
            </div>
        </div>
        <div id="rrGeneratedReport" class="rr-card" style="display:none;">
            <div class="rr-card-title" style="display:flex; justify-content:space-between; align-items:center;">
                Generated Report
                <button class="rr-btn rr-btn-secondary rr-btn-sm" onclick="printReport()"><i class="fa fa-print"></i> Print</button>
            </div>
            <div id="rrReportContent"></div>
        </div>
    </div>
</div>

<div id="rr-print-output"></div>

<script>
(function() {
    var scriptUrl = window.location.pathname;
    if (scriptUrl.indexOf('/PyScript/') > -1) {
        scriptUrl = scriptUrl.replace('/PyScript/', '/PyScriptForm/');
    }

    // Blue Toolbar data comes from sessionStorage (set by PyScript redirect)
    var btOrgId = 0;
    var btPeopleIds = [];
    var _ssBtPeople = sessionStorage.getItem('rrBtPeople');
    if (_ssBtPeople) {
        try {
            btPeopleIds = JSON.parse(_ssBtPeople);
            btOrgId = parseInt(sessionStorage.getItem('rrBtOrg')) || 0;
        } catch(e) {}
        sessionStorage.removeItem('rrBtPeople');
        sessionStorage.removeItem('rrBtOrg');
    }

    var state = {
        currentStep: 1,
        selectedOrgId: null,
        selectedOrgName: '',
        personList: [],
        questions: [],
        availableFields: {},
        template: null,
        savedTemplate: null,
        generatedHtml: '',
        btMode: false,
        btPeopleIds: []
    };

    var TEMPLATES = {
        basic: ''' + tpl_basic_json + ''',
        detailed_form: ''' + tpl_detailed_json + ''',
        custom: ''' + tpl_custom_json + '''
    };

    var fieldIdCounter = 100;
    var sectionIdCounter = 100;

    function ajax(action, params, callback) {
        params = params || {};
        params.action = action;
        if (state.btMode && state.btPeopleIds.length > 0) {
            params.filter_people = state.btPeopleIds.join(',');
        }
        $.ajax({
            url: scriptUrl,
            type: 'POST',
            data: params,
            success: function(response) {
                try {
                    var data = JSON.parse(response);
                    if (callback) callback(data);
                } catch(e) {
                    showToast('Error parsing response', 'danger');
                }
            },
            error: function(xhr, status, error) {
                showToast('Network error: ' + error, 'danger');
            }
        });
    }

    window.showToast = function(msg, cls) {
        var t = document.getElementById('rrToast');
        t.textContent = msg;
        t.className = 'rr-toast rr-toast-' + (cls || 'info') + ' show';
        setTimeout(function() { t.className = 'rr-toast'; }, 3000);
    };

    window.goToStep = function(step) {
        if (step === 1 && state.btMode && state.selectedOrgId) { showToast('Organization was set by Blue Toolbar selection', 'info'); return; }
        if (step > 1 && !state.selectedOrgId) { showToast('Please select an involvement first', 'danger'); return; }
        if (step > 2 && !state.template) { showToast('Please configure the layout first', 'danger'); return; }
        state.currentStep = step;
        document.querySelectorAll('.rr-panel').forEach(function(p) { p.classList.remove('active'); });
        document.getElementById('step' + step).classList.add('active');
        document.querySelectorAll('.rr-step').forEach(function(s) {
            var sn = parseInt(s.getAttribute('data-step'));
            s.classList.remove('active', 'completed');
            if (sn === step) s.classList.add('active');
            else if (sn < step) s.classList.add('completed');
        });
        if (step === 3) refreshPreview();
        if (step === 4) initGeneratePanel();
    };

    // ===== STEP 1 =====
    window.searchOrgs = function() {
        var term = document.getElementById('rrSearchInput').value;
        var progId = document.getElementById('rrProgFilter').value;
        var divId = document.getElementById('rrDivFilter').value;
        document.getElementById('rrOrgList').innerHTML = '<div class="rr-loading"><div class="rr-spinner"></div><br>Searching...</div>';
        ajax('search_orgs', {search_term: term, program_id: progId, division_id: divId}, function(data) {
            if (!data.success) { document.getElementById('rrOrgList').innerHTML = '<div class="rr-empty">Error: ' + data.message + '</div>'; return; }
            if (!data.orgs || data.orgs.length === 0) { document.getElementById('rrOrgList').innerHTML = '<div class="rr-empty"><i class="fa fa-info-circle"></i>No involvements found</div>'; return; }
            var html = '';
            for (var i = 0; i < data.orgs.length; i++) {
                var o = data.orgs[i];
                var sel = (o.id === state.selectedOrgId) ? ' selected' : '';
                html += '<div class="rr-org-item' + sel + '" onclick="selectOrg(' + o.id + ', this)">';
                html += '<div><div class="rr-org-name">' + escHtml(o.name) + '</div>';
                html += '<div class="rr-org-meta">' + escHtml(o.program) + ' &bull; ' + escHtml(o.division) + '</div></div>';
                html += '<div><span class="rr-badge rr-badge-blue">' + o.memberCount + ' people</span> ';
                html += '<span class="rr-badge rr-badge-green">' + o.questionCount + ' questions</span></div>';
                html += '</div>';
            }
            document.getElementById('rrOrgList').innerHTML = html;
        });
    };

    window.selectOrg = function(orgId, el) {
        document.querySelectorAll('.rr-org-item').forEach(function(e) { e.classList.remove('selected'); });
        if (el) el.classList.add('selected');
        state.selectedOrgId = orgId;
        showToast('Loading organization data...', 'info');
        ajax('load_org_data', {org_id: orgId}, function(data) {
            if (!data.success) { showToast('Error: ' + data.message, 'danger'); return; }
            state.selectedOrgName = data.orgName;
            state.personList = data.personList || [];
            state.questions = data.questions || [];
            state.availableFields = data.availableFields || {};
            document.getElementById('rrOrgNameBar').textContent = data.orgName;
            var countLabel = state.btMode ? data.personCount + ' selected' : data.personCount + ' people';
            document.getElementById('rrOrgPeopleCount').textContent = countLabel;
            document.getElementById('rrOrgQuestionCount').textContent = data.questionCount + ' questions';
            document.getElementById('rrSelectedOrgBar').style.display = 'flex';
            var changeBtn = document.getElementById('rrChangeOrgBtn');
            if (changeBtn && state.btMode) changeBtn.style.display = 'none';
            var sel = document.getElementById('rrPreviewPerson');
            sel.innerHTML = '';
            for (var i = 0; i < state.personList.length; i++) {
                var opt = document.createElement('option');
                opt.value = state.personList[i].id;
                opt.textContent = state.personList[i].name;
                sel.appendChild(opt);
            }
            ajax('load_template', {org_id: orgId}, function(tplData) {
                if (tplData.success && tplData.hasSaved) {
                    state.savedTemplate = tplData.template;
                    document.getElementById('rrSavedTplCard').style.display = 'block';
                } else {
                    state.savedTemplate = null;
                    document.getElementById('rrSavedTplCard').style.display = 'none';
                }
                selectTemplate('basic');
                showToast('Organization loaded! Configure layout next.', 'success');
                goToStep(2);
            });
        });
    };

    // ===== STEP 2 =====
    window.selectTemplate = function(tplName) {
        document.querySelectorAll('.rr-template-card').forEach(function(c) { c.classList.remove('selected'); });
        var card = document.querySelector('.rr-template-card[data-tpl="' + tplName + '"]');
        if (card) card.classList.add('selected');
        if (tplName === 'saved' && state.savedTemplate) {
            state.template = JSON.parse(JSON.stringify(state.savedTemplate));
        } else if (TEMPLATES[tplName]) {
            state.template = JSON.parse(JSON.stringify(TEMPLATES[tplName]));
        }
        if (state.template) {
            var go = state.template.globalOptions || {};
            var ps = state.template.printSettings || {};
            setChecked('rrOptHideEmpty', go.hideEmptyFields !== false);
            setChecked('rrOptHideUnanswered', go.hideUnansweredQuestions !== false);
            setChecked('rrOptShowPhoto', go.showPersonPhoto === true);
            setChecked('rrOptOnePerPage', ps.onePersonPerPage !== false);
            setChecked('rrOptShowOrgHeader', ps.showOrgHeader !== false);
            setChecked('rrOptCompactRows', go.compactRows === true);
            document.getElementById('rrOptHeadingColor').value = go.headingColor || '#2c5282';
        }
        renderSections();
    };

    function setChecked(id, val) { var el = document.getElementById(id); if (el) el.checked = val; }

    window.updateGlobalOption = function(key, value) {
        if (!state.template) return;
        if (!state.template.globalOptions) state.template.globalOptions = {};
        state.template.globalOptions[key] = value;
    };
    window.updatePrintSetting = function(key, value) {
        if (!state.template) return;
        if (!state.template.printSettings) state.template.printSettings = {};
        state.template.printSettings[key] = value;
    };

    function renderSections() {
        if (!state.template || !state.template.sections) return;
        var container = document.getElementById('rrSectionsContainer');
        var sections = state.template.sections.sort(function(a, b) { return (a.order || 0) - (b.order || 0); });
        var html = '';
        for (var si = 0; si < sections.length; si++) {
            var sec = sections[si];
            var fields = (sec.fields || []).sort(function(a, b) { return (a.order || 0) - (b.order || 0); });
            html += '<div class="rr-section-editor" data-secidx="' + si + '">';
            html += '<div class="rr-section-head" onclick="toggleSectionBody(this)">';
            html += '<div style="display:flex;align-items:center;gap:8px;">';
            html += '<span style="display:flex;flex-direction:column;gap:0;" onclick="event.stopPropagation()">';
            html += '<button onclick="moveSection(' + si + ',-1)" title="Move up" style="border:none;background:none;cursor:pointer;padding:0;line-height:1;font-size:10px;color:' + (si > 0 ? '#4a5568' : '#e2e8f0') + ';">&#9650;</button>';
            html += '<button onclick="moveSection(' + si + ',1)" title="Move down" style="border:none;background:none;cursor:pointer;padding:0;line-height:1;font-size:10px;color:' + (si < sections.length - 1 ? '#4a5568' : '#e2e8f0') + ';">&#9660;</button>';
            html += '</span>';
            html += '<input type="text" value="' + escAttr(sec.title || '') + '" onchange="updateSectionTitle(' + si + ', this.value)" onclick="event.stopPropagation()" style="border:none;background:transparent;font-weight:600;font-size:14px;width:200px;">';
            html += '<input type="color" value="' + escAttr(sec.headerColor || '#2c5282') + '" onchange="updateSectionColor(' + si + ', this.value)" onclick="event.stopPropagation()" style="width:24px;height:20px;border:none;cursor:pointer;">';
            html += '</div>';
            html += '<div style="display:flex;align-items:center;gap:6px;">';
            html += '<select onchange="updateSectionLayout(' + si + ', this.value)" onclick="event.stopPropagation()" style="font-size:11px;padding:2px 4px;">';
            html += '<option value="single-column"' + (sec.layout === 'single-column' ? ' selected' : '') + '>Single Column</option>';
            html += '<option value="two-column"' + (sec.layout === 'two-column' ? ' selected' : '') + '>Two Column</option>';
            html += '</select>';
            html += '<span style="font-size:11px;color:#718096;">' + fields.length + ' fields</span>';
            html += '<button class="rr-btn rr-btn-danger rr-btn-sm" onclick="event.stopPropagation();removeSection(' + si + ')" title="Remove">&times;</button>';
            html += '</div></div>';
            html += '<div class="rr-section-body' + (openSections[si] ? ' open' : '') + '">';
            if (!fields.length) {
                var tl = (sec.title || '').toLowerCase();
                if (tl.indexOf('question') >= 0 || tl.indexOf('answer') >= 0 || tl.indexOf('registration') >= 0) {
                    html += '<div style="padding:8px;color:#718096;font-size:12px;font-style:italic;">Auto-populates with all registration questions. Add fields to override.</div>';
                }
            }
            for (var fi = 0; fi < fields.length; fi++) {
                var f = fields[fi];
                var mvBtns = '<span style="display:flex;flex-direction:column;gap:0;">'
                    + '<button onclick="moveField(' + si + ',' + fi + ',-1)" title="Move up" style="border:none;background:none;cursor:pointer;padding:0;line-height:1;font-size:9px;color:' + (fi > 0 ? '#4a5568' : '#e2e8f0') + ';">&#9650;</button>'
                    + '<button onclick="moveField(' + si + ',' + fi + ',1)" title="Move down" style="border:none;background:none;cursor:pointer;padding:0;line-height:1;font-size:9px;color:' + (fi < fields.length - 1 ? '#4a5568' : '#e2e8f0') + ';">&#9660;</button>'
                    + '</span>';
                if (f.fieldType === 'static' && f.sourceField === 'text') {
                    html += '<div class="rr-field-item" style="flex-wrap:wrap;">';
                    html += mvBtns;
                    html += '<span class="rr-field-label-text" style="display:flex;align-items:center;gap:4px;">';
                    html += '<i class="fa fa-file-text-o" style="color:#718096;"></i>';
                    html += '<input type="text" value="' + escAttr(f.label || '') + '" onchange="updateFieldLabel(' + si + ',' + fi + ',this.value)" style="border:none;background:transparent;font-size:13px;width:120px;" placeholder="Title (optional)">';
                    html += '</span>';
                    html += '<label style="font-size:11px;white-space:nowrap;"><input type="checkbox"' + (f.visible !== false ? ' checked' : '') + ' onchange="updateFieldVisible(' + si + ',' + fi + ',this.checked)"> Show</label>';
                    html += '<button class="rr-btn rr-btn-danger rr-btn-sm" onclick="removeField(' + si + ',' + fi + ')" style="padding:2px 6px;">&times;</button>';
                    html += '<textarea onchange="updateFieldContent(' + si + ',' + fi + ',this.value)" placeholder="Type your text here (explanation, caution, notes...)" style="width:100%;margin-top:4px;min-height:50px;font-size:12px;padding:6px 8px;border:1px solid #cbd5e0;border-radius:4px;resize:vertical;font-family:inherit;">' + escHtml(f.staticContent || '') + '</textarea>';
                    html += '</div>';
                } else if (f.fieldType === 'static' && f.sourceField === 'separator') {
                    html += '<div class="rr-field-item">';
                    html += mvBtns;
                    html += '<span class="rr-field-label-text" style="color:#718096;font-style:italic;"><i class="fa fa-minus" style="margin-right:4px;"></i>Separator Line</span>';
                    html += '<label style="font-size:11px;white-space:nowrap;"><input type="checkbox"' + (f.visible !== false ? ' checked' : '') + ' onchange="updateFieldVisible(' + si + ',' + fi + ',this.checked)"> Show</label>';
                    html += '<button class="rr-btn rr-btn-danger rr-btn-sm" onclick="removeField(' + si + ',' + fi + ')" style="padding:2px 6px;">&times;</button>';
                    html += '</div>';
                } else {
                    html += '<div class="rr-field-item">';
                    html += mvBtns;
                    html += '<span class="rr-field-label-text">';
                    html += '<input type="text" value="' + escAttr(f.label || '') + '" onchange="updateFieldLabel(' + si + ',' + fi + ',this.value)" style="border:none;background:transparent;font-size:13px;width:140px;">';
                    html += '</span>';
                    html += '<select onchange="updateFieldFormat(' + si + ',' + fi + ',this.value)" style="width:100px;">';
                    html += '<option value="single-line"' + (f.displayFormat === 'single-line' ? ' selected' : '') + '>Inline</option>';
                    html += '<option value="full-width"' + (f.displayFormat === 'full-width' ? ' selected' : '') + '>Full Width</option>';
                    html += '<option value="block"' + (f.displayFormat === 'block' ? ' selected' : '') + '>Block</option>';
                    html += '</select>';
                    html += '<label style="font-size:11px;white-space:nowrap;"><input type="checkbox"' + (f.visible !== false ? ' checked' : '') + ' onchange="updateFieldVisible(' + si + ',' + fi + ',this.checked)"> Show</label>';
                    html += '<button class="rr-btn rr-btn-danger rr-btn-sm" onclick="removeField(' + si + ',' + fi + ')" style="padding:2px 6px;">&times;</button>';
                    html += '</div>';
                }
            }
            html += '<div class="rr-add-field-row">';
            html += '<select id="rrAddField_' + si + '"><option value="">+ Add a field...</option>';
            html += buildFieldOptions();
            html += '</select>';
            html += '<button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addFieldToSection(' + si + ')">Add</button>';
            html += '</div></div></div>';
        }
        container.innerHTML = html;
    }

    function buildFieldOptions() {
        var html = '';
        var af = state.availableFields || {};
        html += '<optgroup label="Person Fields">';
        if (af.person) { for (var i = 0; i < af.person.length; i++) { html += '<option value="person|' + escAttr(af.person[i].sourceField) + '|' + escAttr(af.person[i].label) + '">' + escHtml(af.person[i].label) + '</option>'; } }
        html += '</optgroup><optgroup label="Family Fields">';
        if (af.family) { for (var i = 0; i < af.family.length; i++) { html += '<option value="family|' + escAttr(af.family[i].sourceField) + '|' + escAttr(af.family[i].label) + '">' + escHtml(af.family[i].label) + '</option>'; } }
        html += '</optgroup><optgroup label="Medical/Emergency">';
        if (af.medical) { for (var i = 0; i < af.medical.length; i++) { html += '<option value="medical|' + escAttr(af.medical[i].sourceField) + '|' + escAttr(af.medical[i].label) + '">' + escHtml(af.medical[i].label) + '</option>'; } }
        html += '</optgroup><optgroup label="Registration Questions">';
        if (af.regquestion) { for (var i = 0; i < af.regquestion.length; i++) { var l = af.regquestion[i].label; var sl = l.length > 50 ? l.substring(0,47) + '...' : l; html += '<option value="regquestion|' + escAttr(af.regquestion[i].sourceField) + '|' + escAttr(l) + '">' + escHtml(sl) + '</option>'; } }
        html += '</optgroup><optgroup label="Other"><option value="static|text|">Custom Text / Note</option><option value="static|separator|---">Separator Line</option></optgroup>';
        return html;
    }

    var openSections = {};
    window.toggleSectionBody = function(el) {
        var body = el.nextElementSibling;
        body.classList.toggle('open');
        var si = el.parentElement.getAttribute('data-secidx');
        if (body.classList.contains('open')) { openSections[si] = true; } else { delete openSections[si]; }
    };
    window.updateSectionTitle = function(si, val) { if (state.template && state.template.sections[si]) state.template.sections[si].title = val; };
    window.updateSectionColor = function(si, val) { if (state.template && state.template.sections[si]) state.template.sections[si].headerColor = val; };
    window.updateSectionLayout = function(si, val) { if (state.template && state.template.sections[si]) state.template.sections[si].layout = val; };

    window.moveSection = function(si, dir) {
        var secs = state.template.sections;
        var ni = si + dir;
        if (ni < 0 || ni >= secs.length) return;
        var tmp = secs[si]; secs[si] = secs[ni]; secs[ni] = tmp;
        for (var i = 0; i < secs.length; i++) secs[i].order = i + 1;
        var oSi = openSections[si], oNi = openSections[ni];
        if (oSi) { openSections[ni] = true; } else { delete openSections[ni]; }
        if (oNi) { openSections[si] = true; } else { delete openSections[si]; }
        renderSections();
    };
    window.moveField = function(si, fi, dir) {
        var fields = state.template.sections[si].fields;
        var ni = fi + dir;
        if (ni < 0 || ni >= fields.length) return;
        var tmp = fields[fi]; fields[fi] = fields[ni]; fields[ni] = tmp;
        for (var i = 0; i < fields.length; i++) fields[i].order = i + 1;
        renderSections();
    };
    window.addSection = function() {
        if (!state.template) return;
        sectionIdCounter++;
        var newIdx = state.template.sections.length;
        state.template.sections.push({ sectionId: 'sec_' + sectionIdCounter, title: 'New Section', order: newIdx + 1, layout: 'single-column', headerColor: (state.template.globalOptions || {}).headingColor || '#2c5282', visible: true, fields: [] });
        openSections[newIdx] = true;
        renderSections();
    };
    window.removeSection = function(si) {
        if (!state.template) return;
        delete openSections[si];
        state.template.sections.splice(si, 1);
        for (var i = 0; i < state.template.sections.length; i++) state.template.sections[i].order = i + 1;
        renderSections();
    };
    window.addFieldToSection = function(si) {
        if (!state.template || !state.template.sections[si]) return;
        var sel = document.getElementById('rrAddField_' + si);
        if (!sel || !sel.value) return;
        var parts = sel.value.split('|');
        if (parts.length < 2) return;
        fieldIdCounter++;
        var fields = state.template.sections[si].fields || [];
        var newField = { fieldId: 'fld_' + fieldIdCounter, fieldType: parts[0], sourceField: parts[1], label: parts[2] || '', displayFormat: (parts[0] === 'regquestion' || parts[0] === 'static') ? 'block' : 'single-line', order: fields.length + 1, visible: true, colSpan: 1 };
        if (parts[0] === 'static' && parts[1] === 'text') {
            newField.staticContent = '';
            newField.label = 'Note';
        }
        fields.push(newField);
        state.template.sections[si].fields = fields;
        openSections[si] = true;
        renderSections();
    };
    window.removeField = function(si, fi) {
        if (!state.template || !state.template.sections[si]) return;
        state.template.sections[si].fields.splice(fi, 1);
        for (var i = 0; i < state.template.sections[si].fields.length; i++) state.template.sections[si].fields[i].order = i + 1;
        openSections[si] = true;
        renderSections();
    };
    window.updateFieldLabel = function(si, fi, val) { if (state.template && state.template.sections[si] && state.template.sections[si].fields[fi]) state.template.sections[si].fields[fi].label = val; };
    window.updateFieldFormat = function(si, fi, val) {
        if (state.template && state.template.sections[si] && state.template.sections[si].fields[fi]) {
            state.template.sections[si].fields[fi].displayFormat = val;
            state.template.sections[si].fields[fi].colSpan = (val === 'full-width' || val === 'block') ? 2 : 1;
        }
    };
    window.updateFieldVisible = function(si, fi, val) { if (state.template && state.template.sections[si] && state.template.sections[si].fields[fi]) state.template.sections[si].fields[fi].visible = val; };
    window.updateFieldContent = function(si, fi, val) { if (state.template && state.template.sections[si] && state.template.sections[si].fields[fi]) state.template.sections[si].fields[fi].staticContent = val; };

    // ===== STEP 3 =====
    window.refreshPreview = function() {
        if (!state.selectedOrgId || !state.template) return;
        var personId = document.getElementById('rrPreviewPerson').value || '';
        document.getElementById('rrPreviewFrame').innerHTML = '<div class="rr-loading"><div class="rr-spinner"></div><br>Generating preview...</div>';
        ajax('preview_report', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template), people_id: personId }, function(data) {
            if (data.success) { document.getElementById('rrPreviewFrame').innerHTML = data.html; }
            else { document.getElementById('rrPreviewFrame').innerHTML = '<div class="rr-empty">Error: ' + (data.message || 'Unknown') + '</div>'; }
        });
    };

    // ===== STEP 4 =====
    function initGeneratePanel() {
        var h = '';
        h += '<div class="rr-stat-card"><div class="rr-stat-num">' + state.personList.length + '</div><div class="rr-stat-label">Registrants</div></div>';
        h += '<div class="rr-stat-card"><div class="rr-stat-num">' + state.questions.length + '</div><div class="rr-stat-label">Questions</div></div>';
        h += '<div class="rr-stat-card"><div class="rr-stat-num">' + (state.template ? state.template.sections.length : 0) + '</div><div class="rr-stat-label">Sections</div></div>';
        document.getElementById('rrGenStats').innerHTML = h;
    }

    window.saveTemplate = function() {
        if (!state.selectedOrgId || !state.template) return;
        if (state.selectedOrgId === 'bt_direct') { showToast('Cannot save template without an involvement selected. Templates are saved per-involvement.', 'danger'); return; }
        ajax('save_template', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template) }, function(data) {
            if (data.success) {
                showToast('Template saved!', 'success');
                state.savedTemplate = JSON.parse(JSON.stringify(state.template));
                document.getElementById('rrSavedTplCard').style.display = 'block';
            } else { showToast('Error: ' + data.message, 'danger'); }
        });
    };

    window.generateFullReport = function() {
        if (!state.selectedOrgId || !state.template) return;
        document.getElementById('rrGeneratedReport').style.display = 'block';
        document.getElementById('rrReportContent').innerHTML = '<div class="rr-loading"><div class="rr-spinner"></div><br>Generating report...</div>';
        ajax('generate_report', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template) }, function(data) {
            if (data.success) {
                state.generatedHtml = data.html;
                document.getElementById('rrReportContent').innerHTML = data.html;
                document.getElementById('rrPrintBtn').style.display = 'inline-block';
                showToast('Report generated for ' + (data.personCount || 0) + ' registrants!', 'success');
            } else { document.getElementById('rrReportContent').innerHTML = '<div class="rr-empty">Error: ' + (data.message || 'Unknown') + '</div>'; }
        });
    };

    window.printReport = function() {
        if (!state.generatedHtml) { showToast('Generate the report first', 'danger'); return; }
        var hc = (state.template.globalOptions || {}).headingColor || '#2c5282';
        var ff = (state.template.globalOptions || {}).fontFamily || 'Segoe UI, sans-serif';
        var css = '';
        css += '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}';
        css += 'body{margin:0;padding:20px;font-family:' + ff + ';color:#333}';
        css += '.rr-report{font-family:' + ff + ';color:#333}';
        css += '.rr-person-page{padding:20px 0}';
        css += '.rr-page-break{page-break-after:always;border:none;margin:0;padding:0}';
        css += '.rr-org-header{text-align:center;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid ' + hc + '}';
        css += '.rr-org-header h2{margin:0;font-size:20px;color:' + hc + '}';
        css += '.rr-person-header{display:flex;align-items:center;gap:12px;margin-bottom:16px;padding-bottom:8px;border-bottom:1px solid #e2e8f0}';
        css += '.rr-photo{width:50px;height:50px;border-radius:6px;object-fit:cover}';
        css += '.rr-person-name{font-size:22px;font-weight:700}';
        css += '.rr-section{margin-bottom:16px;page-break-inside:avoid}';
        css += '.rr-section-header{padding:6px 12px;color:#fff;font-weight:700;font-size:14px;border-radius:4px 4px 0 0}';
        css += '.rr-fields-grid{border:1px solid #ccc;border-top:none;border-radius:0 0 4px 4px;padding:10px 12px}';
        css += '.rr-two-col{display:grid;grid-template-columns:1fr 1fr;gap:6px 16px}';
        css += '.rr-one-col{display:grid;grid-template-columns:1fr;gap:6px}';
        css += '.rr-full-width{grid-column:1/-1}';
        css += '.rr-field{padding:3px 0}.rr-field-label{font-weight:600;color:#4a5568;font-size:12px}';
        css += '.rr-field-value{color:#1a202c;font-size:13px}';
        css += '.rr-field-block{grid-column:1/-1;padding:4px 0}';
        css += '.rr-block-value{display:block;padding:4px 8px;background:#f7f7f7;border:1px solid #ddd;border-radius:4px;margin-top:2px;word-wrap:break-word;white-space:pre-wrap;min-height:24px;font-size:13px}';
        css += '.rr-compact .rr-person-page{padding:8px 0}';
        css += '.rr-compact .rr-person-header{margin-bottom:6px;padding-bottom:4px}';
        css += '.rr-compact .rr-section{margin-bottom:6px}';
        css += '.rr-compact .rr-section-header{padding:3px 10px;font-size:13px}';
        css += '.rr-compact .rr-fields-grid{padding:4px 10px}';
        css += '.rr-compact .rr-two-col{gap:1px 12px}';
        css += '.rr-compact .rr-one-col{gap:1px}';
        css += '.rr-compact .rr-field{padding:1px 0}';
        css += '.rr-compact .rr-field-block{padding:1px 0}';
        css += '.rr-compact .rr-block-value{padding:2px 6px;margin-top:1px;min-height:16px}';
        css += '.rr-compact .rr-org-header{margin-bottom:6px;padding-bottom:4px}';
        var pw = window.open('', '_blank');
        if (!pw) { showToast('Popup blocked - please allow popups for this site', 'danger'); return; }
        pw.document.write('<!DOCTYPE html><html><head><title>Registration Report</title>');
        pw.document.write('<style>' + css + '</style>');
        pw.document.write('</head><body>');
        pw.document.write(state.generatedHtml);
        pw.document.write('</body></html>');
        pw.document.close();
        pw.focus();
        setTimeout(function() { pw.print(); }, 300);
    };

    function escHtml(str) {
        if (!str) return '';
        var div = document.createElement('div');
        div.appendChild(document.createTextNode(str));
        return div.innerHTML;
    }
    function escAttr(str) {
        if (!str) return '';
        return String(str).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }

    function loadFilters() {
        ajax('get_filters', {}, function(data) {
            if (!data.success) return;
            var progSel = document.getElementById('rrProgFilter');
            for (var i = 0; i < data.programs.length; i++) {
                var opt = document.createElement('option');
                opt.value = data.programs[i].id;
                opt.textContent = data.programs[i].name;
                progSel.appendChild(opt);
            }
            var divSel = document.getElementById('rrDivFilter');
            for (var i = 0; i < data.divisions.length; i++) {
                var opt = document.createElement('option');
                opt.value = data.divisions[i].id;
                opt.textContent = data.divisions[i].name;
                divSel.appendChild(opt);
            }
        });
    }

    function initApp() {
        // Blue Toolbar mode - has selected people
        if (btPeopleIds.length > 0) {
            state.btMode = true;
            state.btPeopleIds = btPeopleIds;
            loadFilters();

            if (btOrgId > 0) {
                // BT with org - auto-select org (selectOrg navigates to Step 2)
                showToast('Blue Toolbar: Loading ' + btPeopleIds.length + ' selected people...', 'info');
                selectOrg(btOrgId, null);
            } else {
                // BT without org - load people directly, skip to Step 2
                showToast('Loading ' + btPeopleIds.length + ' selected people...', 'info');
                ajax('load_bt_data', {people_ids: btPeopleIds.join(',')}, function(data) {
                    if (!data.success) { showToast('Error: ' + data.message, 'danger'); return; }
                    state.selectedOrgId = 'bt_direct';
                    state.selectedOrgName = data.orgName;
                    state.personList = data.personList || [];
                    state.questions = data.questions || [];
                    state.availableFields = data.availableFields || {};
                    document.getElementById('rrOrgNameBar').textContent = data.personCount + ' Selected People';
                    document.getElementById('rrOrgPeopleCount').textContent = data.personCount + ' people';
                    document.getElementById('rrOrgQuestionCount').textContent = 'No involvement selected';
                    document.getElementById('rrSelectedOrgBar').style.display = 'flex';
                    var changeBtn = document.getElementById('rrChangeOrgBtn');
                    if (changeBtn) changeBtn.style.display = 'none';
                    var sel = document.getElementById('rrPreviewPerson');
                    sel.innerHTML = '';
                    for (var i = 0; i < state.personList.length; i++) {
                        var opt = document.createElement('option');
                        opt.value = state.personList[i].id;
                        opt.textContent = state.personList[i].name;
                        sel.appendChild(opt);
                    }
                    state.savedTemplate = null;
                    document.getElementById('rrSavedTplCard').style.display = 'none';
                    selectTemplate('basic');
                    showToast('People loaded! Person, family & medical fields available. Select an involvement to add registration questions.', 'success');
                    goToStep(2);
                });
            }
            return;
        }

        // Normal mode - load filters, show prompt to search
        loadFilters();
        document.getElementById('rrOrgList').innerHTML = '<div class="rr-empty"><i class="fa fa-search" style="font-size:24px;color:#a0aec0;margin-bottom:8px;display:block;"></i>Search for an involvement by name, or use the program/division filters above.</div>';
    }

    // Wait for jQuery (loaded by TouchPoint page wrapper) before initializing
    (function waitForJQuery() {
        if (typeof $ !== 'undefined') {
            $(document).ready(initApp);
        } else {
            setTimeout(waitForJQuery, 50);
        }
    })();
})();
</script>
</body>
</html>'''

    print html
    model.Form = html
