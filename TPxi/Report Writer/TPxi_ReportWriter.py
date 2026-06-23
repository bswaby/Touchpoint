#roles=Edit
#----------------------------------------------------------------------
# TPxi_RollSheet.py
#
# Configurable Rollsheet Generator for TouchPoint
#
# Allows admins to create reusable rollsheet configurations through a UI,
# then generate and print rollsheets for any ministry area. Each config
# specifies source (program/division or specific orgs), columns to display,
# universal data sources (reg questions, extra values, recreg), layout, and sort options.
#
# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub: https://github.com/bswaby/Touchpoint  (40+ free tools)
# ----------------------------------------------------------------
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to
# support continued development, check out:
#
# DisplayCache - church digital signage that integrates with TouchPoint
# https://displaycache.com
#
# TPxi Go - your church contacts, wherever you work.
# Look up anyone in TouchPoint, log calls and emails from Outlook
# or your phone. No tab switching, no lost context.
# https://tpxigo.com
# ----------------------------------------------------------------

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


Change Log:
v1.6.0 - June 2026
  - Added: Form Summary page -- a Wufoo / MS-Forms style aggregate view.
           "Form summary page (counts & charts)" toggle under Supplemental
           Pages. KPI cards (registrants, unique people, gender, avg age),
           plus bar charts for gender / age bins / subgroups / top ZIPs and
           a per-question answer tally. Pure HTML/CSS bars (no Chart.js) so
           it prints in the popup. Prepended like the cover page
           (printSettings.showSummaryPage); per-person output unchanged when
           off. No new SQL -- aggregates already-fetched data.
  - Added: Customizable summary via a picker (same UX as the Missing-Info /
           Medical item pickers). Choose exactly which items appear and in
           what order: demographics (gender, age bins, marital status,
           state, city, top ZIPs), subgroups, registration questions, the
           same Person/Family/Medical fields as the section builder, and
           Extra Value fields (PeopleExtra). Stored as
           printSettings.summaryItems[] = [{itemType, field, label}]. Empty
           = automatic set. EV names fold into extract_ev_names_from_template
           so they're fetched. render dispatch: _render_summary_item().
  - Added: reorder arrows on summary / missing / medical item chips (no
           more delete-and-readd to change order).
  - Added: "Skip per-person detail" toggle -- output only the summary /
           supplemental pages (printSettings.hidePersonDetail).
  - Fixed: multi-value fields (Approved Medications, Subgroups, Authorized
           Checkout) split their comma-joined lists and tally each item
           separately instead of counting each person's whole list as one
           value (SUMMARY_MULTIVALUE_FIELDS).
  - Fixed: picking a SAVED template didn't restore supplemental-page
           settings (summary/cover/missing/medical toggles, panels and item
           lists) -- the saved-template load path synced only a few global
           options using stale element IDs. Both load paths now share
           syncTemplateOptionsToUI(), so saved supplemental config (incl.
           summaryItems / hidePersonDetail) is restored on load, import, and
           auto-update.

v1.5.4 - June 2026
  - Fixed: Legacy RegistrationData XML fallback referenced an undefined
           `question_set` variable. The NameError was silently swallowed by
           the surrounding try/except, so orgs using the older XML-based
           registration format showed "0 questions" with no clue why.
           Affected installs where the registration data lives in
           RegistrationData.Data rather than RegQuestion/RegAnswer.
v1.5.3 - June 2026
  - Added: "Include dropped registrants" toggle on the selected-involvement bar.
           Surfaces people who answered the org's registration questions but are
           no longer in OrganizationMembers (dropped, moved, etc.). Their reg
           answers persist on the original involvement. Dropped people are
           marked with a red "DROPPED" badge in the report header and "(Dropped)"
           in the preview-person dropdown. Counter ("X dropped") appears next to
           the people count when any dropped registrants exist.
v1.5.2 - May 2026
  - Added: Custom Document Title (template-level field). Replaces the
           "Registration Report" cover-page heading AND the involvement
           name (or "Selected People" in Blue Toolbar runs) shown above
           each person. Persists with Save / Save As / Export.
  - Added: Per-user Blue Toolbar template library. Running ReportWriter
           from the Blue Toolbar without an involvement now saves templates
           to the current user's PeopleExtra (RegReportTemplate) so each
           staffer builds their own personal library. Save / Save As /
           Rename / Delete all work without an org context.
  - Fixed: When picking a saved template card, the Document Title input
           now pre-fills from the saved value instead of staying blank
           (which silently overwrote the stored title on the next save).
  - Smart: When "One person per page" is OFF, the document title renders
           once at the top of the report instead of repeating above each
           person on a continuous page. When ON, per-person header stays
           so every printed page gets its own banner.
v1.5.1 - May 2026
  - Fixed: CSV export missed auto-populated registration questions on Basic template
           (and any template with empty reg-question sections). PDF render auto-populated
           sections whose title contained "question/answer/registration" with one column
           per available question; CSV iterated raw section['fields'] and produced zero
           columns for those sections. Extracted the auto-populator into a shared helper
           (auto_populate_section_fields) so PDF and CSV emit the same columns.
  - Fixed: CSV now respects section.visible and field.visible flags (was ignoring both).
v1.5 - May 2026
  - Added: Multi-template save - keep multiple named report layouts per involvement
           (Save / Save As / Rename / Delete) instead of a single overwriting save
  - Added: Import / Export templates as JSON files for sharing across involvements or churches
  - Added: Mismatch detection on import - case-insensitive auto-remap of question labels and
           subgroup names, plus a warning banner listing anything that doesn't resolve, with
           an "Auto-remove missing fields" one-click cleanup
  - Added: Subgroups field (org-scoped) - shows each person's subgroup memberships within the
           current involvement, with per-field display filter to pick which subgroups to show
  - Added: Approved Medications now pulls from dbo.PersonMedication (canonical, current source)
           in addition to RegAnswer and legacy RecReg booleans, all merged with case-insensitive
           dedup so nothing is missed
  - Fixed: OTC medication answers were tied to a single hardcoded RegQuestionId from one
           involvement; now matches by question label so it works across every org with an
           "Approved over-the-counter medications" question
  - Added: Auto-update mechanism - checks scripts.displaycache.com on page load, shows a banner
           when a newer version is available, one-click update preserves all saved templates
  - Backward-compat: existing single-saved-template installs auto-appear as "Default" with no
           data loss; storage migrates to the new multi-template format on the next save
v1.4 - April 2026
  - Added: CSV export button (Export CSV) alongside Print, using same template layout for columns
  - Added: Page numbers option using @page CSS for true printed page numbering
  - Added: "Allow sections to split across pages" toggle to eliminate blank space from page-break-inside:avoid
  - Added: Marital Status as a person field option (joined from lookup.MaritalStatus, not hardcoded)
  - Fixed: Duplicate registration questions with same label now handled correctly (keyed by RegQuestionId)
  - Fixed: Backward compatibility preserved for saved templates with duplicate question labels
  - Fixed: Registration answers no longer display with surrounding quotes
  - Fixed: JSON array answers (checkboxes) now display as comma-separated text instead of raw JSON
  - Added: "Preserve line breaks in answers" toggle - converts literal \n in answers to line breaks when on, strips them when off
v1.3 - March 2026
  - Added: Cover Page option with summary statistics (total registrants, gender breakdown, org info)
  - Added: Missing Information page with configurable field picker (profile, medical, family, reg questions)
  - Added: Medical summary page with configurable item picker (medical fields, reg questions, keyword search)
  - All supplemental pages use tag-based item pickers instead of hardcoded checkboxes
v1.2 - March 2026
  - Fixed: Unicode encoding crash on special characters (e.g. senor, cafe) via safe_str() transliteration
  - Added: Custody Issue field (boolean from People.CustodyIssue) in Medical/Emergency fields
  - Added: Extra Value field support - type in any Person Extra Value name with optional friendly label
v1.1 - March 2026
  - Fixed: Allergies field showing "True" instead of actual allergy text (MedAllergy is boolean; now displays MedicalDescription)
  - Fixed: Blue Toolbar script reference is now dynamic instead of hardcoded
  - Added: Authorized Checkout field (from PeopleAuthorizedCheckOut table)
  - Added: Approved Medications field (new PersonMedication + lookup.Medication with legacy RecReg boolean fallback)
  - Removed: Medical Conditions field (duplicate of Allergies - both sourced from MedicalDescription)
v1.0 - February 2026
  - Initial release

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

# --- Version / Auto-update -------------------------------------------
APP_VERSION = '1.5.4'
DC_SCRIPT_ID = 'TPxi_ReportWriter'  # ID used on DisplayCache to identify this script
# scripts.displaycache.com is the custom domain used for browser-side version checks.
# workers.dev is used for server-side fetches (bypasses Cloudflare Bot Fight Mode).
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'

def get_script_name():
    """Detect the actual script name TouchPoint installed this as (admin may have
    renamed it). Order: posted script_name, model.URL parse, hardcoded default."""
    try:
        if hasattr(Data, 'script_name') and Data.script_name:
            sn = str(Data.script_name).strip()
            if sn:
                return sn
    except:
        pass
    try:
        url = str(getattr(model, 'URL', '') or '')
        m = re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return DC_SCRIPT_ID

model.Header = 'Registration Report Builder v' + APP_VERSION

# ============================================================================
# UNICODE / ENCODING HELPERS  (IronPython 2.7 safe)
# ============================================================================

def _to_ascii(s):
    """Transliterate a unicode string to pure ASCII.
    IronPython's json.dumps fails on non-ASCII unicode chars, so we must
    convert them to ASCII equivalents (e.g. n->n, e->e) to preserve readability."""
    result = []
    for c in s:
        o = ord(c)
        if o < 128:
            result.append(c)
        else:
            result.append(_LATIN_TO_ASCII.get(o, '?'))
    return ''.join(result)

# Common Latin-1/CP1252 character to ASCII mapping
_LATIN_TO_ASCII = {
    0xc0: 'A', 0xc1: 'A', 0xc2: 'A', 0xc3: 'A', 0xc4: 'A', 0xc5: 'A',
    0xc6: 'AE', 0xc7: 'C', 0xc8: 'E', 0xc9: 'E', 0xca: 'E', 0xcb: 'E',
    0xcc: 'I', 0xcd: 'I', 0xce: 'I', 0xcf: 'I', 0xd0: 'D', 0xd1: 'N',
    0xd2: 'O', 0xd3: 'O', 0xd4: 'O', 0xd5: 'O', 0xd6: 'O', 0xd8: 'O',
    0xd9: 'U', 0xda: 'U', 0xdb: 'U', 0xdc: 'U', 0xdd: 'Y', 0xdf: 'ss',
    0xe0: 'a', 0xe1: 'a', 0xe2: 'a', 0xe3: 'a', 0xe4: 'a', 0xe5: 'a',
    0xe6: 'ae', 0xe7: 'c', 0xe8: 'e', 0xe9: 'e', 0xea: 'e', 0xeb: 'e',
    0xec: 'i', 0xed: 'i', 0xee: 'i', 0xef: 'i', 0xf0: 'd', 0xf1: 'n',
    0xf2: 'o', 0xf3: 'o', 0xf4: 'o', 0xf5: 'o', 0xf6: 'o', 0xf8: 'o',
    0xf9: 'u', 0xfa: 'u', 0xfb: 'u', 0xfc: 'u', 0xfd: 'y', 0xff: 'y',
    0x2018: "'", 0x2019: "'", 0x201c: '"', 0x201d: '"',
    0x2013: '-', 0x2014: '-', 0x2026: '...', 0xa0: ' ',
}

def safe_str(val):
    """Safely convert any value to a pure-ASCII JSON-serializable string.
    IronPython's json.dumps crashes on non-ASCII unicode chars (even valid ones
    like U+00F1 n), so all output must be ASCII. Latin characters are
    transliterated (n->n, e->e) to keep names readable."""
    if val is None:
        return ''
    # 1. Already a Python unicode string - transliterate to ASCII
    try:
        if isinstance(val, unicode):
            return _to_ascii(val)
    except NameError:
        pass
    # 2. Try unicode() - handles .NET System.String values
    try:
        return _to_ascii(unicode(val))
    except:
        pass
    # 3. Try str() then decode to unicode, then to ASCII
    try:
        s = str(val)
        try:
            return _to_ascii(s.decode('utf-8'))
        except:
            pass
        try:
            return _to_ascii(s.decode('latin-1'))
        except:
            pass
        try:
            return _to_ascii(s.decode('cp1252'))
        except:
            pass
        # Byte-by-byte ASCII fallback
        return ''.join(c if ord(c) < 128 else '?' for c in s)
    except:
        pass
    # 4. repr() as absolute last resort
    try:
        return repr(val)
    except:
        return ''

def sanitize_for_json(obj):
    """Recursively walk an object and ensure every value is safe for json.dumps."""
    if obj is None:
        return None
    if isinstance(obj, bool):
        return obj
    if isinstance(obj, (int, float)):
        return obj
    try:
        if isinstance(obj, long):
            return obj
    except NameError:
        pass
    if isinstance(obj, dict):
        return {sanitize_for_json(k): sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize_for_json(item) for item in obj]
    return safe_str(obj)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def clean_answer_value(val, preserve_newlines=False):
    """Clean a registration answer value for display.
    - Strips surrounding quotes
    - Parses JSON arrays into comma-separated text
    - Converts literal \\n to <br> if preserve_newlines is True, otherwise strips them"""
    if not val:
        return ''
    s = safe_str(val).strip()

    # Try to parse as JSON array (checkbox answers like ["A","B","C"])
    if s.startswith('[') and s.endswith(']'):
        try:
            items = json.loads(s)
            if isinstance(items, list):
                cleaned = [safe_str(item).strip().strip('"') for item in items]
                return html_escape(', '.join(cleaned))
        except:
            pass

    # Strip surrounding quotes
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        s = s[1:-1]

    # Handle literal \n in answer text
    s = html_escape(s)
    if preserve_newlines:
        s = s.replace('\\n', '<br>')
    else:
        s = s.replace('\\n', ' ')

    return s

def html_escape(text):
    """Escape HTML special characters"""
    if not text:
        return ''
    s = safe_str(text)
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
                {"fieldId": "fld_18", "fieldType": "medical", "sourceField": "AuthorizedCheckout", "label": "Authorized Checkout", "displayFormat": "block", "order": 2, "visible": True, "colSpan": 1},
                {"fieldId": "fld_19", "fieldType": "medical", "sourceField": "Medications", "label": "Approved Medications", "displayFormat": "block", "order": 3, "visible": True, "colSpan": 1},
                {"fieldId": "fld_20", "fieldType": "person", "sourceField": "SubGroups", "label": "Subgroups", "displayFormat": "block", "order": 4, "visible": True, "colSpan": 1}
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

def build_meds_map(people_ids, people_list):
    """Build {peopleId: [medication names]} from three sources, deduped case-insensitively.

    Sources (all queried; merged, not either/or):
      1. dbo.PersonMedication -> lookup.Medication  (canonical, current)
      2. RegAnswer matched by question label (covers per-org variants of the OTC question)
      3. Legacy RecReg.Tylenol/Advil/Maalox/Robitussin booleans on people_list (carries _Xxx keys)
    """
    if not people_ids:
        return {}
    pid_csv = ','.join(str(pid) for pid in people_ids)

    # peopleId -> {'order': [names], 'seen': set(lowercase)}
    bucket = {}
    def _add(pid, name):
        n = (name or '').strip().strip('"').strip("'").strip()
        if not n:
            return
        if pid not in bucket:
            bucket[pid] = {'order': [], 'seen': set()}
        k = n.lower()
        if k in bucket[pid]['seen']:
            return
        bucket[pid]['seen'].add(k)
        bucket[pid]['order'].append(n)

    # Source 1: PersonMedication
    try:
        pm_sql = """
            SELECT pm.PeopleId, md.Description AS Medication
            FROM dbo.PersonMedication pm
            JOIN lookup.Medication md ON md.Id = pm.MedicationId
            WHERE pm.PeopleId IN ({0})
              AND md.Description IS NOT NULL
            ORDER BY pm.PeopleId, md.Description
        """.format(pid_csv)
        for r in q.QuerySql(pm_sql):
            _add(int(r.PeopleId), safe_str(r.Medication))
    except:
        pass

    # Source 2: RegAnswer (label match across both old hardcoded GUID and newer per-org variants)
    try:
        ra_sql = """
            SELECT rp.PeopleId, ra.AnswerValue
            FROM RegAnswer ra
            INNER JOIN RegPeople rp ON rp.RegPeopleId = ra.RegPeopleId
            INNER JOIN RegQuestion rq ON rq.RegQuestionId = ra.RegQuestionId
            WHERE rp.PeopleId IN ({0})
              AND (
                    ra.RegQuestionId = '8A9F1199-6F1C-480C-8A2D-146EEEAE55B8'
                 OR rq.Label LIKE '%Approved%over-the-counter%'
                 OR rq.Label LIKE '%Approved Over-the-Counter%'
                 OR rq.Label LIKE '%OTC medication%'
              )
              AND ra.AnswerValue IS NOT NULL
              AND LTRIM(RTRIM(CAST(ra.AnswerValue AS NVARCHAR(MAX)))) <> ''
            ORDER BY rp.PeopleId, rp.RegPeopleId DESC
        """.format(pid_csv)
        for r in q.QuerySql(ra_sql):
            pid = int(r.PeopleId)
            av = safe_str(r.AnswerValue)
            if av.startswith('[') and av.endswith(']'):
                av = av[1:-1]
            for med in av.split(','):
                _add(pid, med)
    except:
        pass

    # Source 3: legacy RecReg flags (already loaded onto each person row via _Xxx keys)
    for p in people_list or []:
        try:
            pid = int(p.get('PeopleId') or 0)
        except:
            continue
        if not pid:
            continue
        if p.get('_Tylenol'): _add(pid, 'Tylenol')
        if p.get('_Advil'): _add(pid, 'Advil/Ibuprofen')
        if p.get('_Maalox'): _add(pid, 'Maalox')
        if p.get('_Robitussin'): _add(pid, 'Robitussin')

    return {pid: bucket[pid]['order'] for pid in bucket}


def get_registrant_data(org_id, filter_people_ids=None, include_dropped=False):
    """Get all registrant data for an org in batch queries.
    If filter_people_ids is set, only include those people (Blue Toolbar filter).
    If include_dropped is True, also pull people who have a RegPeople row for this
    org but are no longer in OrganizationMembers (their registration answers
    persist after they're dropped/moved). Dropped people get IsDropped=True."""
    result = {'people': [], 'questions': [], 'familyData': {}}

    filter_clause = ""
    if filter_people_ids:
        filter_clause = " AND p.PeopleId IN ({0})".format(','.join(str(int(pid)) for pid in filter_people_ids))

    # Query 1: Get all org members with person + medical data.
    # When include_dropped, UNION in people who answered registration questions
    # for this org but are no longer current members. We wrap the UNION in a
    # subquery so the outer ORDER BY applies to the combined result.
    if include_dropped:
        people_sql = """
            SELECT * FROM (
                SELECT DISTINCT
                    p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
                    p.BDate, p.Age, p.GenderId, p.EmailAddress, p.CellPhone, p.HomePhone,
                    p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
                    p.FamilyId, ms.Description as MaritalStatus, ISNULL(p.CustodyIssue, 0) as CustodyIssue,
                    rr.MedicalDescription, rr.MedAllergy, rr.emcontact, rr.emphone,
                    rr.doctor, rr.docphone, rr.insurance, rr.policy,
                    ISNULL(rr.Tylenol, 0) as Tylenol, ISNULL(rr.Advil, 0) as Advil,
                    ISNULL(rr.Maalox, 0) as Maalox, ISNULL(rr.Robitussin, 0) as Robitussin,
                    pic.SmallUrl as PhotoUrl,
                    CAST(0 AS BIT) as IsDropped
                FROM OrganizationMembers om
                JOIN People p ON om.PeopleId = p.PeopleId
                LEFT JOIN lookup.MaritalStatus ms ON p.MaritalStatusId = ms.Id
                LEFT JOIN RecReg rr ON rr.PeopleId = p.PeopleId
                LEFT JOIN Picture pic ON pic.PictureId = p.PictureId
                WHERE om.OrganizationId = {0}
                {1}

                UNION

                SELECT DISTINCT
                    p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
                    p.BDate, p.Age, p.GenderId, p.EmailAddress, p.CellPhone, p.HomePhone,
                    p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
                    p.FamilyId, ms.Description as MaritalStatus, ISNULL(p.CustodyIssue, 0) as CustodyIssue,
                    rr.MedicalDescription, rr.MedAllergy, rr.emcontact, rr.emphone,
                    rr.doctor, rr.docphone, rr.insurance, rr.policy,
                    ISNULL(rr.Tylenol, 0) as Tylenol, ISNULL(rr.Advil, 0) as Advil,
                    ISNULL(rr.Maalox, 0) as Maalox, ISNULL(rr.Robitussin, 0) as Robitussin,
                    pic.SmallUrl as PhotoUrl,
                    CAST(1 AS BIT) as IsDropped
                FROM Registration r WITH (NOLOCK)
                JOIN RegPeople rp WITH (NOLOCK) ON r.RegistrationId = rp.RegistrationId
                JOIN People p ON rp.PeopleId = p.PeopleId
                LEFT JOIN lookup.MaritalStatus ms ON p.MaritalStatusId = ms.Id
                LEFT JOIN RecReg rr ON rr.PeopleId = p.PeopleId
                LEFT JOIN Picture pic ON pic.PictureId = p.PictureId
                WHERE r.OrganizationId = {0}
                  AND NOT EXISTS (
                      SELECT 1 FROM OrganizationMembers om2
                      WHERE om2.OrganizationId = {0} AND om2.PeopleId = p.PeopleId
                  )
                {1}
            ) merged
            ORDER BY Name2
        """.format(int(org_id), filter_clause)
    else:
        people_sql = """
            SELECT DISTINCT
                p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
                p.BDate, p.Age, p.GenderId, p.EmailAddress, p.CellPhone, p.HomePhone,
                p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
                p.FamilyId, ms.Description as MaritalStatus, ISNULL(p.CustodyIssue, 0) as CustodyIssue,
                rr.MedicalDescription, rr.MedAllergy, rr.emcontact, rr.emphone,
                rr.doctor, rr.docphone, rr.insurance, rr.policy,
                ISNULL(rr.Tylenol, 0) as Tylenol, ISNULL(rr.Advil, 0) as Advil,
                ISNULL(rr.Maalox, 0) as Maalox, ISNULL(rr.Robitussin, 0) as Robitussin,
                pic.SmallUrl as PhotoUrl,
                CAST(0 AS BIT) as IsDropped
            FROM OrganizationMembers om
            JOIN People p ON om.PeopleId = p.PeopleId
            LEFT JOIN lookup.MaritalStatus ms ON p.MaritalStatusId = ms.Id
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
                'Name': safe_str(r.Name2),
                'FirstName': safe_str(r.FirstName),
                'LastName': safe_str(r.LastName),
                'NickName': safe_str(r.NickName),
                'BDate': fmt_date(r.BDate),
                'Age': str(r.Age) if r.Age else '',
                'Gender': 'Male' if r.GenderId == 1 else 'Female' if r.GenderId == 2 else '',
                'MaritalStatus': safe_str(r.MaritalStatus) if r.MaritalStatus else '',
                'EmailAddress': safe_str(r.EmailAddress),
                'CellPhone': fmt_phone(r.CellPhone),
                'HomePhone': fmt_phone(r.HomePhone),
                'PrimaryAddress': safe_str(r.PrimaryAddress),
                'PrimaryCity': safe_str(r.PrimaryCity),
                'PrimaryState': safe_str(r.PrimaryState),
                'PrimaryZip': safe_str(r.PrimaryZip),
                'FullAddress': '',
                'FamilyId': r.FamilyId,
                'MedicalDescription': safe_str(r.MedicalDescription),
                'MedAllergy': safe_str(r.MedAllergy),
                'emcontact': safe_str(r.emcontact),
                'emphone': fmt_phone(r.emphone),
                'doctor': safe_str(r.doctor),
                'docphone': fmt_phone(r.docphone),
                'insurance': safe_str(r.insurance),
                'policy': safe_str(r.policy),
                'PhotoUrl': safe_str(r.PhotoUrl),
                'CustodyIssue': r.CustodyIssue,
                'IsDropped': bool(getattr(r, 'IsDropped', False)),
                'answers': {},
                '_Tylenol': r.Tylenol,
                '_Advil': r.Advil,
                '_Maalox': r.Maalox,
                '_Robitussin': r.Robitussin
            }
            # Build full address
            addr_parts = []
            if r.PrimaryAddress:
                addr_parts.append(safe_str(r.PrimaryAddress))
            city_state = []
            if r.PrimaryCity:
                city_state.append(safe_str(r.PrimaryCity))
            if r.PrimaryState:
                city_state.append(safe_str(r.PrimaryState))
            if city_state:
                addr_parts.append(', '.join(city_state))
            if r.PrimaryZip:
                addr_parts.append(safe_str(r.PrimaryZip))
            person['FullAddress'] = ', '.join(addr_parts) if addr_parts else ''

            people_list.append(person)
            people_ids.append(r.PeopleId)
            if r.FamilyId:
                family_ids.add(r.FamilyId)
    except Exception as e:
        result['error'] = 'Error loading people: ' + safe_str(e)
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
    question_id_set = set()   # track by RegQuestionId (unique per question)
    question_id_to_key = {}   # RegQuestionId -> answer key
    label_count = {}          # track label occurrences for duplicate detection
    # question_set tracks question keys already added to `questions` so the
    # legacy XML fallback below can dedupe against new-system entries by label.
    # Without it, the fallback NameErrors on `question_set` and the except below
    # silently swallows the error -> 0 questions for legacy-XML orgs.
    question_set = set()

    answers_sql = """
        SELECT
            rp.PeopleId,
            rq.RegQuestionId,
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
            q_text = safe_str(r.Question)
            q_id = r.RegQuestionId

            # First time seeing this RegQuestionId — register it
            if q_id not in question_id_set:
                question_id_set.add(q_id)
                # Check if this label was already used by a different question
                count = label_count.get(q_text, 0)
                label_count[q_text] = count + 1
                if count == 0:
                    # First question with this label — use plain label as key (backward compatible)
                    q_key = q_text
                else:
                    # Duplicate label — append occurrence number to distinguish
                    q_key = '{0} ({1})'.format(q_text, count + 1)
                question_id_to_key[q_id] = q_key
                questions.append({'key': q_key, 'label': q_text})
                question_set.add(q_key)

            # Look up the key for this question and store the answer
            q_key = question_id_to_key.get(q_id, q_text)
            if r.PeopleId in people_map:
                people_map[r.PeopleId]['answers'][q_key] = safe_str(r.Answer)
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
                            questions.append({'key': qtext, 'label': qtext})
                            question_set.add(qtext)
                # Text
                for match in re.finditer(r'<Text[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</Text>', person_xml):
                    qtext, ans = match.group(1), match.group(2).strip()
                    if qtext and qtext not in p['answers']:
                        p['answers'][qtext] = ans
                        if qtext not in question_set:
                            questions.append({'key': qtext, 'label': qtext})
                            question_set.add(qtext)
                # YesNoQuestion
                for match in re.finditer(r'<YesNoQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</YesNoQuestion>', person_xml):
                    qtext, ans_val = match.group(1), match.group(2).strip()
                    if qtext and qtext not in p['answers']:
                        ans = 'Yes' if ans_val == 'True' else 'No' if ans_val == 'False' else ans_val
                        p['answers'][qtext] = ans
                        if qtext not in question_set:
                            questions.append({'key': qtext, 'label': qtext})
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
                    'name': safe_str(r.FirstName) + ' ' + safe_str(r.LastName),
                    'phone': fmt_phone(r.CellPhone),
                    'email': safe_str(r.EmailAddress)
                })
        except:
            pass

    # Query 4: Authorized checkout people
    checkout_map = {}
    try:
        checkout_sql = """
            SELECT ac.PeopleId, p.Name2
            FROM PeopleAuthorizedCheckOut ac
            JOIN People p ON p.PeopleId = ac.AuthorizedPeopleId
            WHERE ac.PeopleId IN ({0})
            ORDER BY ac.PeopleId, p.Name2
        """.format(','.join(str(pid) for pid in people_ids))
        for r in q.QuerySql(checkout_sql):
            pid = r.PeopleId
            if pid not in checkout_map:
                checkout_map[pid] = []
            checkout_map[pid].append(safe_str(r.Name2))
    except:
        pass

    # Query 5: Approved medications (PersonMedication + RegAnswer + RecReg legacy fallback)
    meds_map = build_meds_map(people_ids, people_list)

    # Query 6: Subgroups in this org (MemberTags + OrgMemMemTags). Org-scoped, since
    # the same subgroup name can mean different things in different involvements.
    subgroups_map = {}
    try:
        sg_sql = """
            SELECT ommt.PeopleId, mt.Name
            FROM OrgMemMemTags ommt
            INNER JOIN MemberTags mt ON mt.Id = ommt.MemberTagId
            WHERE ommt.OrgId = {0}
              AND ommt.PeopleId IN ({1})
              AND mt.Name IS NOT NULL
            ORDER BY ommt.PeopleId, mt.Name
        """.format(int(org_id), ','.join(str(pid) for pid in people_ids))
        for r in q.QuerySql(sg_sql):
            pid = int(r.PeopleId)
            name = safe_str(r.Name).strip()
            if not name:
                continue
            if pid not in subgroups_map:
                subgroups_map[pid] = []
            if name not in subgroups_map[pid]:
                subgroups_map[pid].append(name)
    except:
        pass

    # Attach parent data, authorized checkout, and medications to each person
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

        # Authorized checkout
        co_names = checkout_map.get(p['PeopleId'], [])
        p['AuthorizedCheckout'] = ', '.join(co_names) if co_names else ''

        # Approved medications - merged from all three sources with case-insensitive dedup
        med_names = meds_map.get(int(p['PeopleId']), [])
        p['Medications'] = ', '.join(med_names) if med_names else ''

        # Subgroups in this org (org-scoped — subgroup names are unique per involvement)
        sg_names = subgroups_map.get(int(p['PeopleId']), [])
        p['_SubGroupsList'] = sg_names
        p['SubGroups'] = ', '.join(sg_names) if sg_names else ''

    # All subgroups configured for this involvement (from MemberTags, not derived from
    # assignments) — so the picker shows every option even when nobody is in some of them.
    all_subgroups = []
    try:
        sg_all_sql = """
            SELECT DISTINCT Name
            FROM MemberTags
            WHERE OrgId = {0}
              AND Name IS NOT NULL
              AND LTRIM(RTRIM(Name)) <> ''
            ORDER BY Name
        """.format(int(org_id))
        for r in q.QuerySql(sg_all_sql):
            all_subgroups.append(safe_str(r.Name).strip())
    except:
        pass
    result['people'] = people_list
    result['questions'] = questions
    result['familyData'] = family_data
    result['availableSubGroups'] = all_subgroups
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
            p.FamilyId, ms.Description as MaritalStatus, ISNULL(p.CustodyIssue, 0) as CustodyIssue,
            rr.MedicalDescription, rr.MedAllergy, rr.emcontact, rr.emphone,
            rr.doctor, rr.docphone, rr.insurance, rr.policy,
            ISNULL(rr.Tylenol, 0) as Tylenol, ISNULL(rr.Advil, 0) as Advil,
            ISNULL(rr.Maalox, 0) as Maalox, ISNULL(rr.Robitussin, 0) as Robitussin,
            pic.SmallUrl as PhotoUrl
        FROM People p
        LEFT JOIN lookup.MaritalStatus ms ON p.MaritalStatusId = ms.Id
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
                'Name': safe_str(r.Name2),
                'FirstName': safe_str(r.FirstName),
                'LastName': safe_str(r.LastName),
                'NickName': safe_str(r.NickName),
                'BDate': fmt_date(r.BDate),
                'Age': str(r.Age) if r.Age else '',
                'Gender': 'Male' if r.GenderId == 1 else 'Female' if r.GenderId == 2 else '',
                'MaritalStatus': safe_str(r.MaritalStatus) if r.MaritalStatus else '',
                'EmailAddress': safe_str(r.EmailAddress),
                'CellPhone': fmt_phone(r.CellPhone),
                'HomePhone': fmt_phone(r.HomePhone),
                'PrimaryAddress': safe_str(r.PrimaryAddress),
                'PrimaryCity': safe_str(r.PrimaryCity),
                'PrimaryState': safe_str(r.PrimaryState),
                'PrimaryZip': safe_str(r.PrimaryZip),
                'FullAddress': '',
                'FamilyId': r.FamilyId,
                'MedicalDescription': safe_str(r.MedicalDescription),
                'MedAllergy': safe_str(r.MedAllergy),
                'emcontact': safe_str(r.emcontact),
                'emphone': fmt_phone(r.emphone),
                'doctor': safe_str(r.doctor),
                'docphone': fmt_phone(r.docphone),
                'insurance': safe_str(r.insurance),
                'policy': safe_str(r.policy),
                'PhotoUrl': safe_str(r.PhotoUrl),
                'CustodyIssue': r.CustodyIssue,
                'answers': {},
                '_Tylenol': r.Tylenol,
                '_Advil': r.Advil,
                '_Maalox': r.Maalox,
                '_Robitussin': r.Robitussin
            }
            addr_parts = []
            if r.PrimaryAddress:
                addr_parts.append(safe_str(r.PrimaryAddress))
            city_state = []
            if r.PrimaryCity:
                city_state.append(safe_str(r.PrimaryCity))
            if r.PrimaryState:
                city_state.append(safe_str(r.PrimaryState))
            if city_state:
                addr_parts.append(', '.join(city_state))
            if r.PrimaryZip:
                addr_parts.append(safe_str(r.PrimaryZip))
            person['FullAddress'] = ', '.join(addr_parts) if addr_parts else ''

            people_list.append(person)
            if r.FamilyId:
                family_ids.add(r.FamilyId)
    except Exception as e:
        result['error'] = 'Error loading people: ' + safe_str(e)
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
                    'name': safe_str(r.FirstName) + ' ' + safe_str(r.LastName),
                    'phone': fmt_phone(r.CellPhone),
                    'email': safe_str(r.EmailAddress)
                })
        except:
            pass

    # Authorized checkout people
    checkout_map = {}
    try:
        checkout_sql = """
            SELECT ac.PeopleId, p.Name2
            FROM PeopleAuthorizedCheckOut ac
            JOIN People p ON p.PeopleId = ac.AuthorizedPeopleId
            WHERE ac.PeopleId IN ({0})
            ORDER BY ac.PeopleId, p.Name2
        """.format(pid_list)
        for r in q.QuerySql(checkout_sql):
            pid = r.PeopleId
            if pid not in checkout_map:
                checkout_map[pid] = []
            checkout_map[pid].append(safe_str(r.Name2))
    except:
        pass

    # Approved medications (PersonMedication + RegAnswer + RecReg legacy fallback)
    meds_map = build_meds_map(people_ids, people_list)

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

        # Authorized checkout
        co_names = checkout_map.get(p['PeopleId'], [])
        p['AuthorizedCheckout'] = ', '.join(co_names) if co_names else ''

        # Approved medications - merged from all three sources with case-insensitive dedup
        med_names = meds_map.get(int(p['PeopleId']), [])
        p['Medications'] = ', '.join(med_names) if med_names else ''

        # No org context available here — subgroups can't be resolved (they're per-org)
        p['SubGroups'] = ''

    result['people'] = people_list
    result['familyData'] = family_data
    return result

# ============================================================================
# REPORT GENERATION ENGINE
# ============================================================================

def get_field_value(person, field, preserve_newlines=False):
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
        if src == 'SubGroups':
            sg_list = person.get('_SubGroupsList', None)
            if sg_list is None:
                return html_escape(person.get('SubGroups', ''))
            sg_filter = field.get('subgroupFilter') or []
            if sg_filter:
                allowed_lc = set(s.lower() for s in sg_filter if s)
                shown = [s for s in sg_list if s.lower() in allowed_lc]
            else:
                shown = sg_list
            return html_escape(', '.join(shown))
        return html_escape(person.get(src, ''))

    elif ft == 'family':
        return html_escape(person.get(src, ''))

    elif ft == 'medical':
        val = person.get(src, '')
        if src == 'MedAllergy':
            # MedAllergy is a boolean in RecReg - show MedicalDescription text instead
            if str(val).lower().strip() in ['true', '1']:
                desc = person.get('MedicalDescription', '')
                return html_escape(desc) if desc else ''
            if str(val).lower().strip() in ['false', '0']:
                return ''
        if src == 'CustodyIssue':
            if str(val).lower().strip() in ['true', '1']:
                return 'Yes'
            return ''
        if str(val).lower().strip() in ['false', '0']:
            return ''
        return html_escape(val)

    elif ft == 'regquestion':
        raw = person.get('answers', {}).get(src, '')
        return clean_answer_value(raw, preserve_newlines)

    elif ft == 'extravalue':
        return html_escape(person.get('_ev_' + src, ''))

    return ''

def extract_ev_names_from_template(template):
    """Scan template sections AND summary items for extravalue fields and
    return a set of EV names (so both render paths fetch what they need)."""
    ev_names = set()
    for sec in template.get('sections', []):
        for f in sec.get('fields', []):
            if f.get('fieldType') == 'extravalue' and f.get('sourceField'):
                ev_names.add(f['sourceField'])
    ps = template.get('printSettings', {}) or {}
    for it in (ps.get('summaryItems', []) or []):
        if it.get('itemType') == 'extravalue' and it.get('field'):
            ev_names.add(it['field'])
    return ev_names

def fetch_extra_values(people_list, ev_names):
    """Query PeopleExtra for the given field names and attach to each person dict.
    Values are stored as person['_ev_FieldName'] keyed by the template name (preserving case)."""
    if not people_list or not ev_names:
        return
    pids = [p['PeopleId'] for p in people_list]
    pid_map = {}
    for p in people_list:
        pid_map[p['PeopleId']] = p
        for ev in ev_names:
            p['_ev_' + ev] = ''

    # Build a case-insensitive lookup: lowercase DB field -> template field name
    ev_lower_map = {}
    for ev in ev_names:
        ev_lower_map[ev.lower()] = ev

    safe_names = ["'" + n.replace("'", "''") + "'" for n in ev_names]
    ev_sql = """
        SELECT PeopleId, Field, StrValue, Data,
               CAST(DateValue AS NVARCHAR(50)) as DateVal,
               IntValue, BitValue
        FROM PeopleExtra
        WHERE PeopleId IN ({0})
          AND (Field IN ({1}) OR LOWER(Field) IN ({2}))
    """.format(
        ','.join(str(pid) for pid in pids),
        ','.join(safe_names),
        ','.join(["'" + n.replace("'", "''").lower() + "'" for n in ev_names])
    )
    try:
        for r in q.QuerySql(ev_sql):
            if r.PeopleId not in pid_map:
                continue
            # Determine value from whichever column has data
            val = None
            if r.StrValue is not None and str(r.StrValue).strip():
                val = safe_str(r.StrValue)
            elif r.Data is not None and str(r.Data).strip():
                val = safe_str(r.Data)
            elif r.DateVal is not None and str(r.DateVal).strip():
                val = safe_str(r.DateVal)
            elif r.IntValue is not None:
                val = str(r.IntValue)
            elif r.BitValue is not None:
                val = 'Yes' if r.BitValue else 'No'
            if val:
                # Match back to template field name via case-insensitive lookup
                field_name = safe_str(r.Field)
                template_key = ev_lower_map.get(field_name.lower(), field_name)
                pid_map[r.PeopleId]['_ev_' + template_key] = val
    except:
        pass

def auto_populate_section_fields(section, questions):
    """Return the section's fields list with auto-population applied.

    Used when a template section has an empty fields list AND a title that
    looks like a reg-question section (contains "question", "answer", or
    "registration"). In that case we synthesize one regquestion field per
    available question so the section renders meaningfully.

    Shared by the PDF render path and the CSV export so both paths emit
    the same set of columns.
    """
    fields = sorted(section.get('fields', []) or [], key=lambda f: f.get('order', 0))
    if fields:
        return fields
    title_lower = (section.get('title', '') or '').lower()
    if 'question' not in title_lower and 'answer' not in title_lower and 'registration' not in title_lower:
        return fields
    auto_fields = []
    for qi, qitem in enumerate(questions or []):
        q_key = qitem['key'] if isinstance(qitem, dict) else qitem
        q_label = qitem['label'] if isinstance(qitem, dict) else qitem
        auto_fields.append({
            'fieldId': 'auto_q_' + str(qi),
            'fieldType': 'regquestion',
            'sourceField': q_key,
            'label': q_label,
            'displayFormat': 'block',
            'order': qi + 1,
            'visible': True,
            'colSpan': 1
        })
    return auto_fields


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
    preserve_newlines = opts.get('preserveNewlines', False)

    sorted_sections = sorted(sections, key=lambda s: s.get('order', 0))

    if single_person_id:
        pid = safe_int(single_person_id)
        people = [p for p in people if p['PeopleId'] == pid]

    # Custom document title overrides the org-name slot wherever it
    # appears in the report (per-person header, cover page subtitle).
    # Empty title -> fall back to org_name. This is what the user means
    # by "I named my report, it should appear in my report."
    custom_title = ''
    try:
        ct = template.get('reportTitle') if isinstance(template, dict) else None
        custom_title = (ct or '').strip() if ct else ''
    except:
        custom_title = ''
    effective_title = custom_title if custom_title else org_name

    parts = []
    report_class = 'rr-report' + (' rr-compact' if compact_rows else '')
    parts.append('<div class="{0}" style="font-family: {1};">'.format(report_class, html_escape(font_family)))

    # Header strategy depends on one_per_page:
    #   ON  -> show banner above EACH person (every printed page gets a
    #          header). This was the original behavior.
    #   OFF -> show banner ONCE at the top; repeating it above each
    #          person on a continuous page is just noise. This matches
    #          what feels natural when viewing the report on screen.
    show_top_banner = bool(custom_title) and (not one_per_page)
    show_perperson_banner = (show_org_header and effective_title
                             and (one_per_page or not custom_title))

    if show_top_banner:
        parts.append('<h1 style="font-size:24px;color:{0};text-align:center;margin:0 0 16px;padding-bottom:8px;border-bottom:2px solid {0};">{1}</h1>'.format(
            html_escape(heading_color), html_escape(custom_title)))

    for idx, person in enumerate(people):
        page_class = 'rr-person-page'
        if one_per_page and idx < len(people) - 1:
            page_class += ' rr-page-break'

        parts.append('<div class="{0}">'.format(page_class))

        if show_perperson_banner:
            parts.append('<div class="rr-org-header" style="border-bottom-color: {0};">'.format(html_escape(heading_color)))
            parts.append('<h2 style="color: {0};">{1}</h2>'.format(html_escape(heading_color), html_escape(effective_title)))
            parts.append('</div>')

        parts.append('<div class="rr-person-header">')
        if show_photo and person.get('PhotoUrl'):
            parts.append('<img class="rr-photo" src="{0}" alt="Photo" onerror="this.style.display=\'none\'">'.format(html_escape(person['PhotoUrl'])))
        dropped_badge = ''
        if person.get('IsDropped'):
            # Inline styles so the badge survives popup-window printing
            # (TouchPoint page CSS doesn't reach the print popup).
            dropped_badge = (' <span style="display:inline-block;margin-left:8px;padding:2px 8px;'
                             'background:#fed7d7;color:#9b2c2c;border:1px solid #fc8181;'
                             'border-radius:10px;font-size:11px;font-weight:700;vertical-align:middle;'
                             '-webkit-print-color-adjust:exact;print-color-adjust:exact;">'
                             'DROPPED</span>')
        parts.append('<div class="rr-person-name" style="color: {0};">{1}{2}</div>'.format(
            html_escape(heading_color), html_escape(person.get('Name', '')), dropped_badge))
        parts.append('</div>')

        for sec in sorted_sections:
            if not sec.get('visible', True):
                continue

            sec_header_color = sec.get('headerColor', heading_color)
            layout = sec.get('layout', 'single-column')
            fields = auto_populate_section_fields(sec, questions)

            if not fields:
                continue

            # Check if section has any visible values
            if hide_empty:
                has_values = False
                for f in fields:
                    if not f.get('visible', True):
                        continue
                    val = get_field_value(person, f, preserve_newlines)
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

                val = get_field_value(person, f, preserve_newlines)
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
# SUPPLEMENTAL PAGE RENDERERS (Cover, Missing Info, Medical Concerns)
# ============================================================================

# Exclude list for medical values (common non-answers)
_EXCLUDE_MEDICAL = [
    'ma', 'mon', 'mone', 'no ', 'non', 'non ', 'nine',
    'none', 'none.', 'none ', 'none. ', 'none know', 'none know ', 'no allergies', 'no alleriges',
    'no concerns', 'none known', 'none known ', 'none \\nnone', 'nonr', 'nome',
    'no food allergies ', 'null', 'n/a', 'n/', 'n_a',
    'nkda', 'kna', 'na', 'na ', 'nka', 'n-a', 'ns', 'n/s', 'no',
    'no food allergies', '5', ''
]

def _is_meaningful_medical(val):
    """Check if a medical value is meaningful (not a common non-answer)."""
    if not val:
        return False
    return val.strip().lower() not in _EXCLUDE_MEDICAL

# =====================================================================
# Form Summary page (Wufoo / MS-Forms style aggregate view)
# =====================================================================
# PROTOTYPE (Phase 1 slice): a per-involvement aggregate summary --
# demographics strip + per-question answer tallies -- prepended like the
# cover page and gated by printSettings.showSummaryPage. Uses pure
# HTML/CSS horizontal bars (no Chart.js dependency) so it renders the
# same in the page and in the print popup, and inline background colors
# survive print-color-adjust. Operates on the already-fetched
# data['people'] / data['questions'] -- no new SQL.
AGE_BIN_ORDER = ['0-4', '5-9', '10-12', '13-17', '18-24', '25-34',
                 '35-44', '45-54', '55-64', '65+']

# Flat person fields that hold a comma-joined LIST of values (one person can
# have several). When summarized, these are split on commas and each item is
# tallied separately -- otherwise every person's whole list reads as one
# distinct value (e.g. Approved Medications). Other fields tally as-is.
SUMMARY_MULTIVALUE_FIELDS = set(['Medications', 'SubGroups', 'AuthorizedCheckout'])


def _split_answer_options(raw):
    """Split a raw registration answer into individual option strings for
    tallying. Handles JSON-array checkbox answers (["A","B"]) and plain
    single values. Returns a list (possibly empty)."""
    if raw is None:
        return []
    s = safe_str(raw).strip()
    if not s:
        return []
    if s.startswith('[') and s.endswith(']'):
        try:
            items = json.loads(s)
            if isinstance(items, list):
                out = []
                for it in items:
                    v = safe_str(it).strip().strip('"')
                    if v:
                        out.append(v)
                return out
        except:
            pass
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        s = s[1:-1]
    s = s.strip()
    return [s] if s else []


def _age_bin_label(age):
    """Bucket an age into a display label, or None if not a usable age."""
    try:
        a = int(age)
    except:
        return None
    if a < 0:
        return None
    if a <= 4: return '0-4'
    if a <= 9: return '5-9'
    if a <= 12: return '10-12'
    if a <= 17: return '13-17'
    if a <= 24: return '18-24'
    if a <= 34: return '25-34'
    if a <= 44: return '35-44'
    if a <= 54: return '45-54'
    if a <= 64: return '55-64'
    return '65+'


def _summary_bar_rows(pairs, total, bar_color):
    """pairs = ordered list of (label, count). Returns HTML rows: label +
    inline CSS bar + count and %. Pure CSS so it prints in the popup."""
    parts = []
    maxc = 0
    for _, c in pairs:
        if c > maxc:
            maxc = c
    for label, c in pairs:
        pct = (100.0 * c / total) if total else 0
        barw = (100.0 * c / maxc) if maxc else 0
        parts.append(
            '<div style="display:flex;align-items:center;margin:2px 0;font-size:12px;">'
            '<div style="width:38%;padding-right:8px;text-align:right;color:#4a5568;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{0}</div>'
            '<div style="flex:1;background:#edf2f7;border-radius:3px;height:16px;">'
            '<div style="width:{1:.1f}%;background:{2};height:16px;border-radius:3px;"></div>'
            '</div>'
            '<div style="width:90px;padding-left:8px;color:#2d3748;white-space:nowrap;">{3} <span style="color:#a0aec0;">({4:.0f}%)</span></div>'
            '</div>'.format(html_escape(label), barw, bar_color, c, pct))
    return ''.join(parts)


def _summary_block(heading_color, title_html, body_html):
    """Wrap a titled summary section. title_html may contain markup (e.g. the
    'N of M answered' span) so it is NOT escaped here -- callers escape labels."""
    return ('<div style="margin-bottom:22px;break-inside:avoid;page-break-inside:avoid;">'
            '<div style="font-size:14px;font-weight:700;color:#2d3748;border-bottom:2px solid {0};padding-bottom:4px;margin-bottom:8px;">{1}</div>'
            '{2}</div>').format(heading_color, title_html, body_html)


def _tally_person_field(people, field):
    """Count distinct non-empty values of a flat person dict field.
    Returns (counts_dict, answered_count). For known multi-value fields
    (comma-joined lists like Approved Medications), each list item is counted
    separately; answered still counts people, so a bar % reads as 'X% of
    respondents have this item'."""
    counts = {}
    answered = 0
    multi = field in SUMMARY_MULTIVALUE_FIELDS
    for p in people:
        v = safe_str(p.get(field, '')).strip()
        if not v:
            continue
        answered += 1
        if multi:
            seen = set()
            for tok in v.split(','):
                tok = tok.strip()
                if tok and tok not in seen:
                    seen.add(tok)
                    counts[tok] = counts.get(tok, 0) + 1
        else:
            counts[v] = counts.get(v, 0) + 1
    return counts, answered


def _render_summary_item(item, people, heading_color):
    """Render one CONFIGURED summary item -> an HTML block, or '' if there's
    nothing to show. item = {itemType, field, label}. itemType is one of:
      demographic  -- field Age (bins), PrimaryZip (top), or any person field tally
      subgroup     -- field '_all' (counts per involvement subgroup)
      regquestion  -- field = question key (answer tally; choice vs free-text)
      person       -- field = any flat person field (generic value tally)
    Mirrors the auto sections so customized and automatic modes look identical."""
    it = (item.get('itemType') or '').strip()
    field = (item.get('field') or '').strip()
    label = item.get('label') or field
    total = len(people)

    # Age distribution (binned)
    if it == 'demographic' and field == 'Age':
        bin_counts = {}
        have = 0
        for p in people:
            lbl = _age_bin_label(p.get('Age'))
            if lbl:
                have += 1
                bin_counts[lbl] = bin_counts.get(lbl, 0) + 1
        if not have:
            return ''
        pairs = [(b, bin_counts[b]) for b in AGE_BIN_ORDER if b in bin_counts]
        return _summary_block(heading_color, html_escape(label), _summary_bar_rows(pairs, have, '#48bb78'))

    # Top ZIP codes
    if it == 'demographic' and field in ('PrimaryZip', 'Zip'):
        zc = {}
        for p in people:
            z = safe_str(p.get('PrimaryZip')).strip()[:5]
            if z:
                zc[z] = zc.get(z, 0) + 1
        if not zc:
            return ''
        pairs = sorted(zc.items(), key=lambda kv: (-kv[1], kv[0]))[:12]
        return _summary_block(heading_color, html_escape(label), _summary_bar_rows(pairs, total, '#ed8936'))

    # Subgroup membership counts
    if it == 'subgroup':
        sc = {}
        for p in people:
            for sg in (p.get('_SubGroupsList') or []):
                sc[sg] = sc.get(sg, 0) + 1
        if not sc:
            return ''
        pairs = sorted(sc.items(), key=lambda kv: (-kv[1], kv[0]))
        return _summary_block(heading_color, html_escape(label), _summary_bar_rows(pairs, total, '#9f7aea'))

    # Registration question answer tally
    if it == 'regquestion':
        opt_counts = {}
        answered = 0
        long_text_hits = 0
        for p in people:
            optslist = _split_answer_options(p.get('answers', {}).get(field, ''))
            if not optslist:
                continue
            answered += 1
            for o in optslist:
                if len(o) > 60:
                    long_text_hits += 1
                opt_counts[o] = opt_counts.get(o, 0) + 1
        if answered == 0:
            return ''
        distinct = len(opt_counts)
        is_freetext = (distinct > 12) or (long_text_hits > 0 and distinct > 5)
        head = '{0} <span style="font-weight:400;color:#a0aec0;font-size:11px;">({1} of {2} answered)</span>'.format(
            html_escape(label), answered, total)
        if is_freetext:
            toppairs = sorted(opt_counts.items(), key=lambda kv: (-kv[1], kv[0]))[:8]
            body = '<div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">Free-text &mdash; {0} distinct responses. Most common:</div>'.format(distinct)
            body += _summary_bar_rows(toppairs, answered, '#718096')
        else:
            pairs = sorted(opt_counts.items(), key=lambda kv: (-kv[1], kv[0]))
            body = _summary_bar_rows(pairs, answered, '#4299e1')
        return _summary_block(heading_color, head, body)

    # Extra Value tally (PeopleExtra). Values were fetched onto the person dict
    # as '_ev_<Field>' by fetch_extra_values (driven by extract_ev_names).
    if it == 'extravalue':
        counts = {}
        answered = 0
        for p in people:
            v = safe_str(p.get('_ev_' + field, '')).strip()
            if not v:
                continue
            answered += 1
            counts[v] = counts.get(v, 0) + 1
        if answered == 0:
            return ''
        pairs = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
        if len(pairs) > 15:
            pairs = pairs[:15]
        head = '{0} <span style="font-weight:400;color:#a0aec0;font-size:11px;">({1} of {2})</span>'.format(
            html_escape(label), answered, total)
        return _summary_block(heading_color, head, _summary_bar_rows(pairs, answered, '#38b2ac'))

    # Generic person-field / demographic value tally (Gender, MaritalStatus,
    # State, City, or any other flat field).
    if it in ('demographic', 'person'):
        counts, answered = _tally_person_field(people, field)
        if answered == 0:
            return ''
        pairs = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
        if len(pairs) > 15:
            pairs = pairs[:15]
        head = '{0} <span style="font-weight:400;color:#a0aec0;font-size:11px;">({1} of {2})</span>'.format(
            html_escape(label), answered, total)
        return _summary_block(heading_color, head, _summary_bar_rows(pairs, answered, '#4299e1'))

    return ''


def render_summary_page(people, questions, org_name, template):
    """Aggregate summary page: KPIs + demographics + per-question tallies.
    Prepended like the cover page; gated by printSettings.showSummaryPage.
    If printSettings.summaryItems[] is set, only those items render (in order);
    otherwise an automatic set (gender/age/subgroups/ZIPs/all questions)."""
    opts = template.get('globalOptions', {})
    heading_color = opts.get('headingColor', '#2c5282')
    total = len(people)

    custom_title = ''
    try:
        ct = template.get('reportTitle') if isinstance(template, dict) else None
        custom_title = (ct or '').strip() if ct else ''
    except:
        custom_title = ''
    title_text = custom_title if custom_title else 'Registration Summary'

    import datetime
    now_str = datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')

    parts = []
    parts.append('<div class="rr-summary-page" style="page-break-after:always;padding:30px;">')
    parts.append('<h1 style="font-size:26px;color:{0};margin:0 0 4px 0;">{1}</h1>'.format(
        heading_color, html_escape(title_text)))
    if org_name:
        parts.append('<div style="font-size:16px;color:#4a5568;margin-bottom:2px;">{0}</div>'.format(html_escape(org_name)))
    parts.append('<div style="font-size:12px;color:#a0aec0;margin-bottom:18px;">Summary &bull; {0} registrants &bull; generated {1}</div>'.format(total, html_escape(now_str)))

    if total == 0:
        parts.append('<div style="color:#718096;">No registrants to summarize.</div></div>')
        return ''.join(parts)

    # KPI cards
    unique_ids = set(p.get('PeopleId') for p in people)
    male = sum(1 for p in people if p.get('Gender') == 'Male')
    female = sum(1 for p in people if p.get('Gender') == 'Female')
    unknown = total - male - female
    ages = []
    for p in people:
        try:
            if p.get('Age') not in (None, ''):
                ages.append(int(p.get('Age')))
        except:
            pass
    avg_age = (sum(ages) / float(len(ages))) if ages else None

    def kpi(label, value):
        return ('<div style="flex:1;min-width:104px;background:#f7fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px 14px;">'
                '<div style="font-size:22px;font-weight:700;color:{0};">{1}</div>'
                '<div style="font-size:11px;color:#718096;text-transform:uppercase;letter-spacing:.04em;">{2}</div>'
                '</div>').format(heading_color, value, html_escape(label))

    parts.append('<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px;">')
    parts.append(kpi('Registrants', total))
    parts.append(kpi('Unique People', len(unique_ids)))
    parts.append(kpi('Male', male))
    parts.append(kpi('Female', female))
    if unknown:
        parts.append(kpi('Unspecified', unknown))
    if avg_age is not None:
        parts.append(kpi('Avg Age', '{0:.0f}'.format(avg_age)))
    parts.append('</div>')

    # Customized mode: if the template lists explicit summary items, render
    # only those (in the admin's order) and stop. Empty list -> automatic mode.
    ps_cfg = template.get('printSettings', {}) or {}
    summary_items = ps_cfg.get('summaryItems', []) or []
    if summary_items:
        for sit in summary_items:
            blk = _render_summary_item(sit, people, heading_color)
            if blk:
                parts.append(blk)
        parts.append('</div>')
        return ''.join(parts)

    def block(title_html, body_html):
        return _summary_block(heading_color, title_html, body_html)

    # Gender
    gpairs = [('Male', male), ('Female', female)]
    if unknown:
        gpairs.append(('Unspecified', unknown))
    parts.append(block('Gender', _summary_bar_rows(gpairs, total, '#4299e1')))

    # Age bins
    bin_counts = {}
    have_age = 0
    for p in people:
        lbl = _age_bin_label(p.get('Age'))
        if lbl:
            have_age += 1
            bin_counts[lbl] = bin_counts.get(lbl, 0) + 1
    if have_age:
        agepairs = [(b, bin_counts[b]) for b in AGE_BIN_ORDER if b in bin_counts]
        parts.append(block('Age Distribution', _summary_bar_rows(agepairs, have_age, '#48bb78')))

    # Subgroups
    sg_counts = {}
    for p in people:
        for sg in (p.get('_SubGroupsList') or []):
            sg_counts[sg] = sg_counts.get(sg, 0) + 1
    if sg_counts:
        sgpairs = sorted(sg_counts.items(), key=lambda kv: (-kv[1], kv[0]))
        parts.append(block('Subgroups', _summary_bar_rows(sgpairs, total, '#9f7aea')))

    # Top ZIPs
    zip_counts = {}
    for p in people:
        z = safe_str(p.get('PrimaryZip')).strip()[:5]
        if z:
            zip_counts[z] = zip_counts.get(z, 0) + 1
    if zip_counts:
        zippairs = sorted(zip_counts.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
        parts.append(block('Top ZIP Codes', _summary_bar_rows(zippairs, total, '#ed8936')))

    # Per-question tallies
    CHOICE_MAX_DISTINCT = 12
    for qd in questions:
        key = qd.get('key')
        label = qd.get('label', key)
        opt_counts = {}
        answered = 0
        long_text_hits = 0
        for p in people:
            raw = p.get('answers', {}).get(key, '')
            optslist = _split_answer_options(raw)
            if not optslist:
                continue
            answered += 1
            for o in optslist:
                if len(o) > 60:
                    long_text_hits += 1
                opt_counts[o] = opt_counts.get(o, 0) + 1
        if answered == 0:
            continue
        distinct = len(opt_counts)
        is_freetext = (distinct > CHOICE_MAX_DISTINCT) or (long_text_hits > 0 and distinct > 5)
        head = '{0} <span style="font-weight:400;color:#a0aec0;font-size:11px;">({1} of {2} answered)</span>'.format(
            html_escape(label), answered, total)
        if is_freetext:
            toppairs = sorted(opt_counts.items(), key=lambda kv: (-kv[1], kv[0]))[:8]
            body = '<div style="font-size:11px;color:#a0aec0;margin-bottom:4px;">Free-text &mdash; {0} distinct responses. Most common:</div>'.format(distinct)
            body += _summary_bar_rows(toppairs, answered, '#718096')
        else:
            qpairs = sorted(opt_counts.items(), key=lambda kv: (-kv[1], kv[0]))
            body = _summary_bar_rows(qpairs, answered, '#4299e1')
        parts.append(block(head, body))

    parts.append('</div>')
    return ''.join(parts)


def render_cover_page(people, org_name, template):
    """Render a cover page with summary statistics."""
    opts = template.get('globalOptions', {})
    heading_color = opts.get('headingColor', '#2c5282')

    male_count = 0
    female_count = 0
    unknown_count = 0
    for p in people:
        g = p.get('Gender', '')
        if g == 'Male':
            male_count += 1
        elif g == 'Female':
            female_count += 1
        else:
            unknown_count += 1

    total = len(people)

    # Count people with medical info
    medical_count = 0
    allergy_count = 0
    for p in people:
        has_med = False
        desc = p.get('MedicalDescription', '')
        if _is_meaningful_medical(desc):
            allergy_count += 1
            has_med = True
        if p.get('Medications', ''):
            has_med = True
        if has_med:
            medical_count += 1

    import datetime
    now_str = datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')

    parts = []
    parts.append('<div class="rr-cover-page" style="page-break-after:always;padding:50px 30px;text-align:center;">')
    parts.append('<div style="position:absolute;top:20px;right:20px;color:#c53030;font-weight:bold;font-size:14px;border:2px solid #c53030;padding:4px 12px;background:#fed7d7;">CONFIDENTIAL</div>')
    # Custom title from template settings, falls back to "Registration Report".
    custom_title = ''
    try:
        ct = template.get('reportTitle') if isinstance(template, dict) else None
        custom_title = (ct or '').strip() if ct else ''
    except:
        custom_title = ''
    title_text = custom_title if custom_title else 'Registration Report'
    parts.append('<h1 style="font-size:32px;color:{0};margin-bottom:8px;">{1}</h1>'.format(
        heading_color, html_escape(title_text)))
    if org_name:
        parts.append('<h2 style="font-size:22px;color:#4a5568;margin-bottom:30px;">{0}</h2>'.format(html_escape(org_name)))

    parts.append('<div style="margin:30px auto;padding:24px;background:#f7fafc;border-radius:10px;max-width:500px;text-align:left;">')
    parts.append('<h3 style="color:#4a5568;margin:0 0 16px 0;text-align:center;border-bottom:2px solid #e2e8f0;padding-bottom:8px;">Report Summary</h3>')
    parts.append('<div style="font-size:16px;line-height:2;">')
    parts.append('<div><strong>Total Registrants:</strong> {0}</div>'.format(total))
    parts.append('<div><strong>Male:</strong> {0}</div>'.format(male_count))
    parts.append('<div><strong>Female:</strong> {0}</div>'.format(female_count))
    if unknown_count > 0:
        parts.append('<div><strong>Gender Not Specified:</strong> {0}</div>'.format(unknown_count))
    if allergy_count > 0:
        parts.append('<div style="color:#c53030;"><strong>With Allergies:</strong> {0}</div>'.format(allergy_count))
    if medical_count > 0:
        parts.append('<div><strong>With Medical Info:</strong> {0}</div>'.format(medical_count))
    parts.append('<div style="margin-top:16px;padding-top:12px;border-top:1px solid #e2e8f0;">')
    parts.append('<strong>Report Generated:</strong><br>{0}'.format(now_str))
    parts.append('</div>')
    parts.append('</div></div>')
    parts.append('</div>')
    return ''.join(parts)


def _get_item_value(person, item_cfg):
    """Get the value of a configured item from person data.
    item_cfg = {itemType, field, label}
    Returns the string value (may be empty)."""
    it = item_cfg.get('itemType', '')
    field = item_cfg.get('field', '')

    if it == 'regquestion':
        return person.get('answers', {}).get(field, '')

    # person, family, medical all stored flat on the person dict
    val = person.get(field, '')

    # Special handling for boolean-style fields
    if field == 'MedAllergy':
        raw = str(val).lower().strip()
        if raw in ['true', '1']:
            return person.get('MedicalDescription', '')
        return ''
    if field == 'CustodyIssue':
        if str(val).lower().strip() in ['true', '1']:
            return 'Yes'
        return ''

    return str(val).strip() if val else ''


def render_missing_info_page(people, template):
    """Render a page highlighting people with missing information.
    Uses printSettings.missingInfoItems[] to decide which fields to check."""
    ps = template.get('printSettings', {})
    items_config = ps.get('missingInfoItems', [])

    # Default items if none configured
    if not items_config:
        items_config = [
            {'itemType': 'medical', 'field': 'emcontact', 'label': 'Emergency Contact'},
            {'itemType': 'medical', 'field': 'insurance', 'label': 'Insurance'},
            {'itemType': 'person', 'field': 'PhotoUrl', 'label': 'Profile Photo'},
            {'itemType': 'person', 'field': 'Age', 'label': 'Age/Birthdate'},
        ]

    missing_items = []
    for p in people:
        missing_fields = []
        for item_cfg in items_config:
            label = item_cfg.get('label', item_cfg.get('field', ''))
            val = _get_item_value(p, item_cfg)
            if not val:
                missing_fields.append(label)

        if missing_fields:
            missing_items.append({
                'name': p.get('Name', ''),
                'age': p.get('Age', '') or 'N/A',
                'missing': ', '.join(missing_fields)
            })

    parts = []
    parts.append('<div class="rr-missing-page" style="page-break-after:always;padding:20px;">')
    parts.append('<div style="position:absolute;top:10px;right:20px;color:#c53030;font-weight:bold;font-size:14px;border:1px solid #c53030;padding:3px 10px;background:#fed7d7;">CONFIDENTIAL</div>')

    if missing_items:
        parts.append('<h2 style="color:#f59e0b;border-bottom:3px solid #f59e0b;padding-bottom:10px;margin-top:0;">Missing Information - Action Required</h2>')
        parts.append('<p style="color:#92400e;font-weight:bold;margin-bottom:16px;">The following {0} people have incomplete information:</p>'.format(len(missing_items)))
        parts.append('<table style="width:100%;border-collapse:collapse;">')
        parts.append('<thead><tr style="background:#fef3c7;">')
        parts.append('<th style="padding:8px;text-align:left;border:1px solid #f59e0b;">Name</th>')
        parts.append('<th style="padding:8px;text-align:center;border:1px solid #f59e0b;width:80px;">Age</th>')
        parts.append('<th style="padding:8px;text-align:left;border:1px solid #f59e0b;">Missing Information</th>')
        parts.append('</tr></thead><tbody>')

        for item in sorted(missing_items, key=lambda x: x['name']):
            parts.append('<tr>')
            parts.append('<td style="padding:6px;border:1px solid #e2e8f0;font-weight:bold;">{0}</td>'.format(html_escape(item['name'])))
            parts.append('<td style="padding:6px;border:1px solid #e2e8f0;text-align:center;">{0}</td>'.format(html_escape(item['age'])))
            parts.append('<td style="padding:6px;border:1px solid #e2e8f0;color:#92400e;font-weight:500;">{0}</td>'.format(html_escape(item['missing'])))
            parts.append('</tr>')

        parts.append('</tbody></table>')
    else:
        parts.append('<h2 style="color:#059669;border-bottom:3px solid #059669;padding-bottom:10px;margin-top:0;">Data Completeness</h2>')
        parts.append('<div style="margin-top:30px;padding:30px;background:#d1fae5;border-radius:10px;text-align:center;">')
        parts.append('<div style="font-size:48px;color:#059669;margin-bottom:16px;">&#10003;</div>')
        parts.append('<h3 style="color:#065f46;margin:0 0 8px 0;">All Information Complete!</h3>')
        parts.append('<p style="color:#047857;font-size:16px;margin:0;">Every registrant has the selected fields filled in.</p>')
        parts.append('</div>')

    parts.append('</div>')
    return ''.join(parts)


def render_medical_page(people, template, questions):
    """Render a medical summary page.
    Uses printSettings.medicalItems[] to decide which fields/questions/keywords to include."""
    ps = template.get('printSettings', {})
    items_config = ps.get('medicalItems', [])

    # Default items if none configured
    if not items_config:
        items_config = [
            {'itemType': 'medical', 'field': 'MedAllergy', 'label': 'Allergies'},
            {'itemType': 'medical', 'field': 'Medications', 'label': 'Medications'},
        ]

    entries = []  # list of {name, age, items: [{label, value}]}

    for p in people:
        items = []
        for item_cfg in items_config:
            it = item_cfg.get('itemType', '')
            field = item_cfg.get('field', '')
            label = item_cfg.get('label', field)

            if it == 'keyword':
                # Scan all answers for keyword match
                kw = field.lower()
                answers = p.get('answers', {})
                for q_text, a_val in answers.items():
                    if not a_val:
                        continue
                    if kw in a_val.lower() or kw in q_text.lower():
                        items.append({'label': q_text, 'value': a_val})
            else:
                val = _get_item_value(p, item_cfg)
                if val and _is_meaningful_medical(val):
                    items.append({'label': label, 'value': val})

        if items:
            entries.append({
                'name': p.get('Name', ''),
                'age': p.get('Age', '') or 'N/A',
                'items': items
            })

    parts = []
    parts.append('<div class="rr-medical-page" style="page-break-after:always;padding:20px;">')
    parts.append('<div style="position:absolute;top:10px;right:20px;color:#c53030;font-weight:bold;font-size:14px;border:1px solid #c53030;padding:3px 10px;background:#fed7d7;">CONFIDENTIAL</div>')

    if entries:
        parts.append('<h2 style="color:#c53030;border-bottom:3px solid #c53030;padding-bottom:10px;margin-top:0;">Medical Summary</h2>')
        parts.append('<table style="width:100%;border-collapse:collapse;">')
        parts.append('<thead><tr style="background:#fed7d7;">')
        parts.append('<th style="padding:8px;text-align:left;border:1px solid #fc8181;">Name</th>')
        parts.append('<th style="padding:8px;text-align:center;border:1px solid #fc8181;width:60px;">Age</th>')
        parts.append('<th style="padding:8px;text-align:left;border:1px solid #fc8181;">Item</th>')
        parts.append('<th style="padding:8px;text-align:left;border:1px solid #fc8181;">Details</th>')
        parts.append('</tr></thead><tbody>')

        for c in sorted(entries, key=lambda x: x['name']):
            for i, item in enumerate(c['items']):
                parts.append('<tr>')
                if i == 0:
                    parts.append('<td style="padding:6px;border:1px solid #e2e8f0;font-weight:bold;vertical-align:top;" rowspan="{0}">{1}</td>'.format(
                        len(c['items']), html_escape(c['name'])))
                    parts.append('<td style="padding:6px;border:1px solid #e2e8f0;text-align:center;vertical-align:top;" rowspan="{0}">{1}</td>'.format(
                        len(c['items']), html_escape(c['age'])))
                parts.append('<td style="padding:6px;border:1px solid #e2e8f0;color:#c53030;font-weight:600;font-size:12px;">{0}</td>'.format(html_escape(item['label'])))
                parts.append('<td style="padding:6px;border:1px solid #e2e8f0;">{0}</td>'.format(html_escape(item['value'])))
                parts.append('</tr>')

        parts.append('</tbody></table>')
    else:
        parts.append('<h2 style="color:#059669;border-bottom:3px solid #059669;padding-bottom:10px;margin-top:0;">Medical Summary</h2>')
        parts.append('<div style="margin-top:30px;padding:30px;background:#d1fae5;border-radius:10px;text-align:center;">')
        parts.append('<div style="font-size:48px;color:#059669;margin-bottom:16px;">&#10003;</div>')
        parts.append('<h3 style="color:#065f46;margin:0 0 8px 0;">No Medical Items Found</h3>')
        parts.append('<p style="color:#047857;font-size:16px;margin:0;">No data was found for the selected medical items in this group.</p>')
        parts.append('</div>')

    parts.append('</div>')
    return ''.join(parts)


# ============================================================================
# AJAX HANDLER
# ============================================================================
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -------------------------------------------------------------------------
    # Template storage helpers (multi-template + backward compat)
    # Storage shape (new "v2"): {"_format": "v2", "templates": {name: tpl}, "currentName": name}
    # Old shape: a single template object (has 'sections' or 'templateName' at top level).
    # On read: detect format, return list. On save: always write new format.
    # -------------------------------------------------------------------------
    def _read_template_store(org_id):
        try:
            saved = model.ExtraValueTextOrg(org_id, 'RegReportTemplate')
        except:
            saved = None
        if not saved:
            return {'_format': 'v2', 'templates': {}, 'currentName': None}
        try:
            data = json.loads(saved)
        except:
            return {'_format': 'v2', 'templates': {}, 'currentName': None}
        if isinstance(data, dict) and data.get('_format') == 'v2' and isinstance(data.get('templates'), dict):
            return data
        # Legacy single-template format - wrap it as {"Default": <old>}
        if isinstance(data, dict) and ('sections' in data or 'templateName' in data):
            return {'_format': 'v2', 'templates': {'Default': data}, 'currentName': 'Default'}
        return {'_format': 'v2', 'templates': {}, 'currentName': None}

    def _write_template_store(org_id, store):
        store['_format'] = 'v2'
        if 'templates' not in store or not isinstance(store.get('templates'), dict):
            store['templates'] = {}
        model.AddExtraValueTextOrg(org_id, 'RegReportTemplate', json.dumps(store))

    # Per-user templates, used when running from Blue Toolbar with no
    # involvement context. Same shape as the org store; lives on
    # PeopleExtra against the running user's PeopleId so each staff
    # member builds their own personal library without stomping on
    # anyone else.
    USER_TEMPLATE_FIELD = 'RegReportTemplate'

    def _current_user_pid():
        try:
            pid = model.UserPeopleId
            return int(pid) if pid else 0
        except:
            return 0

    def _read_user_template_store(people_id):
        if not people_id:
            return {'_format': 'v2', 'templates': {}, 'currentName': None}
        try:
            saved = model.ExtraValueText(people_id, USER_TEMPLATE_FIELD)
        except:
            saved = None
        if not saved:
            return {'_format': 'v2', 'templates': {}, 'currentName': None}
        try:
            data = json.loads(saved)
        except:
            return {'_format': 'v2', 'templates': {}, 'currentName': None}
        if isinstance(data, dict) and data.get('_format') == 'v2' and isinstance(data.get('templates'), dict):
            return data
        if isinstance(data, dict) and ('sections' in data or 'templateName' in data):
            return {'_format': 'v2', 'templates': {'Default': data}, 'currentName': 'Default'}
        return {'_format': 'v2', 'templates': {}, 'currentName': None}

    def _write_user_template_store(people_id, store):
        if not people_id:
            return False
        store['_format'] = 'v2'
        if 'templates' not in store or not isinstance(store.get('templates'), dict):
            store['templates'] = {}
        try:
            model.AddExtraValueText(people_id, USER_TEMPLATE_FIELD, json.dumps(store))
            return True
        except:
            return False

    def _is_user_scope(org_id_raw):
        """Returns True when org_id is the Blue Toolbar sentinel rather
        than a real involvement id. Lets the load/save/delete/rename
        actions route to per-user storage in that case."""
        if org_id_raw is None:
            return True
        s = str(org_id_raw).strip()
        if not s:
            return True
        if s == 'bt_direct' or s == 'user' or s.startswith('user_'):
            return True
        # Anything not parseable as int is treated as user-scope (safer
        # than dropping the save silently).
        try:
            int(s)
            return False
        except:
            return True

    def _load_store(org_id_raw):
        """Pick org-store or user-store based on org_id. Returns
        (store, scope_label) where scope_label is 'org' or 'user' for
        diagnostics."""
        if _is_user_scope(org_id_raw):
            return _read_user_template_store(_current_user_pid()), 'user'
        return _read_template_store(int(org_id_raw)), 'org'

    def _save_store(org_id_raw, store):
        if _is_user_scope(org_id_raw):
            return _write_user_template_store(_current_user_pid(), store)
        _write_template_store(int(org_id_raw), store)
        return True

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
                    'name': safe_str(r.OrganizationName),
                    'division': safe_str(r.DivisionName),
                    'program': safe_str(r.ProgramName),
                    'memberCount': r.MemberCount or 0,
                    'questionCount': r.QuestionCount or 0
                })
            print json.dumps({'success': True, 'orgs': orgs})
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(e)})

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
                programs.append({'id': r.Id, 'name': safe_str(r.Name)})

            div_sql = """
                SELECT DISTINCT d.Id, d.Name, d.ProgId
                FROM Division d
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE o.OrganizationStatusId = 30
                ORDER BY d.Name
            """
            divisions = []
            for r in q.QuerySql(div_sql):
                divisions.append({'id': r.Id, 'name': safe_str(r.Name), 'programId': r.ProgId})

            print json.dumps({'success': True, 'programs': programs, 'divisions': divisions})
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Load Org Data (questions + people + person field list)
    # -------------------------------------------------------------------------
    elif action == 'load_org_data':
        org_id = getattr(Data, 'org_id', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))
        include_dropped = str(getattr(Data, 'include_dropped', '') or '').lower() in ('1', 'true', 'yes')
        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                org_info = q.QuerySqlTop1("SELECT OrganizationName FROM Organizations WHERE OrganizationId = {0}".format(org_id))
                org_name = safe_str(org_info.OrganizationName) if org_info else 'Unknown'

                data = get_registrant_data(org_id, filter_ids, include_dropped=include_dropped)
                if 'error' in data:
                    print json.dumps({'success': False, 'message': data['error']})
                else:
                    person_list = []
                    dropped_count = 0
                    for p in data['people']:
                        is_dropped = bool(p.get('IsDropped', False))
                        if is_dropped:
                            dropped_count += 1
                        person_list.append({
                            'id': p['PeopleId'],
                            'name': p['Name'],
                            'dropped': is_dropped
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
                        {'sourceField': 'MaritalStatus', 'label': 'Marital Status', 'fieldType': 'person'},
                        {'sourceField': 'PrimaryAddress', 'label': 'Full Address', 'fieldType': 'person'},
                        {'sourceField': 'SubGroups', 'label': 'Subgroups', 'fieldType': 'person'}
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
                        {'sourceField': 'CustodyIssue', 'label': 'Custody Issue', 'fieldType': 'medical'},
                        {'sourceField': 'AuthorizedCheckout', 'label': 'Authorized Checkout', 'fieldType': 'medical'},
                        {'sourceField': 'Medications', 'label': 'Approved Medications', 'fieldType': 'medical'}
                    ]

                    question_fields = []
                    for qitem in data['questions']:
                        q_key = qitem['key'] if isinstance(qitem, dict) else qitem
                        q_label = qitem['label'] if isinstance(qitem, dict) else qitem
                        question_fields.append({
                            'sourceField': q_key,
                            'label': q_label,
                            'fieldType': 'regquestion'
                        })

                    print json.dumps({
                        'success': True,
                        'orgName': org_name,
                        'personCount': len(data['people']),
                        'droppedCount': dropped_count,
                        'questionCount': len(data['questions']),
                        'questions': data['questions'],
                        'personList': person_list,
                        'availableSubGroups': data.get('availableSubGroups', []),
                        'availableFields': {
                            'person': person_fields,
                            'family': family_fields,
                            'medical': medical_fields,
                            'regquestion': question_fields
                        }
                    })
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

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
                        {'sourceField': 'MaritalStatus', 'label': 'Marital Status', 'fieldType': 'person'},
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
                        {'sourceField': 'CustodyIssue', 'label': 'Custody Issue', 'fieldType': 'medical'},
                        {'sourceField': 'AuthorizedCheckout', 'label': 'Authorized Checkout', 'fieldType': 'medical'},
                        {'sourceField': 'Medications', 'label': 'Approved Medications', 'fieldType': 'medical'}
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
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Load Saved Template(s)
    # Returns: {success, templates: {name: tpl}, names: [sortedNames], currentName, currentTemplate}
    # If a specific 'name' is given, currentTemplate is that one; else first/saved currentName.
    # -------------------------------------------------------------------------
    elif action == 'load_template':
        org_id = getattr(Data, 'org_id', '')
        requested_name = getattr(Data, 'tpl_name', '') or getattr(Data, 'name', '')
        try:
            store, scope = _load_store(org_id)
            templates = store.get('templates', {}) or {}
            names = sorted(templates.keys(), key=lambda s: s.lower())
            current_name = None
            if requested_name and str(requested_name) in templates:
                current_name = str(requested_name)
            elif store.get('currentName') and store['currentName'] in templates:
                current_name = store['currentName']
            elif names:
                current_name = names[0]
            current_template = templates.get(current_name) if current_name else None
            print json.dumps({
                'success': True,
                'templates': templates,
                'names': names,
                'currentName': current_name,
                'currentTemplate': current_template,
                'hasSaved': len(names) > 0,
                'scope': scope
            })
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Save Template (under a name; creates or overwrites)
    # -------------------------------------------------------------------------
    elif action == 'save_template':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        name = getattr(Data, 'tpl_name', '') or getattr(Data, 'name', '') or 'Default'
        if not template_json:
            print json.dumps({'success': False, 'message': 'Template required'})
        else:
            try:
                parsed = json.loads(template_json)
                store, scope = _load_store(org_id)
                templates = store.get('templates', {}) or {}
                templates[str(name)] = parsed
                store['templates'] = templates
                store['currentName'] = str(name)
                ok = _save_store(org_id, store)
                if not ok:
                    print json.dumps({'success': False, 'message': 'Could not save -- no current user context for personal templates'})
                else:
                    print json.dumps({
                        'success': True,
                        'message': 'Saved as "' + str(name) + '"'
                                   + (' (personal)' if scope == 'user' else ''),
                        'name': str(name),
                        'names': sorted(templates.keys(), key=lambda s: s.lower()),
                        'scope': scope
                    })
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Delete Template by name
    # -------------------------------------------------------------------------
    elif action == 'delete_template':
        org_id = getattr(Data, 'org_id', '')
        name = getattr(Data, 'tpl_name', '') or getattr(Data, 'name', '')
        if not name:
            print json.dumps({'success': False, 'message': 'Template name required'})
        else:
            try:
                store, scope = _load_store(org_id)
                templates = store.get('templates', {}) or {}
                if str(name) in templates:
                    del templates[str(name)]
                    store['templates'] = templates
                    if store.get('currentName') == str(name):
                        store['currentName'] = sorted(templates.keys(), key=lambda s: s.lower())[0] if templates else None
                    _save_store(org_id, store)
                print json.dumps({
                    'success': True,
                    'names': sorted(templates.keys(), key=lambda s: s.lower()),
                    'scope': scope
                })
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Rename Template
    # -------------------------------------------------------------------------
    elif action == 'rename_template':
        org_id = getattr(Data, 'org_id', '')
        old_name = getattr(Data, 'tpl_old_name', '') or getattr(Data, 'old_name', '')
        new_name = getattr(Data, 'tpl_new_name', '') or getattr(Data, 'new_name', '')
        if not old_name or not new_name:
            print json.dumps({'success': False, 'message': 'Old name and new name required'})
        else:
            try:
                store, scope = _load_store(org_id)
                templates = store.get('templates', {}) or {}
                old_name = str(old_name); new_name = str(new_name)
                if old_name not in templates:
                    print json.dumps({'success': False, 'message': 'Template "' + old_name + '" not found'})
                elif new_name in templates and new_name != old_name:
                    print json.dumps({'success': False, 'message': 'Template "' + new_name + '" already exists'})
                else:
                    templates[new_name] = templates.pop(old_name)
                    store['templates'] = templates
                    if store.get('currentName') == old_name:
                        store['currentName'] = new_name
                    _save_store(org_id, store)
                    print json.dumps({
                        'success': True,
                        'name': new_name,
                        'names': sorted(templates.keys(), key=lambda s: s.lower()),
                        'scope': scope
                    })
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Preview Report (single person)
    # -------------------------------------------------------------------------
    elif action == 'preview_report':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        people_id = getattr(Data, 'people_id', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))
        include_dropped = str(getattr(Data, 'include_dropped', '') or '').lower() in ('1', 'true', 'yes')

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
                    org_name = safe_str(org_info.OrganizationName) if org_info else ''
                    data = get_registrant_data(org_id, filter_ids, include_dropped=include_dropped)
                ev_names = extract_ev_names_from_template(template)
                if ev_names:
                    fetch_extra_values(data['people'], ev_names)

                # Supplemental pages in preview (shown before the person preview)
                ps = template.get('printSettings', {})
                prefix_html = ''
                if ps.get('showSummaryPage', False):
                    prefix_html += render_summary_page(data['people'], data.get('questions', []), org_name, template)
                if ps.get('showCoverPage', False):
                    prefix_html += render_cover_page(data['people'], org_name, template)
                if ps.get('showMissingInfoPage', False):
                    prefix_html += render_missing_info_page(data['people'], template)
                if ps.get('showMedicalPage', False):
                    prefix_html += render_medical_page(data['people'], template, data.get('questions', []))

                if ps.get('hidePersonDetail', False):
                    html_out = prefix_html if prefix_html else '<div style="padding:30px;color:#718096;">No summary/supplemental pages enabled. Turn on a summary page, or turn off "Skip per-person detail".</div>'
                else:
                    html_out = prefix_html + render_report_html(data['people'], template, org_name, data.get('questions', []), people_id)
                print json.dumps({'success': True, 'html': html_out})
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Generate Full Report (all registrants)
    # -------------------------------------------------------------------------
    elif action == 'generate_report':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))
        include_dropped = str(getattr(Data, 'include_dropped', '') or '').lower() in ('1', 'true', 'yes')

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
                    org_name = safe_str(org_info.OrganizationName) if org_info else ''
                    data = get_registrant_data(org_id, filter_ids, include_dropped=include_dropped)
                ev_names = extract_ev_names_from_template(template)
                if ev_names:
                    fetch_extra_values(data['people'], ev_names)

                # Build supplemental pages based on print settings
                ps = template.get('printSettings', {})
                prefix_html = ''
                if ps.get('showSummaryPage', False):
                    prefix_html += render_summary_page(data['people'], data.get('questions', []), org_name, template)
                if ps.get('showCoverPage', False):
                    prefix_html += render_cover_page(data['people'], org_name, template)
                if ps.get('showMissingInfoPage', False):
                    prefix_html += render_missing_info_page(data['people'], template)
                if ps.get('showMedicalPage', False):
                    prefix_html += render_medical_page(data['people'], template, data.get('questions', []))

                if ps.get('hidePersonDetail', False):
                    html_out = prefix_html if prefix_html else '<div style="padding:30px;color:#718096;">No summary/supplemental pages enabled. Turn on a summary page, or turn off "Skip per-person detail".</div>'
                else:
                    html_out = prefix_html + render_report_html(data['people'], template, org_name, data.get('questions', []))
                print json.dumps({'success': True, 'html': html_out, 'personCount': len(data['people'])})
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Get Person List (for selector dropdown)
    # -------------------------------------------------------------------------
    elif action == 'export_csv':
        org_id = getattr(Data, 'org_id', '')
        template_json = getattr(Data, 'template_json', '')
        filter_ids = parse_filter_people(getattr(Data, 'filter_people', ''))
        include_dropped = str(getattr(Data, 'include_dropped', '') or '').lower() in ('1', 'true', 'yes')

        if not org_id or not template_json:
            print json.dumps({'success': False, 'message': 'Organization and template required'})
        else:
            try:
                template = json.loads(template_json)
                if str(org_id) == 'bt_direct':
                    data = get_people_data_direct(filter_ids)
                else:
                    org_id = int(org_id)
                    data = get_registrant_data(org_id, filter_ids, include_dropped=include_dropped)
                ev_names = extract_ev_names_from_template(template)
                if ev_names:
                    fetch_extra_values(data['people'], ev_names)

                # Build CSV headers and field definitions from template sections.
                # Use the same auto-populate helper as the PDF render path so
                # sections that rely on auto-populated reg questions (empty
                # fields + "question/answer/registration" in the title) emit
                # the same set of columns in the CSV.
                questions = data.get('questions', [])
                csv_headers = []
                csv_fields = []
                for section in template.get('sections', []):
                    if not section.get('visible', True):
                        continue
                    for field in auto_populate_section_fields(section, questions):
                        if not field.get('visible', True):
                            continue
                        ft = field.get('fieldType', '')
                        if ft == 'static':
                            continue  # Skip separators and static text in CSV
                        label = field.get('label', field.get('sourceField', ''))
                        if label:
                            csv_headers.append(label)
                            csv_fields.append(field)

                # Build CSV rows using the same get_field_value() as report rendering
                csv_rows = []
                for person in data.get('people', []):
                    row = []
                    for field in csv_fields:
                        val = get_field_value(person, field)
                        # Strip HTML entities back to plain text for CSV
                        val = val.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&#39;', "'")
                        row.append(val)
                    csv_rows.append(row)

                # Build CSV string
                def csv_escape(v):
                    s = safe_str(v)
                    if ',' in s or '"' in s or '\n' in s:
                        return '"' + s.replace('"', '""') + '"'
                    return s

                csv_lines = []
                csv_lines.append(','.join(csv_escape(h) for h in csv_headers))
                for row in csv_rows:
                    csv_lines.append(','.join(csv_escape(c) for c in row))
                csv_text = '\n'.join(csv_lines)

                print json.dumps(sanitize_for_json({'success': True, 'csv': csv_text, 'personCount': len(csv_rows)}))
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

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
                    persons.append({'id': r.PeopleId, 'name': safe_str(r.Name2)})
                print json.dumps({'success': True, 'persons': persons})
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(e)})

    # -------------------------------------------------------------------------
    # Apply Update — fetch latest script from DisplayCache and overwrite the
    # installed Python content. Triggered by "Update Now" banner button.
    # -------------------------------------------------------------------------
    elif action == 'apply_update':
        new_code = ''
        try:
            fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
            new_code = str(model.RestGet(fetch_url, {}))
        except Exception as fe:
            print json.dumps({'success': False, 'message': 'Failed to fetch update: ' + safe_str(fe)})
        else:
            if not new_code or len(new_code) < 200:
                print json.dumps({'success': False, 'message': 'Invalid or empty script code received'})
            else:
                target_name = get_script_name() or DC_SCRIPT_ID
                try:
                    model.WriteContentPython(target_name, new_code)
                    print json.dumps({'success': True, 'message': 'Updated ' + target_name + '. Reload the page.'})
                except Exception as we:
                    print json.dumps({'success': False, 'message': 'Write failed: ' + safe_str(we)})

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
        <div id="rrAppUpdateBanner" style="display:none;background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;align-items:center;gap:10px;"></div>
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
            <span class="rr-badge" id="rrOrgDroppedCount" style="display:none;background:#fed7d7;color:#9b2c2c;"></span>
            <span class="rr-badge rr-badge-green" id="rrOrgQuestionCount"></span>
            <label id="rrIncludeDroppedWrap" style="display:none;margin:0 0 0 8px;align-items:center;gap:6px;font-size:12px;color:#4a5568;cursor:pointer;" title="Also include people who answered registration questions but were later dropped or moved out of this involvement. Registration answers persist on the original involvement.">
                <input type="checkbox" id="rrIncludeDropped" onchange="toggleIncludeDropped(this.checked)" style="margin:0;">
                Include dropped registrants
            </label>
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
                <span id="rrSavedTplContainer"></span>
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
                        <span class="rr-option-label">Preserve line breaks in answers</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptPreserveNewlines" onchange="updateGlobalOption('preserveNewlines', this.checked)"><span class="rr-toggle-slider"></span></label>
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
                        <span class="rr-option-label">Show page numbers</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptPageNumbers" onchange="updatePrintSetting('showPageNumbers', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Allow sections to split across pages</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptSplitSections" onchange="updatePrintSetting('allowSectionSplit', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Compact row spacing</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptCompactRows" onchange="updateGlobalOption('compactRows', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Heading color</span>
                        <input type="color" id="rrOptHeadingColor" value="#2c5282" onchange="updateGlobalOption('headingColor', this.value)">
                    </div>
                    <div class="rr-option-row" style="align-items:flex-start;">
                        <span class="rr-option-label" style="padding-top:6px;">Document title</span>
                        <input type="text" id="rrOptReportTitle"
                               placeholder="Registration Report"
                               oninput="updateReportTitle(this.value)"
                               style="flex:1;min-width:200px;padding:5px 8px;border:1px solid #cbd5e0;border-radius:4px;font-size:13px;">
                    </div>
                    <div class="rr-option-row" style="margin-top:-4px;">
                        <span style="flex:1;font-size:11px;color:#718096;font-style:italic;">Title that prints on the report and the cover page. Replaces the involvement name (or "Selected People" in Blue Toolbar runs). Leave blank to use the default.</span>
                    </div>
                </div>

                <div class="rr-card">
                    <div class="rr-card-title">Supplemental Pages</div>
                    <div style="font-size:12px;color:#718096;margin-bottom:10px;">These pages are prepended to the report when generating/printing.</div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Skip per-person detail <span style="font-weight:400;color:#a0aec0;">(print only the pages below)</span></span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptSkipDetail" onchange="updatePrintSetting('hidePersonDetail', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Form summary page (counts &amp; charts)</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptSummaryPage" onchange="updatePrintSetting('showSummaryPage', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div id="rrSummaryPanel" style="display:none;padding:8px 12px;background:#f7fafc;border:1px solid #e2e8f0;border-radius:4px;margin:4px 0 8px 0;">
                        <div style="font-size:12px;font-weight:600;color:#4a5568;margin-bottom:6px;">Items to summarize <span style="font-weight:400;color:#a0aec0;">(leave empty for automatic: gender, age, subgroups, ZIPs &amp; every question)</span>:</div>
                        <div id="rrSummaryItemsList"></div>
                        <div style="display:flex;gap:6px;margin-top:6px;">
                            <select id="rrSummaryItemPicker" style="flex:1;font-size:12px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:4px;"></select>
                            <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addSupplementalItem('summary')">Add</button>
                        </div>
                        <div id="rrSummaryEvRow" style="display:none;margin-top:6px;padding:6px;background:#fff;border:1px dashed #cbd5e0;border-radius:4px;">
                            <div style="font-size:11px;color:#4a5568;margin-bottom:4px;">Extra Value field name (exact) + optional label:</div>
                            <div style="display:flex;gap:6px;">
                                <input type="text" id="rrSummaryEvName" placeholder="EV field name" style="flex:1;font-size:12px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:4px;"/>
                                <input type="text" id="rrSummaryEvLabel" placeholder="Label (optional)" style="flex:1;font-size:12px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:4px;"/>
                                <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addSummaryEvField()">Add</button>
                            </div>
                        </div>
                    </div>
                    <div class="rr-option-row">
                        <span class="rr-option-label">Cover page with summary</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptCoverPage" onchange="updatePrintSetting('showCoverPage', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>

                    <div class="rr-option-row">
                        <span class="rr-option-label">Missing information page</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptMissingInfo" onchange="updatePrintSetting('showMissingInfoPage', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div id="rrMissingInfoPanel" style="display:none;padding:8px 12px;background:#f7fafc;border:1px solid #e2e8f0;border-radius:4px;margin:4px 0 8px 0;">
                        <div style="font-size:12px;font-weight:600;color:#4a5568;margin-bottom:6px;">Fields to check (people missing these will be listed):</div>
                        <div id="rrMissingItemsList"></div>
                        <div style="display:flex;gap:6px;margin-top:6px;">
                            <select id="rrMissingItemPicker" style="flex:1;font-size:12px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:4px;"></select>
                            <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addSupplementalItem('missing')">Add</button>
                        </div>
                    </div>

                    <div class="rr-option-row">
                        <span class="rr-option-label">Medical summary page</span>
                        <label class="rr-toggle"><input type="checkbox" id="rrOptMedicalPage" onchange="updatePrintSetting('showMedicalPage', this.checked)"><span class="rr-toggle-slider"></span></label>
                    </div>
                    <div id="rrMedicalPanel" style="display:none;padding:8px 12px;background:#f7fafc;border:1px solid #e2e8f0;border-radius:4px;margin:4px 0 8px 0;">
                        <div style="font-size:12px;font-weight:600;color:#4a5568;margin-bottom:6px;">Items to include on medical page:</div>
                        <div id="rrMedicalItemsList"></div>
                        <div style="display:flex;gap:6px;margin-top:6px;">
                            <select id="rrMedicalItemPicker" style="flex:1;font-size:12px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:4px;"></select>
                            <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addSupplementalItem('medical')">Add</button>
                        </div>
                        <div id="rrMedKeywordRow" style="display:none;margin-top:6px;">
                            <div style="display:flex;gap:6px;">
                                <input type="text" id="rrMedKeywordInput" placeholder="Enter keyword (e.g. asthma)" style="flex:1;padding:4px 6px;border:1px solid #cbd5e0;border-radius:4px;font-size:12px;">
                                <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addKeywordItem()">Add Keyword</button>
                                <button class="rr-btn rr-btn-sm" onclick="document.getElementById('rrMedKeywordRow').style.display='none'" style="padding:2px 8px;">Cancel</button>
                            </div>
                        </div>
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
            <div class="rr-generate-actions" style="flex-wrap:wrap;gap:8px;">
                <div style="display:flex;gap:6px;align-items:center;">
                    <select id="rrSavedTplPicker" onchange="onSavedTemplatePicked(this.value)" style="padding:6px 8px;border:1px solid #cbd5e0;border-radius:4px;font-size:13px;min-width:200px;">
                        <option value="">(no saved templates)</option>
                    </select>
                    <button class="rr-btn rr-btn-primary" id="rrSaveBtn" onclick="saveTemplate()" title="Overwrite the currently loaded saved template"><i class="fa fa-save"></i> Save</button>
                    <button class="rr-btn rr-btn-secondary" onclick="saveTemplateAs()" title="Save current configuration under a new name"><i class="fa fa-plus"></i> Save As&hellip;</button>
                    <button class="rr-btn rr-btn-secondary" id="rrRenameBtn" onclick="renameTemplate()" title="Rename the currently loaded saved template"><i class="fa fa-pencil"></i> Rename</button>
                    <button class="rr-btn rr-btn-danger" id="rrDeleteBtn" onclick="deleteTemplate()" title="Delete the currently loaded saved template"><i class="fa fa-trash"></i> Delete</button>
                </div>
                <div style="display:flex;gap:6px;align-items:center;">
                    <button class="rr-btn rr-btn-secondary" onclick="exportTemplate()" title="Download the current configuration as a JSON file"><i class="fa fa-download"></i> Export</button>
                    <button class="rr-btn rr-btn-secondary" onclick="openImportTemplate()" title="Load a configuration from a JSON file or pasted text"><i class="fa fa-upload"></i> Import</button>
                </div>
                <div style="flex-basis:100%;height:0;"></div>
                <button class="rr-btn rr-btn-success" onclick="generateFullReport()"><i class="fa fa-file-text-o"></i> Generate Report</button>
                <button class="rr-btn rr-btn-secondary" id="rrPrintBtn" style="display:none;" onclick="printReport()"><i class="fa fa-print"></i> Print Report</button>
                <button class="rr-btn rr-btn-secondary" id="rrCsvBtn" style="display:none;" onclick="exportCsv()"><i class="fa fa-download"></i> Export CSV</button>
            </div>
            <div id="rrImportPanel" style="display:none;margin-top:10px;padding:12px;background:#f0f4ff;border:1px solid #cbd5e0;border-radius:6px;">
                <div style="font-weight:600;margin-bottom:6px;">Import Template</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:6px;">
                    <input type="file" id="rrImportFile" accept=".json,application/json" style="font-size:12px;">
                    <span style="font-size:12px;color:#666;">or paste JSON below:</span>
                </div>
                <textarea id="rrImportText" placeholder='Paste a previously exported template JSON here...' style="width:100%;min-height:80px;font-size:12px;font-family:monospace;padding:6px;border:1px solid #cbd5e0;border-radius:4px;"></textarea>
                <div style="margin-top:6px;display:flex;gap:6px;">
                    <button class="rr-btn rr-btn-primary rr-btn-sm" onclick="doImportTemplate()"><i class="fa fa-check"></i> Import</button>
                    <button class="rr-btn rr-btn-sm" onclick="document.getElementById('rrImportPanel').style.display='none';" style="padding:4px 10px;">Cancel</button>
                </div>
            </div>
            <div id="rrImportWarnings" style="display:none;margin-top:10px;"></div>
        </div>
        <div id="rrGeneratedReport" class="rr-card" style="display:none;">
            <div class="rr-card-title" style="display:flex; justify-content:space-between; align-items:center;">
                Generated Report
                <span>
                <button class="rr-btn rr-btn-secondary rr-btn-sm" onclick="printReport()"><i class="fa fa-print"></i> Print</button>
                <button class="rr-btn rr-btn-secondary rr-btn-sm" onclick="exportCsv()"><i class="fa fa-download"></i> CSV</button>
                </span>
            </div>
            <div id="rrReportContent"></div>
        </div>
    </div>
</div>

<div id="rr-print-output"></div>

<script>
(function() {
    var APP_VERSION = ''' + json.dumps(APP_VERSION) + ''';
    var DC_SCRIPT_ID = ''' + json.dumps(DC_SCRIPT_ID) + ''';
    var DC_API_BASE = ''' + json.dumps(DC_API_BASE) + ''';
    // SCRIPT_NAME is detected client-side from the URL the browser actually loaded —
    // this is bulletproof regardless of admin renames or model.URL quirks server-side.
    // Falls back to the server-rendered value, then the hardcoded default.
    var SCRIPT_NAME = (function() {
        try {
            var m = (window.location.pathname || '').match(/\\/PyScript(?:Form)?\\/([^\\/?#]+)/);
            if (m && m[1]) return m[1];
        } catch(e) {}
        return ''' + json.dumps(get_script_name()) + ''';
    })();
    var APP_UPDATE_AVAILABLE = false;
    var APP_LATEST_VERSION = '';

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
        availableSubGroups: [],
        template: null,
        savedTemplate: null,
        savedTemplates: {},
        savedNames: [],
        currentSavedName: null,
        generatedHtml: '',
        btMode: false,
        btPeopleIds: [],
        includeDropped: false
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
        // Always echo the script name so server handlers (e.g. apply_update) write to the
        // right Python content slot regardless of how the admin renamed the install.
        if (typeof SCRIPT_NAME !== 'undefined' && SCRIPT_NAME) {
            params.script_name = SCRIPT_NAME;
        }
        if (state.btMode && state.btPeopleIds.length > 0) {
            params.filter_people = state.btPeopleIds.join(',');
        }
        // Only the org-based actions consume include_dropped. Sending it on
        // every call is harmless — unrecognized params are ignored server-side.
        if (state.includeDropped) {
            params.include_dropped = '1';
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

    function applyOrgLoadResponse(data) {
        state.selectedOrgName = data.orgName;
        state.personList = data.personList || [];
        state.questions = data.questions || [];
        state.availableFields = data.availableFields || {};
        state.availableSubGroups = data.availableSubGroups || [];
        document.getElementById('rrOrgNameBar').textContent = data.orgName;
        var countLabel = state.btMode ? data.personCount + ' selected' : data.personCount + ' people';
        document.getElementById('rrOrgPeopleCount').textContent = countLabel;
        var droppedEl = document.getElementById('rrOrgDroppedCount');
        var droppedCount = data.droppedCount || 0;
        if (droppedCount > 0) {
            droppedEl.textContent = droppedCount + ' dropped';
            droppedEl.style.display = '';
        } else {
            droppedEl.style.display = 'none';
        }
        document.getElementById('rrOrgQuestionCount').textContent = data.questionCount + ' questions';
        document.getElementById('rrSelectedOrgBar').style.display = 'flex';
        // The include-dropped toggle only makes sense for org-based runs, not Blue Toolbar.
        var dropWrap = document.getElementById('rrIncludeDroppedWrap');
        if (dropWrap) dropWrap.style.display = state.btMode ? 'none' : 'inline-flex';
        var dropCb = document.getElementById('rrIncludeDropped');
        if (dropCb) dropCb.checked = !!state.includeDropped;
        var changeBtn = document.getElementById('rrChangeOrgBtn');
        if (changeBtn && state.btMode) changeBtn.style.display = 'none';
        var sel = document.getElementById('rrPreviewPerson');
        sel.innerHTML = '';
        for (var i = 0; i < state.personList.length; i++) {
            var opt = document.createElement('option');
            opt.value = state.personList[i].id;
            opt.textContent = state.personList[i].name + (state.personList[i].dropped ? ' (Dropped)' : '');
            sel.appendChild(opt);
        }
    }

    window.toggleIncludeDropped = function(checked) {
        state.includeDropped = !!checked;
        if (!state.selectedOrgId || state.btMode) return;
        showToast('Reloading involvement data...', 'info');
        ajax('load_org_data', {org_id: state.selectedOrgId}, function(data) {
            if (!data.success) { showToast('Error: ' + data.message, 'danger'); return; }
            applyOrgLoadResponse(data);
            // Refresh preview if user has already configured a template
            if (state.currentStep >= 3 && state.template) refreshPreview();
            showToast(state.includeDropped
                ? 'Now including ' + (data.droppedCount || 0) + ' dropped registrants'
                : 'Showing active members only', 'success');
        });
    };

    window.selectOrg = function(orgId, el) {
        document.querySelectorAll('.rr-org-item').forEach(function(e) { e.classList.remove('selected'); });
        if (el) el.classList.add('selected');
        state.selectedOrgId = orgId;
        showToast('Loading organization data...', 'info');
        ajax('load_org_data', {org_id: orgId}, function(data) {
            if (!data.success) { showToast('Error: ' + data.message, 'danger'); return; }
            applyOrgLoadResponse(data);
            ajax('load_template', {org_id: orgId}, function(tplData) {
                if (tplData.success) {
                    state.savedTemplates = tplData.templates || {};
                    state.savedNames = tplData.names || [];
                    state.currentSavedName = tplData.currentName || null;
                    state.savedTemplate = tplData.currentTemplate || null;
                } else {
                    state.savedTemplates = {};
                    state.savedNames = [];
                    state.currentSavedName = null;
                    state.savedTemplate = null;
                }
                renderSavedTemplateCards();
                selectTemplate('basic');
                showToast('Organization loaded! Configure layout next.', 'success');
                goToStep(2);
            });
        });
    };

    // ===== STEP 2 =====
    function renderSavedTemplateCards() {
        var container = document.getElementById('rrSavedTplContainer');
        if (!container) return;
        var html = '';
        for (var i = 0; i < state.savedNames.length; i++) {
            var nm = state.savedNames[i];
            html += '<div class="rr-template-card" data-tpl="saved" data-name="' + escAttr(nm) + '" onclick="selectSavedTemplate(this.dataset.name)">';
            html += '<h4><i class="fa fa-bookmark-o"></i> ' + escHtml(nm) + '</h4>';
            html += '<p>Saved template for this org</p>';
            html += '</div>';
        }
        container.innerHTML = html;
    }

    window.selectSavedTemplate = function(name) {
        if (!state.savedTemplates || !state.savedTemplates[name]) return;
        state.currentSavedName = name;
        state.savedTemplate = state.savedTemplates[name];
        document.querySelectorAll('.rr-template-card').forEach(function(c) { c.classList.remove('selected'); });
        var card = document.querySelector('.rr-template-card[data-tpl="saved"][data-name="' + name.replace(/"/g, '\\\\"') + '"]');
        if (card) card.classList.add('selected');
        state.template = JSON.parse(JSON.stringify(state.savedTemplate));
        applyLoadedTemplateToUI();
    };

    function applyLoadedTemplateToUI() {
        if (!state.template) return;
        // Full sync (global options + supplemental pages + title) -- shared
        // with selectTemplate. Previously this only synced a handful of
        // controls using stale element IDs, so saved templates lost their
        // supplemental-page settings (summary/cover/missing/medical) on load.
        syncTemplateOptionsToUI();
        renderSections();
        refreshSaveControls();
    }

    window.selectTemplate = function(tplName) {
        document.querySelectorAll('.rr-template-card').forEach(function(c) { c.classList.remove('selected'); });
        var card = document.querySelector('.rr-template-card[data-tpl="' + tplName + '"]');
        if (card) card.classList.add('selected');
        if (tplName === 'saved' && state.savedTemplate) {
            state.template = JSON.parse(JSON.stringify(state.savedTemplate));
        } else if (TEMPLATES[tplName]) {
            state.template = JSON.parse(JSON.stringify(TEMPLATES[tplName]));
            state.currentSavedName = null;
        }
        if (state.template) {
            syncTemplateOptionsToUI();
        }
        renderSections();
    };

    // Sync every Step-2 option control (global options + supplemental pages)
    // from state.template. Shared by selectTemplate (built-in templates) AND
    // applyLoadedTemplateToUI (saved templates) so picking a SAVED template
    // restores its supplemental-page settings -- summary/cover/missing/medical
    // toggles, panels, and item lists -- not just the section layout.
    function syncTemplateOptionsToUI() {
        if (!state.template) return;
        var go = state.template.globalOptions || {};
        var ps = state.template.printSettings || {};
        setChecked('rrOptHideEmpty', go.hideEmptyFields !== false);
        setChecked('rrOptHideUnanswered', go.hideUnansweredQuestions !== false);
        setChecked('rrOptShowPhoto', go.showPersonPhoto === true);
        setChecked('rrOptPreserveNewlines', go.preserveNewlines === true);
        setChecked('rrOptOnePerPage', ps.onePersonPerPage !== false);
        setChecked('rrOptShowOrgHeader', ps.showOrgHeader !== false);
        setChecked('rrOptCompactRows', go.compactRows === true);
        setChecked('rrOptPageNumbers', ps.showPageNumbers === true);
        setChecked('rrOptSplitSections', ps.allowSectionSplit === true);
        var hcEl = document.getElementById('rrOptHeadingColor');
        if (hcEl) hcEl.value = go.headingColor || '#2c5282';
        // Document title is a template-root field; pre-fill from there.
        var rtEl = document.getElementById('rrOptReportTitle');
        if (rtEl) rtEl.value = state.template.reportTitle || '';
        // Supplemental page toggles
        setChecked('rrOptSkipDetail', ps.hidePersonDetail === true);
        setChecked('rrOptSummaryPage', ps.showSummaryPage === true);
        setChecked('rrOptCoverPage', ps.showCoverPage === true);
        setChecked('rrOptMissingInfo', ps.showMissingInfoPage === true);
        setChecked('rrOptMedicalPage', ps.showMedicalPage === true);
        var summaryPanel = document.getElementById('rrSummaryPanel');
        if (summaryPanel) summaryPanel.style.display = ps.showSummaryPage ? 'block' : 'none';
        var missingPanel = document.getElementById('rrMissingInfoPanel');
        if (missingPanel) missingPanel.style.display = ps.showMissingInfoPage ? 'block' : 'none';
        var medPanel = document.getElementById('rrMedicalPanel');
        if (medPanel) medPanel.style.display = ps.showMedicalPage ? 'block' : 'none';
        // Build pickers and render item lists
        buildSupplementalPicker('summary');
        buildSupplementalPicker('missing');
        buildSupplementalPicker('medical');
        renderSupplementalItems('summary');
        renderSupplementalItems('missing');
        renderSupplementalItems('medical');
    }

    function setChecked(id, val) { var el = document.getElementById(id); if (el) el.checked = val; }

    // Lives at template root (not globalOptions) so it round-trips with
    // export/import as a first-class field.
    window.updateReportTitle = function(value) {
        if (!state.template) return;
        state.template.reportTitle = (value || '').trim();
    };

    window.updateGlobalOption = function(key, value) {
        if (!state.template) return;
        if (!state.template.globalOptions) state.template.globalOptions = {};
        state.template.globalOptions[key] = value;
    };
    window.updatePrintSetting = function(key, value) {
        if (!state.template) return;
        if (!state.template.printSettings) state.template.printSettings = {};
        state.template.printSettings[key] = value;
        if (key === 'showSummaryPage') {
            var el = document.getElementById('rrSummaryPanel');
            if (el) el.style.display = value ? 'block' : 'none';
            if (value) { buildSupplementalPicker('summary'); renderSupplementalItems('summary'); }
        }
        if (key === 'showMissingInfoPage') {
            var el = document.getElementById('rrMissingInfoPanel');
            if (el) el.style.display = value ? 'block' : 'none';
            if (value) { buildSupplementalPicker('missing'); renderSupplementalItems('missing'); }
        }
        if (key === 'showMedicalPage') {
            var el = document.getElementById('rrMedicalPanel');
            if (el) el.style.display = value ? 'block' : 'none';
            if (value) { buildSupplementalPicker('medical'); renderSupplementalItems('medical'); }
        }
    };

    // Build picker options for supplemental page item pickers
    function buildSupplementalPicker(pageType) {
        var pickerId = pageType === 'missing' ? 'rrMissingItemPicker' : (pageType === 'medical' ? 'rrMedicalItemPicker' : 'rrSummaryItemPicker');
        var sel = document.getElementById(pickerId);
        if (!sel) return;
        var af = state.availableFields || {};
        // Summary picker offers aggregatable items (demographics, subgroups,
        // registration questions) plus the same field sources as the
        // section/field builder. Free-text-ish fields fall back to a
        // top-values list in the renderer.
        if (pageType === 'summary') {
            var shtml = '<option value="">+ Add an item...</option>';
            shtml += '<optgroup label="Demographics (computed)">';
            shtml += '<option value="demographic|Gender|Gender">Gender</option>';
            shtml += '<option value="demographic|Age|Age Distribution">Age Distribution</option>';
            shtml += '<option value="demographic|MaritalStatus|Marital Status">Marital Status</option>';
            shtml += '<option value="demographic|PrimaryState|State">State</option>';
            shtml += '<option value="demographic|PrimaryCity|City">City</option>';
            shtml += '<option value="demographic|PrimaryZip|Top ZIP Codes">Top ZIP Codes</option>';
            shtml += '</optgroup>';
            shtml += '<optgroup label="Groups"><option value="subgroup|_all|Subgroups">Subgroups</option></optgroup>';
            var skipPerson = {Gender:1, Age:1, MaritalStatus:1, SubGroups:1};
            if (af.person && af.person.length) {
                shtml += '<optgroup label="Person Fields">';
                for (var pi = 0; pi < af.person.length; pi++) {
                    if (skipPerson[af.person[pi].sourceField]) continue;
                    shtml += '<option value="person|' + escAttr(af.person[pi].sourceField) + '|' + escAttr(af.person[pi].label) + '">' + escHtml(af.person[pi].label) + '</option>';
                }
                shtml += '</optgroup>';
            }
            if (af.family && af.family.length) {
                shtml += '<optgroup label="Family Fields">';
                for (var fyi = 0; fyi < af.family.length; fyi++) {
                    shtml += '<option value="person|' + escAttr(af.family[fyi].sourceField) + '|' + escAttr(af.family[fyi].label) + '">' + escHtml(af.family[fyi].label) + '</option>';
                }
                shtml += '</optgroup>';
            }
            if (af.medical && af.medical.length) {
                shtml += '<optgroup label="Medical / Emergency">';
                for (var mdi = 0; mdi < af.medical.length; mdi++) {
                    shtml += '<option value="person|' + escAttr(af.medical[mdi].sourceField) + '|' + escAttr(af.medical[mdi].label) + '">' + escHtml(af.medical[mdi].label) + '</option>';
                }
                shtml += '</optgroup>';
            }
            if (af.regquestion && af.regquestion.length > 0) {
                shtml += '<optgroup label="Registration Questions">';
                for (var si = 0; si < af.regquestion.length; si++) {
                    var sql2 = af.regquestion[si].label;
                    var sqs = sql2.length > 50 ? sql2.substring(0,47) + '...' : sql2;
                    shtml += '<option value="regquestion|' + escAttr(af.regquestion[si].sourceField) + '|' + escAttr(sql2) + '">' + escHtml(sqs) + '</option>';
                }
                shtml += '</optgroup>';
            }
            shtml += '<optgroup label="Other"><option value="extravalue|_prompt_|">Extra Value Field...</option></optgroup>';
            sel.innerHTML = shtml;
            return;
        }
        var html = '<option value="">+ Add an item...</option>';
        html += '<optgroup label="Person Fields">';
        var personItems = [
            {f: 'PhotoUrl', l: 'Profile Photo'},
            {f: 'Age', l: 'Age/Birthdate'},
            {f: 'EmailAddress', l: 'Email'},
            {f: 'CellPhone', l: 'Cell Phone'},
            {f: 'PrimaryAddress', l: 'Address'},
            {f: 'Gender', l: 'Gender'}
        ];
        for (var i = 0; i < personItems.length; i++) {
            html += '<option value="person|' + personItems[i].f + '|' + escAttr(personItems[i].l) + '">' + escHtml(personItems[i].l) + '</option>';
        }
        html += '</optgroup><optgroup label="Family">';
        html += '<option value="family|Parents|Parent(s)/Guardian(s)">Parent(s)/Guardian(s)</option>';
        html += '</optgroup><optgroup label="Medical / Emergency">';
        var medItems = [
            {f: 'emcontact', l: 'Emergency Contact'},
            {f: 'emphone', l: 'Emergency Phone'},
            {f: 'doctor', l: 'Doctor'},
            {f: 'docphone', l: 'Doctor Phone'},
            {f: 'insurance', l: 'Insurance'},
            {f: 'policy', l: 'Policy #'},
            {f: 'MedAllergy', l: 'Allergies'},
            {f: 'CustodyIssue', l: 'Custody Issue'},
            {f: 'AuthorizedCheckout', l: 'Authorized Checkout'},
            {f: 'Medications', l: 'Medications'}
        ];
        for (var i = 0; i < medItems.length; i++) {
            html += '<option value="medical|' + medItems[i].f + '|' + escAttr(medItems[i].l) + '">' + escHtml(medItems[i].l) + '</option>';
        }
        html += '</optgroup>';
        if (af.regquestion && af.regquestion.length > 0) {
            html += '<optgroup label="Registration Questions">';
            for (var i = 0; i < af.regquestion.length; i++) {
                var ql = af.regquestion[i].label;
                var qs = ql.length > 50 ? ql.substring(0,47) + '...' : ql;
                html += '<option value="regquestion|' + escAttr(af.regquestion[i].sourceField) + '|' + escAttr(ql) + '">' + escHtml(qs) + '</option>';
            }
            html += '</optgroup>';
        }
        if (pageType === 'medical') {
            html += '<optgroup label="Keyword Search"><option value="keyword|_prompt_|">Keyword (scan all answers)...</option></optgroup>';
        }
        sel.innerHTML = html;
    }

    function renderSupplementalItems(pageType) {
        if (!state.template || !state.template.printSettings) return;
        var listKey = pageType === 'missing' ? 'missingInfoItems' : (pageType === 'medical' ? 'medicalItems' : 'summaryItems');
        var containerId = pageType === 'missing' ? 'rrMissingItemsList' : (pageType === 'medical' ? 'rrMedicalItemsList' : 'rrSummaryItemsList');
        var container = document.getElementById(containerId);
        if (!container) return;
        var items = state.template.printSettings[listKey] || [];
        if (!items.length) {
            container.innerHTML = '<div style="font-size:11px;color:#a0aec0;padding:4px 0;font-style:italic;">No items added. Use the picker below to add fields.</div>';
            return;
        }
        var html = '';
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var tagColor = '#2c5282'; var tagBg = '#ebf8ff';
            if (item.itemType === 'medical') { tagColor = '#c53030'; tagBg = '#fff5f5'; }
            else if (item.itemType === 'regquestion') { tagColor = '#276749'; tagBg = '#f0fff4'; }
            else if (item.itemType === 'keyword') { tagColor = '#6b46c1'; tagBg = '#faf5ff'; }
            else if (item.itemType === 'family') { tagColor = '#b7791f'; tagBg = '#fffff0'; }
            else if (item.itemType === 'extravalue') { tagColor = '#2c7a7b'; tagBg = '#e6fffa'; }
            else if (item.itemType === 'subgroup') { tagColor = '#6b46c1'; tagBg = '#faf5ff'; }
            html += '<div style="display:inline-flex;align-items:center;gap:3px;margin:2px 4px 2px 0;padding:3px 8px;background:' + tagBg + ';border:1px solid ' + tagColor + ';border-radius:12px;font-size:11px;color:' + tagColor + ';">';
            // Reorder arrows (disabled at the ends). Order is the render order.
            var upDis = (i === 0) ? 'opacity:.25;cursor:default;' : 'cursor:pointer;';
            var dnDis = (i === items.length - 1) ? 'opacity:.25;cursor:default;' : 'cursor:pointer;';
            html += '<button data-pagetype="' + pageType + '" data-idx="' + i + '" onclick="moveSupplementalItem(this.dataset.pagetype,parseInt(this.dataset.idx),-1)" style="border:none;background:none;color:' + tagColor + ';font-size:11px;padding:0 1px;line-height:1;' + upDis + '" title="Move up">&#9650;</button>';
            html += '<button data-pagetype="' + pageType + '" data-idx="' + i + '" onclick="moveSupplementalItem(this.dataset.pagetype,parseInt(this.dataset.idx),1)" style="border:none;background:none;color:' + tagColor + ';font-size:11px;padding:0 1px;line-height:1;' + dnDis + '" title="Move down">&#9660;</button>';
            html += '<span>' + escHtml(item.label || item.field) + '</span>';
            html += '<button data-pagetype="' + pageType + '" data-idx="' + i + '" onclick="removeSupplementalItem(this.dataset.pagetype,parseInt(this.dataset.idx))" style="border:none;background:none;cursor:pointer;color:' + tagColor + ';font-size:14px;padding:0 2px;line-height:1;" title="Remove">&times;</button>';
            html += '</div>';
        }
        container.innerHTML = html;
    }

    window.moveSupplementalItem = function(pageType, idx, dir) {
        if (!state.template || !state.template.printSettings) return;
        var listKey = pageType === 'missing' ? 'missingInfoItems' : (pageType === 'medical' ? 'medicalItems' : 'summaryItems');
        var items = state.template.printSettings[listKey] || [];
        var ni = idx + dir;
        if (idx < 0 || idx >= items.length || ni < 0 || ni >= items.length) return;
        var tmp = items[idx]; items[idx] = items[ni]; items[ni] = tmp;
        renderSupplementalItems(pageType);
    };

    window.addSupplementalItem = function(pageType) {
        if (!state.template) return;
        if (!state.template.printSettings) state.template.printSettings = {};
        var listKey = pageType === 'missing' ? 'missingInfoItems' : (pageType === 'medical' ? 'medicalItems' : 'summaryItems');
        if (!state.template.printSettings[listKey]) state.template.printSettings[listKey] = [];
        var pickerId = pageType === 'missing' ? 'rrMissingItemPicker' : (pageType === 'medical' ? 'rrMedicalItemPicker' : 'rrSummaryItemPicker');
        var sel = document.getElementById(pickerId);
        if (!sel || !sel.value) return;
        var parts = sel.value.split('|');
        if (parts.length < 2) return;
        // Keyword prompt
        if (parts[0] === 'keyword' && parts[1] === '_prompt_') {
            document.getElementById('rrMedKeywordRow').style.display = 'block';
            var inp = document.getElementById('rrMedKeywordInput');
            if (inp) { inp.value = ''; inp.focus(); }
            sel.value = '';
            return;
        }
        // Extra Value prompt (summary): reveal the name/label input row
        if (parts[0] === 'extravalue' && parts[1] === '_prompt_') {
            var evRow = document.getElementById('rrSummaryEvRow');
            if (evRow) evRow.style.display = 'block';
            var evN = document.getElementById('rrSummaryEvName');
            var evL = document.getElementById('rrSummaryEvLabel');
            if (evN) { evN.value = ''; evN.focus(); }
            if (evL) evL.value = '';
            sel.value = '';
            return;
        }
        // Check for duplicate
        var existing = state.template.printSettings[listKey];
        for (var i = 0; i < existing.length; i++) {
            if (existing[i].itemType === parts[0] && existing[i].field === parts[1]) {
                showToast('Already added', 'info');
                sel.value = '';
                return;
            }
        }
        existing.push({itemType: parts[0], field: parts[1], label: parts[2] || parts[1]});
        sel.value = '';
        renderSupplementalItems(pageType);
    };

    window.addKeywordItem = function() {
        if (!state.template) return;
        if (!state.template.printSettings) state.template.printSettings = {};
        if (!state.template.printSettings.medicalItems) state.template.printSettings.medicalItems = [];
        var inp = document.getElementById('rrMedKeywordInput');
        var kw = inp ? inp.value.trim() : '';
        if (!kw) { if (inp) inp.focus(); return; }
        var existing = state.template.printSettings.medicalItems;
        for (var i = 0; i < existing.length; i++) {
            if (existing[i].itemType === 'keyword' && existing[i].field.toLowerCase() === kw.toLowerCase()) {
                showToast('Keyword already added', 'info');
                return;
            }
        }
        existing.push({itemType: 'keyword', field: kw, label: 'Keyword: ' + kw});
        document.getElementById('rrMedKeywordRow').style.display = 'none';
        if (inp) inp.value = '';
        renderSupplementalItems('medical');
    };

    window.addSummaryEvField = function() {
        if (!state.template) return;
        if (!state.template.printSettings) state.template.printSettings = {};
        if (!state.template.printSettings.summaryItems) state.template.printSettings.summaryItems = [];
        var nameInput = document.getElementById('rrSummaryEvName');
        var labelInput = document.getElementById('rrSummaryEvLabel');
        var evName = nameInput ? nameInput.value.trim() : '';
        if (!evName) { if (nameInput) nameInput.focus(); return; }
        var lbl = (labelInput && labelInput.value.trim()) ? labelInput.value.trim() : evName;
        var existing = state.template.printSettings.summaryItems;
        for (var i = 0; i < existing.length; i++) {
            if (existing[i].itemType === 'extravalue' && existing[i].field === evName) {
                showToast('Already added', 'info');
                return;
            }
        }
        existing.push({itemType: 'extravalue', field: evName, label: lbl});
        var row = document.getElementById('rrSummaryEvRow');
        if (row) row.style.display = 'none';
        renderSupplementalItems('summary');
    };

    window.removeSupplementalItem = function(pageType, idx) {
        if (!state.template || !state.template.printSettings) return;
        var listKey = pageType === 'missing' ? 'missingInfoItems' : (pageType === 'medical' ? 'medicalItems' : 'summaryItems');
        var items = state.template.printSettings[listKey] || [];
        if (idx >= 0 && idx < items.length) {
            items.splice(idx, 1);
        }
        renderSupplementalItems(pageType);
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
                    html += '<div class="rr-field-item" style="flex-wrap:wrap;">';
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
                    if (f.sourceField === 'SubGroups' && (state.availableSubGroups || []).length > 0) {
                        var sgFilter = f.subgroupFilter || [];
                        var summary = sgFilter.length === 0 ? 'All' : sgFilter.length + ' picked';
                        html += '<button type="button" class="rr-btn rr-btn-sm" onclick="toggleSubGroupFilter(' + si + ',' + fi + ')" style="padding:2px 8px;font-size:11px;" title="Pick which subgroups appear in this field">Subgroups: ' + escHtml(summary) + ' \\u25BE</button>';
                    }
                    html += '<button class="rr-btn rr-btn-danger rr-btn-sm" onclick="removeField(' + si + ',' + fi + ')" style="padding:2px 6px;">&times;</button>';
                    if (f.sourceField === 'SubGroups' && (state.availableSubGroups || []).length > 0) {
                        var fid = 'rrSgFilter_' + si + '_' + fi;
                        var sgFilter2 = f.subgroupFilter || [];
                        var sgSet = {};
                        for (var sgi = 0; sgi < sgFilter2.length; sgi++) sgSet[sgFilter2[sgi]] = true;
                        html += '<div id="' + fid + '" style="display:none;width:100%;margin-top:4px;padding:6px 8px;background:#f0f4ff;border:1px solid #cbd5e0;border-radius:4px;">';
                        html += '<div style="font-size:12px;font-weight:600;margin-bottom:4px;color:#2c5282;">Show only these subgroups (none = all):</div>';
                        html += '<label style="font-size:11px;display:inline-block;margin-right:10px;cursor:pointer;"><input type="checkbox"' + (sgFilter2.length === 0 ? ' checked' : '') + ' onchange="setSubGroupFilterAll(' + si + ',' + fi + ',this.checked)"> <strong>All</strong></label>';
                        for (var asgi = 0; asgi < state.availableSubGroups.length; asgi++) {
                            var asg = state.availableSubGroups[asgi];
                            html += '<label style="font-size:11px;display:inline-block;margin-right:10px;cursor:pointer;white-space:nowrap;"><input type="checkbox" data-sg="' + escAttr(asg) + '"' + (sgSet[asg] ? ' checked' : '') + ' onchange="toggleSubGroupPick(' + si + ',' + fi + ',this.dataset.sg,this.checked)"> ' + escHtml(asg) + '</label>';
                        }
                        html += '</div>';
                    }
                    html += '</div>';
                }
            }
            html += '<div class="rr-add-field-row">';
            html += '<select id="rrAddField_' + si + '" onchange="onFieldSelectChange(' + si + ')">';
            html += '<option value="">+ Add a field...</option>';
            html += buildFieldOptions();
            html += '</select>';
            html += '<button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addFieldToSection(' + si + ')">Add</button>';
            html += '</div>';
            html += '<div id="rrEvInputs_' + si + '" style="display:none;padding:6px 8px;background:#f0f4ff;border:1px solid #cbd5e0;border-radius:4px;margin-top:4px;">';
            html += '<div style="font-size:12px;font-weight:600;margin-bottom:4px;color:#2c5282;">Add Extra Value Field</div>';
            html += '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">';
            html += '<input type="text" id="rrEvName_' + si + '" placeholder="Extra Value name (exact)" style="flex:1;min-width:140px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:3px;font-size:12px;">';
            html += '<input type="text" id="rrEvLabel_' + si + '" placeholder="Display label (optional)" style="flex:1;min-width:140px;padding:4px 6px;border:1px solid #cbd5e0;border-radius:3px;font-size:12px;">';
            html += '<button class="rr-btn rr-btn-primary rr-btn-sm" onclick="addEvField(' + si + ')">Add</button>';
            html += '<button class="rr-btn rr-btn-sm" onclick="hideEvInputs(' + si + ')" style="padding:2px 8px;">Cancel</button>';
            html += '</div></div>';
            html += '</div></div>';
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
        html += '</optgroup><optgroup label="Other"><option value="extravalue|_prompt_|">Extra Value Field...</option><option value="static|text|">Custom Text / Note</option><option value="static|separator|---">Separator Line</option></optgroup>';
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
    window.onFieldSelectChange = function(si) {
        var sel = document.getElementById('rrAddField_' + si);
        var evPanel = document.getElementById('rrEvInputs_' + si);
        if (!sel || !evPanel) return;
        var parts = sel.value.split('|');
        if (parts[0] === 'extravalue' && parts[1] === '_prompt_') {
            evPanel.style.display = 'block';
            sel.value = '';
            var nameInput = document.getElementById('rrEvName_' + si);
            if (nameInput) nameInput.focus();
        } else {
            evPanel.style.display = 'none';
        }
    };
    window.hideEvInputs = function(si) {
        var evPanel = document.getElementById('rrEvInputs_' + si);
        if (evPanel) evPanel.style.display = 'none';
        var nameInput = document.getElementById('rrEvName_' + si);
        var labelInput = document.getElementById('rrEvLabel_' + si);
        if (nameInput) nameInput.value = '';
        if (labelInput) labelInput.value = '';
    };
    window.addEvField = function(si) {
        if (!state.template || !state.template.sections[si]) return;
        var nameInput = document.getElementById('rrEvName_' + si);
        var labelInput = document.getElementById('rrEvLabel_' + si);
        var evName = nameInput ? nameInput.value.trim() : '';
        if (!evName) { if (nameInput) nameInput.focus(); return; }
        var friendlyLabel = labelInput ? labelInput.value.trim() : '';
        if (!friendlyLabel) friendlyLabel = evName;
        fieldIdCounter++;
        var fields = state.template.sections[si].fields || [];
        fields.push({ fieldId: 'fld_' + fieldIdCounter, fieldType: 'extravalue', sourceField: evName, label: friendlyLabel, displayFormat: 'single-line', order: fields.length + 1, visible: true, colSpan: 1 });
        state.template.sections[si].fields = fields;
        openSections[si] = true;
        hideEvInputs(si);
        renderSections();
    };
    window.addFieldToSection = function(si) {
        if (!state.template || !state.template.sections[si]) return;
        var sel = document.getElementById('rrAddField_' + si);
        if (!sel || !sel.value) return;
        var parts = sel.value.split('|');
        if (parts.length < 2) return;

        // Extra Value is handled by the inline panel, skip here
        if (parts[0] === 'extravalue') return;

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
        sel.value = '';
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

    // ----- Subgroup filter (Option 1: display-only filter) -----
    window.toggleSubGroupFilter = function(si, fi) {
        var el = document.getElementById('rrSgFilter_' + si + '_' + fi);
        if (el) el.style.display = (el.style.display === 'none' || !el.style.display) ? 'block' : 'none';
    };
    window.setSubGroupFilterAll = function(si, fi, checked) {
        if (!state.template || !state.template.sections[si] || !state.template.sections[si].fields[fi]) return;
        if (checked) {
            state.template.sections[si].fields[fi].subgroupFilter = [];
            renderSections();
        }
    };
    window.toggleSubGroupPick = function(si, fi, sgName, checked) {
        if (!state.template || !state.template.sections[si] || !state.template.sections[si].fields[fi]) return;
        var f = state.template.sections[si].fields[fi];
        var arr = f.subgroupFilter || [];
        var idx = arr.indexOf(sgName);
        if (checked && idx < 0) arr.push(sgName);
        if (!checked && idx >= 0) arr.splice(idx, 1);
        f.subgroupFilter = arr;
        // Update the button summary text without collapsing the picker
        renderSections();
        // Re-open the picker that was collapsed by re-render
        var el = document.getElementById('rrSgFilter_' + si + '_' + fi);
        if (el) el.style.display = 'block';
    };

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

    function refreshSaveControls() {
        var sel = document.getElementById('rrSavedTplPicker');
        if (!sel) return;
        var html = '';
        if (!state.savedNames || state.savedNames.length === 0) {
            html = '<option value="">(no saved templates)</option>';
        } else {
            html = '<option value="">-- pick a saved template --</option>';
            for (var i = 0; i < state.savedNames.length; i++) {
                var nm = state.savedNames[i];
                var sel2 = state.currentSavedName === nm ? ' selected' : '';
                html += '<option value="' + escAttr(nm) + '"' + sel2 + '>' + escHtml(nm) + '</option>';
            }
        }
        sel.innerHTML = html;
        var hasCurrent = !!state.currentSavedName;
        var saveBtn = document.getElementById('rrSaveBtn');
        var renameBtn = document.getElementById('rrRenameBtn');
        var deleteBtn = document.getElementById('rrDeleteBtn');
        if (saveBtn) saveBtn.disabled = !hasCurrent;
        if (renameBtn) renameBtn.disabled = !hasCurrent;
        if (deleteBtn) deleteBtn.disabled = !hasCurrent;
    }

    window.onSavedTemplatePicked = function(name) {
        if (!name || !state.savedTemplates[name]) {
            state.currentSavedName = null;
            refreshSaveControls();
            return;
        }
        state.currentSavedName = name;
        state.savedTemplate = state.savedTemplates[name];
        state.template = JSON.parse(JSON.stringify(state.savedTemplate));
        applyLoadedTemplateToUI();
        showToast('Loaded "' + name + '"', 'success');
    };

    window.saveTemplate = function() {
        if (!state.selectedOrgId || !state.template) return;
        // bt_direct routes server-side to per-user PeopleExtra storage.
        var name = state.currentSavedName;
        if (!name) { return saveTemplateAs(); }
        ajax('save_template', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template), tpl_name: name }, function(data) {
            if (data.success) {
                state.savedTemplates[name] = JSON.parse(JSON.stringify(state.template));
                state.savedNames = data.names || state.savedNames;
                state.currentSavedName = name;
                state.savedTemplate = state.savedTemplates[name];
                renderSavedTemplateCards();
                refreshSaveControls();
                showToast('Saved "' + name + '"', 'success');
            } else { showToast('Error: ' + data.message, 'danger'); }
        });
    };

    window.saveTemplateAs = function() {
        if (!state.selectedOrgId || !state.template) return;
        // bt_direct routes server-side to per-user PeopleExtra storage.
        var name = prompt('Save as (template name):', state.currentSavedName || 'My Report');
        if (name === null) return;
        name = (name || '').trim();
        if (!name) { showToast('Name cannot be blank', 'warning'); return; }
        if (state.savedTemplates[name] && !confirm('A template named "' + name + '" already exists. Overwrite?')) return;
        ajax('save_template', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template), tpl_name: name }, function(data) {
            if (data.success) {
                state.savedTemplates[name] = JSON.parse(JSON.stringify(state.template));
                state.savedNames = data.names || state.savedNames;
                state.currentSavedName = name;
                state.savedTemplate = state.savedTemplates[name];
                renderSavedTemplateCards();
                refreshSaveControls();
                showToast('Saved as "' + name + '"', 'success');
            } else { showToast('Error: ' + data.message, 'danger'); }
        });
    };

    window.renameTemplate = function() {
        if (!state.currentSavedName) return;
        var oldName = state.currentSavedName;
        var newName = prompt('Rename "' + oldName + '" to:', oldName);
        if (newName === null) return;
        newName = (newName || '').trim();
        if (!newName || newName === oldName) return;
        ajax('rename_template', { org_id: state.selectedOrgId, tpl_old_name: oldName, tpl_new_name: newName }, function(data) {
            if (data.success) {
                state.savedTemplates[newName] = state.savedTemplates[oldName];
                delete state.savedTemplates[oldName];
                state.savedNames = data.names || state.savedNames;
                state.currentSavedName = newName;
                renderSavedTemplateCards();
                refreshSaveControls();
                showToast('Renamed to "' + newName + '"', 'success');
            } else { showToast('Error: ' + data.message, 'danger'); }
        });
    };

    window.deleteTemplate = function() {
        if (!state.currentSavedName) return;
        var name = state.currentSavedName;
        if (!confirm('Delete saved template "' + name + '"? This cannot be undone.')) return;
        ajax('delete_template', { org_id: state.selectedOrgId, tpl_name: name }, function(data) {
            if (data.success) {
                delete state.savedTemplates[name];
                state.savedNames = data.names || [];
                state.currentSavedName = null;
                state.savedTemplate = null;
                renderSavedTemplateCards();
                refreshSaveControls();
                showToast('Deleted "' + name + '"', 'success');
            } else { showToast('Error: ' + data.message, 'danger'); }
        });
    };

    // ===== Export / Import =====
    window.exportTemplate = function() {
        if (!state.template) { showToast('No template loaded', 'warning'); return; }
        var payload = {
            _exportFormat: 'reportwriter-v1',
            exportedAt: new Date().toISOString(),
            sourceOrgName: state.selectedOrgName || '',
            templateName: state.currentSavedName || 'Untitled',
            template: state.template
        };
        var json = JSON.stringify(payload, null, 2);
        var safeName = (state.currentSavedName || 'template').replace(/[^a-z0-9_\\-]+/gi, '_');
        var orgPart = (state.selectedOrgName || 'org').replace(/[^a-z0-9_\\-]+/gi, '_').substring(0, 30);
        var fname = orgPart + '-' + safeName + '.json';
        try {
            var blob = new Blob([json], {type: 'application/json'});
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url; a.download = fname;
            document.body.appendChild(a); a.click();
            setTimeout(function() { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
        } catch(e) { showToast('Export failed: ' + e.message, 'danger'); }
    };

    window.openImportTemplate = function() {
        var p = document.getElementById('rrImportPanel');
        if (p) p.style.display = (p.style.display === 'block') ? 'none' : 'block';
        var w = document.getElementById('rrImportWarnings');
        if (w) { w.style.display = 'none'; w.innerHTML = ''; }
        var fileEl = document.getElementById('rrImportFile');
        if (fileEl) {
            fileEl.value = '';
            fileEl.onchange = function(e) {
                var f = e.target.files && e.target.files[0];
                if (!f) return;
                var reader = new FileReader();
                reader.onload = function(ev) {
                    document.getElementById('rrImportText').value = ev.target.result || '';
                };
                reader.readAsText(f);
            };
        }
    };

    function _normalizeForMatch(s) { return (s || '').toString().trim().toLowerCase(); }

    function validateImportedTemplate(tpl) {
        // Returns {missingQuestions: [{label, fieldRef}], missingSubgroups: [{name, fieldRef}], missingExtras: [...]}.
        // Auto-remaps subgroup names case-insensitively if a trim-match exists in availableSubGroups.
        var result = {missingQuestions: [], missingSubgroups: [], missingExtras: []};
        if (!tpl || !tpl.sections) return result;

        // Build lookup sets from current org context
        var qKeys = {};
        var qLabels = {};
        for (var qi = 0; qi < (state.questions || []).length; qi++) {
            var qit = state.questions[qi];
            var k = (qit && qit.key) ? qit.key : qit;
            var l = (qit && qit.label) ? qit.label : qit;
            qKeys[_normalizeForMatch(k)] = k;
            qLabels[_normalizeForMatch(l)] = k;
        }
        var sgMap = {};
        for (var sgi = 0; sgi < (state.availableSubGroups || []).length; sgi++) {
            sgMap[_normalizeForMatch(state.availableSubGroups[sgi])] = state.availableSubGroups[sgi];
        }

        for (var si = 0; si < tpl.sections.length; si++) {
            var sec = tpl.sections[si];
            for (var fi = 0; fi < (sec.fields || []).length; fi++) {
                var f = sec.fields[fi];
                if (f.fieldType === 'regquestion' && f.sourceField) {
                    var matched = qKeys[_normalizeForMatch(f.sourceField)] || qLabels[_normalizeForMatch(f.label)] || qLabels[_normalizeForMatch(f.sourceField)];
                    if (matched) {
                        f.sourceField = matched;
                    } else {
                        result.missingQuestions.push({label: f.label || f.sourceField, sectionIdx: si, fieldIdx: fi});
                    }
                } else if (f.sourceField === 'SubGroups' && f.subgroupFilter && f.subgroupFilter.length > 0) {
                    var keptFilter = [];
                    for (var fsi = 0; fsi < f.subgroupFilter.length; fsi++) {
                        var orig = f.subgroupFilter[fsi];
                        var canon = sgMap[_normalizeForMatch(orig)];
                        if (canon) {
                            keptFilter.push(canon);
                        } else {
                            result.missingSubgroups.push({name: orig, sectionIdx: si, fieldIdx: fi});
                            keptFilter.push(orig); // preserve so re-import to original involvement still works
                        }
                    }
                    f.subgroupFilter = keptFilter;
                } else if (f.fieldType === 'extravalue' && f.sourceField) {
                    // We can't easily verify EVs here without an extra round-trip; assume the renderer
                    // will produce empty for missing ones (consistent with existing behavior).
                }
            }
        }
        return result;
    }

    window.doImportTemplate = function() {
        var raw = document.getElementById('rrImportText').value || '';
        raw = raw.trim();
        if (!raw) { showToast('Paste a JSON template or choose a file first', 'warning'); return; }
        var parsed;
        try { parsed = JSON.parse(raw); } catch(e) { showToast('Invalid JSON: ' + e.message, 'danger'); return; }
        // Accept either the export envelope or a bare template
        var tpl = (parsed && parsed._exportFormat === 'reportwriter-v1' && parsed.template) ? parsed.template : parsed;
        if (!tpl || !tpl.sections) { showToast('Not a valid report template', 'danger'); return; }
        var suggestedName = (parsed && parsed.templateName) ? parsed.templateName : 'Imported';
        var nameForUI = prompt('Save imported template as:', suggestedName);
        if (nameForUI === null) return;
        nameForUI = (nameForUI || '').trim();
        if (!nameForUI) { showToast('Name cannot be blank', 'warning'); return; }

        var validation = validateImportedTemplate(tpl);

        // Apply to UI
        state.template = JSON.parse(JSON.stringify(tpl));
        state.currentSavedName = nameForUI;
        applyLoadedTemplateToUI();
        document.getElementById('rrImportPanel').style.display = 'none';

        // Auto-save to server too
        if (state.selectedOrgId && state.selectedOrgId !== 'bt_direct') {
            ajax('save_template', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template), tpl_name: nameForUI }, function(data) {
                if (data.success) {
                    state.savedTemplates[nameForUI] = JSON.parse(JSON.stringify(state.template));
                    state.savedNames = data.names || state.savedNames;
                    state.savedTemplate = state.savedTemplates[nameForUI];
                    renderSavedTemplateCards();
                    refreshSaveControls();
                }
            });
        }

        renderImportWarnings(validation);
        var totalMissing = validation.missingQuestions.length + validation.missingSubgroups.length;
        showToast(totalMissing === 0 ? 'Imported successfully' : 'Imported with ' + totalMissing + ' mismatch(es)', totalMissing === 0 ? 'success' : 'warning');
    };

    function renderImportWarnings(v) {
        var box = document.getElementById('rrImportWarnings');
        if (!box) return;
        if (v.missingQuestions.length === 0 && v.missingSubgroups.length === 0) {
            box.style.display = 'none'; box.innerHTML = '';
            return;
        }
        var h = '<div style="background:#fff7ed;border:1px solid #fb923c;border-radius:6px;padding:10px 14px;">';
        h += '<div style="font-weight:600;color:#9a3412;margin-bottom:6px;"><i class="fa fa-exclamation-triangle"></i> Imported with mismatches</div>';
        h += '<div style="font-size:13px;color:#4a4a4a;margin-bottom:6px;">The fields below reference items that don\\'t exist in this involvement. They will render empty until adjusted.</div>';
        if (v.missingQuestions.length > 0) {
            h += '<div style="margin-top:6px;"><strong>Missing questions (' + v.missingQuestions.length + '):</strong><ul style="margin:4px 0 0 20px;font-size:12px;">';
            for (var i = 0; i < v.missingQuestions.length; i++) {
                h += '<li>' + escHtml(v.missingQuestions[i].label) + '</li>';
            }
            h += '</ul></div>';
        }
        if (v.missingSubgroups.length > 0) {
            var seen = {}; var uniq = [];
            for (var j = 0; j < v.missingSubgroups.length; j++) {
                var n = v.missingSubgroups[j].name;
                if (!seen[n]) { seen[n] = true; uniq.push(n); }
            }
            h += '<div style="margin-top:6px;"><strong>Subgroup filter names not in this involvement (' + uniq.length + '):</strong><ul style="margin:4px 0 0 20px;font-size:12px;">';
            for (var k = 0; k < uniq.length; k++) {
                h += '<li>' + escHtml(uniq[k]) + '</li>';
            }
            h += '</ul></div>';
        }
        h += '<div style="margin-top:8px;display:flex;gap:6px;">';
        h += '<button class="rr-btn rr-btn-sm rr-btn-secondary" onclick="autoRemoveMissingFields()" style="font-size:12px;"><i class="fa fa-eraser"></i> Auto-remove missing fields</button>';
        h += '<button class="rr-btn rr-btn-sm" onclick="document.getElementById(\\'rrImportWarnings\\').style.display=\\'none\\';" style="font-size:12px;padding:4px 10px;">Dismiss</button>';
        h += '</div></div>';
        box.innerHTML = h;
        box.style.display = 'block';
    }

    window.autoRemoveMissingFields = function() {
        if (!state.template) return;
        var qKeys = {};
        for (var qi = 0; qi < (state.questions || []).length; qi++) {
            var qit = state.questions[qi];
            var k = (qit && qit.key) ? qit.key : qit;
            qKeys[_normalizeForMatch(k)] = true;
        }
        var sgMap = {};
        for (var sgi = 0; sgi < (state.availableSubGroups || []).length; sgi++) {
            sgMap[_normalizeForMatch(state.availableSubGroups[sgi])] = true;
        }
        var removedFields = 0;
        for (var si = 0; si < state.template.sections.length; si++) {
            var sec = state.template.sections[si];
            var keep = [];
            for (var fi = 0; fi < (sec.fields || []).length; fi++) {
                var f = sec.fields[fi];
                if (f.fieldType === 'regquestion' && f.sourceField && !qKeys[_normalizeForMatch(f.sourceField)]) {
                    removedFields++; continue;
                }
                if (f.sourceField === 'SubGroups' && f.subgroupFilter && f.subgroupFilter.length > 0) {
                    var keptFilter = [];
                    for (var fsi = 0; fsi < f.subgroupFilter.length; fsi++) {
                        if (sgMap[_normalizeForMatch(f.subgroupFilter[fsi])]) keptFilter.push(f.subgroupFilter[fsi]);
                    }
                    f.subgroupFilter = keptFilter;
                }
                keep.push(f);
            }
            sec.fields = keep;
        }
        renderSections();
        document.getElementById('rrImportWarnings').style.display = 'none';
        showToast('Removed ' + removedFields + ' field(s) referencing missing questions', 'success');
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
                document.getElementById('rrCsvBtn').style.display = 'inline-block';
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
        var allowSplit = ((state.template.printSettings || {}).allowSectionSplit === true);
        css += '.rr-section{margin-bottom:16px' + (allowSplit ? '' : ';page-break-inside:avoid') + '}';
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
        css += '.rr-cover-page{position:relative}';
        css += '.rr-missing-page{position:relative}';
        css += '.rr-medical-page{position:relative}';
        css += '.rr-cover-page table,.rr-missing-page table,.rr-medical-page table{font-size:13px}';
        // Page numbers via @page CSS (actual printed page numbers)
        var showPageNums = ((state.template.printSettings || {}).showPageNumbers === true);
        if (showPageNums) {
            css += '@page{margin-bottom:30px;@bottom-center{content:"Page " counter(page);font-size:10px;color:#999}}';
            // Fallback for browsers that don't support @page @bottom-center
            css += '@media print{body::after{content:"";display:block;height:0}}';
        }
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

    window.exportCsv = function() {
        if (!state.selectedOrgId || !state.template) { showToast('Generate the report first', 'danger'); return; }
        showToast('Exporting CSV...', 'info');
        ajax('export_csv', { org_id: state.selectedOrgId, template_json: JSON.stringify(state.template) }, function(data) {
            if (data.success) {
                var blob = new Blob([data.csv], { type: 'text/csv;charset=utf-8;' });
                var link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = 'report_export.csv';
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                showToast('CSV exported (' + (data.personCount || 0) + ' records)', 'success');
            } else { showToast('CSV export failed: ' + (data.message || 'Unknown error'), 'danger'); }
        });
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

    // --- App version / auto-update ---------------------------------------
    function checkForAppUpdate() {
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', DC_API_BASE + '/script-versions', true);
            xhr.timeout = 5000;
            xhr.onreadystatechange = function() {
                if (xhr.readyState !== 4) return;
                if (xhr.status !== 200) return;
                try {
                    var versions = JSON.parse(xhr.responseText);
                    var latest = versions[DC_SCRIPT_ID];
                    if (latest && latest !== APP_VERSION) {
                        APP_UPDATE_AVAILABLE = true;
                        APP_LATEST_VERSION = latest;
                        renderAppUpdateBanner();
                    }
                } catch(e) {}
            };
            xhr.send();
        } catch(e) {}
    }

    function renderAppUpdateBanner() {
        var b = document.getElementById('rrAppUpdateBanner');
        if (!b || !APP_UPDATE_AVAILABLE) return;
        var h = '';
        h += '<div style="font-size:18px">&#128640;</div>';
        h += '<div style="flex:1;font-size:12px;color:#0078d4">';
        h += '<strong>Report Writer update available</strong>';
        h += ' &mdash; you have <code>v' + escHtml(APP_VERSION) + '</code>, latest is <code>v' + escHtml(APP_LATEST_VERSION) + '</code>. Your saved templates are preserved.';
        h += '</div>';
        h += '<button id="rrAppUpdateBtn" class="rr-btn rr-btn-primary" onclick="applyAppUpdate()" style="white-space:nowrap;">Update Now</button>';
        b.innerHTML = h;
        b.style.display = 'flex';
    }

    window.applyAppUpdate = function() {
        if (!confirm('Update Report Writer from v' + APP_VERSION + ' to v' + APP_LATEST_VERSION + '?\\n\\nYour saved templates and report configurations are stored separately and will be preserved.')) return;
        var btn = document.getElementById('rrAppUpdateBtn');
        if (btn) { btn.disabled = true; btn.innerHTML = 'Updating...'; }
        $.post(scriptUrl, {action: 'apply_update', script_name: SCRIPT_NAME}, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    alert(d.message || 'Updated! Reloading...');
                    window.location.reload(true);
                } else {
                    alert('Update failed: ' + (d.message || 'Unknown error'));
                    if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
                }
            } catch(e) {
                alert('Update failed (could not parse response)');
                if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
            }
        }).fail(function() {
            alert('Update failed (network error)');
            if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
        });
    };

    function initApp() {
        // Kick off background version check; banner appears if update available
        checkForAppUpdate();
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
                    // Load this user's personal templates (PeopleExtra-
                    // backed) so Blue Toolbar runs aren't a clean slate
                    // every time.
                    ajax('load_template', {org_id: 'bt_direct'}, function(tplData) {
                        if (tplData && tplData.success) {
                            state.savedTemplates = tplData.templates || {};
                            state.savedNames = tplData.names || [];
                            state.currentSavedName = tplData.currentName || null;
                            state.savedTemplate = tplData.currentTemplate || null;
                        } else {
                            state.savedTemplates = {};
                            state.savedNames = [];
                            state.currentSavedName = null;
                            state.savedTemplate = null;
                        }
                        renderSavedTemplateCards();
                        selectTemplate('basic');
                        showToast('People loaded! Personal templates available. Person, family & medical fields ready -- select an involvement to add registration questions.', 'success');
                        goToStep(2);
                    });
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
