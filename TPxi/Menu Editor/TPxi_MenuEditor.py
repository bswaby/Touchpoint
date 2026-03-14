#Roles=Admin
# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub:  https://github.com/bswaby/Touchpoint
# ---------------------------------------------------------------
# Support: These tools are free because they should be. If they've
#          saved you time, consider DisplayCache — church digital
#          signage that integrates with TouchPoint.
#          https://displaycache.com
# ---------------------------------------------------------------

"""
TPxi Menu Editor
=================
A visual UI for managing TouchPoint menu XML configurations stored in
Special Content (Text Content).

Manages:
- ReportsMenuAdmin, ReportsMenuFinance, ReportsMenuInvolvements,
  ReportsMenuPeople (column-based report menus)
- CustomReports (blue toolbar and custom report definitions)

Features:
- Tab-based navigation between all 5 XML files
- Add, edit, delete headers and report items
- Reorder items with up/down arrows and drag-and-drop
- Move items between columns
- Set roles/permissions from dropdown
- CustomReports: manage report types, column definitions, org-specific fields
- Raw XML editor for advanced editing
- Dirty state tracking with unsaved changes warning

Version: 1.0
Date: March 2025

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python "TPxi_MenuEditor" and paste all this code
4. Add to CustomReports:
   <Report name="TPxi_MenuEditor" type="PyScript" role="Admin" />
   
   or

   Use this tool to add it
--Upload Instructions End--
"""

import json
import re
import datetime

# --- IronPython Unicode Safety ---

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
    """Safely convert any value to a pure-ASCII JSON-serializable string."""
    if val is None:
        return ''
    try:
        if isinstance(val, unicode):
            return _to_ascii(val)
    except NameError:
        pass
    try:
        return _to_ascii(unicode(val))
    except:
        pass
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
        return ''.join(c if ord(c) < 128 else '?' for c in s)
    except:
        pass
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


# --- Constants ---

CONTENT_NAMES = {
    'admin': 'ReportsMenuAdmin.xml',
    'finance': 'ReportsMenuFinance.xml',
    'involvements': 'ReportsMenuInvolvements.xml',
    'people': 'ReportsMenuPeople.xml',
    'custom': 'CustomReports'
}

def get_all_roles():
    """Pull all roles from TouchPoint's Roles table, with a hardcoded fallback."""
    fallback = [
        '', 'Access', 'Edit', 'Admin', 'SuperAdmin', 'Developer',
        'Finance', 'FinanceAdmin', 'FinanceViewOnly',
        'ManageTransactions', 'PastoralCare', 'FMC',
        'Checkin', 'OrgLeadersOnly', 'ApplicationReview',
        'Attendance', 'ManageGroups', 'Coupon', 'Coupon2',
        'CreditCheck', 'Delete', 'Design', 'ManageEmails',
        'ManageOrgMembers', 'Membership', 'Manager', 'ManageSchedule',
        'ManageTasks', 'ManageVolunteers', 'Reports',
        'Staff', 'Prayer', 'VoIP'
    ]
    try:
        sql = "SELECT RoleName FROM dbo.Roles ORDER BY RoleName"
        rows = q.QuerySql(sql)
        db_roles = ['']
        for row in rows:
            rn = safe_str(row.RoleName).strip()
            if rn:
                db_roles.append(rn)
        if len(db_roles) > 1:
            return db_roles
    except:
        pass
    return fallback

COMMON_ROLES = get_all_roles()

BACKUP_LOG_NAME = 'MenuEditor_BackupLog'
MAX_BACKUPS_PER_CONTENT = 10


# --- Backup / Restore Helpers ---

def get_backup_log():
    """Load the backup log from content storage."""
    try:
        log_str = model.TextContent(BACKUP_LOG_NAME)
        if log_str:
            return json.loads(log_str)
    except:
        pass
    return {'backups': []}

def save_backup_log(log):
    """Save the backup log to content storage."""
    model.WriteContentText(BACKUP_LOG_NAME, json.dumps(sanitize_for_json(log)), "")

def create_backup(content_name, xml_str, user_name):
    """Create a backup of the current content before saving."""
    now = datetime.datetime.now()
    timestamp = now.strftime('%Y%m%d_%H%M%S')
    backup_name = 'MenuEditor_Bak_%s_%s' % (content_name, timestamp)

    # Write the backup content
    model.WriteContentText(backup_name, xml_str, "")

    # Update backup log
    log = get_backup_log()
    entry = {
        'backup_name': backup_name,
        'content_name': content_name,
        'timestamp': now.strftime('%Y-%m-%d %H:%M:%S'),
        'user': user_name
    }
    log['backups'].insert(0, entry)

    # Prune old backups per content name (keep MAX_BACKUPS_PER_CONTENT)
    content_backups = [b for b in log['backups'] if b['content_name'] == content_name]
    if len(content_backups) > MAX_BACKUPS_PER_CONTENT:
        old_backups = content_backups[MAX_BACKUPS_PER_CONTENT:]
        for old in old_backups:
            log['backups'].remove(old)
            # Optionally clear old backup content (write empty)
            try:
                model.WriteContentText(old['backup_name'], '', "")
            except:
                pass

    save_backup_log(log)
    return entry


# --- XML Parsing Helpers ---

def parse_reports_menu(xml_str):
    """Parse ReportsMenu XML into JSON-friendly dict."""
    data = {}
    if not xml_str or not xml_str.strip():
        data = {'Column1': [], 'Column2': [], 'Column3': [], 'Column4': []}
        return data
    try:
        for col_key in ['Column1', 'Column2', 'Column3', 'Column4']:
            pattern = '<%s>(.*?)</%s>' % (col_key, col_key)
            match = re.search(pattern, xml_str, re.DOTALL)
            if not match:
                continue
            data[col_key] = []
            col_content = match.group(1)
            items = re.finditer(r'<(Header|Report)\s*(.*?)>(.*?)</\1>', col_content, re.DOTALL)
            for m in items:
                tag = m.group(1)
                attrs_str = m.group(2).strip()
                text = m.group(3).strip()
                item = {
                    'item_type': 'header' if tag == 'Header' else 'report',
                    'text': safe_str(text),
                    'roles': ''
                }
                roles_match = re.search(r'roles="([^"]*)"', attrs_str)
                if roles_match:
                    item['roles'] = safe_str(roles_match.group(1))
                if tag == 'Report':
                    link_match = re.search(r'link="([^"]*)"', attrs_str)
                    item['link'] = safe_str(link_match.group(1)) if link_match else ''
                data[col_key].append(item)
    except Exception as e:
        pass
    return data


def build_reports_menu_xml(data):
    """Convert JSON dict back to ReportsMenu XML string."""
    lines = ['<ReportsMenu>']
    for col_key in ['Column1', 'Column2', 'Column3', 'Column4']:
        if col_key not in data:
            continue
        items = data.get(col_key, [])
        lines.append('  <%s>' % col_key)
        for item in items:
            item_type = item.get('item_type', 'report')
            text = item.get('text', '')
            roles = item.get('roles', '')
            text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            if item_type == 'header':
                if roles:
                    lines.append('    <Header roles="%s">%s</Header>' % (roles, text))
                else:
                    lines.append('    <Header>%s</Header>' % text)
            else:
                link = item.get('link', '')
                attrs = ''
                if roles:
                    attrs += ' roles="%s"' % roles
                if link:
                    attrs += ' link="%s"' % link
                lines.append('    <Report%s>%s</Report>' % (attrs, text))
        lines.append('  </%s>' % col_key)
    lines.append('</ReportsMenu>')
    return '\n'.join(lines)


def parse_custom_reports(xml_str):
    """Parse CustomReports XML into JSON-friendly list."""
    reports = []
    if not xml_str or not xml_str.strip():
        return reports
    try:
        pos = 0
        while pos < len(xml_str):
            idx = xml_str.find('<Report', pos)
            if idx == -1:
                break
            end_self = xml_str.find('/>', idx)
            end_open = xml_str.find('>', idx)
            if end_self != -1 and (end_open == -1 or end_self <= end_open):
                tag_str = xml_str[idx:end_self + 2]
                report = parse_report_tag(tag_str, '')
                reports.append(report)
                pos = end_self + 2
            elif end_open != -1:
                if xml_str[end_open - 1] == '/':
                    tag_str = xml_str[idx:end_open + 1]
                    report = parse_report_tag(tag_str, '')
                    reports.append(report)
                    pos = end_open + 1
                else:
                    close_idx = xml_str.find('</Report>', end_open)
                    if close_idx != -1:
                        tag_str = xml_str[idx:end_open + 1]
                        inner = xml_str[end_open + 1:close_idx]
                        report = parse_report_tag(tag_str, inner)
                        reports.append(report)
                        pos = close_idx + 9
                    else:
                        pos = end_open + 1
            else:
                pos = idx + 7
    except Exception as e:
        pass
    return reports


def parse_report_tag(tag_str, inner_content):
    """Parse a single Report tag and its inner content."""
    report = {
        'name': '',
        'item_type': '',
        'item_role': '',
        'showOnOrgId': '',
        'url': '',
        'columns': []
    }
    name_match = re.search(r'name="([^"]*)"', tag_str)
    if name_match:
        report['name'] = safe_str(name_match.group(1))
    type_match = re.search(r'type="([^"]*)"', tag_str)
    if type_match:
        report['item_type'] = safe_str(type_match.group(1))
    role_match = re.search(r'role="([^"]*)"', tag_str)
    if role_match:
        report['item_role'] = safe_str(role_match.group(1))
    org_match = re.search(r'showOnOrgId="([^"]*)"', tag_str)
    if org_match:
        report['showOnOrgId'] = safe_str(org_match.group(1))
    url_match = re.search(r'url="([^"]*)"', tag_str)
    if url_match:
        report['url'] = safe_str(url_match.group(1))
    if inner_content and inner_content.strip():
        col_iter = re.finditer(r'<Column\s+(.*?)\s*/>', inner_content)
        for cm in col_iter:
            col_attrs = cm.group(1)
            col = {}
            for attr_match in re.finditer(r'(\w+)="([^"]*)"', col_attrs):
                col[safe_str(attr_match.group(1))] = safe_str(attr_match.group(2))
            report['columns'].append(col)
    return report


def build_custom_reports_xml(reports):
    """Convert JSON list back to CustomReports XML string."""
    lines = ['<CustomReports>']
    for report in reports:
        name = report.get('name', '')
        rtype = report.get('item_type', '')
        role = report.get('item_role', '')
        org_id = report.get('showOnOrgId', '')
        url = report.get('url', '')
        columns = report.get('columns', [])
        attrs = ' name="%s"' % name
        if rtype:
            attrs += ' type="%s"' % rtype
        if role:
            attrs += ' role="%s"' % role
        if org_id:
            attrs += ' showOnOrgId="%s"' % org_id
        if url:
            attrs += ' url="%s"' % url.replace('&', '&amp;')
        if not columns:
            lines.append('  <Report%s />' % attrs)
        else:
            lines.append('  <Report%s>' % attrs)
            for col in columns:
                col_attrs = ''
                for key in ['field', 'smallgroup', 'description', 'flag']:
                    if key in col and col[key]:
                        col_attrs += ' %s="%s"' % (key, col[key])
                if 'name' in col:
                    col_attrs += ' name="%s"' % col['name']
                if 'orgid' in col and col['orgid']:
                    col_attrs += ' orgid="%s"' % col['orgid']
                if 'disabled' in col:
                    col_attrs += ' disabled="%s"' % col['disabled']
                lines.append('    <Column%s />' % col_attrs)
            lines.append('  </Report>')
    lines.append('</CustomReports>')
    return '\n'.join(lines)


# ============================================================================
# AJAX HANDLER
# ============================================================================
model.Header = 'Menu Editor'

if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -------------------------------------------------------------------------
    # Load Content (parse XML to JSON)
    # -------------------------------------------------------------------------
    if action == 'load_content':
        tab_key = str(Data.tab_key) if hasattr(Data, 'tab_key') else ''
        content_name = CONTENT_NAMES.get(tab_key, '')
        if content_name:
            try:
                xml_str = model.TextContent(content_name)
                if not xml_str:
                    xml_str = ''
            except:
                xml_str = ''
            if tab_key == 'custom':
                data = parse_custom_reports(xml_str)
            else:
                data = parse_reports_menu(xml_str)
            print json.dumps(sanitize_for_json({'success': True, 'data': data}))
        else:
            print json.dumps({'success': False, 'message': 'Invalid tab'})

    # -------------------------------------------------------------------------
    # Save Content (JSON back to XML) - creates backup first
    # -------------------------------------------------------------------------
    elif action == 'save_content':
        tab_key = str(Data.tab_key) if hasattr(Data, 'tab_key') else ''
        content_name = CONTENT_NAMES.get(tab_key, '')
        json_data = str(Data.json_data) if hasattr(Data, 'json_data') else '{}'
        # Decode HTML entities from client-side encoding
        json_data = json_data.replace('&lt;', '<').replace('&gt;', '>')
        json_data = json_data.replace('&quot;', '"').replace('&#39;', "'")
        json_data = json_data.replace('&amp;', '&')
        try:
            # Get current user for audit trail
            user_name = ''
            try:
                uid = model.UserPeopleId
                if uid:
                    p = model.GetPerson(uid)
                    if p:
                        user_name = safe_str(p.Name)
            except:
                user_name = 'Unknown'

            # Backup current content before overwriting
            try:
                current_xml = model.TextContent(content_name)
                if current_xml and current_xml.strip():
                    create_backup(content_name, current_xml, user_name)
            except:
                pass  # Don't fail the save if backup fails

            data = json.loads(json_data)
            if tab_key == 'custom':
                xml_str = build_custom_reports_xml(data)
            else:
                xml_str = build_reports_menu_xml(data)
            model.WriteContentText(content_name, xml_str, "")
            print json.dumps({'success': True, 'message': 'Saved successfully (backup created)'})
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(str(e))})

    # -------------------------------------------------------------------------
    # Load Raw XML
    # -------------------------------------------------------------------------
    elif action == 'load_raw':
        tab_key = str(Data.tab_key) if hasattr(Data, 'tab_key') else ''
        content_name = CONTENT_NAMES.get(tab_key, '')
        try:
            xml_str = model.TextContent(content_name)
            if not xml_str:
                xml_str = ''
        except:
            xml_str = ''
        print json.dumps(sanitize_for_json({'success': True, 'raw': xml_str}))

    # -------------------------------------------------------------------------
    # Save Raw XML - creates backup first
    # -------------------------------------------------------------------------
    elif action == 'save_raw':
        tab_key = str(Data.tab_key) if hasattr(Data, 'tab_key') else ''
        content_name = CONTENT_NAMES.get(tab_key, '')
        raw_xml = str(Data.raw_xml) if hasattr(Data, 'raw_xml') else ''
        raw_xml = raw_xml.replace('&lt;', '<').replace('&gt;', '>')
        raw_xml = raw_xml.replace('&quot;', '"').replace('&#39;', "'")
        raw_xml = raw_xml.replace('&amp;', '&')
        try:
            # Get current user for audit trail
            user_name = ''
            try:
                uid = model.UserPeopleId
                if uid:
                    p = model.GetPerson(uid)
                    if p:
                        user_name = safe_str(p.Name)
            except:
                user_name = 'Unknown'

            # Backup current content before overwriting
            try:
                current_xml = model.TextContent(content_name)
                if current_xml and current_xml.strip():
                    create_backup(content_name, current_xml, user_name)
            except:
                pass

            model.WriteContentText(content_name, raw_xml, "")
            print json.dumps({'success': True, 'message': 'Saved raw XML (backup created)'})
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(str(e))})

    # -------------------------------------------------------------------------
    # List Backups for a content name
    # -------------------------------------------------------------------------
    elif action == 'list_backups':
        tab_key = str(Data.tab_key) if hasattr(Data, 'tab_key') else ''
        content_name = CONTENT_NAMES.get(tab_key, '')
        try:
            log = get_backup_log()
            backups = [b for b in log['backups'] if b['content_name'] == content_name]
            print json.dumps(sanitize_for_json({'success': True, 'backups': backups}))
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(str(e))})

    # -------------------------------------------------------------------------
    # Restore from Backup
    # -------------------------------------------------------------------------
    elif action == 'restore_backup':
        backup_name = str(Data.backup_name) if hasattr(Data, 'backup_name') else ''
        tab_key = str(Data.tab_key) if hasattr(Data, 'tab_key') else ''
        content_name = CONTENT_NAMES.get(tab_key, '')
        try:
            if not backup_name or not content_name:
                print json.dumps({'success': False, 'message': 'Missing backup name or tab'})
            else:
                # Get current user
                user_name = ''
                try:
                    uid = model.UserPeopleId
                    if uid:
                        p = model.GetPerson(uid)
                        if p:
                            user_name = safe_str(p.Name)
                except:
                    user_name = 'Unknown'

                # Backup the CURRENT content before restoring (safety net)
                try:
                    current_xml = model.TextContent(content_name)
                    if current_xml and current_xml.strip():
                        create_backup(content_name, current_xml, user_name + ' (pre-restore)')
                except:
                    pass

                # Read backup content and restore
                backup_xml = model.TextContent(backup_name)
                if backup_xml:
                    model.WriteContentText(content_name, backup_xml, "")
                    print json.dumps({'success': True, 'message': 'Restored from backup'})
                else:
                    print json.dumps({'success': False, 'message': 'Backup content is empty or not found'})
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(str(e))})

    # -------------------------------------------------------------------------
    # Preview Backup Content
    # -------------------------------------------------------------------------
    elif action == 'preview_backup':
        backup_name = str(Data.backup_name) if hasattr(Data, 'backup_name') else ''
        try:
            backup_xml = model.TextContent(backup_name)
            if backup_xml is None:
                backup_xml = ''
            print json.dumps(sanitize_for_json({'success': True, 'raw': backup_xml}))
        except Exception as e:
            print json.dumps({'success': False, 'message': safe_str(str(e))})

    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + action})

# ============================================================================
# MAIN PAGE RENDER (GET request)
# ============================================================================
else:
    # Build role options for dropdowns
    role_opts = ''
    for r in COMMON_ROLES:
        label = r if r else '(none)'
        role_opts += '<option value="%s">%s</option>' % (r, label)

    type_opts = '<option value="">Custom Report (column-based data export)</option>'
    for t in [('PyScript', 'Python Script'), ('SqlReport', 'SQL Report'), ('OrgSearchSqlReport', 'Org Search SQL Report'), ('URL', 'URL Link')]:
        type_opts += '<option value="%s">%s</option>' % (t[0], t[1])

    html = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Auto-redirect to PyScriptForm (required for AJAX to work) -->
    <script>
    if (window.location.pathname.indexOf('/PyScript/') > -1) {
        window.location.href = window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search;
    }
    </script>
    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <style>
/* Menu Editor - Scoped CSS (me- prefix) */
:root {
    --me-primary: #2196F3;
    --me-success: #4CAF50;
    --me-danger: #f44336;
    --me-warning: #ff9800;
    --me-dark: #333;
    --me-light-bg: #f8f9fa;
    --me-border: #ddd;
}
.me-root {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    max-width: 1400px;
    margin: 0 auto;
    padding: 10px 20px;
}

/* Tabs */
.me-tabs {
    display: flex;
    gap: 2px;
    margin-bottom: 0;
    border-bottom: 3px solid var(--me-primary);
}
.me-tab {
    padding: 10px 20px;
    cursor: pointer;
    background: #e0e0e0;
    border-radius: 6px 6px 0 0;
    font-weight: 600;
    font-size: 14px;
    transition: all 0.2s;
    position: relative;
    user-select: none;
}
.me-tab:hover { background: #bbdefb; }
.me-tab.active { background: var(--me-primary); color: #fff; }
.me-tab .me-dirty-dot {
    display: none;
    width: 8px; height: 8px;
    background: var(--me-danger);
    border-radius: 50%;
    position: absolute;
    top: 6px; right: 6px;
}
.me-tab.dirty .me-dirty-dot { display: block; }

/* Toolbar */
.me-toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 16px;
    background: var(--me-light-bg);
    border: 1px solid var(--me-border);
    border-top: none;
}
.me-toolbar-left { display: flex; gap: 8px; align-items: center; }
.me-toolbar-right { display: flex; gap: 8px; }

/* Buttons */
.me-btn {
    padding: 6px 14px;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
    font-weight: 500;
    transition: all 0.15s;
    display: inline-flex;
    align-items: center;
    gap: 4px;
}
.me-btn-primary { background: #2196F3 !important; color: #fff !important; border: 1px solid #1976D2 !important; }
.me-btn-primary:hover { background: #1565C0 !important; }
.me-btn-success { background: #4CAF50 !important; color: #fff !important; border: 1px solid #388E3C !important; }
.me-btn-success:hover { background: #2E7D32 !important; }
.me-btn-danger { background: #f44336 !important; color: #fff !important; border: 1px solid #d32f2f !important; }
.me-btn-danger:hover { background: #c62828 !important; }
.me-btn-outline { background: #fff; border: 1px solid #ccc; color: #333; }
.me-btn-outline:hover { background: #f0f0f0; }
.me-btn-sm { padding: 3px 8px; font-size: 11px; }
.me-btn-icon { padding: 4px 7px; font-size: 13px; line-height: 1; min-width: 26px; text-align: center; }

/* Content area */
.me-content {
    border: 1px solid var(--me-border);
    border-top: none;
    min-height: 400px;
    background: #fff;
}
.me-loading { text-align: center; padding: 60px; color: #999; font-size: 16px; }
.me-empty { text-align: center; padding: 30px; color: #bbb; font-size: 13px; font-style: italic; }

/* Column sections (vertical stacked layout) */
.me-columns {
    padding: 12px;
}
.me-column-section {
    border: 1px solid #e0e0e0;
    border-radius: 6px;
    margin-bottom: 12px;
    background: #fff;
    overflow: hidden;
}
.me-col-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 14px;
    background: #f0f4f8;
    border-bottom: 1px solid #e0e0e0;
    cursor: pointer;
}
.me-col-header:hover { background: #e8eef4; }
.me-col-title { font-weight: 700; font-size: 14px; color: #444; }
.me-col-title .me-col-count { font-weight: 400; font-size: 12px; color: #888; margin-left: 6px; }
.me-col-actions { display: flex; gap: 4px; }
.me-col-body { padding: 8px; }
.me-col-body.collapsed { display: none; }

/* Menu items - horizontal row layout */
.me-item {
    display: flex;
    align-items: center;
    border: 1px solid #e8e8e8;
    border-radius: 4px;
    margin-bottom: 4px;
    padding: 6px 12px;
    background: #fff;
    transition: all 0.15s;
    cursor: grab;
    gap: 10px;
}
.me-item:hover { border-color: #90caf9; background: #f8fbff; }
.me-item-header { background: #e3f2fd; border-color: #bbdefb; }
.me-item-type-col {
    flex: 0 0 24px;
    text-align: center;
}
.me-item-text-col {
    flex: 1 1 auto;
    min-width: 0;
    font-size: 13px;
    font-weight: 600;
    color: #333;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.me-item-link-col {
    flex: 0 1 280px;
    min-width: 0;
    font-size: 11px;
    color: #1976D2;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.me-item-link-col a { color: #1976D2; text-decoration: none; }
.me-item-link-col a:hover { text-decoration: underline; }
.me-item-roles-col {
    flex: 0 0 auto;
}
.me-item-actions-col {
    flex: 0 0 auto;
    display: flex;
    gap: 2px;
    white-space: nowrap;
}
.me-badge {
    display: inline-block;
    padding: 2px 7px;
    border-radius: 3px;
    font-size: 10px;
    font-weight: 600;
    background: #e8eaf6;
    color: #3f51b5;
    white-space: nowrap;
}
.me-badge-type { background: #fff3e0; color: #e65100; }
.me-badge-h { background: #e3f2fd; color: #1565c0; font-size: 9px; padding: 1px 5px; }
.me-badge-r { background: #f3e5f5; color: #7b1fa2; font-size: 9px; padding: 1px 5px; }

/* Drag styles */
.me-item.dragging { opacity: 0.4; }
.me-item.drag-over { border-top: 3px solid var(--me-primary); margin-top: -3px; }
.me-col-body.drag-over-col { background: #e3f2fd; }

/* Column move dropdown */
.me-move-select {
    padding: 2px 4px;
    font-size: 11px;
    border: 1px solid #ccc;
    border-radius: 3px;
    background: #fff;
    cursor: pointer;
}

/* Custom Reports table */
.me-cr-table { width: 100%; border-collapse: collapse; }
.me-cr-table th {
    text-align: left;
    padding: 10px 12px;
    background: #f5f5f5;
    border-bottom: 2px solid var(--me-border);
    font-size: 13px;
    font-weight: 600;
    color: #555;
}
.me-cr-table td { padding: 8px 12px; border-bottom: 1px solid #eee; font-size: 13px; vertical-align: top; }
.me-cr-table tr:hover { background: #fafafa; }
.me-cr-expand { cursor: pointer; color: var(--me-primary); font-size: 14px; }
.me-cr-columns { background: #f9f9f9; padding: 8px 12px 8px 40px; }
.me-cr-col-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 4px 8px;
    margin: 2px 0;
    background: #fff;
    border: 1px solid #eee;
    border-radius: 3px;
    font-size: 12px;
}
.me-cr-col-item:hover { border-color: #90caf9; }

/* Blue Toolbar group headers */
.me-bt-group {
    margin-bottom: 16px;
    border: 1px solid #ddd;
    border-radius: 6px;
    overflow: hidden;
}
.me-bt-group-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 16px;
    background: #1a3a5c;
    color: #fff;
    font-size: 15px;
    font-weight: 700;
    cursor: pointer;
    user-select: none;
}
.me-bt-group-header:hover { background: #224a72; }
.me-bt-group-header .me-bt-count {
    font-size: 12px;
    font-weight: 400;
    background: rgba(255,255,255,0.2);
    padding: 2px 8px;
    border-radius: 10px;
}
.me-bt-group-body { display: block; }
.me-bt-group.collapsed .me-bt-group-body { display: none; }
.me-bt-group-header .me-bt-chevron { transition: transform 0.2s; }
.me-bt-group.collapsed .me-bt-chevron { transform: rotate(-90deg); }

/* Duplicate detection */
.me-dup-banner {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 10px 16px;
    margin-bottom: 12px;
    background: #fff3cd;
    border: 1px solid #ffc107;
    border-left: 4px solid #ff9800;
    border-radius: 4px;
    font-size: 13px;
    color: #664d03;
}
.me-dup-banner i { font-size: 18px; color: #ff9800; }
.me-dup-banner strong { color: #856404; }
.me-dup-row { background: #fff8e1 !important; }
.me-dup-row:hover { background: #fff3cd !important; }
.me-dup-badge {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    padding: 2px 8px;
    background: #ff9800;
    color: #fff;
    border-radius: 10px;
    font-size: 11px;
    font-weight: 600;
    margin-left: 8px;
}
.me-dup-badge i { font-size: 10px; }

/* Modal */
.me-modal-overlay {
    display: none;
    position: fixed;
    top: 0; left: 0;
    width: 100%; height: 100%;
    background: rgba(0,0,0,0.5);
    z-index: 9999;
    justify-content: center;
    align-items: flex-start;
    padding-top: 80px;
}
.me-modal-overlay.show { display: flex; }
.me-modal {
    background: #fff;
    border-radius: 8px;
    width: 550px;
    max-height: 80vh;
    overflow-y: auto;
    box-shadow: 0 8px 32px rgba(0,0,0,0.2);
}
.me-modal-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 14px 20px;
    border-bottom: 1px solid #eee;
}
.me-modal-title { font-size: 16px; font-weight: 700; }
.me-modal-close { cursor: pointer; font-size: 22px; color: #999; background: none; border: none; line-height: 1; }
.me-modal-body { padding: 20px; }
.me-modal-footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    padding: 14px 20px;
    border-top: 1px solid #eee;
    background: #f8f9fa;
}
.me-modal-footer .me-btn { padding: 8px 20px; font-size: 14px; font-weight: 600; }

/* Form fields */
.me-field { margin-bottom: 14px; }
.me-field label { display: block; font-size: 12px; font-weight: 600; color: #555; margin-bottom: 4px; }
.me-field input, .me-field select, .me-field textarea {
    width: 100%;
    padding: 8px 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    font-size: 13px;
    box-sizing: border-box;
}
.me-field input:focus, .me-field select:focus, .me-field textarea:focus {
    border-color: var(--me-primary);
    outline: none;
    box-shadow: 0 0 0 2px rgba(33,150,243,0.15);
}
.me-field .me-help { font-size: 11px; color: #999; margin-top: 2px; }

/* Toast */
.me-toast {
    position: fixed;
    bottom: 24px; right: 24px;
    padding: 12px 24px;
    border-radius: 6px;
    color: #fff;
    font-size: 14px;
    font-weight: 500;
    z-index: 99999;
    opacity: 0;
    transition: opacity 0.3s;
    pointer-events: none;
}
.me-toast.show { opacity: 1; }
.me-toast-success { background: var(--me-success); }
.me-toast-error { background: var(--me-danger); }

/* Backup list */
.me-backup-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 12px;
    border: 1px solid #eee;
    border-radius: 4px;
    margin-bottom: 6px;
    background: #fff;
}
.me-backup-item:hover { background: #f8fbff; border-color: #90caf9; }
.me-backup-info { flex: 1; }
.me-backup-date { font-weight: 600; font-size: 13px; color: #333; }
.me-backup-user { font-size: 12px; color: #666; margin-top: 2px; }
.me-backup-actions { display: flex; gap: 6px; }

/* Tutorial */
.me-tutorial-step {
    display: flex;
    gap: 14px;
    padding: 14px 0;
    border-bottom: 1px solid #eee;
}
.me-tutorial-step:last-child { border-bottom: none; }
.me-tutorial-num {
    flex-shrink: 0;
    width: 32px; height: 32px;
    background: #1a3a5c;
    color: #fff;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 14px;
}
.me-tutorial-body h4 { margin: 0 0 4px 0; font-size: 14px; color: #333; }
.me-tutorial-body p { margin: 0; font-size: 13px; color: #555; line-height: 1.5; }
.me-tutorial-body .me-tip {
    margin-top: 6px;
    padding: 6px 10px;
    background: #fff3cd;
    border-left: 3px solid #ff9800;
    border-radius: 3px;
    font-size: 12px;
    color: #664d03;
}
.me-tutorial-section {
    margin-top: 16px;
    padding-top: 12px;
    border-top: 2px solid #1a3a5c;
}
.me-tutorial-section h3 {
    margin: 0 0 8px 0;
    font-size: 15px;
    color: #1a3a5c;
}
.me-tutorial-kbd {
    display: inline-block;
    padding: 1px 6px;
    background: #eee;
    border: 1px solid #ccc;
    border-radius: 3px;
    font-family: monospace;
    font-size: 11px;
}

/* Raw editor */
.me-raw-editor {
    width: 100%;
    min-height: 400px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    padding: 12px;
    border: 1px solid #ccc;
    border-radius: 4px;
    box-sizing: border-box;
    tab-size: 2;
    white-space: pre;
}
    </style>
</head>
<body>

<div class="me-root" id="meRoot">

    <!-- Tabs -->
    <div class="me-tabs" id="meTabs">
        <div class="me-tab active" data-tab="admin" onclick="switchTab('admin')"><i class="bi bi-gear"></i> Admin<span class="me-dirty-dot"></span></div>
        <div class="me-tab" data-tab="finance" onclick="switchTab('finance')"><i class="bi bi-currency-dollar"></i> Finance<span class="me-dirty-dot"></span></div>
        <div class="me-tab" data-tab="involvements" onclick="switchTab('involvements')"><i class="bi bi-people"></i> Involvements<span class="me-dirty-dot"></span></div>
        <div class="me-tab" data-tab="people" onclick="switchTab('people')"><i class="bi bi-person"></i> People<span class="me-dirty-dot"></span></div>
        <div class="me-tab" data-tab="custom" onclick="switchTab('custom')"><i class="bi bi-wrench"></i> Blue Toolbar<span class="me-dirty-dot"></span></div>
    </div>

    <!-- Toolbar -->
    <div class="me-toolbar">
        <div class="me-toolbar-left">
            <span id="meStatus" style="font-size:12px;color:#888;"></span>
        </div>
        <div class="me-toolbar-right">
            <button class="me-btn me-btn-outline" onclick="showTutorial()" title="How to use this tool"><i class="bi bi-question-circle"></i> Help</button>
            <button class="me-btn me-btn-outline" onclick="showBackups()"><i class="bi bi-clock-history"></i> Backups</button>
            <button class="me-btn me-btn-outline" onclick="toggleRawEditor()" id="btnRaw"><i class="bi bi-code-slash"></i> Raw XML</button>
            <button class="me-btn me-btn-outline" onclick="reloadContent()"><i class="bi bi-arrow-clockwise"></i> Reload</button>
            <button class="me-btn me-btn-outline" onclick="discardChanges()" id="btnDiscard" style="display:none;color:#f44336;border-color:#f44336;"><i class="bi bi-x-lg"></i> Discard Changes</button>
            <button class="me-btn me-btn-success" onclick="saveContent()" id="btnSave"><i class="bi bi-check-lg"></i> Save Changes</button>
        </div>
    </div>

    <!-- Content Area -->
    <div class="me-content" id="meContent">
        <div class="me-loading"><i class="bi bi-hourglass-split"></i> Loading...</div>
    </div>

</div><!-- End me-root -->

<!-- ==================== MODALS ==================== -->

<!-- Item Edit Modal (for ReportsMenu headers/reports) -->
<div class="me-modal-overlay" id="modalItem">
    <div class="me-modal">
        <div class="me-modal-header">
            <span class="me-modal-title" id="modalItemTitle">Edit Item</span>
            <button class="me-modal-close" onclick="closeModal('modalItem')">&times;</button>
        </div>
        <div class="me-modal-body">
            <input type="hidden" id="modalItemCol" />
            <input type="hidden" id="modalItemIdx" />
            <input type="hidden" id="modalItemType" />
            <div class="me-field">
                <label>Display Text</label>
                <input type="text" id="modalItemText" placeholder="Menu item text" />
            </div>
            <div class="me-field" id="fieldLink" style="display:none;">
                <label>Link URL</label>
                <input type="text" id="modalItemLink" placeholder="/PyScript/ScriptName" />
                <div class="me-help">e.g. /PyScript/MyScript or /RunScript/MyScript/</div>
            </div>
            <div class="me-field">
                <label>Roles (comma-separated or pick from list)</label>
                <select id="modalItemRoleSelect" onchange="addRoleFromSelect('modalItemRoleSelect','modalItemRoles')">
                    <option value="">-- Add a role --</option>
                    ''' + role_opts + '''
                </select>
                <input type="text" id="modalItemRoles" placeholder="Admin,Edit" style="margin-top:4px;" />
                <div class="me-help">Leave blank for no role restriction</div>
            </div>
        </div>
        <div class="me-modal-footer">
            <button class="me-btn me-btn-outline" onclick="closeModal('modalItem')">Cancel</button>
            <button class="me-btn me-btn-primary" onclick="saveItemModal()"><i class="bi bi-check-lg"></i> Save</button>
        </div>
    </div>
</div>

<!-- Custom Report Edit Modal -->
<div class="me-modal-overlay" id="modalReport">
    <div class="me-modal">
        <div class="me-modal-header">
            <span class="me-modal-title" id="modalReportTitle">Edit Report</span>
            <button class="me-modal-close" onclick="closeModal('modalReport')">&times;</button>
        </div>
        <div class="me-modal-body">
            <input type="hidden" id="modalReportIdx" />
            <div class="me-field">
                <label>Report Name</label>
                <input type="text" id="modalReportName" placeholder="MyReport" />
            </div>
            <div class="me-field">
                <label>Type</label>
                <select id="modalReportType" onchange="updateTypeHint()">
                    ''' + type_opts + '''
                </select>
                <div class="me-help" id="typeHint">This determines which group the report appears under in the Blue Toolbar.</div>
            </div>
            <div class="me-field">
                <label>Role</label>
                <select id="modalReportRoleSelect" onchange="addRoleFromSelect('modalReportRoleSelect','modalReportRole')">
                    <option value="">-- Select a role --</option>
                    ''' + role_opts + '''
                </select>
                <input type="text" id="modalReportRole" placeholder="Edit" style="margin-top:4px;" />
            </div>
            <div class="me-field">
                <label>Show On Org ID</label>
                <input type="text" id="modalReportOrgId" placeholder="e.g. 123 (leave blank for all)" />
                <div class="me-help">Only shows this report when viewing this specific org</div>
            </div>
            <div class="me-field" id="fieldUrl">
                <label>URL (for URL type only)</label>
                <input type="text" id="modalReportUrl" placeholder="/Export2/..." />
            </div>
        </div>
        <div class="me-modal-footer">
            <button class="me-btn me-btn-outline" onclick="closeModal('modalReport')">Cancel</button>
            <button class="me-btn me-btn-primary" onclick="saveReportModal()"><i class="bi bi-check-lg"></i> Save</button>
        </div>
    </div>
</div>

<!-- Column Edit Modal (for CustomReports column definitions) -->
<div class="me-modal-overlay" id="modalColumn">
    <div class="me-modal">
        <div class="me-modal-header">
            <span class="me-modal-title" id="modalColumnTitle">Edit Column</span>
            <button class="me-modal-close" onclick="closeModal('modalColumn')">&times;</button>
        </div>
        <div class="me-modal-body">
            <input type="hidden" id="modalColReportIdx" />
            <input type="hidden" id="modalColIdx" />
            <div class="me-field">
                <label>Column Type</label>
                <select id="modalColType" onchange="updateColFields()">
                    <option value="simple">Simple Field (Name, Age, etc.)</option>
                    <option value="ExtraValueText">Extra Value Text</option>
                    <option value="FamilyExtraValueText">Family Extra Value Text</option>
                    <option value="SmallGroup">Small Group</option>
                    <option value="StatusFlag">Status Flag</option>
                    <option value="OrgField">Org-Specific Field (AmountDue, etc.)</option>
                </select>
            </div>
            <div class="me-field" id="colFieldName">
                <label>Column Name</label>
                <input type="text" id="modalColName" placeholder="First, Last, Age, AmountDue, etc." />
            </div>
            <div class="me-field" id="colFieldField" style="display:none;">
                <label>Field Name</label>
                <input type="text" id="modalColField" placeholder="Extra value field name" />
            </div>
            <div class="me-field" id="colFieldSmallgroup" style="display:none;">
                <label>Small Group Name</label>
                <input type="text" id="modalColSmallgroup" placeholder="Group name" />
            </div>
            <div class="me-field" id="colFieldDescription" style="display:none;">
                <label>Description</label>
                <input type="text" id="modalColDescription" placeholder="Flag description" />
            </div>
            <div class="me-field" id="colFieldFlag" style="display:none;">
                <label>Flag ID</label>
                <input type="text" id="modalColFlag" placeholder="e.g. 22" />
            </div>
            <div class="me-field" id="colFieldOrgId" style="display:none;">
                <label>Org ID</label>
                <input type="text" id="modalColOrgId" placeholder="e.g. 123" />
            </div>
            <div class="me-field" id="colFieldDisabled" style="display:none;">
                <label>Disabled</label>
                <select id="modalColDisabled">
                    <option value="">Not set</option>
                    <option value="true">true</option>
                    <option value="false">false</option>
                </select>
            </div>
        </div>
        <div class="me-modal-footer">
            <button class="me-btn me-btn-outline" onclick="closeModal('modalColumn')">Cancel</button>
            <button class="me-btn me-btn-primary" onclick="saveColumnModal()"><i class="bi bi-check-lg"></i> Save</button>
        </div>
    </div>
</div>

<!-- Confirm Delete Modal -->
<div class="me-modal-overlay" id="modalConfirm">
    <div class="me-modal" style="width:400px;">
        <div class="me-modal-header">
            <span class="me-modal-title">Confirm Delete</span>
            <button class="me-modal-close" onclick="closeModal('modalConfirm')">&times;</button>
        </div>
        <div class="me-modal-body">
            <p id="confirmMessage">Are you sure you want to delete this item?</p>
        </div>
        <div class="me-modal-footer">
            <button class="me-btn me-btn-outline" onclick="closeModal('modalConfirm')">Cancel</button>
            <button class="me-btn me-btn-danger" id="confirmDeleteBtn"><i class="bi bi-trash"></i> Delete</button>
        </div>
    </div>
</div>

<!-- Backup/Restore Modal -->
<div class="me-modal-overlay" id="modalBackups">
    <div class="me-modal" style="width:650px;">
        <div class="me-modal-header">
            <span class="me-modal-title"><i class="bi bi-clock-history"></i> Backup History</span>
            <button class="me-modal-close" onclick="closeModal('modalBackups')">&times;</button>
        </div>
        <div class="me-modal-body" id="backupList" style="max-height:400px;overflow-y:auto;">
            <div class="me-loading">Loading backups...</div>
        </div>
        <div class="me-modal-footer">
            <span style="flex:1;font-size:11px;color:#888;">Backups are created automatically before each save</span>
            <button class="me-btn me-btn-outline" onclick="closeModal('modalBackups')">Close</button>
        </div>
    </div>
</div>

<!-- Backup Preview Modal -->
<div class="me-modal-overlay" id="modalPreview">
    <div class="me-modal" style="width:700px;">
        <div class="me-modal-header">
            <span class="me-modal-title" id="previewTitle">Backup Preview</span>
            <button class="me-modal-close" onclick="closeModal('modalPreview')">&times;</button>
        </div>
        <div class="me-modal-body">
            <textarea class="me-raw-editor" id="previewContent" readonly style="min-height:300px;background:#f9f9f9;"></textarea>
        </div>
        <div class="me-modal-footer">
            <button class="me-btn me-btn-outline" onclick="closeModal('modalPreview')">Close</button>
            <button class="me-btn me-btn-danger" id="previewRestoreBtn"><i class="bi bi-arrow-counterclockwise"></i> Restore This Backup</button>
        </div>
    </div>
</div>

<!-- Tutorial Modal -->
<div class="me-modal-overlay" id="modalTutorial">
    <div class="me-modal" style="width:650px;max-height:85vh;">
        <div class="me-modal-header" style="background:#1a3a5c;color:#fff;">
            <span class="me-modal-title"><i class="bi bi-book" style="margin-right:6px;"></i> Menu Editor Guide</span>
            <button class="me-modal-close" onclick="closeModal('modalTutorial')" style="color:#fff;">&times;</button>
        </div>
        <div class="me-modal-body" style="max-height:70vh;overflow-y:auto;">

            <h3 style="margin:0 0 6px;font-size:15px;color:#1a3a5c;">What is this tool?</h3>
            <p style="font-size:13px;color:#555;margin:0 0 16px;line-height:1.5;">
                The Menu Editor lets you visually manage the XML menu configurations that control what appears in TouchPoint's <strong>Reports menus</strong> (Admin, Finance, Involvements, People) and the <strong>Blue Toolbar</strong> dropdown. No more hand-editing XML!
            </p>

            <div class="me-tutorial-section" style="margin-top:0;">
                <h3>Getting Started</h3>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num">1</div>
                    <div class="me-tutorial-body">
                        <h4>Pick a tab</h4>
                        <p>Click a tab at the top to choose which menu to edit. The first four tabs (Admin, Finance, Involvements, People) manage their respective Reports menus. The <strong>Blue Toolbar</strong> tab manages the toolbar that appears on search results.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num">2</div>
                    <div class="me-tutorial-body">
                        <h4>Browse the items</h4>
                        <p><strong>Reports menus:</strong> Items are organized in collapsible column sections (Column 1-4) matching how they appear in TouchPoint. Blue cards are <strong>Headers</strong> (section titles), white cards are <strong>Reports</strong> (clickable links).</p>
                        <p style="margin-top:4px;"><strong>Blue Toolbar:</strong> Items are grouped by type &mdash; SQL Reports, Python Scripts, Other Reports, and Custom Reports &mdash; just like the toolbar dropdown itself.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num">3</div>
                    <div class="me-tutorial-body">
                        <h4>Make changes, then save</h4>
                        <p>A red dot on the tab means you have unsaved changes. Click <strong>Save Changes</strong> (green button) to write your changes back to TouchPoint. A backup is created automatically before every save.</p>
                        <div class="me-tip"><i class="bi bi-info-circle"></i> After saving, wait about 1 minute then refresh your browser to see the updated menus in TouchPoint.</div>
                    </div>
                </div>
            </div>

            <div class="me-tutorial-section">
                <h3>Reports Menu Tabs (Admin, Finance, Involvements, People)</h3>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-plus"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Adding items</h4>
                        <p>Each column section has <strong>+ Header</strong> and <strong>+ Report</strong> buttons. Headers are section titles (displayed in bold in the menu). Reports are the clickable links underneath.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-pencil"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Editing items</h4>
                        <p>Click the pencil icon on any item to edit its text, link URL, and role restrictions. The role dropdown provides common TouchPoint roles, or you can type custom roles separated by commas.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-arrows-move"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Reordering &amp; moving</h4>
                        <p>Use the <i class="bi bi-arrow-up"></i> <i class="bi bi-arrow-down"></i> arrows to reorder items within a column. Use the <strong>Move to...</strong> dropdown to move an item to a different column.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-trash"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Deleting items</h4>
                        <p>Click the red trash icon to delete an item. You will be asked to confirm before deletion.</p>
                    </div>
                </div>
            </div>

            <div class="me-tutorial-section">
                <h3>Blue Toolbar Tab</h3>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-grid"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Grouped by type</h4>
                        <p>Items are grouped into the same categories TouchPoint uses: <strong>SQL Reports</strong>, <strong>Python Scripts</strong>, <strong>Other Reports</strong> (URL, OrgSearchSqlReport), and <strong>Custom Reports</strong> (column-based data exports). Click a group header to collapse/expand it.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-table"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Custom Reports &amp; columns</h4>
                        <p>Custom Reports define which data columns appear when users export from a search. Click the <i class="bi bi-caret-right-fill"></i> arrow to expand a report and see its columns. You can add, edit, reorder, and remove columns.</p>
                        <p style="margin-top:4px;">Column types include: Simple fields (Name, Email, Age, etc.), ExtraValueText, FamilyExtraValueText, StatusFlag (with flag ID), SmallGroup (with org ID), and more.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-eye"></i></div>
                    <div class="me-tutorial-body">
                        <h4>showOnOrgId</h4>
                        <p>If a Custom Report has a <strong>showOnOrgId</strong>, it only appears in the blue toolbar when viewing that specific organization. Leave blank (or 0) to show it everywhere.</p>
                    </div>
                </div>
            </div>

            <div class="me-tutorial-section">
                <h3>Other Features</h3>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-code-slash"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Raw XML</h4>
                        <p>Click <strong>Raw XML</strong> to switch to a text editor where you can edit the XML directly. Useful for advanced users or bulk changes. Click <strong>Visual</strong> to switch back.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-clock-history"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Backups &amp; restore</h4>
                        <p>Every time you save, a backup of the previous version is created automatically (up to 10 per file). Click <strong>Backups</strong> to see the history. You can preview any backup and restore it with one click. A safety backup of the current version is also created before restoring.</p>
                    </div>
                </div>

                <div class="me-tutorial-step">
                    <div class="me-tutorial-num"><i class="bi bi-arrow-clockwise"></i></div>
                    <div class="me-tutorial-body">
                        <h4>Reload</h4>
                        <p>Click <strong>Reload</strong> to re-fetch the current tab's content from TouchPoint, discarding any unsaved changes.</p>
                    </div>
                </div>
            </div>

        </div>
        <div class="me-modal-footer">
            <button class="me-btn me-btn-primary" onclick="closeModal('modalTutorial')"><i class="bi bi-check-lg"></i> Got it!</button>
        </div>
    </div>
</div>

<!-- Toast -->
<div class="me-toast" id="meToast"></div>

<!-- ==================== JAVASCRIPT ==================== -->
<script>
console.log('=== MENU EDITOR JS STARTING ===');

var scriptUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
var currentTab = 'admin';
var menuData = {};
var isDirty = {};
var rawMode = false;
var collapsedCols = {}; // Track collapsed column sections

// ============================================================================
// Init
// ============================================================================
function initApp() {
    console.log('Menu Editor initialized, scriptUrl:', scriptUrl);
    loadContent('admin');
}

// ============================================================================
// Tab Switching
// ============================================================================
function switchTab(tab) {
    if (isDirty[currentTab]) {
        if (!confirm('You have unsaved changes on this tab. Switch anyway?')) return;
    }
    currentTab = tab;
    var tabs = document.querySelectorAll('.me-tab');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].className = tabs[i].dataset.tab === tab ? 'me-tab active' : 'me-tab';
        if (isDirty[tabs[i].dataset.tab]) tabs[i].classList.add('dirty');
    }
    rawMode = false;
    document.getElementById('btnRaw').innerHTML = '<i class="bi bi-code-slash"></i> Raw XML';

    if (menuData[tab]) {
        renderTab();
    } else {
        loadContent(tab);
    }
}

// ============================================================================
// AJAX Operations
// ============================================================================
function ajax(action, params, callback) {
    var data = $.extend({action: action}, params);
    $.ajax({
        url: scriptUrl,
        type: 'POST',
        data: data,
        success: function(resp) {
            try {
                var r = JSON.parse(resp);
                callback(r);
            } catch(e) {
                console.error('Parse error:', e, resp);
                showToast('Response parse error', 'error');
            }
        },
        error: function(xhr, status, err) {
            console.error('AJAX error:', status, err);
            showToast('Network error: ' + status, 'error');
        }
    });
}

function loadContent(tab) {
    document.getElementById('meContent').innerHTML = '<div class="me-loading"><i class="bi bi-hourglass-split"></i> Loading...</div>';
    ajax('load_content', {tab_key: tab}, function(r) {
        if (r.success) {
            menuData[tab] = r.data;
            isDirty[tab] = false;
            updateDirtyDots();
            renderTab();
            setStatus('Loaded');
        } else {
            setStatus('Error: ' + (r.message || 'Unknown'));
        }
    });
}

function saveContent() {
    if (rawMode) { saveRawContent(); return; }
    var data = menuData[currentTab];
    if (!data) return;

    var jsonStr = JSON.stringify(data);
    // Encode HTML entities for ASP.NET
    jsonStr = jsonStr.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                     .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

    var btn = document.getElementById('btnSave');
    btn.innerHTML = '<i class="bi bi-hourglass-split"></i> Saving...';
    btn.disabled = true;

    ajax('save_content', {tab_key: currentTab, json_data: jsonStr}, function(r) {
        btn.innerHTML = '<i class="bi bi-check-lg"></i> Save Changes';
        btn.disabled = false;
        if (r.success) {
            isDirty[currentTab] = false;
            updateDirtyDots();
            showToast('Saved successfully!', 'success');
        } else {
            showToast('Save failed: ' + (r.message || 'Unknown'), 'error');
        }
    });
}

function reloadContent() {
    if (isDirty[currentTab]) {
        if (!confirm('Discard unsaved changes and reload?')) return;
    }
    delete menuData[currentTab];
    isDirty[currentTab] = false;
    updateDirtyDots();
    loadContent(currentTab);
}

function discardChanges() {
    // Find all dirty tabs
    var dirtyTabs = [];
    for (var k in isDirty) { if (isDirty[k]) dirtyTabs.push(k); }
    if (dirtyTabs.length === 0) return;

    var msg = dirtyTabs.length === 1
        ? 'Discard unsaved changes on the ' + dirtyTabs[0] + ' tab?'
        : 'Discard unsaved changes on ' + dirtyTabs.length + ' tabs (' + dirtyTabs.join(', ') + ')?';
    if (!confirm(msg)) return;

    for (var i = 0; i < dirtyTabs.length; i++) {
        delete menuData[dirtyTabs[i]];
        isDirty[dirtyTabs[i]] = false;
    }
    updateDirtyDots();
    if (rawMode) toggleRawEditor();
    loadContent(currentTab);
    showToast('Changes discarded', 'info');
}

// ============================================================================
// Raw XML Editor
// ============================================================================
// Build XML from in-memory data (client-side)
function buildReportsMenuXml(data) {
    var allCols = ['Column1', 'Column2', 'Column3', 'Column4'];
    var lines = ['<ReportsMenu>'];
    for (var c = 0; c < allCols.length; c++) {
        var ck = allCols[c];
        if (!data.hasOwnProperty(ck)) continue;
        var items = data[ck] || [];
        lines.push('  <' + ck + '>');
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var text = (item.text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            if (item.item_type === 'header') {
                if (item.roles) {
                    lines.push('    <Header roles="' + item.roles + '">' + text + '</Header>');
                } else {
                    lines.push('    <Header>' + text + '</Header>');
                }
            } else {
                var attrs = '';
                if (item.roles) attrs += ' roles="' + item.roles + '"';
                if (item.link) attrs += ' link="' + item.link + '"';
                lines.push('    <Report' + attrs + '>' + text + '</Report>');
            }
        }
        lines.push('  </' + ck + '>');
    }
    lines.push('</ReportsMenu>');
    return lines.join('\\n');
}

function buildCustomReportsXml(data) {
    var lines = ['<CustomReports>'];
    for (var i = 0; i < data.length; i++) {
        var r = data[i];
        var attrs = ' name="' + (r.name || '') + '"';
        if (r.item_type) attrs += ' type="' + r.item_type + '"';
        if (r.item_role) attrs += ' role="' + r.item_role + '"';
        if (r.showOnOrgId) attrs += ' showOnOrgId="' + r.showOnOrgId + '"';
        if (r.url) attrs += ' url="' + r.url + '"';
        var cols = r.columns || [];
        if (cols.length === 0) {
            lines.push('  <Report' + attrs + ' />');
        } else {
            lines.push('  <Report' + attrs + '>');
            for (var j = 0; j < cols.length; j++) {
                var col = cols[j];
                var ca = ' name="' + (col.name || '') + '"';
                if (col.field) ca += ' field="' + col.field + '"';
                if (col.description) ca += ' description="' + col.description + '"';
                if (col.flag) ca += ' flag="' + col.flag + '"';
                if (col.orgid) ca += ' orgid="' + col.orgid + '"';
                if (col.smallgroup) ca += ' smallgroup="' + col.smallgroup + '"';
                if (col.disabled) ca += ' disabled="' + col.disabled + '"';
                lines.push('    <Column' + ca + ' />');
            }
            lines.push('  </Report>');
        }
    }
    lines.push('</CustomReports>');
    return lines.join('\\n');
}

function dataToXml() {
    var data = menuData[currentTab];
    if (!data) return '';
    if (currentTab === 'custom') {
        return buildCustomReportsXml(data);
    } else {
        return buildReportsMenuXml(data);
    }
}

function toggleRawEditor() {
    if (rawMode) {
        rawMode = false;
        document.getElementById('btnRaw').innerHTML = '<i class="bi bi-code-slash"></i> Raw XML';
        if (menuData[currentTab]) {
            renderTab();
        } else {
            loadContent(currentTab);
        }
        return;
    }
    rawMode = true;
    document.getElementById('btnRaw').innerHTML = '<i class="bi bi-layout-text-window-reverse"></i> Visual Editor';

    // If there are unsaved changes, build XML from in-memory data
    if (isDirty[currentTab] && menuData[currentTab]) {
        var xml = dataToXml();
        showRawEditor(xml);
    } else {
        // Otherwise fetch from server
        document.getElementById('meContent').innerHTML = '<div class="me-loading"><i class="bi bi-hourglass-split"></i> Loading raw XML...</div>';
        ajax('load_raw', {tab_key: currentTab}, function(r) {
            if (r.success) {
                showRawEditor(r.raw || '');
            }
        });
    }
}

function showRawEditor(xml) {
    var escaped = xml.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    document.getElementById('meContent').innerHTML =
        '<div style="padding:12px;">' +
        '<p style="font-size:12px;color:#888;margin:0 0 8px;"><i class="bi bi-info-circle"></i> Editing raw XML directly. Changes here will update the visual editor when you switch back.</p>' +
        '<textarea class="me-raw-editor" id="rawEditor">' + escaped + '</textarea>' +
        '</div>';
}

function saveRawContent() {
    var raw = document.getElementById('rawEditor').value;
    var encoded = raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                     .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

    var btn = document.getElementById('btnSave');
    btn.innerHTML = '<i class="bi bi-hourglass-split"></i> Saving...';
    btn.disabled = true;

    ajax('save_raw', {tab_key: currentTab, raw_xml: encoded}, function(r) {
        btn.innerHTML = '<i class="bi bi-check-lg"></i> Save Changes';
        btn.disabled = false;
        if (r.success) {
            delete menuData[currentTab];
            isDirty[currentTab] = false;
            updateDirtyDots();
            showToast('Raw XML saved!', 'success');
        } else {
            showToast('Save failed: ' + (r.message || ''), 'error');
        }
    });
}

// ============================================================================
// Rendering
// ============================================================================
function renderTab() {
    if (currentTab === 'custom') {
        renderCustomReports();
    } else {
        renderReportsMenu();
    }
}

function renderReportsMenu() {
    var data = menuData[currentTab];
    if (!data) return;

    var allCols = ['Column1', 'Column2', 'Column3', 'Column4'];
    var cols = [];
    for (var ci = 0; ci < allCols.length; ci++) {
        if (data.hasOwnProperty(allCols[ci])) cols.push(allCols[ci]);
    }
    var h = '<div class="me-columns">';

    for (var c = 0; c < cols.length; c++) {
        var ck = cols[c];
        var items = data[ck] || [];
        var isCollapsed = collapsedCols[currentTab + '_' + ck] ? true : false;

        h += '<div class="me-column-section">';

        // Section header - clickable to collapse
        h += '<div class="me-col-header" data-col="' + ck + '" onclick="toggleColSection(this.dataset.col)">';
        h += '<span class="me-col-title"><i class="bi bi-caret-' + (isCollapsed ? 'right' : 'down') + '-fill" id="colIcon_' + ck + '"></i> ' + ck.replace('Column', 'Column ');
        h += '<span class="me-col-count">(' + items.length + ' item' + (items.length !== 1 ? 's' : '') + ')</span></span>';
        h += '<div class="me-col-actions" onclick="event.stopPropagation()">';
        h += '<button class="me-btn me-btn-sm me-btn-primary" data-col="' + ck + '" data-itype="header" onclick="addItem(this.dataset.col,this.dataset.itype)" title="Add Header"><i class="bi bi-plus"></i> Header</button>';
        h += '<button class="me-btn me-btn-sm me-btn-outline" data-col="' + ck + '" data-itype="report" onclick="addItem(this.dataset.col,this.dataset.itype)" title="Add Report"><i class="bi bi-plus"></i> Report</button>';
        h += '<button class="me-btn me-btn-sm me-btn-danger" data-col="' + ck + '" onclick="deleteColumn(this.dataset.col)" title="Delete Column" style="margin-left:4px;"><i class="bi bi-trash"></i></button>';
        h += '</div></div>';

        // Section body
        h += '<div class="me-col-body' + (isCollapsed ? ' collapsed' : '') + '" id="colBody_' + ck + '" data-col="' + ck + '"';
        h += ' ondragover="colDragOver(event)" ondragleave="colDragLeave(event)" ondrop="colDrop(event,this.dataset.col)">';

        if (items.length === 0) {
            h += '<div class="me-empty">Empty column - use buttons above to add items</div>';
        }

        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var isH = item.item_type === 'header';
            h += '<div class="me-item' + (isH ? ' me-item-header' : '') + '" draggable="true" data-col="' + ck + '" data-idx="' + i + '"';
            h += ' ondragstart="dragStart(event)" ondragover="dragOver(event)" ondrop="dragDrop(event)" ondragend="dragEnd(event)" ondragleave="dragLeave(event)">';

            // Type badge
            h += '<div class="me-item-type-col">';
            h += isH ? '<span class="me-badge me-badge-h">H</span>' : '<span class="me-badge me-badge-r">R</span>';
            h += '</div>';

            // Text
            h += '<div class="me-item-text-col" title="' + escHtml(item.text || '') + '">';
            h += escHtml(item.text || '(untitled)');
            h += '</div>';

            // Link (reports only)
            h += '<div class="me-item-link-col">';
            if (item.link) h += '<a href="' + escHtml(item.link) + '" target="_blank" title="' + escHtml(item.link) + '">' + escHtml(item.link) + '</a>';
            h += '</div>';

            // Roles
            h += '<div class="me-item-roles-col">';
            if (item.roles) h += '<span class="me-badge">' + escHtml(item.roles) + '</span>';
            h += '</div>';

            // Actions
            h += '<div class="me-item-actions-col">';
            h += '<button class="me-btn me-btn-icon me-btn-outline me-btn-sm" data-col="' + ck + '" data-idx="' + i + '" onclick="editItem(this.dataset.col,this.dataset.idx)" title="Edit"><i class="bi bi-pencil"></i></button>';
            h += '<button class="me-btn me-btn-icon me-btn-outline me-btn-sm" data-col="' + ck + '" data-idx="' + i + '" onclick="moveItem(this.dataset.col,this.dataset.idx,-1)" title="Move up"><i class="bi bi-arrow-up"></i></button>';
            h += '<button class="me-btn me-btn-icon me-btn-outline me-btn-sm" data-col="' + ck + '" data-idx="' + i + '" onclick="moveItem(this.dataset.col,this.dataset.idx,1)" title="Move down"><i class="bi bi-arrow-down"></i></button>';
            // Move to column dropdown
            h += '<select class="me-move-select" data-col="' + ck + '" data-idx="' + i + '" onchange="if(this.value){moveToCol(this.dataset.col,this.dataset.idx,this.value);this.value=&quot;&quot;}" title="Move to column">';
            h += '<option value="">Move to...</option>';
            for (var cc = 0; cc < cols.length; cc++) {
                if (cols[cc] !== ck) h += '<option value="' + cols[cc] + '">' + cols[cc].replace('Column','Col ') + '</option>';
            }
            h += '</select>';
            h += '<button class="me-btn me-btn-icon me-btn-danger me-btn-sm" data-col="' + ck + '" data-idx="' + i + '" onclick="deleteItem(this.dataset.col,this.dataset.idx)" title="Delete"><i class="bi bi-trash"></i></button>';
            h += '</div>';

            h += '</div>'; // end me-item
        }
        h += '</div>'; // end me-col-body
        h += '</div>'; // end me-column-section
    }

    // Add Column button if fewer than 4 columns
    if (cols.length < 4) {
        h += '<div style="text-align:center;padding:12px;">';
        h += '<button class="me-btn me-btn-outline" onclick="addNewColumn()"><i class="bi bi-plus-lg"></i> Add Column</button>';
        h += '</div>';
    }

    h += '</div>';
    document.getElementById('meContent').innerHTML = h;
}

function addNewColumn() {
    var data = menuData[currentTab];
    if (!data) return;
    var allCols = ['Column1', 'Column2', 'Column3', 'Column4'];
    for (var i = 0; i < allCols.length; i++) {
        if (!data.hasOwnProperty(allCols[i])) {
            data[allCols[i]] = [];
            markDirty();
            renderReportsMenu();
            showToast(allCols[i].replace('Column', 'Column ') + ' added', 'success');
            return;
        }
    }
    alert('All 4 columns already exist.');
}

function toggleColSection(colKey) {
    var key = currentTab + '_' + colKey;
    collapsedCols[key] = !collapsedCols[key];
    var body = document.getElementById('colBody_' + colKey);
    var icon = document.getElementById('colIcon_' + colKey);
    if (body) {
        if (collapsedCols[key]) {
            body.classList.add('collapsed');
            if (icon) icon.className = 'bi bi-caret-right-fill';
        } else {
            body.classList.remove('collapsed');
            if (icon) icon.className = 'bi bi-caret-down-fill';
        }
    }
}

function deleteColumn(colKey) {
    var data = menuData[currentTab];
    if (!data) return;
    var items = data[colKey] || [];
    if (items.length > 0) {
        alert('Cannot delete ' + colKey.replace('Column', 'Column ') + ' because it still has ' + items.length + ' item(s). Please move or delete all items in this column first.');
        return;
    }
    showConfirm('Delete ' + colKey.replace('Column', 'Column ') + '? This will remove the empty column from the XML.', function() {
        delete data[colKey];
        markDirty();
        renderReportsMenu();
        showToast('Column deleted', 'success');
    });
}

function getBtCategory(r) {
    var t = (r.item_type || '').toLowerCase();
    if (t === 'sqlreport') return 'sql';
    if (t === 'pyscript') return 'python';
    if (t === 'url' || t === 'orgsearchsqlreport') return 'other';
    // No type = custom column report; or has columns
    return 'custom';
}

var btGroupLabels = {
    sql: 'SQL Reports',
    python: 'Python Scripts',
    other: 'Other Reports',
    custom: 'Custom Reports'
};
var btGroupIcons = {
    sql: 'bi-database',
    python: 'bi-code-slash',
    other: 'bi-link-45deg',
    custom: 'bi-table'
};
var btGroupOrder = ['sql', 'python', 'other', 'custom'];

function renderCustomReports() {
    var data = menuData[currentTab];
    if (!data) return;

    // Build duplicate name map
    var nameCounts = {};
    for (var di = 0; di < data.length; di++) {
        var dn = (data[di].name || '').toLowerCase();
        if (!nameCounts[dn]) nameCounts[dn] = [];
        nameCounts[dn].push(di);
    }
    var dupNames = {};
    var totalDups = 0;
    for (var dk in nameCounts) {
        if (nameCounts[dk].length > 1) {
            for (var dj = 0; dj < nameCounts[dk].length; dj++) {
                dupNames[nameCounts[dk][dj]] = nameCounts[dk].length;
            }
            totalDups += nameCounts[dk].length;
        }
    }

    // Group items by category, preserving original index
    var groups = {};
    for (var g = 0; g < btGroupOrder.length; g++) groups[btGroupOrder[g]] = [];
    for (var i = 0; i < data.length; i++) {
        var cat = getBtCategory(data[i]);
        groups[cat].push(i);
    }

    var h = '<div style="padding:8px 12px;">';

    // Duplicate warning banner
    if (totalDups > 0) {
        var dupGroupCount = 0;
        for (var dk2 in nameCounts) { if (nameCounts[dk2].length > 1) dupGroupCount++; }
        h += '<div class="me-dup-banner">';
        h += '<i class="bi bi-exclamation-triangle-fill"></i>';
        h += '<div><strong>Duplicate reports detected!</strong> Found ' + dupGroupCount + ' report name(s) with duplicates (' + totalDups + ' total entries). ';
        h += 'Duplicates are highlighted in yellow below.</div>';
        h += '</div>';
    }

    h += '<div style="margin-bottom:10px;display:flex;align-items:center;gap:10px;">';
    h += '<button class="me-btn me-btn-primary" onclick="addCustomReport()"><i class="bi bi-plus-lg"></i> Add Report</button>';
    h += '<span style="font-size:12px;color:#888;">' + data.length + ' report(s) total</span>';
    h += '</div>';

    for (var gi = 0; gi < btGroupOrder.length; gi++) {
        var gk = btGroupOrder[gi];
        var items = groups[gk];
        if (items.length === 0) continue;

        h += '<div class="me-bt-group" id="btGroup_' + gk + '">';
        h += '<div class="me-bt-group-header" onclick="toggleBtGroup(this.parentNode)">';
        h += '<span><i class="bi ' + btGroupIcons[gk] + '" style="margin-right:8px;"></i>' + btGroupLabels[gk] + '</span>';
        h += '<span style="display:flex;align-items:center;gap:8px;"><span class="me-bt-count">' + items.length + '</span><i class="bi bi-chevron-down me-bt-chevron"></i></span>';
        h += '</div>';
        h += '<div class="me-bt-group-body">';
        h += '<table class="me-cr-table"><thead><tr>';
        if (gk === 'custom') {
            h += '<th style="width:30px;"></th><th>Name</th><th>Role</th><th>OrgId</th><th>Cols</th><th style="width:140px;">Actions</th>';
        } else {
            h += '<th>Name</th><th>Type</th><th>Role</th><th style="width:140px;">Actions</th>';
        }
        h += '</tr></thead><tbody>';

        for (var ii = 0; ii < items.length; ii++) {
            var idx = items[ii];
            var r = data[idx];
            var cc = (r.columns || []).length;

            var isDup = dupNames.hasOwnProperty(idx);
            var dupClass = isDup ? ' class="me-dup-row"' : '';
            var dupBadge = isDup ? ' <span class="me-dup-badge">' + dupNames[idx] + 'x</span>' : '';

            if (gk === 'custom') {
                h += '<tr' + dupClass + '>';
                h += '<td>';
                if (cc > 0) {
                    h += '<span class="me-cr-expand" onclick="toggleExpand(' + idx + ')" id="expandIcon' + idx + '"><i class="bi bi-caret-right-fill"></i></span>';
                } else {
                    h += '<span class="me-cr-expand" onclick="toggleExpand(' + idx + ')" id="expandIcon' + idx + '" style="color:#ccc;"><i class="bi bi-caret-right"></i></span>';
                }
                h += '</td>';
                h += '<td><strong>' + escHtml(r.name || '') + '</strong>' + dupBadge + '</td>';
                h += '<td>' + (r.item_role ? '<span class="me-badge">' + escHtml(r.item_role) + '</span>' : '') + '</td>';
                h += '<td>' + escHtml(r.showOnOrgId || '') + '</td>';
                h += '<td>' + cc + '</td>';
            } else {
                h += '<tr' + dupClass + '>';
                h += '<td><strong>' + escHtml(r.name || '') + '</strong>' + dupBadge + '</td>';
                h += '<td>' + (r.item_type ? '<span class="me-badge me-badge-type">' + escHtml(r.item_type) + '</span>' : '') + '</td>';
                h += '<td>' + (r.item_role ? '<span class="me-badge">' + escHtml(r.item_role) + '</span>' : '') + '</td>';
            }

            h += '<td>';
            h += '<button class="me-btn me-btn-icon me-btn-outline me-btn-sm" onclick="editCustomReport(' + idx + ')" title="Edit"><i class="bi bi-pencil"></i></button> ';
            if (idx > 0) h += '<button class="me-btn me-btn-icon me-btn-outline me-btn-sm" onclick="moveCr(' + idx + ',-1)" title="Move up"><i class="bi bi-arrow-up"></i></button> ';
            if (idx < data.length - 1) h += '<button class="me-btn me-btn-icon me-btn-outline me-btn-sm" onclick="moveCr(' + idx + ',1)" title="Move down"><i class="bi bi-arrow-down"></i></button> ';
            h += '<button class="me-btn me-btn-icon me-btn-danger me-btn-sm" onclick="deleteCr(' + idx + ')" title="Delete"><i class="bi bi-trash"></i></button>';
            h += '</td></tr>';

            // Expandable columns row (only for custom)
            if (gk === 'custom') {
                var colSpan = 6;
                h += '<tr id="crCols' + idx + '" style="display:none;"><td colspan="' + colSpan + '" class="me-cr-columns">';
                h += '<div style="margin-bottom:6px;display:flex;justify-content:space-between;align-items:center;">';
                h += '<strong style="font-size:12px;">Columns (' + cc + ')</strong> ';
                h += '<button class="me-btn me-btn-sm me-btn-primary" onclick="addColumn(' + idx + ')"><i class="bi bi-plus"></i> Add Column</button>';
                h += '</div>';
                if (cc > 0) {
                    for (var j = 0; j < r.columns.length; j++) {
                        var col = r.columns[j];
                        h += '<div class="me-cr-col-item">';
                        h += '<span>' + escHtml(fmtCol(col)) + '</span>';
                        h += '<span style="white-space:nowrap;">';
                        h += '<button class="me-btn me-btn-icon me-btn-sm me-btn-outline" onclick="editColumn(' + idx + ',' + j + ')" title="Edit"><i class="bi bi-pencil"></i></button>';
                        if (j > 0) h += '<button class="me-btn me-btn-icon me-btn-sm me-btn-outline" onclick="moveCol(' + idx + ',' + j + ',-1)" title="Up"><i class="bi bi-arrow-up"></i></button>';
                        if (j < r.columns.length - 1) h += '<button class="me-btn me-btn-icon me-btn-sm me-btn-outline" onclick="moveCol(' + idx + ',' + j + ',1)" title="Down"><i class="bi bi-arrow-down"></i></button>';
                        h += '<button class="me-btn me-btn-icon me-btn-sm me-btn-danger" onclick="deleteCol(' + idx + ',' + j + ')" title="Delete"><i class="bi bi-trash"></i></button>';
                        h += '</span></div>';
                    }
                } else {
                    h += '<div class="me-empty">No columns defined</div>';
                }
                h += '</td></tr>';
            }
        }

        h += '</tbody></table>';
        h += '</div></div>';
    }

    h += '</div>';
    document.getElementById('meContent').innerHTML = h;
}

function toggleBtGroup(el) {
    el.classList.toggle('collapsed');
}

function fmtCol(col) {
    var n = col.name || '';
    if (col.smallgroup) return n + ': "' + col.smallgroup + '"' + (col.orgid ? ' (org:' + col.orgid + ')' : '');
    if (col.flag) return n + ': "' + (col.description || '') + '" (flag:' + col.flag + ')';
    if (col.field) return n + ': ' + col.field + (col.disabled ? ' [disabled:' + col.disabled + ']' : '');
    if (col.orgid) return n + ' (org:' + col.orgid + ')';
    return n;
}

function toggleExpand(idx) {
    var row = document.getElementById('crCols' + idx);
    var icon = document.getElementById('expandIcon' + idx);
    if (row.style.display === 'none') {
        row.style.display = '';
        icon.innerHTML = '<i class="bi bi-caret-down-fill"></i>';
    } else {
        row.style.display = 'none';
        icon.innerHTML = '<i class="bi bi-caret-right-fill"></i>';
    }
}

// ============================================================================
// ReportsMenu Item Operations
// ============================================================================
function addItem(colKey, itemType) {
    document.getElementById('modalItemTitle').textContent = 'Add ' + (itemType === 'header' ? 'Header' : 'Report');
    document.getElementById('modalItemCol').value = colKey;
    document.getElementById('modalItemIdx').value = '-1';
    document.getElementById('modalItemType').value = itemType;
    document.getElementById('modalItemText').value = '';
    document.getElementById('modalItemLink').value = '';
    document.getElementById('modalItemRoles').value = '';
    document.getElementById('fieldLink').style.display = itemType === 'report' ? 'block' : 'none';
    openModal('modalItem');
    document.getElementById('modalItemText').focus();
}

function editItem(colKey, idx) {
    idx = parseInt(idx);
    var item = menuData[currentTab][colKey][idx];
    document.getElementById('modalItemTitle').textContent = 'Edit ' + (item.item_type === 'header' ? 'Header' : 'Report');
    document.getElementById('modalItemCol').value = colKey;
    document.getElementById('modalItemIdx').value = idx;
    document.getElementById('modalItemType').value = item.item_type;
    document.getElementById('modalItemText').value = item.text || '';
    document.getElementById('modalItemLink').value = item.link || '';
    document.getElementById('modalItemRoles').value = item.roles || '';
    document.getElementById('fieldLink').style.display = item.item_type === 'report' ? 'block' : 'none';
    openModal('modalItem');
}

function saveItemModal() {
    var col = document.getElementById('modalItemCol').value;
    var idx = parseInt(document.getElementById('modalItemIdx').value);
    var itemType = document.getElementById('modalItemType').value;
    var text = document.getElementById('modalItemText').value.trim();
    var link = document.getElementById('modalItemLink').value.trim();
    var roles = document.getElementById('modalItemRoles').value.trim();

    if (!text) { alert('Display text is required'); return; }
    if (itemType === 'report' && !link) { alert('Link URL is required for reports'); return; }

    var item = { item_type: itemType, text: text, roles: roles };
    if (itemType === 'report') item.link = link;

    if (idx === -1) {
        menuData[currentTab][col].push(item);
    } else {
        menuData[currentTab][col][idx] = item;
    }
    markDirty();
    closeModal('modalItem');
    renderTab();
}

function deleteItem(colKey, idx) {
    idx = parseInt(idx);
    var item = menuData[currentTab][colKey][idx];
    showConfirm('Delete "' + (item.text || 'this item') + '"?', function() {
        menuData[currentTab][colKey].splice(idx, 1);
        markDirty();
        renderTab();
    });
}

function moveItem(colKey, idx, dir) {
    idx = parseInt(idx);
    dir = parseInt(dir);
    var items = menuData[currentTab][colKey];
    var newIdx = idx + dir;
    if (newIdx < 0 || newIdx >= items.length) return;
    var temp = items[idx];
    items[idx] = items[newIdx];
    items[newIdx] = temp;
    markDirty();
    renderTab();
}

function moveToCol(fromCol, idx, toCol) {
    idx = parseInt(idx);
    var item = menuData[currentTab][fromCol].splice(idx, 1)[0];
    menuData[currentTab][toCol].push(item);
    markDirty();
    renderTab();
}

// ============================================================================
// Custom Reports Operations
// ============================================================================
var typeHints = {
    '': 'Custom Report: Defines data columns for exporting from search results. Appears under "Custom Reports" in the toolbar.',
    'PyScript': 'Python Script: Runs a Python script stored in Special Content. Appears under "Python Scripts" in the toolbar.',
    'SqlReport': 'SQL Report: Runs a SQL report stored in Special Content. Appears under "SQL Reports" in the toolbar.',
    'OrgSearchSqlReport': 'Org Search SQL Report: SQL report that runs on organization search results. Appears under "Other Reports" in the toolbar.',
    'URL': 'URL Link: Opens a URL (e.g. document merge). Appears under "Other Reports" in the toolbar.'
};

function updateTypeHint() {
    var t = document.getElementById('modalReportType').value;
    var hint = document.getElementById('typeHint');
    hint.textContent = typeHints[t] || '';
    // Show/hide URL field based on type
    var urlField = document.getElementById('fieldUrl');
    if (urlField) urlField.style.display = (t === 'URL') ? 'block' : 'none';
    // Show/hide OrgId field - relevant for custom reports and OrgSearchSqlReport
    var orgField = document.getElementById('modalReportOrgId');
    if (orgField) orgField.closest('.me-field').style.display = (t === '' || t === 'OrgSearchSqlReport') ? 'block' : 'none';
}

function addCustomReport() {
    document.getElementById('modalReportTitle').textContent = 'Add Report';
    document.getElementById('modalReportIdx').value = '-1';
    document.getElementById('modalReportName').value = '';
    document.getElementById('modalReportType').value = '';
    document.getElementById('modalReportRole').value = '';
    document.getElementById('modalReportOrgId').value = '';
    document.getElementById('modalReportUrl').value = '';
    updateTypeHint();
    openModal('modalReport');
    document.getElementById('modalReportName').focus();
}

function editCustomReport(idx) {
    var r = menuData[currentTab][idx];
    document.getElementById('modalReportTitle').textContent = 'Edit Report';
    document.getElementById('modalReportIdx').value = idx;
    document.getElementById('modalReportName').value = r.name || '';
    document.getElementById('modalReportType').value = r.item_type || '';
    document.getElementById('modalReportRole').value = r.item_role || '';
    document.getElementById('modalReportOrgId').value = r.showOnOrgId || '';
    document.getElementById('modalReportUrl').value = r.url || '';
    updateTypeHint();
    openModal('modalReport');
}

function saveReportModal() {
    var idx = parseInt(document.getElementById('modalReportIdx').value);
    var name = document.getElementById('modalReportName').value.trim();
    if (!name) { alert('Report name is required'); return; }

    var report = {
        name: name,
        item_type: document.getElementById('modalReportType').value,
        item_role: document.getElementById('modalReportRole').value.trim(),
        showOnOrgId: document.getElementById('modalReportOrgId').value.trim(),
        url: document.getElementById('modalReportUrl').value.trim(),
        columns: []
    };

    if (idx === -1) {
        menuData[currentTab].push(report);
    } else {
        report.columns = menuData[currentTab][idx].columns || [];
        menuData[currentTab][idx] = report;
    }
    markDirty();
    closeModal('modalReport');
    renderTab();
}

function deleteCr(idx) {
    var r = menuData[currentTab][idx];
    showConfirm('Delete report "' + (r.name || '') + '"?', function() {
        menuData[currentTab].splice(idx, 1);
        markDirty();
        renderTab();
    });
}

function moveCr(idx, dir) {
    var data = menuData[currentTab];
    var ni = idx + dir;
    if (ni < 0 || ni >= data.length) return;
    var temp = data[idx]; data[idx] = data[ni]; data[ni] = temp;
    markDirty();
    renderTab();
}

// ============================================================================
// Column Operations (within CustomReports)
// ============================================================================
function addColumn(reportIdx) {
    document.getElementById('modalColumnTitle').textContent = 'Add Column';
    document.getElementById('modalColReportIdx').value = reportIdx;
    document.getElementById('modalColIdx').value = '-1';
    document.getElementById('modalColType').value = 'simple';
    document.getElementById('modalColName').value = '';
    document.getElementById('modalColField').value = '';
    document.getElementById('modalColSmallgroup').value = '';
    document.getElementById('modalColDescription').value = '';
    document.getElementById('modalColFlag').value = '';
    document.getElementById('modalColOrgId').value = '';
    document.getElementById('modalColDisabled').value = '';
    updateColFields();
    openModal('modalColumn');
}

function editColumn(reportIdx, colIdx) {
    var col = menuData[currentTab][reportIdx].columns[colIdx];
    document.getElementById('modalColumnTitle').textContent = 'Edit Column';
    document.getElementById('modalColReportIdx').value = reportIdx;
    document.getElementById('modalColIdx').value = colIdx;

    // Detect column type
    var ct = 'simple';
    if (col.name === 'SmallGroup') ct = 'SmallGroup';
    else if (col.name === 'StatusFlag') ct = 'StatusFlag';
    else if (col.name === 'ExtraValueText') ct = 'ExtraValueText';
    else if (col.name === 'FamilyExtraValueText') ct = 'FamilyExtraValueText';
    else if (col.orgid && col.name !== 'SmallGroup') ct = 'OrgField';

    document.getElementById('modalColType').value = ct;
    document.getElementById('modalColName').value = col.name || '';
    document.getElementById('modalColField').value = col.field || '';
    document.getElementById('modalColSmallgroup').value = col.smallgroup || '';
    document.getElementById('modalColDescription').value = col.description || '';
    document.getElementById('modalColFlag').value = col.flag || '';
    document.getElementById('modalColOrgId').value = col.orgid || '';
    document.getElementById('modalColDisabled').value = col.disabled || '';
    updateColFields();
    openModal('modalColumn');
}

function updateColFields() {
    var t = document.getElementById('modalColType').value;
    var show = function(id, vis) { document.getElementById(id).style.display = vis ? 'block' : 'none'; };
    show('colFieldName', t === 'simple' || t === 'OrgField');
    show('colFieldField', t === 'ExtraValueText' || t === 'FamilyExtraValueText');
    show('colFieldSmallgroup', t === 'SmallGroup');
    show('colFieldDescription', t === 'StatusFlag');
    show('colFieldFlag', t === 'StatusFlag');
    show('colFieldOrgId', t === 'SmallGroup' || t === 'OrgField');
    show('colFieldDisabled', t === 'ExtraValueText' || t === 'FamilyExtraValueText');

    // Auto-set name for typed columns
    var nf = document.getElementById('modalColName');
    var autoNames = ['SmallGroup', 'StatusFlag', 'ExtraValueText', 'FamilyExtraValueText'];
    if (autoNames.indexOf(t) > -1) nf.value = t;
    else if (t === 'simple' && autoNames.indexOf(nf.value) > -1) nf.value = '';
}

function saveColumnModal() {
    var ri = parseInt(document.getElementById('modalColReportIdx').value);
    var ci = parseInt(document.getElementById('modalColIdx').value);
    var ct = document.getElementById('modalColType').value;
    var col = {};

    if (ct === 'simple') {
        var n = document.getElementById('modalColName').value.trim();
        if (!n) { alert('Column name is required'); return; }
        col.name = n;
    } else if (ct === 'ExtraValueText' || ct === 'FamilyExtraValueText') {
        var f = document.getElementById('modalColField').value.trim();
        if (!f) { alert('Field name is required'); return; }
        col.field = f;
        col.name = ct;
        var d = document.getElementById('modalColDisabled').value;
        if (d) col.disabled = d;
    } else if (ct === 'SmallGroup') {
        var sg = document.getElementById('modalColSmallgroup').value.trim();
        if (!sg) { alert('Small group name is required'); return; }
        col.smallgroup = sg;
        col.name = 'SmallGroup';
        var oid = document.getElementById('modalColOrgId').value.trim();
        if (oid) col.orgid = oid;
    } else if (ct === 'StatusFlag') {
        var fl = document.getElementById('modalColFlag').value.trim();
        if (!fl) { alert('Flag ID is required'); return; }
        col.description = document.getElementById('modalColDescription').value.trim();
        col.flag = fl;
        col.name = 'StatusFlag';
    } else if (ct === 'OrgField') {
        var n2 = document.getElementById('modalColName').value.trim();
        if (!n2) { alert('Column name is required'); return; }
        col.name = n2;
        var oid2 = document.getElementById('modalColOrgId').value.trim();
        if (oid2) col.orgid = oid2;
    }

    var report = menuData[currentTab][ri];
    if (!report.columns) report.columns = [];
    if (ci === -1) { report.columns.push(col); }
    else { report.columns[ci] = col; }

    markDirty();
    closeModal('modalColumn');
    renderTab();
    // Re-expand
    setTimeout(function() { expandRow(ri); }, 50);
}

function deleteCol(ri, ci) {
    menuData[currentTab][ri].columns.splice(ci, 1);
    markDirty();
    renderTab();
    setTimeout(function() { expandRow(ri); }, 50);
}

function moveCol(ri, ci, dir) {
    var cols = menuData[currentTab][ri].columns;
    var ni = ci + dir;
    if (ni < 0 || ni >= cols.length) return;
    var tmp = cols[ci]; cols[ci] = cols[ni]; cols[ni] = tmp;
    markDirty();
    renderTab();
    setTimeout(function() { expandRow(ri); }, 50);
}

function expandRow(ri) {
    var row = document.getElementById('crCols' + ri);
    var icon = document.getElementById('expandIcon' + ri);
    if (row) row.style.display = '';
    if (icon) icon.innerHTML = '<i class="bi bi-caret-down-fill"></i>';
}

// ============================================================================
// Drag & Drop (for ReportsMenu items)
// ============================================================================
var dragData = null;

function dragStart(e) {
    var el = e.target.closest('.me-item');
    if (!el) return;
    dragData = { col: el.dataset.col, idx: parseInt(el.dataset.idx) };
    el.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', '');
}

function dragOver(e) {
    e.preventDefault();
    var el = e.target.closest('.me-item');
    if (el) el.classList.add('drag-over');
}

function dragLeave(e) {
    var el = e.target.closest('.me-item');
    if (el) el.classList.remove('drag-over');
}

function dragDrop(e) {
    e.preventDefault();
    var el = e.target.closest('.me-item');
    if (!el || !dragData) return;
    el.classList.remove('drag-over');

    var toCol = el.dataset.col;
    var toIdx = parseInt(el.dataset.idx);
    var fromCol = dragData.col;
    var fromIdx = dragData.idx;
    if (fromCol === toCol && fromIdx === toIdx) return;

    var item = menuData[currentTab][fromCol].splice(fromIdx, 1)[0];
    if (fromCol === toCol && fromIdx < toIdx) toIdx--;
    menuData[currentTab][toCol].splice(toIdx, 0, item);
    markDirty();
    renderTab();
}

function dragEnd(e) {
    dragData = null;
    var items = document.querySelectorAll('.me-item');
    for (var k = 0; k < items.length; k++) items[k].classList.remove('dragging', 'drag-over');
    var cols = document.querySelectorAll('.me-col-body');
    for (var k = 0; k < cols.length; k++) cols[k].classList.remove('drag-over-col');
}

// Drop on empty column
function colDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('drag-over-col');
}
function colDragLeave(e) {
    e.currentTarget.classList.remove('drag-over-col');
}
function colDrop(e, colKey) {
    e.currentTarget.classList.remove('drag-over-col');
    if (!dragData) return;
    // Only handle if dropped on column background, not on an item
    if (e.target.closest('.me-item')) return;
    e.preventDefault();

    var item = menuData[currentTab][dragData.col].splice(dragData.idx, 1)[0];
    menuData[currentTab][colKey].push(item);
    markDirty();
    renderTab();
}

// ============================================================================
// Role Helper
// ============================================================================
function addRoleFromSelect(selId, inputId) {
    var sel = document.getElementById(selId);
    var input = document.getElementById(inputId);
    if (!sel.value) return;
    var current = input.value.trim();
    if (current && current.indexOf(sel.value) === -1) {
        input.value = current + ',' + sel.value;
    } else if (!current) {
        input.value = sel.value;
    }
    sel.value = '';
}

// ============================================================================
// Utility Functions
// ============================================================================
function markDirty() {
    isDirty[currentTab] = true;
    updateDirtyDots();
}

function updateDirtyDots() {
    var tabs = document.querySelectorAll('.me-tab');
    var anyDirty = false;
    for (var i = 0; i < tabs.length; i++) {
        var t = tabs[i].dataset.tab;
        if (isDirty[t]) { tabs[i].classList.add('dirty'); anyDirty = true; }
        else { tabs[i].classList.remove('dirty'); }
    }
    var discardBtn = document.getElementById('btnDiscard');
    if (discardBtn) discardBtn.style.display = anyDirty ? '' : 'none';
}

function openModal(id) { document.getElementById(id).classList.add('show'); }
function closeModal(id) { document.getElementById(id).classList.remove('show'); }

function showTutorial() { openModal('modalTutorial'); }

function showConfirm(msg, onConfirm) {
    document.getElementById('confirmMessage').textContent = msg;
    document.getElementById('confirmDeleteBtn').onclick = function() {
        closeModal('modalConfirm');
        onConfirm();
    };
    openModal('modalConfirm');
}

function escHtml(s) {
    if (!s) return '';
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function truncate(s, max) {
    if (!s || s.length <= max) return s;
    return s.substring(0, max) + '...';
}

function setStatus(msg) { document.getElementById('meStatus').textContent = msg; }

function showToast(msg, kind) {
    var t = document.getElementById('meToast');
    t.textContent = msg;
    t.className = 'me-toast me-toast-' + (kind || 'success') + ' show';
    setTimeout(function() { t.classList.remove('show'); }, 3000);
}

// ============================================================================
// Backup / Restore
// ============================================================================
function showBackups() {
    openModal('modalBackups');
    document.getElementById('backupList').innerHTML = '<div class="me-loading"><i class="bi bi-hourglass-split"></i> Loading backups...</div>';
    ajax('list_backups', {tab_key: currentTab}, function(r) {
        if (r.success) {
            var bk = r.backups || [];
            if (bk.length === 0) {
                document.getElementById('backupList').innerHTML = '<div class="me-empty">No backups found for this tab. Backups are created automatically when you save changes.</div>';
                return;
            }
            var h = '';
            for (var i = 0; i < bk.length; i++) {
                var b = bk[i];
                h += '<div class="me-backup-item">';
                h += '<div class="me-backup-info">';
                h += '<div class="me-backup-date"><i class="bi bi-clock"></i> ' + escHtml(b.timestamp) + '</div>';
                h += '<div class="me-backup-user"><i class="bi bi-person"></i> ' + escHtml(b.user || 'Unknown') + '</div>';
                h += '</div>';
                h += '<div class="me-backup-actions">';
                h += '<button class="me-btn me-btn-sm me-btn-outline" data-bname="' + escHtml(b.backup_name) + '" data-btime="' + escHtml(b.timestamp) + '" data-buser="' + escHtml(b.user || '') + '" onclick="previewBackup(this.dataset.bname,this.dataset.btime,this.dataset.buser)"><i class="bi bi-eye"></i> Preview</button>';
                h += '<button class="me-btn me-btn-sm me-btn-primary" data-bname="' + escHtml(b.backup_name) + '" data-btime="' + escHtml(b.timestamp) + '" onclick="restoreBackup(this.dataset.bname,this.dataset.btime)"><i class="bi bi-arrow-counterclockwise"></i> Restore</button>';
                h += '</div></div>';
            }
            document.getElementById('backupList').innerHTML = h;
        } else {
            document.getElementById('backupList').innerHTML = '<div class="me-empty">Error loading backups: ' + escHtml(r.message || '') + '</div>';
        }
    });
}

function previewBackup(backupName, timestamp, user) {
    document.getElementById('previewTitle').textContent = 'Backup from ' + timestamp + (user ? ' by ' + user : '');
    document.getElementById('previewContent').value = 'Loading...';
    document.getElementById('previewRestoreBtn').onclick = function() {
        closeModal('modalPreview');
        restoreBackup(backupName, timestamp);
    };
    openModal('modalPreview');

    ajax('preview_backup', {backup_name: backupName}, function(r) {
        if (r.success) {
            document.getElementById('previewContent').value = r.raw || '(empty)';
        } else {
            document.getElementById('previewContent').value = 'Error: ' + (r.message || 'Could not load backup');
        }
    });
}

function restoreBackup(backupName, timestamp) {
    if (!confirm('Restore backup from ' + timestamp + '?\\n\\nThe current content will be backed up automatically before restoring.')) return;

    ajax('restore_backup', {backup_name: backupName, tab_key: currentTab}, function(r) {
        if (r.success) {
            closeModal('modalBackups');
            closeModal('modalPreview');
            delete menuData[currentTab];
            isDirty[currentTab] = false;
            updateDirtyDots();
            loadContent(currentTab);
            showToast('Restored from backup!', 'success');
        } else {
            showToast('Restore failed: ' + (r.message || ''), 'error');
        }
    });
}

// Warn before leaving with unsaved changes
window.onbeforeunload = function() {
    for (var k in isDirty) {
        if (isDirty[k]) return 'You have unsaved changes.';
    }
};

// ============================================================================
// Wait for jQuery then init
// ============================================================================
(function waitForJQuery() {
    if (typeof $ !== 'undefined') {
        $(document).ready(initApp);
    } else {
        setTimeout(waitForJQuery, 50);
    }
})();

</script>
</body>
</html>
'''
    # Output for both PyScript (print) and PyScriptForm (model.Form)
    print html
    model.Form = html
