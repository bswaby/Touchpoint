#Roles=Admin
# Written By: Ben Swaby (TPxi Software, LLC)
# Email: bswaby@fbchtn.org                                                                                                      
# Website: https://tpxisoftware.com
# GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)                                                                
# ----------------------------------------------------------------                                                              
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to                                                                
# support continued development, check out:                                                                                     
#
# DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)                                                  
# https://displaycache.com                                
#
# TPxi Go(TM) - your church contacts, wherever you work.
# Look up anyone in TouchPoint(R), log calls and emails from Outlook                                                            
# or your phone. No tab switching, no lost context.
# https://tpxigo.com                                                                                                            
# ----------------------------------------------------------------

"""
TPxi QuickLinks Admin
======================
Part 2 of 2 - Admin UI for managing QuickLinks configuration.

  Part 1: TPxi_QuickLinks (the widget that displays on the homepage)
  Part 2: TPxi_QuickLinksAdmin (this script - the admin UI to manage it)

Both scripts share the same JSON config stored in Special Content
as "QuickLinksConfig". This admin UI writes the config; the widget reads it.

Features:
- Unified tree view for categories, subgroups, and links
- Drag-and-drop reordering (replaces manual priority weights)
- Icon picker with ~150 Font Awesome icons
- Role-based permission management
- Nested subgroups (popup subgroups inside inline subgroups)
- Backup/restore with version history
- Import from starter defaults
- Raw JSON editor
- Count query builder with live testing
- Guided tour for first-time users
- Inline help tooltips on all form fields

Version: 2.0
Date: March 2026

--Upload Instructions Start--
1. Admin ~ Advanced ~ Special Content ~ Python
2. New Python Script File named "TPxi_QuickLinksAdmin"
3. Add to CustomReports:
   <Report name="TPxi_QuickLinksAdmin" type="PyScript" role="Admin" />
4. Also upload TPxi_QuickLinks as a widget script (see Part 1)
--Upload Instructions End--
"""

import json
import datetime
import traceback

# --- IronPython Unicode Safety ---

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

def _to_ascii(s):
    result = []
    for c in s:
        o = ord(c)
        if o < 128:
            result.append(c)
        else:
            result.append(_LATIN_TO_ASCII.get(o, '?'))
    return ''.join(result)

def safe_str(val):
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

def decode_html_entities(s):
    if not s:
        return s
    s = s.replace('&lt;', '<').replace('&gt;', '>')
    s = s.replace('&quot;', '"').replace('&#39;', "'")
    s = s.replace('&amp;', '&')
    return s

# --- Constants ---
CONTENT_NAME = "QuickLinksConfig"
BACKUP_LOG_NAME = "QuickLinksConfig_BackupLog"
MAX_BACKUPS = 10

# --- Backup Helpers ---

def get_backup_log():
    try:
        log_str = model.TextContent(BACKUP_LOG_NAME)
        if log_str:
            return json.loads(log_str)
    except:
        pass
    return {'backups': []}

def save_backup_log(log):
    model.WriteContentText(BACKUP_LOG_NAME, json.dumps(sanitize_for_json(log)), "")

def create_backup(content_str, user_name):
    now = datetime.datetime.now()
    timestamp = now.strftime('%Y%m%d_%H%M%S')
    backup_name = 'QuickLinks_Bak_%s' % timestamp
    model.WriteContentText(backup_name, content_str, "")
    log = get_backup_log()
    entry = {
        'backup_name': safe_str(backup_name),
        'timestamp': now.strftime('%Y-%m-%d %H:%M:%S'),
        'user': safe_str(user_name)
    }
    log['backups'].insert(0, entry)
    if len(log['backups']) > MAX_BACKUPS:
        old = log['backups'][MAX_BACKUPS:]
        log['backups'] = log['backups'][:MAX_BACKUPS]
        for o in old:
            try:
                model.WriteContentText(o['backup_name'], '', "")
            except:
                pass
    save_backup_log(log)
    return entry

def get_user_name():
    try:
        uid = model.UserPeopleId
        if uid:
            p = model.GetPerson(uid)
            if p:
                return safe_str(p.Name)
    except:
        pass
    return 'Unknown'

def get_all_roles():
    fallback = ['', 'Access', 'Edit', 'Admin', 'SuperAdmin', 'Developer',
                'Finance', 'FinanceAdmin', 'ManageTransactions', 'PastoralCare',
                'IT-Team', 'Security-Doors', 'Beta', 'ManageGroups',
                'ManageOrgMembers', 'BackgroundCheck', 'ManageApplication',
                'Staff', 'BackgroundCheckRun', 'ViewApplication',
                'ProgramManager', 'SquareReports']
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

# --- Access Check ---
if not model.UserIsInRole("Admin"):
    denied_html = "<h3>Access Denied</h3><p>You need the Admin role to access QuickLinks Admin.</p>"
    print denied_html
    model.Form = denied_html
else:
    # --- AJAX Handler ---
    if model.HttpMethod == "post":
        action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

        if action == 'load_config':
            try:
                config_str = model.TextContent(CONTENT_NAME)
                if not config_str:
                    config_str = '{}'
                config = json.loads(config_str)
                print json.dumps(sanitize_for_json({'success': True, 'data': config}))
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'save_config':
            import base64
            json_b64 = str(Data.tpl_json_b64) if hasattr(Data, 'tpl_json_b64') else ''
            try:
                json_data = base64.b64decode(json_b64)
            except:
                json_data = '{}'
            try:
                current = model.TextContent(CONTENT_NAME)
                if current and current.strip():
                    create_backup(current, get_user_name())
                data = json.loads(json_data)
                data['version'] = 1
                model.WriteContentText(CONTENT_NAME, json.dumps(sanitize_for_json(data)), "")
                print json.dumps({'success': True, 'message': 'Saved successfully'})
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'list_backups':
            try:
                log = get_backup_log()
                print json.dumps(sanitize_for_json({'success': True, 'backups': log.get('backups', [])}))
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'restore_backup':
            bname = str(Data.tpl_backup_name) if hasattr(Data, 'tpl_backup_name') else ''
            bname = decode_html_entities(bname)
            try:
                backup_content = model.TextContent(bname)
                if not backup_content:
                    print json.dumps({'success': False, 'message': 'Backup not found'})
                else:
                    current = model.TextContent(CONTENT_NAME)
                    if current and current.strip():
                        create_backup(current, get_user_name() + ' (pre-restore)')
                    model.WriteContentText(CONTENT_NAME, backup_content, "")
                    print json.dumps({'success': True, 'message': 'Restored successfully'})
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'preview_backup':
            bname = str(Data.tpl_backup_name) if hasattr(Data, 'tpl_backup_name') else ''
            bname = decode_html_entities(bname)
            try:
                backup_content = model.TextContent(bname)
                if not backup_content:
                    print json.dumps({'success': False, 'message': 'Backup not found'})
                else:
                    print json.dumps(sanitize_for_json({'success': True, 'content': backup_content}))
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'get_roles':
            try:
                roles = get_all_roles()
                print json.dumps(sanitize_for_json({'success': True, 'roles': roles}))
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'test_query':
            import base64
            tpl_sql_b64 = str(Data.tpl_sql_b64) if hasattr(Data, 'tpl_sql_b64') else ''
            try:
                tpl_sql = base64.b64decode(tpl_sql_b64)
            except:
                tpl_sql = ''
            tpl_org = str(Data.tpl_org) if hasattr(Data, 'tpl_org') else '0'
            try:
                if '{0}' in tpl_sql:
                    tpl_sql = tpl_sql.replace('{0}', str(int(tpl_org)))
                if '{}' in tpl_sql:
                    uid = model.UserPeopleId
                    tpl_sql = tpl_sql.replace('{}', str(uid))
                result = q.QuerySqlInt(tpl_sql)
                print json.dumps({'success': True, 'count': result})
            except Exception as e:
                print json.dumps({'success': False, 'message': safe_str(str(e))})

        elif action == 'import_hardcoded':
            print json.dumps({'success': True, 'message': 'Use client-side import'})

        else:
            print json.dumps({'success': False, 'message': 'Unknown action: ' + safe_str(action)})

    else:
        # --- Render UI ---
        html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<style>
:root {
    --ql-primary: #2196F3;
    --ql-primary-dark: #1976D2;
    --ql-primary-light: #BBDEFB;
    --ql-success: #4CAF50;
    --ql-success-dark: #388E3C;
    --ql-danger: #f44336;
    --ql-danger-dark: #d32f2f;
    --ql-warning: #FF9800;
    --ql-info: #00BCD4;
    --ql-bg: #f5f7fa;
    --ql-card: #ffffff;
    --ql-border: #e0e4e8;
    --ql-text: #333;
    --ql-text-muted: #777;
    --ql-shadow: 0 2px 8px rgba(0,0,0,0.08);
    --ql-radius: 8px;
}
* { box-sizing: border-box; }
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
#ql-admin-root {
    max-width: 1100px; margin: 0 auto; padding: 16px;
    color: var(--ql-text); background: var(--ql-bg);
    border-radius: var(--ql-radius);
}
.ql-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 16px; padding: 16px 20px;
    background: linear-gradient(135deg, var(--ql-primary), var(--ql-primary-dark));
    color: white; border-radius: var(--ql-radius);
    box-shadow: var(--ql-shadow);
}
.ql-header h2 { margin: 0; font-size: 1.4em; }
.ql-header-actions { display: flex; gap: 8px; }
.ql-btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 8px 16px; border: none; border-radius: 6px;
    font-size: 13px; font-weight: 600; cursor: pointer;
    transition: all 0.2s; text-decoration: none;
}
.ql-btn:hover { transform: translateY(-1px); box-shadow: 0 2px 6px rgba(0,0,0,0.15); }
.ql-btn-primary { background: var(--ql-primary); color: white; }
.ql-btn-primary:hover { background: var(--ql-primary-dark); }
.ql-btn-success { background: var(--ql-success); color: white; }
.ql-btn-success:hover { background: var(--ql-success-dark); }
.ql-btn-danger { background: var(--ql-danger); color: white; }
.ql-btn-danger:hover { background: var(--ql-danger-dark); }
.ql-btn-outline {
    background: white; color: var(--ql-primary);
    border: 1px solid var(--ql-primary);
}
.ql-btn-outline:hover { background: var(--ql-primary-light); }
.ql-btn-sm { padding: 5px 10px; font-size: 12px; }
.ql-btn-white { background: rgba(255,255,255,0.2); color: white; border: 1px solid rgba(255,255,255,0.4); }
.ql-btn-white:hover { background: rgba(255,255,255,0.3); }

/* Tabs */
.ql-tabs {
    display: flex; gap: 0; border-bottom: 2px solid var(--ql-border);
    margin-bottom: 16px; overflow-x: auto;
}
.ql-tab {
    padding: 10px 20px; cursor: pointer; font-weight: 600;
    color: var(--ql-text-muted); border-bottom: 3px solid transparent;
    transition: all 0.2s; white-space: nowrap; font-size: 13px;
    background: none; border-top: none; border-left: none; border-right: none;
}
.ql-tab:hover { color: var(--ql-primary); }
.ql-tab.active {
    color: var(--ql-primary); border-bottom-color: var(--ql-primary);
}
.ql-tab-content { display: none; }
.ql-tab-content.active { display: block; }

/* Card */
.ql-card {
    background: var(--ql-card); border-radius: var(--ql-radius);
    box-shadow: var(--ql-shadow); margin-bottom: 16px; overflow: hidden;
}
.ql-card-header {
    padding: 12px 16px; background: #fafbfc;
    border-bottom: 1px solid var(--ql-border);
    display: flex; align-items: center; justify-content: space-between;
}
.ql-card-header h3 { margin: 0; font-size: 1em; }
.ql-card-body { padding: 16px; }

/* Table */
.ql-table {
    width: 100%; border-collapse: collapse; font-size: 13px;
}
.ql-table th {
    text-align: left; padding: 10px 12px; background: #f8f9fa;
    border-bottom: 2px solid var(--ql-border); font-weight: 600;
    color: var(--ql-text-muted); font-size: 11px; text-transform: uppercase;
}
.ql-table td {
    padding: 8px 12px; border-bottom: 1px solid #f0f0f0;
    vertical-align: middle;
}
.ql-table tr:hover { background: #f8f9ff; }
.ql-table tr.drag-over { border-top: 2px solid var(--ql-primary); }

/* Drag handle */
.ql-drag { cursor: grab; color: #ccc; font-size: 16px; }
.ql-drag:active { cursor: grabbing; }

/* Badge */
.ql-badge {
    display: inline-block; padding: 2px 8px; border-radius: 12px;
    font-size: 11px; font-weight: 600; margin: 1px 2px;
}
.ql-badge-blue { background: var(--ql-primary-light); color: var(--ql-primary-dark); }
.ql-badge-green { background: #C8E6C9; color: #2E7D32; }
.ql-badge-orange { background: #FFE0B2; color: #E65100; }
.ql-badge-gray { background: #eee; color: #666; }
.ql-badge-red { background: #FFCDD2; color: #C62828; }

/* Icon preview */
.ql-icon-preview {
    width: 32px; height: 32px; display: inline-flex;
    align-items: center; justify-content: center;
    background: #f0f4f8; border-radius: 6px; color: var(--ql-primary);
    font-size: 16px;
}

/* Modal */
.ql-modal-overlay {
    display: none; position: fixed; top: 0; left: 0;
    width: 100%; height: 100%; background: rgba(0,0,0,0.5);
    z-index: 10000; justify-content: center; align-items: flex-start;
    padding-top: 40px; overflow-y: auto;
}
.ql-modal-overlay.active { display: flex; }
.ql-modal {
    background: white; border-radius: var(--ql-radius);
    box-shadow: 0 8px 32px rgba(0,0,0,0.2); width: 90%;
    max-width: 700px; max-height: 85vh; overflow-y: auto;
    margin-bottom: 40px;
}
.ql-modal-header {
    padding: 16px 20px; border-bottom: 1px solid var(--ql-border);
    display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; background: white; z-index: 1;
}
.ql-modal-header h3 { margin: 0; font-size: 1.1em; }
.ql-modal-close {
    background: none; border: none; font-size: 22px;
    cursor: pointer; color: #999; padding: 0 4px;
}
.ql-modal-close:hover { color: #333; }
.ql-modal-body { padding: 20px; }
.ql-modal-footer {
    padding: 12px 20px; border-top: 1px solid var(--ql-border);
    display: flex; justify-content: flex-end; gap: 8px;
    position: sticky; bottom: 0; background: white;
}

/* Form */
.ql-form-group { margin-bottom: 14px; }
.ql-form-group label {
    display: block; margin-bottom: 4px; font-weight: 600;
    font-size: 12px; color: var(--ql-text-muted);
}
.ql-form-group input[type="text"],
.ql-form-group input[type="number"],
.ql-form-group input[type="date"],
.ql-form-group select,
.ql-form-group textarea {
    width: 100%; padding: 8px 12px; border: 1px solid var(--ql-border);
    border-radius: 6px; font-size: 13px; font-family: inherit;
}
.ql-form-group textarea { resize: vertical; min-height: 80px; }
.ql-form-row { display: flex; gap: 12px; }
.ql-form-row .ql-form-group { flex: 1; }
.ql-form-check {
    display: flex; align-items: center; gap: 6px;
    margin-bottom: 6px; font-size: 13px;
}
.ql-form-check input[type="checkbox"] { margin: 0; }

/* Icon Picker */
.ql-icon-picker-btn {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 8px 14px; border: 1px dashed var(--ql-border);
    border-radius: 6px; cursor: pointer; background: #fafbfc;
    font-size: 13px; transition: all 0.2s;
}
.ql-icon-picker-btn:hover { border-color: var(--ql-primary); background: #f0f6ff; }
.ql-icon-picker-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(70px, 1fr));
    gap: 6px; max-height: 300px; overflow-y: auto;
    padding: 8px; border: 1px solid var(--ql-border); border-radius: 6px;
    margin-top: 8px;
}
.ql-icon-picker-tile {
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; padding: 8px 4px; border-radius: 6px;
    cursor: pointer; transition: all 0.15s; border: 2px solid transparent;
    font-size: 10px; text-align: center; min-height: 60px;
}
.ql-icon-picker-tile:hover { background: var(--ql-primary-light); }
.ql-icon-picker-tile.selected { border-color: var(--ql-primary); background: #e3f2fd; }
.ql-icon-picker-tile i { font-size: 20px; margin-bottom: 4px; color: #555; }
.ql-icon-picker-search {
    width: 100%; padding: 8px 12px; border: 1px solid var(--ql-border);
    border-radius: 6px; font-size: 13px; margin-bottom: 8px;
}
.ql-icon-picker-cats {
    display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 8px;
}
.ql-icon-picker-cat {
    padding: 3px 10px; border-radius: 12px; font-size: 11px;
    cursor: pointer; border: 1px solid var(--ql-border);
    background: white; transition: all 0.15s;
}
.ql-icon-picker-cat:hover, .ql-icon-picker-cat.active {
    background: var(--ql-primary); color: white; border-color: var(--ql-primary);
}

/* Roles selector */
.ql-roles-tags {
    display: flex; flex-wrap: wrap; gap: 4px; padding: 6px;
    border: 1px solid var(--ql-border); border-radius: 6px;
    min-height: 36px; background: white;
}
.ql-role-tag {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 2px 8px; background: var(--ql-primary-light);
    color: var(--ql-primary-dark); border-radius: 12px;
    font-size: 12px; font-weight: 500;
}
.ql-role-tag-remove {
    cursor: pointer; font-weight: bold; font-size: 14px;
    line-height: 1; opacity: 0.6;
}
.ql-role-tag-remove:hover { opacity: 1; }
.ql-roles-input {
    border: none; outline: none; font-size: 12px;
    flex: 1; min-width: 80px; padding: 2px 4px;
}
.ql-roles-dropdown {
    position: absolute; background: white; border: 1px solid var(--ql-border);
    border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    max-height: 200px; overflow-y: auto; z-index: 100; display: none;
    width: 200px;
}
.ql-roles-dropdown.show { display: block; }
.ql-roles-option {
    padding: 6px 12px; cursor: pointer; font-size: 12px;
}
.ql-roles-option:hover { background: #f0f4ff; }

/* Toast */
.ql-toast-container {
    position: fixed; top: 20px; right: 20px; z-index: 20000;
    display: flex; flex-direction: column; gap: 8px;
}
.ql-toast {
    padding: 12px 20px; border-radius: 8px; color: white;
    font-size: 13px; font-weight: 500; box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    animation: ql-toast-in 0.3s ease;
    max-width: 350px;
}
.ql-toast-success { background: var(--ql-success); }
.ql-toast-error { background: var(--ql-danger); }
.ql-toast-info { background: var(--ql-info); }
@keyframes ql-toast-in {
    from { opacity: 0; transform: translateX(40px); }
    to { opacity: 1; transform: translateX(0); }
}

/* Dirty indicator */
.ql-dirty-dot {
    display: none; width: 10px; height: 10px; background: var(--ql-danger);
    border-radius: 50%; margin-left: 8px;
}
.ql-dirty-dot.show { display: inline-block; }

/* Filter bar */
.ql-filter-bar {
    display: flex; gap: 8px; align-items: center; margin-bottom: 12px;
    flex-wrap: wrap;
}
.ql-filter-bar select, .ql-filter-bar input {
    padding: 6px 10px; border: 1px solid var(--ql-border);
    border-radius: 6px; font-size: 13px;
}

/* Truncate */
.ql-truncate {
    max-width: 200px; overflow: hidden; text-overflow: ellipsis;
    white-space: nowrap; display: inline-block;
}
.ql-truncate-sm { max-width: 120px; }

/* Tree View */
.ql-tree-cat {
    border: 1px solid var(--ql-border); border-radius: var(--ql-radius);
    margin-bottom: 8px; overflow: hidden;
}
.ql-tree-cat-header {
    display: flex; align-items: center; gap: 8px; padding: 10px 12px;
    background: #e8f0fe; cursor: pointer; user-select: none;
}
.ql-tree-cat-header:hover { background: #d2e3fc; }
.ql-tree-cat-header .ql-drag { margin-right: 2px; }
.ql-tree-cat-header .ql-tree-chevron { transition: transform 0.2s; font-size: 12px; color: #666; width: 14px; }
.ql-tree-cat-header .ql-tree-chevron.collapsed { transform: rotate(-90deg); }
.ql-tree-cat-title { font-weight: 700; font-size: 13px; flex: 1; display: flex; align-items: center; gap: 6px; }
.ql-tree-cat-actions { display: flex; gap: 4px; align-items: center; }
.ql-tree-cat-content { padding: 0; }
.ql-tree-cat-content.collapsed { display: none; }

.ql-tree-sg {
    border-top: 1px solid #f0f0f0;
}
.ql-tree-sg-header {
    display: flex; align-items: center; gap: 8px; padding: 8px 12px 8px 32px;
    background: #fff8e1; cursor: pointer; user-select: none; font-size: 12px;
}
.ql-tree-sg-header:hover { background: #fff3cd; }
.ql-tree-sg-title { font-weight: 600; flex: 1; display: flex; align-items: center; gap: 6px; }
.ql-tree-sg-actions { display: flex; gap: 4px; align-items: center; }
.ql-tree-sg-content { }
.ql-tree-sg-content.collapsed { display: none; }

.ql-tree-icon {
    display: flex; align-items: center; gap: 8px; padding: 6px 12px 6px 52px;
    border-top: 1px solid #f8f8f8; font-size: 12px; transition: background 0.1s;
}
.ql-tree-icon:hover { background: #f8f9ff; }
.ql-tree-icon.ungrouped { padding-left: 32px; }
.ql-tree-icon .ql-icon-preview { width: 26px; height: 26px; font-size: 13px; flex-shrink: 0; }
.ql-tree-icon-label { font-weight: 600; min-width: 120px; }
.ql-tree-icon-link { color: var(--ql-text-muted); max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ql-tree-icon-meta { flex: 1; display: flex; gap: 4px; flex-wrap: wrap; align-items: center; }
.ql-tree-icon-actions { display: flex; gap: 4px; flex-shrink: 0; }
.ql-tree-ungrouped-label { padding: 6px 12px 4px 32px; font-size: 11px; color: var(--ql-text-muted); font-weight: 600; text-transform: uppercase; border-top: 1px solid #f0f0f0; }

/* Nested subgroups (child inside parent) */
.ql-tree-sg .ql-tree-sg { margin-left: 16px; border-left: 2px solid #ffe0b2; }
.ql-tree-sg .ql-tree-sg .ql-tree-sg-header { padding-left: 16px; background: #fff8e1; font-size: 11px; }
.ql-tree-sg .ql-tree-sg .ql-tree-icon { padding-left: 36px; }

.ql-tree-drop-above { border-top: 2px solid var(--ql-primary) !important; }
.ql-tree-drop-below { border-bottom: 2px solid var(--ql-primary) !important; }
.ql-tree-dragging { opacity: 0.4; }
.ql-tree-count { background: #e0e0e0; color: #555; font-size: 10px; padding: 1px 6px; border-radius: 8px; font-weight: 600; }

/* JSON editor */
#ql-json-editor {
    width: 100%; min-height: 400px; font-family: 'Consolas', 'Monaco', monospace;
    font-size: 12px; padding: 12px; border: 1px solid var(--ql-border);
    border-radius: 6px; resize: vertical;
}

/* Loading spinner */
.ql-spinner {
    display: inline-block; width: 16px; height: 16px;
    border: 2px solid rgba(255,255,255,0.3); border-top-color: white;
    border-radius: 50%; animation: ql-spin 0.6s linear infinite;
}
@keyframes ql-spin { to { transform: rotate(360deg); } }

/* Empty state */
.ql-empty {
    text-align: center; padding: 40px; color: var(--ql-text-muted);
}
.ql-empty i { font-size: 40px; margin-bottom: 12px; display: block; opacity: 0.3; }

/* Tutorial overlay */
.ql-tour-overlay {
    display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.6); z-index: 30000;
}
.ql-tour-overlay.active { display: block; }
.ql-tour-spotlight {
    position: absolute; border-radius: 8px;
    box-shadow: 0 0 0 9999px rgba(0,0,0,0.55); z-index: 30001;
    transition: all 0.3s ease;
}
.ql-tour-card {
    position: absolute; background: white; border-radius: 12px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3); padding: 20px 24px;
    max-width: 380px; width: 90%; z-index: 30002;
    animation: ql-tour-pop 0.25s ease;
}
@keyframes ql-tour-pop {
    from { opacity: 0; transform: scale(0.9); }
    to { opacity: 1; transform: scale(1); }
}
.ql-tour-card h4 { margin: 0 0 8px; font-size: 15px; color: var(--ql-primary); }
.ql-tour-card p { margin: 0 0 14px; font-size: 13px; color: #444; line-height: 1.5; }
.ql-tour-step { font-size: 11px; color: var(--ql-text-muted); margin-bottom: 10px; }
.ql-tour-actions { display: flex; justify-content: space-between; align-items: center; }
.ql-tour-actions .ql-btn { min-width: 70px; justify-content: center; }

/* Inline help tooltip */
.ql-help {
    display: inline-flex; align-items: center; justify-content: center;
    width: 15px; height: 15px; border-radius: 50%; background: #ddd;
    color: #666; font-size: 10px; font-weight: 700; cursor: help;
    margin-left: 4px; position: relative; vertical-align: middle;
    line-height: 1;
}
.ql-help:hover { background: var(--ql-primary); color: white; }
.ql-help-tip {
    display: none; position: fixed; background: #333; color: white;
    padding: 8px 12px; border-radius: 6px; font-size: 11px; font-weight: 400;
    white-space: normal; width: 240px; z-index: 50000; line-height: 1.4;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2); text-align: left;
    pointer-events: none;
}

/* Responsive */
@media (max-width: 768px) {
    #ql-admin-root { padding: 8px; }
    .ql-header { flex-direction: column; gap: 8px; }
    .ql-form-row { flex-direction: column; gap: 0; }
    .ql-modal { width: 95%; }
    .ql-table { font-size: 12px; }
    .ql-table th, .ql-table td { padding: 6px 8px; }
}
</style>
</head>
<body>
<div id="ql-admin-root">
<div class="ql-toast-container" id="ql-toasts"></div>

<div class="ql-header">
    <div>
        <h2><i class="fa fa-th"></i> QuickLinks Admin</h2>
    </div>
    <div class="ql-header-actions">
        <span class="ql-dirty-dot" id="ql-dirty-dot" title="Unsaved changes"></span>
        <button class="ql-btn ql-btn-white" onclick="QL.startTour()"><i class="fa fa-question-circle"></i> Tour</button>
        <button class="ql-btn ql-btn-white" onclick="QL.loadConfig()"><i class="fa fa-refresh"></i> Reload</button>
        <button class="ql-btn ql-btn-success" onclick="QL.saveConfig()"><i class="fa fa-save"></i> Save</button>
    </div>
</div>

<div class="ql-tabs">
    <button class="ql-tab active" data-tab="structure" onclick="QL.switchTab('structure')"><i class="fa fa-sitemap"></i> Structure</button>
    <button class="ql-tab" data-tab="queries" onclick="QL.switchTab('queries')"><i class="fa fa-database"></i> Count Queries</button>
    <button class="ql-tab" data-tab="settings" onclick="QL.switchTab('settings')"><i class="fa fa-cog"></i> Settings & Backup</button>
</div>

<!-- Structure Tab -->
<div class="ql-tab-content active" id="tab-structure">
    <div class="ql-card">
        <div class="ql-card-header">
            <h3><i class="fa fa-sitemap"></i> Link Structure</h3>
            <div style="display:flex;gap:6px;align-items:center;">
                <input type="text" id="ql-tree-search" placeholder="Search links..." oninput="QL.renderStructure()" style="padding:5px 10px;border:1px solid var(--ql-border);border-radius:6px;font-size:12px;width:180px;">
                <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.expandAll()"><i class="fa fa-expand"></i> Expand</button>
                <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.collapseAll()"><i class="fa fa-compress"></i> Collapse</button>
                <button class="ql-btn ql-btn-primary ql-btn-sm" onclick="QL.addCategory()"><i class="fa fa-plus"></i> Add Category</button>
            </div>
        </div>
        <div class="ql-card-body" style="padding:8px;" id="ql-tree-root"></div>
    </div>
</div>

<!-- Queries Tab -->
<div class="ql-tab-content" id="tab-queries">
    <div class="ql-card">
        <div class="ql-card-header">
            <h3>Count Queries</h3>
            <button class="ql-btn ql-btn-primary ql-btn-sm" onclick="QL.addQuery()"><i class="fa fa-plus"></i> Add Query</button>
        </div>
        <div class="ql-card-body" style="padding:0;">
            <table class="ql-table">
                <thead><tr>
                    <th>Query Name</th>
                    <th>Description</th>
                    <th>Uses Current User</th>
                    <th style="width:140px;">Actions</th>
                </tr></thead>
                <tbody id="ql-q-tbody"></tbody>
            </table>
        </div>
    </div>
</div>

<!-- Settings Tab -->
<div class="ql-tab-content" id="tab-settings">
    <div class="ql-card">
        <div class="ql-card-header">
            <h3><i class="fa fa-layer-group"></i> Subgroups</h3>
            <button class="ql-btn ql-btn-primary ql-btn-sm" onclick="QL.addSubgroup()"><i class="fa fa-plus"></i> Add Subgroup</button>
        </div>
        <div class="ql-card-body" style="padding:0;">
            <p style="padding:12px 16px 0;margin:0;font-size:12px;color:var(--ql-text-muted);">
                Subgroups are <strong>shared definitions</strong> &mdash; any icon in any category can reference a subgroup by name.
                For example, both Finance and Tools can have icons assigned to a "Reports" subgroup. The display name, icon, and popup
                setting are defined here once.
            </p>
            <div style="padding:0;" id="ql-sg-manage-area"></div>
        </div>
    </div>

    <div class="ql-card">
        <div class="ql-card-header"><h3>Settings</h3></div>
        <div class="ql-card-body">
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Auto-Subgroup Threshold</label>
                    <input type="number" id="ql-setting-threshold" value="8" min="1" max="50" onchange="QL.updateSetting('auto_subgroup_threshold', parseInt(this.value))">
                    <small style="color:var(--ql-text-muted);font-size:11px;">Categories with more visible icons than this will use subgroups</small>
                </div>
                <div class="ql-form-group">
                    <label>Hide Small Categories (min icons)</label>
                    <input type="number" id="ql-setting-hide-small" value="0" min="0" max="20" onchange="QL.updateSetting('hide_small_categories', parseInt(this.value))">
                    <small style="color:var(--ql-text-muted);font-size:11px;">0 = never hide</small>
                </div>
            </div>
        </div>
    </div>

    <div class="ql-card">
        <div class="ql-card-header">
            <h3>Import / Export</h3>
        </div>
        <div class="ql-card-body">
            <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;">
                <button class="ql-btn ql-btn-outline" onclick="QL.importHardcoded()"><i class="fa fa-download"></i> Import from Hardcoded Defaults</button>
                <button class="ql-btn ql-btn-outline" onclick="QL.exportJson()"><i class="fa fa-file-export"></i> Export JSON</button>
            </div>
        </div>
    </div>

    <div class="ql-card">
        <div class="ql-card-header"><h3>Raw JSON Editor</h3></div>
        <div class="ql-card-body">
            <textarea id="ql-json-editor"></textarea>
            <div style="margin-top:8px;display:flex;gap:8px;">
                <button class="ql-btn ql-btn-primary ql-btn-sm" onclick="QL.applyJson()"><i class="fa fa-check"></i> Apply JSON</button>
                <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.refreshJsonEditor()"><i class="fa fa-refresh"></i> Refresh from Config</button>
            </div>
        </div>
    </div>

    <div class="ql-card">
        <div class="ql-card-header"><h3>Backups</h3></div>
        <div class="ql-card-body">
            <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.loadBackups()" style="margin-bottom:12px;"><i class="fa fa-list"></i> Load Backup List</button>
            <div id="ql-backup-list"></div>
        </div>
    </div>
</div>

<!-- Category Edit Modal -->
<div class="ql-modal-overlay" id="modal-category">
    <div class="ql-modal">
        <div class="ql-modal-header">
            <h3 id="modal-cat-title">Edit Category</h3>
            <button class="ql-modal-close" onclick="QL.closeModal('modal-category')">&times;</button>
        </div>
        <div class="ql-modal-body">
            <input type="hidden" id="cat-edit-idx">
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Category ID <span class="ql-help">?<span class="ql-help-tip">A unique key used internally (e.g. "general", "finance"). Cannot contain spaces. Icons reference this ID.</span></span></label>
                    <input type="text" id="cat-id" placeholder="e.g. general">
                </div>
                <div class="ql-form-group">
                    <label>Display Name <span class="ql-help">?<span class="ql-help-tip">The name shown in the widget header for this category.</span></span></label>
                    <input type="text" id="cat-name" placeholder="e.g. General">
                </div>
            </div>
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Icon</label>
                    <div class="ql-icon-picker-btn" onclick="QL.openIconPicker('cat-icon-val')">
                        <i id="cat-icon-preview" class="fa fa-church"></i>
                        <span id="cat-icon-label">fa-church</span>
                    </div>
                    <input type="hidden" id="cat-icon-val" value="fa-church">
                </div>
                <div class="ql-form-group">
                    <label>Expanded by Default <span class="ql-help">?<span class="ql-help-tip">When "Yes", this category starts open in the widget. Users can still toggle it.</span></span></label>
                    <select id="cat-expanded">
                        <option value="true">Yes</option>
                        <option value="false">No</option>
                    </select>
                </div>
            </div>
            <div class="ql-form-group">
                <label>Required Roles <span class="ql-help">?<span class="ql-help-tip">Which TouchPoint roles can see this item. Leave empty so everyone can see it. Add multiple roles - users with ANY of them get access.</span></span></label>
                <div id="cat-roles-container" style="position:relative;"></div>
            </div>
            <input type="hidden" id="cat-sort" value="0">
        </div>
        <div class="ql-modal-footer">
            <button class="ql-btn ql-btn-outline" onclick="QL.closeModal('modal-category')">Cancel</button>
            <button class="ql-btn ql-btn-primary" onclick="QL.saveCategory()">Save</button>
        </div>
    </div>
</div>

<!-- Icon Edit Modal -->
<div class="ql-modal-overlay" id="modal-icon">
    <div class="ql-modal" style="max-width:750px;">
        <div class="ql-modal-header">
            <h3 id="modal-icon-title">Edit Icon</h3>
            <button class="ql-modal-close" onclick="QL.closeModal('modal-icon')">&times;</button>
        </div>
        <div class="ql-modal-body">
            <input type="hidden" id="icon-edit-idx">
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Label <span class="ql-help">?<span class="ql-help-tip">The text shown below the icon in the widget. Keep it short (1-2 words).</span></span></label>
                    <input type="text" id="icon-label" placeholder="e.g. Hospital">
                </div>
                <div class="ql-form-group">
                    <label>Category <span class="ql-help">?<span class="ql-help-tip">Which category this link appears in. Links are grouped by category in the widget.</span></span></label>
                    <select id="icon-category"></select>
                </div>
            </div>
            <div class="ql-form-group">
                <label>Link URL <span class="ql-help">?<span class="ql-help-tip">Where the icon navigates. Use relative paths for TouchPoint (e.g. /PyScript/MyScript) or full URLs for external links.</span></span></label>
                <input type="text" id="icon-link" placeholder="/PyScript/MyScript">
            </div>
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Icon</label>
                    <div class="ql-icon-picker-btn" onclick="QL.openIconPicker('icon-icon-val')">
                        <i id="icon-icon-preview" class="fa fa-link"></i>
                        <span id="icon-icon-label">fa-link</span>
                    </div>
                    <input type="hidden" id="icon-icon-val" value="fa-link">
                    <input type="hidden" id="icon-priority" value="50">
                </div>
            </div>
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Org ID <span class="ql-help">?<span class="ql-help-tip">TouchPoint Organization ID. If set with a count query, the org ID is passed to the query. If set alone, shows the org's member count as a badge.</span></span></label>
                    <input type="number" id="icon-orgid" placeholder="e.g. 819">
                </div>
                <div class="ql-form-group">
                    <label>Count Query <span class="ql-help">?<span class="ql-help-tip">Select a named SQL query (defined in Settings tab) to show a count badge on this icon. Use {0} for org ID or {} for current user ID in the query.</span></span></label>
                    <select id="icon-query">
                        <option value="">None</option>
                    </select>
                </div>
            </div>
            <div class="ql-form-row">
                <div class="ql-form-group">
                    <label>Subgroup <span class="ql-help">?<span class="ql-help-tip">Assign this icon to a subgroup. When a category has many icons, subgroups organize them into sections. Manage subgroups in the Settings tab.</span></span></label>
                    <select id="icon-subgroup">
                        <option value="">None</option>
                    </select>
                </div>
                <div class="ql-form-group">
                    <label>Expiration <span class="ql-help">?<span class="ql-help-tip">After this date, the icon automatically hides from the widget. Useful for seasonal events like VBS.</span></span></label>
                    <input type="date" id="icon-expiration">
                </div>
            </div>
            <div class="ql-form-group">
                <label>Required Roles <span class="ql-help">?<span class="ql-help-tip">Which TouchPoint roles can see this item. Leave empty so everyone can see it. Add multiple roles - users with ANY of them get access.</span></span></label>
                <div id="icon-roles-container" style="position:relative;"></div>
            </div>
        </div>
        <div class="ql-modal-footer">
            <button class="ql-btn ql-btn-outline" onclick="QL.closeModal('modal-icon')">Cancel</button>
            <button class="ql-btn ql-btn-primary" onclick="QL.saveIcon()">Save</button>
        </div>
    </div>
</div>

<!-- Subgroup Edit Modal -->
<div class="ql-modal-overlay" id="modal-subgroup">
    <div class="ql-modal" style="max-width:500px;">
        <div class="ql-modal-header">
            <h3 id="modal-sg-title">Edit Subgroup</h3>
            <button class="ql-modal-close" onclick="QL.closeModal('modal-subgroup')">&times;</button>
        </div>
        <div class="ql-modal-body">
            <input type="hidden" id="sg-edit-key">
            <div class="ql-form-group">
                <label>Subgroup ID <span class="ql-help">?<span class="ql-help-tip">A unique key for this subgroup. Icons reference this ID in their Subgroup field. Use simple names without spaces (e.g. "Reports").</span></span></label>
                <input type="text" id="sg-id" placeholder="e.g. Reports">
            </div>
            <div class="ql-form-group">
                <label>Display Name</label>
                <input type="text" id="sg-name" placeholder="e.g. Financial Reports">
            </div>
            <div class="ql-form-group">
                <label>Icon</label>
                <div class="ql-icon-picker-btn" onclick="QL.openIconPicker('sg-icon-val')">
                    <i id="sg-icon-preview" class="fa fa-folder"></i>
                    <span id="sg-icon-label">fa-folder</span>
                </div>
                <input type="hidden" id="sg-icon-val" value="fa-folder">
            </div>
            <div class="ql-form-group">
                <label>Popup Mode <span class="ql-help">?<span class="ql-help-tip">Inline: shows icons directly in the category. Popup: shows as a single clickable tile that opens a popup with all its icons.</span></span></label>
                <select id="sg-popup">
                    <option value="false">No (inline)</option>
                    <option value="true">Yes (collapsible popup)</option>
                </select>
            </div>
            <div class="ql-form-group">
                <label>Parent Subgroup (optional - nests this inside another subgroup in admin view)</label>
                <select id="sg-parent">,
                    <option value="">None (top-level)</option>
                </select>
                <small style="color:var(--ql-text-muted);font-size:11px;">Widget rendering is unaffected - this only organizes the admin tree</small>
            </div>
        </div>
        <div class="ql-modal-footer">
            <button class="ql-btn ql-btn-outline" onclick="QL.closeModal('modal-subgroup')">Cancel</button>
            <button class="ql-btn ql-btn-primary" onclick="QL.saveSubgroup()">Save</button>
        </div>
    </div>
</div>

<!-- Query Edit Modal -->
<div class="ql-modal-overlay" id="modal-query">
    <div class="ql-modal" style="max-width:650px;">
        <div class="ql-modal-header">
            <h3 id="modal-q-title">Edit Count Query</h3>
            <button class="ql-modal-close" onclick="QL.closeModal('modal-query')">&times;</button>
        </div>
        <div class="ql-modal-body">
            <input type="hidden" id="q-edit-key">
            <div class="ql-form-group">
                <label>Query Name (key)</label>
                <input type="text" id="q-name" placeholder="e.g. member_count">
            </div>
            <div class="ql-form-group">
                <label>Description</label>
                <input type="text" id="q-desc" placeholder="e.g. Member count for org">
            </div>
            <div class="ql-form-group">
                <label>SQL (use {0} for org_id, {} for current user PeopleId)</label>
                <textarea id="q-sql" style="min-height:120px;font-family:Consolas,Monaco,monospace;font-size:12px;"></textarea>
            </div>
            <div class="ql-form-check">
                <input type="checkbox" id="q-uses-user">
                <label for="q-uses-user" style="margin:0;font-weight:normal;font-size:13px;">Uses current user ID ({})</label>
            </div>
        </div>
        <div class="ql-modal-footer">
            <button class="ql-btn ql-btn-outline" onclick="QL.closeModal('modal-query')">Cancel</button>
            <button class="ql-btn ql-btn-primary" onclick="QL.saveQuery()">Save</button>
        </div>
    </div>
</div>

<!-- Icon Picker Modal -->
<div class="ql-modal-overlay" id="modal-iconpicker">
    <div class="ql-modal" style="max-width:600px;">
        <div class="ql-modal-header">
            <h3>Select Icon</h3>
            <button class="ql-modal-close" onclick="QL.closeModal('modal-iconpicker')">&times;</button>
        </div>
        <div class="ql-modal-body">
            <input type="text" class="ql-icon-picker-search" id="ip-search" placeholder="Search icons..." oninput="QL.filterIconPicker()">
            <div class="ql-icon-picker-cats" id="ip-cats"></div>
            <div class="ql-icon-picker-grid" id="ip-grid"></div>
            <div class="ql-form-group" style="margin-top:12px;">
                <label>Or enter custom FA class:</label>
                <input type="text" id="ip-custom" placeholder="e.g. fas fa-user-graduate">
            </div>
        </div>
        <div class="ql-modal-footer">
            <button class="ql-btn ql-btn-outline" onclick="QL.closeModal('modal-iconpicker')">Cancel</button>
            <button class="ql-btn ql-btn-primary" onclick="QL.selectIconFromPicker()">Select</button>
        </div>
    </div>
</div>

<!-- Backup Preview Modal -->
<div class="ql-modal-overlay" id="modal-backup-preview">
    <div class="ql-modal" style="max-width:700px;">
        <div class="ql-modal-header">
            <h3>Backup Preview</h3>
            <button class="ql-modal-close" onclick="QL.closeModal('modal-backup-preview')">&times;</button>
        </div>
        <div class="ql-modal-body">
            <textarea id="backup-preview-content" style="width:100%;min-height:400px;font-family:Consolas,Monaco,monospace;font-size:11px;" readonly></textarea>
        </div>
        <div class="ql-modal-footer">
            <button class="ql-btn ql-btn-outline" onclick="QL.closeModal('modal-backup-preview')">Close</button>
        </div>
    </div>
</div>

</div>

<!-- Tutorial Overlay -->
<div class="ql-tour-overlay" id="ql-tour-overlay" onclick="QL.endTour()">
    <div class="ql-tour-spotlight" id="ql-tour-spotlight"></div>
    <div class="ql-tour-card" id="ql-tour-card" onclick="event.stopPropagation()">
        <div class="ql-tour-step" id="ql-tour-step"></div>
        <h4 id="ql-tour-title"></h4>
        <p id="ql-tour-text"></p>
        <div class="ql-tour-actions">
            <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.endTour()">Skip</button>
            <div>
                <button class="ql-btn ql-btn-outline ql-btn-sm" id="ql-tour-prev" onclick="QL.tourStep(-1)"><i class="fa fa-arrow-left"></i> Back</button>
                <button class="ql-btn ql-btn-primary ql-btn-sm" id="ql-tour-next" onclick="QL.tourStep(1)">Next <i class="fa fa-arrow-right"></i></button>
            </div>
        </div>
    </div>
</div>

<script>
(function(){
// --- PyScript/PyScriptForm redirect ---
var loc = window.location.pathname;
if (loc.indexOf('/PyScript/') !== -1 && loc.indexOf('/PyScriptForm/') === -1) {
    window.location.href = loc.replace('/PyScript/', '/PyScriptForm/');
    return;
}

var scriptName = 'TPxi_QuickLinksAdmin';
var parts = loc.split('/');
for (var i = 0; i < parts.length; i++) {
    if ((parts[i] === 'PyScriptForm' || parts[i] === 'PyScript') && i + 1 < parts.length) {
        scriptName = parts[i + 1].split('?')[0];
        break;
    }
}
var scriptUrl = '/PyScriptForm/' + scriptName;

// --- Icon Picker Data ---
var FA_ICONS = {
    "General": [
        {i:"fa-home",l:"Home"},{i:"fa-church",l:"Church"},{i:"fa-globe",l:"Globe"},{i:"fa-briefcase",l:"Briefcase"},
        {i:"fa-star",l:"Star"},{i:"fa-flag",l:"Flag"},{i:"fa-heart",l:"Heart"},{i:"fa-bookmark",l:"Bookmark"},
        {i:"fa-building",l:"Building"},{i:"fa-university",l:"University"},{i:"fa-map-marker-alt",l:"Map Pin"},
        {i:"fa-calendar",l:"Calendar"},{i:"fa-calendar-alt",l:"Calendar Alt"},{i:"fa-calendar-check",l:"Calendar Check"},
        {i:"fa-clock",l:"Clock"},{i:"fa-bell",l:"Bell"},{i:"fa-bullhorn",l:"Bullhorn"},{i:"fa-gift",l:"Gift"},
        {i:"fa-trophy",l:"Trophy"},{i:"fa-child",l:"Child"},{i:"fa-baby",l:"Baby"},{i:"fa-cross",l:"Cross"},
        {i:"fa-hands-praying",l:"Praying"},{i:"fa-dove",l:"Dove"},{i:"fa-book",l:"Book"},
        {i:"fa-book-open",l:"Book Open"},{i:"fa-bible",l:"Bible"},{i:"fa-graduation-cap",l:"Grad Cap"},
        {i:"fa-bus",l:"Bus"},{i:"fa-car",l:"Car"},{i:"fa-plane",l:"Plane"},{i:"fa-anchor",l:"Anchor"}
    ],
    "People": [
        {i:"fa-user",l:"User"},{i:"fa-users",l:"Users"},{i:"fa-user-plus",l:"User Plus"},
        {i:"fa-user-minus",l:"User Minus"},{i:"fa-user-check",l:"User Check"},{i:"fa-user-lock",l:"User Lock"},
        {i:"fa-user-shield",l:"User Shield"},{i:"fa-user-cog",l:"User Cog"},{i:"fa-user-slash",l:"User Slash"},
        {i:"fa-user-tie",l:"User Tie"},{i:"fa-user-graduate",l:"Graduate"},{i:"fa-user-nurse",l:"Nurse"},
        {i:"fa-people-group",l:"Group"},{i:"fa-person",l:"Person"},{i:"fa-person-chalkboard",l:"Chalkboard"},
        {i:"fa-users-cog",l:"Users Cog"},{i:"fa-handshake",l:"Handshake"},{i:"fa-hand-holding-heart",l:"Holding Heart"},
        {i:"fa-chalkboard-teacher",l:"Teacher"}
    ],
    "Communication": [
        {i:"fa-envelope",l:"Envelope"},{i:"fa-envelope-open-text",l:"Envelope Open"},{i:"fa-paper-plane",l:"Paper Plane"},
        {i:"fa-phone",l:"Phone"},{i:"fa-comment",l:"Comment"},{i:"fa-comments",l:"Comments"},
        {i:"fa-bullseye",l:"Bullseye"},{i:"fa-at",l:"At"},{i:"fa-inbox",l:"Inbox"},
        {i:"fa-share",l:"Share"},{i:"fa-rss",l:"RSS"},{i:"fa-wifi",l:"WiFi"},
        {i:"fa-headphones",l:"Headphones"},{i:"fa-microphone",l:"Microphone"}
    ],
    "Finance": [
        {i:"fa-dollar-sign",l:"Dollar"},{i:"fa-credit-card",l:"Credit Card"},{i:"fa-wallet",l:"Wallet"},
        {i:"fa-hand-holding-usd",l:"Holding USD"},{i:"fa-coins",l:"Coins"},{i:"fa-money-bill",l:"Money Bill"},
        {i:"fa-file-invoice-dollar",l:"Invoice"},{i:"fa-cash-register",l:"Cash Register"},
        {i:"fa-building-columns",l:"Columns"},{i:"fa-receipt",l:"Receipt"},{i:"fa-piggy-bank",l:"Piggy Bank"},
        {i:"fa-chart-pie",l:"Chart Pie"},{i:"fa-calculator",l:"Calculator"}
    ],
    "Charts": [
        {i:"fa-chart-bar",l:"Chart Bar"},{i:"fa-chart-line",l:"Chart Line"},{i:"fa-chart-area",l:"Chart Area"},
        {i:"fa-table",l:"Table"},{i:"fa-database",l:"Database"},{i:"fa-gauge",l:"Gauge"},
        {i:"fa-tachometer",l:"Tachometer"},{i:"fa-signal",l:"Signal"},{i:"fa-list",l:"List"},
        {i:"fa-list-check",l:"List Check"},{i:"fa-bars",l:"Bars"},{i:"fa-project-diagram",l:"Diagram"}
    ],
    "Medical": [
        {i:"fa-hospital",l:"Hospital"},{i:"fa-first-aid",l:"First Aid"},{i:"fa-heartbeat",l:"Heartbeat"},
        {i:"fa-stethoscope",l:"Stethoscope"},{i:"fa-pills",l:"Pills"},{i:"fa-ambulance",l:"Ambulance"},
        {i:"fa-medkit",l:"Medkit"}
    ],
    "Security": [
        {i:"fa-lock",l:"Lock"},{i:"fa-unlock",l:"Unlock"},{i:"fa-shield-alt",l:"Shield"},
        {i:"fa-key",l:"Key"},{i:"fa-eye",l:"Eye"},{i:"fa-eye-slash",l:"Eye Slash"},
        {i:"fa-fire",l:"Fire"},{i:"fa-camera",l:"Camera"},{i:"fa-fingerprint",l:"Fingerprint"},
        {i:"fa-id-card",l:"ID Card"},{i:"fa-user-secret",l:"User Secret"}
    ],
    "Technology": [
        {i:"fa-cogs",l:"Cogs"},{i:"fa-cog",l:"Cog"},{i:"fa-code",l:"Code"},{i:"fa-server",l:"Server"},
        {i:"fa-tools",l:"Tools"},{i:"fa-toolbox",l:"Toolbox"},{i:"fa-plug",l:"Plug"},{i:"fa-print",l:"Print"},
        {i:"fa-tag",l:"Tag"},{i:"fa-tags",l:"Tags"},{i:"fa-desktop",l:"Desktop"},{i:"fa-laptop",l:"Laptop"},
        {i:"fa-mobile-alt",l:"Mobile"},{i:"fa-battery-full",l:"Battery"},{i:"fa-bolt",l:"Bolt"},
        {i:"fa-cloud",l:"Cloud"},{i:"fa-folder",l:"Folder"},{i:"fa-hdd",l:"HDD"},
        {i:"fa-windows",l:"Windows"},{i:"fa-network-wired",l:"Network"}
    ],
    "UI": [
        {i:"fa-check",l:"Check"},{i:"fa-check-circle",l:"Check Circle"},{i:"fa-check-square",l:"Check Square"},
        {i:"fa-plus",l:"Plus"},{i:"fa-plus-circle",l:"Plus Circle"},{i:"fa-minus",l:"Minus"},
        {i:"fa-edit",l:"Edit"},{i:"fa-pen-to-square",l:"Pen Square"},{i:"fa-trash",l:"Trash"},
        {i:"fa-download",l:"Download"},{i:"fa-upload",l:"Upload"},{i:"fa-file-alt",l:"File"},
        {i:"fa-file-export",l:"File Export"},{i:"fa-arrow-right",l:"Arrow Right"},{i:"fa-arrow-left",l:"Arrow Left"},
        {i:"fa-arrows-alt",l:"Arrows"},{i:"fa-sign-in-alt",l:"Sign In"},{i:"fa-sign-out-alt",l:"Sign Out"},
        {i:"fa-refresh",l:"Refresh"},{i:"fa-sync",l:"Sync"},{i:"fa-search",l:"Search"},
        {i:"fa-filter",l:"Filter"},{i:"fa-sort",l:"Sort"},{i:"fa-ellipsis-h",l:"Ellipsis"},
        {i:"fa-th",l:"Grid"},{i:"fa-circle-o-notch",l:"Spinner"},{i:"fa-hourglass-half",l:"Hourglass"},
        {i:"fa-question-circle",l:"Question"},{i:"fa-info-circle",l:"Info"},{i:"fa-exclamation-triangle",l:"Warning"},
        {i:"fa-layer-group",l:"Layers"},{i:"fa-cubes",l:"Cubes"},{i:"fa-tasks",l:"Tasks"},
        {i:"fa-history",l:"History"},{i:"fa-flask",l:"Flask"}
    ]
};

// --- DEFAULT CONFIG (starter example for new churches) ---
var DEFAULT_CONFIG = {
  "version": 1,
  "settings": {
    "auto_subgroup_threshold": 8,
    "hide_small_categories": 0
  },
  "categories": [
    {"id":"general","icon":"fa-church","name":"General","expanded":true,"roles":null,"sort_order":0},
    {"id":"reports","icon":"fa-chart-bar","name":"Reports","expanded":true,"roles":["Edit"],"sort_order":1},
    {"id":"finance","icon":"fa-university","name":"Finance","expanded":false,"roles":["Finance"],"sort_order":2},
    {"id":"tools","icon":"fa-tools","name":"Tools","expanded":false,"roles":["Admin"],"sort_order":3}
  ],
  "icons": [
    {"id":"general_emails","category_id":"general","icon":"fa-envelope","label":"Emails","link":"Person2/0#tab-receivedemails","org_id":null,"query_name":null,"roles":null,"priority":900,"subgroup":null,"expiration_date":null},
    {"id":"general_tasks","category_id":"general","icon":"fa-check-square","label":"Tasks","link":"TaskNoteIndex?v=Action","org_id":null,"query_name":"taskCount","roles":null,"priority":899,"subgroup":null,"expiration_date":null},
    {"id":"reports_attendance","category_id":"reports","icon":"fa-chart-line","label":"Weekly Attendance","link":"/PyScript/TPxi_WeeklyAttendance","org_id":null,"query_name":null,"roles":["Admin"],"priority":900,"subgroup":null,"expiration_date":null},
    {"id":"finance_giving","category_id":"finance","icon":"fa-dollar-sign","label":"Giving Dashboard","link":"/PyScript/TPxi_GivingDashboard","org_id":null,"query_name":null,"roles":["Finance"],"priority":900,"subgroup":null,"expiration_date":null},
    {"id":"tools_data_quality","category_id":"tools","icon":"fa-database","label":"Data Quality","link":"/PyScript/TPxi_DataQualityDashboard","org_id":null,"query_name":null,"roles":["Admin"],"priority":900,"subgroup":null,"expiration_date":null},
    {"id":"tools_quicklinks_admin","category_id":"tools","icon":"fa-th","label":"QuickLinks Admin","link":"/PyScript/TPxi_QuickLinksAdmin","org_id":null,"query_name":null,"roles":["Admin"],"priority":899,"subgroup":null,"expiration_date":null}
  ],
  "subgroups": {},
  "count_queries": {
    "member_count":{"sql":"SELECT COUNT(*) FROM OrganizationMembers WHERE OrganizationId = {0}","description":"Member count for org","uses_current_user":false},
    "taskCount":{"sql":"SELECT COUNT(*) FROM TaskNote WHERE AssigneeId = {} AND CompletedDate IS NULL AND StatusId NOT IN (5,1)","description":"Open tasks for current user","uses_current_user":true}
  }
};

// --- Main Application ---
var QL = {
    config: { version:1, settings:{auto_subgroup_threshold:8,hide_small_categories:0}, categories:[], icons:[], subgroups:{}, count_queries:{} },
    dirty: false,
    allRoles: [],
    _iconPickerTarget: null,
    _iconPickerSelected: null,
    _iconPickerCat: 'All',

    // --- AJAX ---
    ajax: function(action, params, cb) {
        params = params || {};
        params.action = action;
        var data = '';
        for (var k in params) {
            if (params.hasOwnProperty(k)) {
                if (data) data += '&';
                data += encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
            }
        }
        var xhr = new XMLHttpRequest();
        xhr.open('POST', scriptUrl, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                try {
                    var r = JSON.parse(xhr.responseText);
                    cb(r);
                } catch(e) {
                    cb({success:false, message:'Parse error: ' + e.message + ' | Response: ' + (xhr.responseText||'').substring(0,200)});
                }
            }
        };
        xhr.send(data);
    },

    encodeForPost: function(s) {
        return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
    },

    // --- Toast ---
    toast: function(msg, type) {
        type = type || 'info';
        var c = document.getElementById('ql-toasts');
        var t = document.createElement('div');
        t.className = 'ql-toast ql-toast-' + type;
        t.textContent = msg;
        c.appendChild(t);
        setTimeout(function(){ if(t.parentNode) t.parentNode.removeChild(t); }, 3500);
    },

    // --- Dirty tracking ---
    markDirty: function() {
        QL.dirty = true;
        var dot = document.getElementById('ql-dirty-dot');
        if (dot) dot.className = 'ql-dirty-dot show';
    },
    clearDirty: function() {
        QL.dirty = false;
        var dot = document.getElementById('ql-dirty-dot');
        if (dot) dot.className = 'ql-dirty-dot';
    },

    // --- Tab switching ---
    switchTab: function(tab) {
        var tabs = document.querySelectorAll('.ql-tab');
        var contents = document.querySelectorAll('.ql-tab-content');
        for (var i=0;i<tabs.length;i++) tabs[i].className = 'ql-tab' + (tabs[i].getAttribute('data-tab')===tab?' active':'');
        for (var j=0;j<contents.length;j++) contents[j].className = 'ql-tab-content' + (contents[j].id==='tab-'+tab?' active':'');
    },

    // --- Modal ---
    openModal: function(id) {
        document.getElementById(id).className = 'ql-modal-overlay active';
    },
    closeModal: function(id) {
        document.getElementById(id).className = 'ql-modal-overlay';
    },

    // --- Load/Save Config ---
    loadConfig: function() {
        QL.ajax('load_config', {}, function(r) {
            if (r.success) {
                var d = r.data || {};
                QL.config = {
                    version: d.version || 1,
                    settings: d.settings || {auto_subgroup_threshold:8, hide_small_categories:0},
                    categories: d.categories || [],
                    icons: d.icons || [],
                    subgroups: d.subgroups || {},
                    count_queries: d.count_queries || {}
                };
                QL.clearDirty();
                QL.renderAll();
                QL.toast('Config loaded', 'success');
            } else {
                QL.toast('Load failed: ' + (r.message||''), 'error');
            }
        });
        QL.ajax('get_roles', {}, function(r) {
            if (r.success) QL.allRoles = r.roles || [];
        });
    },

    saveConfig: function() {
        QL.recomputePriorities();
        var jsonStr = JSON.stringify(QL.config);
        var encoded = btoa(unescape(encodeURIComponent(jsonStr)));
        QL.ajax('save_config', {tpl_json_b64: encoded}, function(r) {
            if (r.success) {
                QL.clearDirty();
                QL.toast('Saved successfully!', 'success');
            } else {
                QL.toast('Save failed: ' + (r.message||''), 'error');
            }
        });
    },

    _treeExpanded: {},

    renderAll: function() {
        QL.renderStructure();
        QL.renderSubgroupList();
        QL.renderQueries();
        QL.renderSettings();
        QL.refreshJsonEditor();
    },

    // --- Helpers ---
    makeId: function(cat, label) {
        return (cat + '_' + label).toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/_+$/,'');
    },
    getCatName: function(catId) {
        var cats = QL.config.categories;
        for (var i=0;i<cats.length;i++) { if (cats[i].id === catId) return cats[i].name; }
        return catId;
    },
    renderRolesBadges: function(roles) {
        if (!roles || roles.length === 0) return '<span class="ql-badge ql-badge-gray">Everyone</span>';
        var h = '';
        for (var i=0;i<roles.length;i++) {
            h += '<span class="ql-badge ql-badge-blue">' + QL.esc(roles[i]) + '</span>';
        }
        return h;
    },
    esc: function(s) {
        if (!s) return '';
        var d = document.createElement('div');
        d.appendChild(document.createTextNode(s));
        return d.innerHTML;
    },

    // ===================== STRUCTURE TREE =====================
    renderStructure: function() {
        var root = document.getElementById('ql-tree-root');
        if (!root) return;

        // Preserve current expand/collapse state from DOM before re-rendering
        if (!QL._skipDomPreserve) {
            var allCollapsed = root.querySelectorAll('.ql-tree-cat-content, .ql-tree-sg-content');
            for (var ei = 0; ei < allCollapsed.length; ei++) {
                var el = allCollapsed[ei];
                var elId = el.id;
                if (elId && elId.indexOf('tree-') === 0) {
                    var key = elId.substring(5);
                    QL._treeExpanded[key] = !el.classList.contains('collapsed');
                }
            }
        }
        QL._skipDomPreserve = false;

        var cats = QL.config.categories || [];
        var icons = QL.config.icons || [];
        var sgs = QL.config.subgroups || {};
        var search = ((document.getElementById('ql-tree-search') || {}).value || '').toLowerCase();

        if (cats.length === 0) {
            root.innerHTML = '<div class="ql-empty"><i class="fa fa-sitemap"></i>No categories yet. Click Add Category to start.</div>';
            return;
        }

        var h = '';
        for (var ci = 0; ci < cats.length; ci++) {
            var cat = cats[ci];
            var catIcons = [];
            for (var ii = 0; ii < icons.length; ii++) {
                if (icons[ii].category_id === cat.id) catIcons.push({idx: ii, icon: icons[ii]});
            }

            // Group by subgroup
            var grouped = {};
            var ungrouped = [];
            var sgSet = {};
            for (var gi = 0; gi < catIcons.length; gi++) {
                var ic = catIcons[gi];
                var sg = ic.icon.subgroup;
                if (sg) {
                    if (!grouped[sg]) { grouped[sg] = []; sgSet[sg] = true; }
                    grouped[sg].push(ic);
                } else {
                    ungrouped.push(ic);
                }
            }
            // Also include parent subgroups that have children in this category
            for (var pk in sgs) {
                if (!sgs.hasOwnProperty(pk)) continue;
                if (sgSet[pk]) continue; // already included
                // Check if any child subgroup with this as parent has icons in this category
                for (var ck in sgSet) {
                    if (sgs[ck] && sgs[ck].parent_subgroup === pk) {
                        sgSet[pk] = true;
                        if (!grouped[pk]) grouped[pk] = [];
                        break;
                    }
                }
            }

            // Sort subgroups by sort_order from subgroup definitions
            var sgOrder = Object.keys(sgSet);
            sgOrder.sort(function(a, b) {
                var sa = (sgs[a] && sgs[a].sort_order != null) ? sgs[a].sort_order : 999;
                var sb = (sgs[b] && sgs[b].sort_order != null) ? sgs[b].sort_order : 999;
                return sa - sb;
            });

            // Filter by search
            var totalVisible = 0;
            if (search) {
                for (var sk in grouped) {
                    grouped[sk] = grouped[sk].filter(function(x) { return x.icon.label.toLowerCase().indexOf(search) !== -1 || (x.icon.link||'').toLowerCase().indexOf(search) !== -1; });
                    totalVisible += grouped[sk].length;
                }
                ungrouped = ungrouped.filter(function(x) { return x.icon.label.toLowerCase().indexOf(search) !== -1 || (x.icon.link||'').toLowerCase().indexOf(search) !== -1; });
                totalVisible += ungrouped.length;
                if (totalVisible === 0 && cat.name.toLowerCase().indexOf(search) === -1) continue;
            } else {
                totalVisible = catIcons.length;
            }

            var catExpKey = 'cat_' + cat.id;
            var catExpanded = QL._treeExpanded[catExpKey] !== undefined ? QL._treeExpanded[catExpKey] : (cat.expanded !== false);

            h += '<div class="ql-tree-cat" data-type="cat" data-idx="'+ci+'">';

            // Category header
            h += '<div class="ql-tree-cat-header" draggable="true" ondragstart="QL.treeDragStart(event,&#39;cat&#39;,'+ci+')" ondragover="QL.treeDragOver(event,&#39;cat&#39;,'+ci+')" ondragleave="QL.treeDragLeave(event)" ondrop="QL.treeDrop(event,&#39;cat&#39;,'+ci+')">';
            h += '<span class="ql-drag" style="cursor:grab;"><i class="fa fa-grip-vertical"></i></span>';
            h += '<i class="fa fa-chevron-down ql-tree-chevron'+(catExpanded?'':' collapsed')+'" onclick="QL.toggleTree(&#39;'+catExpKey+'&#39;,event)"></i>';
            h += '<span class="ql-icon-preview" style="width:28px;height:28px;font-size:14px;"><i class="fa '+QL.esc(cat.icon)+'"></i></span>';
            h += '<span class="ql-tree-cat-title" onclick="QL.toggleTree(&#39;'+catExpKey+'&#39;,event)">'+QL.esc(cat.name)+' <span class="ql-tree-count">'+totalVisible+'</span></span>';
            h += '<span class="ql-tree-cat-actions" draggable="false">';
            h += QL.renderRolesBadges(cat.roles);
            h += ' <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.addIcon(&#39;'+QL.esc(cat.id)+'&#39;)" title="Add Link"><i class="fa fa-plus"></i></button>';
            h += ' <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.addSubgroupForCat(&#39;'+QL.esc(cat.id)+'&#39;)" title="Add Subgroup"><i class="fa fa-layer-group"></i></button>';
            h += ' <button class="ql-btn ql-btn-outline ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.editCategory('+ci+')" title="Edit Category"><i class="fa fa-edit"></i></button>';
            h += ' <button class="ql-btn ql-btn-danger ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.deleteCategory('+ci+')" title="Delete Category"><i class="fa fa-trash"></i></button>';
            h += '</span>';
            h += '</div>';

            // Category content
            h += '<div class="ql-tree-cat-content'+(catExpanded?'':' collapsed')+'" id="tree-'+catExpKey+'">';

            // Build child subgroup map (parent_subgroup -> [child keys])
            var sgChildren = {};
            for (var sci = 0; sci < sgOrder.length; sci++) {
                var scKey = sgOrder[sci];
                var scDef = sgs[scKey];
                if (scDef && scDef.parent_subgroup && sgSet[scDef.parent_subgroup]) {
                    if (!sgChildren[scDef.parent_subgroup]) sgChildren[scDef.parent_subgroup] = [];
                    sgChildren[scDef.parent_subgroup].push(scKey);
                }
            }

            // Render subgroups (skip children - they render inside parents)
            for (var si = 0; si < sgOrder.length; si++) {
                var sgKey = sgOrder[si];
                var sgDef = sgs[sgKey] || {name: sgKey, icon: 'fa-folder', popup: false};
                // Skip child subgroups - they render inside their parent
                if (sgDef.parent_subgroup && sgSet[sgDef.parent_subgroup]) continue;
                var sgIcons = grouped[sgKey] || [];
                // Count child subgroup icons too
                var childKeys = sgChildren[sgKey] || [];
                var childIconCount = 0;
                for (var cci = 0; cci < childKeys.length; cci++) {
                    childIconCount += (grouped[childKeys[cci]] || []).length;
                }
                if (sgIcons.length === 0 && childIconCount === 0) continue;

                h += QL._renderSubgroupBlock(cat.id, sgKey, sgDef, sgIcons, grouped, sgChildren, sgs);
            }

            // Ungrouped
            if (ungrouped.length > 0) {
                if (sgOrder.length > 0) {
                    h += '<div class="ql-tree-ungrouped-label">Ungrouped</div>';
                }
                for (var ui = 0; ui < ungrouped.length; ui++) {
                    h += QL._renderTreeIcon(ungrouped[ui], sgOrder.length === 0);
                }
            }

            h += '</div></div>';
        }

        root.innerHTML = h;
    },

    _renderTreeIcon: function(item, isUngrouped) {
        var ic = item.icon;
        var idx = item.idx;
        var cls = 'ql-tree-icon' + (isUngrouped ? ' ungrouped' : '');
        var h = '<div class="'+cls+'" draggable="true" data-type="icon" data-idx="'+idx+'" ondragstart="QL.treeDragStart(event,&#39;icon&#39;,'+idx+')" ondragover="QL.treeDragOver(event,&#39;icon&#39;,'+idx+')" ondragleave="QL.treeDragLeave(event)" ondrop="QL.treeDrop(event,&#39;icon&#39;,'+idx+')">';
        h += '<span class="ql-drag"><i class="fa fa-grip-vertical"></i></span>';
        h += '<span class="ql-icon-preview"><i class="fa '+QL.esc(ic.icon)+'"></i></span>';
        h += '<span class="ql-tree-icon-label">'+QL.esc(ic.label)+'</span>';
        h += '<span class="ql-tree-icon-link" title="'+QL.esc(ic.link)+'">'+QL.esc(ic.link)+'</span>';
        h += '<span class="ql-tree-icon-meta">';
        h += QL.renderRolesBadges(ic.roles);
        if (ic.query_name) h += '<span class="ql-badge ql-badge-green" title="Count query: '+QL.esc(ic.query_name)+'"><i class="fa fa-database"></i> '+QL.esc(ic.query_name)+'</span>';
        if (ic.org_id) h += '<span class="ql-badge ql-badge-gray">Org:'+ic.org_id+'</span>';
        if (ic.expiration_date) h += '<span class="ql-badge ql-badge-red" title="Expires: '+QL.esc(ic.expiration_date)+'"><i class="fa fa-clock"></i> '+QL.esc(ic.expiration_date)+'</span>';
        h += '</span>';
        h += '<span class="ql-tree-icon-actions" draggable="false">';
        h += '<button class="ql-btn ql-btn-outline ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.editIcon('+idx+')" title="Edit"><i class="fa fa-edit"></i></button>';
        h += ' <button class="ql-btn ql-btn-danger ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.deleteIcon('+idx+')" title="Delete"><i class="fa fa-trash"></i></button>';
        h += '</span>';
        h += '</div>';
        return h;
    },

    _renderSubgroupBlock: function(catId, sgKey, sgDef, sgIcons, grouped, sgChildren, sgs) {
        var sgExpKey = 'sg_' + catId + '_' + sgKey;
        var sgExpanded = QL._treeExpanded[sgExpKey] !== undefined ? QL._treeExpanded[sgExpKey] : true;
        var childKeys = sgChildren[sgKey] || [];
        // Sort child keys by sort_order
        childKeys.sort(function(a, b) {
            var sa = (sgs[a] && sgs[a].sort_order != null) ? sgs[a].sort_order : 999;
            var sb = (sgs[b] && sgs[b].sort_order != null) ? sgs[b].sort_order : 999;
            return sa - sb;
        });
        var totalCount = sgIcons.length;
        for (var ci = 0; ci < childKeys.length; ci++) totalCount += (grouped[childKeys[ci]] || []).length;

        var h = '<div class="ql-tree-sg" data-sg-key="'+QL.esc(sgKey)+'" data-cat-id="'+QL.esc(catId)+'">';
        h += '<div class="ql-tree-sg-header" draggable="true" ondragstart="QL.sgDragStart(event,&#39;'+QL.esc(catId)+'&#39;,&#39;'+QL.esc(sgKey)+'&#39;)" ondragover="QL.sgDragOver(event,&#39;'+QL.esc(catId)+'&#39;,&#39;'+QL.esc(sgKey)+'&#39;)" ondragleave="QL.treeDragLeave(event)" ondrop="QL.sgDrop(event,&#39;'+QL.esc(catId)+'&#39;,&#39;'+QL.esc(sgKey)+'&#39;)">';
        h += '<span class="ql-drag" style="cursor:grab;"><i class="fa fa-grip-vertical"></i></span>';
        h += '<i class="fa fa-chevron-down ql-tree-chevron'+(sgExpanded?'':' collapsed')+'" onclick="QL.toggleTree(&#39;'+sgExpKey+'&#39;,event)"></i>';
        h += '<span class="ql-icon-preview" style="width:22px;height:22px;font-size:11px;background:#fff3e0;color:#e65100;"><i class="fa '+QL.esc(sgDef.icon)+'"></i></span>';
        h += '<span class="ql-tree-sg-title" onclick="QL.toggleTree(&#39;'+sgExpKey+'&#39;,event)">'+QL.esc(sgDef.name)+' <span class="ql-tree-count">'+totalCount+'</span>';
        if (sgDef.popup) h += ' <span class="ql-badge ql-badge-orange" style="font-size:9px;">popup</span>';
        h += '</span>';
        h += '<span class="ql-tree-sg-actions" draggable="false">';
        h += '<button class="ql-btn ql-btn-outline ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.editSubgroup(&#39;'+QL.esc(sgKey)+'&#39;)" title="Edit Subgroup"><i class="fa fa-edit"></i></button>';
        h += ' <button class="ql-btn ql-btn-danger ql-btn-sm" onclick="event.stopPropagation();event.preventDefault();QL.deleteSubgroup(&#39;'+QL.esc(sgKey)+'&#39;)" title="Delete Subgroup"><i class="fa fa-trash"></i></button>';
        h += '</span>';
        h += '</div>';

        h += '<div class="ql-tree-sg-content'+(sgExpanded?'':' collapsed')+'" id="tree-'+sgExpKey+'">';

        // Render this subgroup's own icons
        for (var sii = 0; sii < sgIcons.length; sii++) {
            h += QL._renderTreeIcon(sgIcons[sii], false);
        }

        // Render child subgroups nested inside
        for (var chi = 0; chi < childKeys.length; chi++) {
            var childKey = childKeys[chi];
            var childDef = sgs[childKey] || {name: childKey, icon: 'fa-folder', popup: false};
            var childIcons = grouped[childKey] || [];
            if (childIcons.length === 0) continue;
            h += QL._renderSubgroupBlock(catId, childKey, childDef, childIcons, grouped, sgChildren, sgs);
        }

        h += '</div></div>';
        return h;
    },

    toggleTree: function(key, e) {
        if (e) e.stopPropagation();
        var cur = QL._treeExpanded[key];
        if (cur === undefined) cur = true;
        QL._treeExpanded[key] = !cur;
        var el = document.getElementById('tree-' + key);
        if (el) el.classList.toggle('collapsed');
        // Update chevron
        if (e) {
            var chevron = e.target.closest('.ql-tree-cat-header,.ql-tree-sg-header');
            if (chevron) {
                var ch = chevron.querySelector('.ql-tree-chevron');
                if (ch) ch.classList.toggle('collapsed');
            }
        }
    },

    expandAll: function() {
        QL._skipDomPreserve = true;
        var cats = QL.config.categories || [];
        var sgs = QL.config.subgroups || {};
        for (var i=0;i<cats.length;i++) {
            QL._treeExpanded['cat_'+cats[i].id] = true;
            for (var k in sgs) {
                if (sgs.hasOwnProperty(k)) QL._treeExpanded['sg_'+cats[i].id+'_'+k] = true;
            }
        }
        QL.renderStructure();
    },

    collapseAll: function() {
        QL._skipDomPreserve = true;
        // Reset all to collapsed
        QL._treeExpanded = {};
        var cats = QL.config.categories || [];
        var sgs = QL.config.subgroups || {};
        for (var i=0;i<cats.length;i++) {
            QL._treeExpanded['cat_'+cats[i].id] = false;
            for (var k in sgs) {
                if (sgs.hasOwnProperty(k)) QL._treeExpanded['sg_'+cats[i].id+'_'+k] = false;
            }
        }
        QL.renderStructure();
    },

    addSubgroupForCat: function(catId) {
        QL.addSubgroup();
    },

    recomputePriorities: function() {
        var cats = QL.config.categories || [];
        var icons = QL.config.icons || [];
        var sgs = QL.config.subgroups || {};

        // Update category sort_order from position
        for (var ci=0;ci<cats.length;ci++) {
            cats[ci].sort_order = ci;
        }

        // Recompute icon priorities based on their position in the icons array
        // Icons earlier in the array for a given category get higher priority
        for (var ci2=0;ci2<cats.length;ci2++) {
            var catId = cats[ci2].id;
            var priority = 900;
            for (var ii=0;ii<icons.length;ii++) {
                if (icons[ii].category_id === catId) {
                    icons[ii].priority = priority;
                    priority--;
                }
            }
        }

        // Ensure subgroups have sort_order (preserve existing or assign from current key order)
        var sgKeys = Object.keys(sgs);
        for (var si=0;si<sgKeys.length;si++) {
            if (sgs[sgKeys[si]].sort_order == null) {
                sgs[sgKeys[si]].sort_order = si;
            }
        }
    },

    // Subgroup drag-and-drop
    _sgDragCat: null,
    _sgDragKey: null,

    sgDragStart: function(e, catId, sgKey) {
        QL._dragType = 'subgroup';
        QL._sgDragCat = catId;
        QL._sgDragKey = sgKey;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'sg:' + sgKey);
    },

    sgDragOver: function(e, catId, sgKey) {
        if (QL._dragType === 'subgroup') {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            var el = e.target.closest('.ql-tree-sg-header');
            if (el) el.classList.add('ql-tree-drop-above');
        }
    },

    sgDrop: function(e, catId, targetSgKey) {
        e.preventDefault();
        var els = document.querySelectorAll('.ql-tree-drop-above');
        for (var i=0;i<els.length;i++) els[i].classList.remove('ql-tree-drop-above');

        if (QL._dragType !== 'subgroup') return;
        var srcKey = QL._sgDragKey;
        if (srcKey === targetSgKey) return;

        var sgs = QL.config.subgroups;
        var srcDef = sgs[srcKey];
        var tgtDef = sgs[targetSgKey];
        if (!srcDef || !tgtDef) return;

        // Determine which group of siblings to reorder
        // Siblings = subgroups sharing the same parent_subgroup value
        var srcParent = srcDef.parent_subgroup || null;
        var tgtParent = tgtDef.parent_subgroup || null;

        // Only allow reorder among siblings (same parent)
        if (srcParent !== tgtParent) return;

        // Build ordered list of sibling keys by current sort_order
        var siblingKeys = [];
        for (var k in sgs) {
            if (!sgs.hasOwnProperty(k)) continue;
            var p = sgs[k].parent_subgroup || null;
            if (p === srcParent) siblingKeys.push(k);
        }
        siblingKeys.sort(function(a, b) {
            var sa = (sgs[a].sort_order != null) ? sgs[a].sort_order : 999;
            var sb = (sgs[b].sort_order != null) ? sgs[b].sort_order : 999;
            return sa - sb;
        });

        // Remove srcKey from list and insert before targetSgKey
        var newOrder = [];
        for (var i = 0; i < siblingKeys.length; i++) {
            if (siblingKeys[i] === srcKey) continue;
            if (siblingKeys[i] === targetSgKey) newOrder.push(srcKey);
            newOrder.push(siblingKeys[i]);
        }

        // Reassign sort_order from new positions (for all subgroups, preserving non-siblings)
        // First, collect all keys sorted
        var allKeys = Object.keys(sgs);
        allKeys.sort(function(a, b) {
            var sa = (sgs[a].sort_order != null) ? sgs[a].sort_order : 999;
            var sb = (sgs[b].sort_order != null) ? sgs[b].sort_order : 999;
            return sa - sb;
        });

        // Rebuild full order: replace siblings in their new order, keep others in place
        var finalOrder = [];
        var sibIdx = 0;
        for (var j = 0; j < allKeys.length; j++) {
            var kj = allKeys[j];
            var pj = sgs[kj].parent_subgroup || null;
            if (pj === srcParent) {
                // Replace with reordered sibling
                if (sibIdx < newOrder.length) {
                    finalOrder.push(newOrder[sibIdx]);
                    sibIdx++;
                }
            } else {
                finalOrder.push(kj);
            }
        }

        // Assign sequential sort_order
        for (var fi = 0; fi < finalOrder.length; fi++) {
            sgs[finalOrder[fi]].sort_order = fi;
        }

        QL.markDirty();
        QL.renderStructure();
    },

    addCategory: function() {
        document.getElementById('modal-cat-title').textContent = 'Add Category';
        document.getElementById('cat-edit-idx').value = '-1';
        document.getElementById('cat-id').value = '';
        document.getElementById('cat-name').value = '';
        document.getElementById('cat-icon-val').value = 'fa-church';
        document.getElementById('cat-icon-preview').className = 'fa fa-church';
        document.getElementById('cat-icon-label').textContent = 'fa-church';
        document.getElementById('cat-expanded').value = 'true';
        document.getElementById('cat-sort').value = QL.config.categories.length;
        QL.initRolesSelector('cat-roles-container', []);
        QL.openModal('modal-category');
    },

    editCategory: function(idx) {
        var c = QL.config.categories[idx];
        document.getElementById('modal-cat-title').textContent = 'Edit Category';
        document.getElementById('cat-edit-idx').value = idx;
        document.getElementById('cat-id').value = c.id || '';
        document.getElementById('cat-name').value = c.name || '';
        document.getElementById('cat-icon-val').value = c.icon || 'fa-church';
        document.getElementById('cat-icon-preview').className = 'fa ' + (c.icon||'fa-church');
        document.getElementById('cat-icon-label').textContent = c.icon || 'fa-church';
        document.getElementById('cat-expanded').value = c.expanded ? 'true' : 'false';
        document.getElementById('cat-sort').value = c.sort_order != null ? c.sort_order : idx;
        QL.initRolesSelector('cat-roles-container', c.roles || []);
        QL.openModal('modal-category');
    },

    saveCategory: function() {
        var idx = parseInt(document.getElementById('cat-edit-idx').value);
        var obj = {
            id: document.getElementById('cat-id').value.trim(),
            icon: document.getElementById('cat-icon-val').value,
            name: document.getElementById('cat-name').value.trim(),
            expanded: document.getElementById('cat-expanded').value === 'true',
            roles: QL.getRolesFromSelector('cat-roles-container'),
            sort_order: parseInt(document.getElementById('cat-sort').value) || 0
        };
        if (!obj.id || !obj.name) { QL.toast('ID and Name are required','error'); return; }
        var roles = obj.roles;
        if (roles && roles.length === 0) obj.roles = null;
        if (idx === -1) {
            QL.config.categories.push(obj);
        } else {
            QL.config.categories[idx] = obj;
        }
        QL.markDirty();
        QL.closeModal('modal-category');
        QL.renderStructure();
    },

    deleteCategory: function(idx) {
        if (!confirm('Delete category "' + QL.config.categories[idx].name + '"?')) return;
        QL.config.categories.splice(idx, 1);
        QL.markDirty();
        QL.renderStructure();
    },

    // ===================== ICONS (CRUD) =====================
    addIcon: function(catId) {
        catId = catId || '';
        document.getElementById('modal-icon-title').textContent = 'Add Link';
        document.getElementById('icon-edit-idx').value = '-1';
        document.getElementById('icon-label').value = '';
        document.getElementById('icon-link').value = '';
        document.getElementById('icon-icon-val').value = 'fa-link';
        document.getElementById('icon-icon-preview').className = 'fa fa-link';
        document.getElementById('icon-icon-label').textContent = 'fa-link';
        document.getElementById('icon-priority').value = '50';
        document.getElementById('icon-orgid').value = '';
        document.getElementById('icon-expiration').value = '';
        QL._populateIconModalDropdowns(catId, '', '');
        QL.initRolesSelector('icon-roles-container', []);
        QL.openModal('modal-icon');
    },

    editIcon: function(idx) {
        var ic = QL.config.icons[idx];
        document.getElementById('modal-icon-title').textContent = 'Edit Icon';
        document.getElementById('icon-edit-idx').value = idx;
        document.getElementById('icon-label').value = ic.label || '';
        document.getElementById('icon-link').value = ic.link || '';
        document.getElementById('icon-icon-val').value = ic.icon || 'fa-link';
        document.getElementById('icon-icon-preview').className = 'fa ' + (ic.icon||'fa-link');
        document.getElementById('icon-icon-label').textContent = ic.icon || 'fa-link';
        document.getElementById('icon-priority').value = ic.priority || 50;
        document.getElementById('icon-orgid').value = ic.org_id || '';
        document.getElementById('icon-expiration').value = ic.expiration_date || '';
        QL._populateIconModalDropdowns(ic.category_id || '', ic.query_name || '', ic.subgroup || '');
        QL.initRolesSelector('icon-roles-container', ic.roles || []);
        QL.openModal('modal-icon');
    },

    _populateIconModalDropdowns: function(catVal, queryVal, sgVal) {
        var catSel = document.getElementById('icon-category');
        var h = '';
        var cats = QL.config.categories || [];
        for (var i=0;i<cats.length;i++) {
            h += '<option value="'+QL.esc(cats[i].id)+'"'+(catVal===cats[i].id?' selected':'')+'>'+QL.esc(cats[i].name)+'</option>';
        }
        catSel.innerHTML = h;

        var qSel = document.getElementById('icon-query');
        var qh = '<option value="">None</option>';
        var cq = QL.config.count_queries || {};
        for (var k in cq) {
            if (cq.hasOwnProperty(k)) {
                qh += '<option value="'+QL.esc(k)+'"'+(queryVal===k?' selected':'')+'>'+QL.esc(k)+'</option>';
            }
        }
        qSel.innerHTML = qh;

        var sgSel = document.getElementById('icon-subgroup');
        var sgh = '<option value="">None</option>';
        var sgs = QL.config.subgroups || {};
        for (var s in sgs) {
            if (sgs.hasOwnProperty(s)) {
                sgh += '<option value="'+QL.esc(s)+'"'+(sgVal===s?' selected':'')+'>'+QL.esc(s)+' ('+QL.esc(sgs[s].name)+')</option>';
            }
        }
        sgSel.innerHTML = sgh;
    },

    saveIcon: function() {
        var idx = parseInt(document.getElementById('icon-edit-idx').value);
        var label = document.getElementById('icon-label').value.trim();
        var catId = document.getElementById('icon-category').value;
        var orgVal = document.getElementById('icon-orgid').value.trim();
        var obj = {
            id: QL.makeId(catId, label),
            category_id: catId,
            icon: document.getElementById('icon-icon-val').value,
            label: label,
            link: document.getElementById('icon-link').value.trim(),
            org_id: orgVal ? parseInt(orgVal) : null,
            query_name: document.getElementById('icon-query').value || null,
            roles: QL.getRolesFromSelector('icon-roles-container'),
            priority: parseInt(document.getElementById('icon-priority').value) || 50,
            subgroup: document.getElementById('icon-subgroup').value || null,
            expiration_date: document.getElementById('icon-expiration').value || null
        };
        if (!obj.label || !obj.link) { QL.toast('Label and Link are required','error'); return; }
        if (obj.roles && obj.roles.length === 0) obj.roles = null;
        if (idx === -1) {
            QL.config.icons.push(obj);
        } else {
            QL.config.icons[idx] = obj;
        }
        QL.markDirty();
        QL.closeModal('modal-icon');
        QL.renderStructure();
    },

    deleteIcon: function(idx) {
        if (!confirm('Delete link "' + QL.config.icons[idx].label + '"?')) return;
        QL.config.icons.splice(idx, 1);
        QL.markDirty();
        QL.renderStructure();
    },

    // ===================== SUBGROUPS (CRUD) =====================
    renderSubgroupList: function() {
        var area = document.getElementById('ql-sg-manage-area');
        if (!area) return;
        var sgs = QL.config.subgroups || {};
        var keys = Object.keys(sgs);
        // Sort by sort_order
        keys.sort(function(a, b) {
            var sa = (sgs[a].sort_order != null) ? sgs[a].sort_order : 999;
            var sb = (sgs[b].sort_order != null) ? sgs[b].sort_order : 999;
            return sa - sb;
        });
        if (keys.length === 0) {
            area.innerHTML = '<div style="padding:16px;text-align:center;color:var(--ql-text-muted);font-size:13px;">No subgroups defined yet.</div>';
            return;
        }
        // Count icons per subgroup (direct)
        var iconCounts = {};
        var icons = QL.config.icons || [];
        for (var i = 0; i < icons.length; i++) {
            var sg = icons[i].subgroup;
            if (sg) iconCounts[sg] = (iconCounts[sg] || 0) + 1;
        }
        // Find which categories use each subgroup (direct)
        var catUsage = {};
        for (var j = 0; j < icons.length; j++) {
            var sg2 = icons[j].subgroup;
            var catId = icons[j].category_id;
            if (sg2 && catId) {
                if (!catUsage[sg2]) catUsage[sg2] = {};
                catUsage[sg2][catId] = true;
            }
        }
        // Roll up child subgroup counts and usage to parents
        for (var ck in sgs) {
            if (!sgs.hasOwnProperty(ck)) continue;
            var parentKey = sgs[ck].parent_subgroup;
            if (parentKey && sgs[parentKey]) {
                // Add child icon count to parent
                iconCounts[parentKey] = (iconCounts[parentKey] || 0) + (iconCounts[ck] || 0);
                // Merge child category usage into parent
                if (catUsage[ck]) {
                    if (!catUsage[parentKey]) catUsage[parentKey] = {};
                    for (var cu in catUsage[ck]) {
                        if (catUsage[ck].hasOwnProperty(cu)) catUsage[parentKey][cu] = true;
                    }
                }
            }
        }

        var h = '<table class="ql-table" style="font-size:12px;"><thead><tr>';
        h += '<th style="width:40px;">Icon</th><th>ID</th><th>Name</th><th>Type</th><th>Parent</th><th>Used In</th><th>Icons</th><th style="width:90px;">Actions</th>';
        h += '</tr></thead><tbody>';
        for (var ki = 0; ki < keys.length; ki++) {
            var k = keys[ki];
            var sg = sgs[k];
            var cats = catUsage[k] ? Object.keys(catUsage[k]) : [];
            var catNames = [];
            for (var cn = 0; cn < cats.length; cn++) catNames.push(QL.getCatName(cats[cn]));

            h += '<tr>';
            h += '<td><span class="ql-icon-preview" style="width:24px;height:24px;font-size:12px;"><i class="fa '+QL.esc(sg.icon)+'"></i></span></td>';
            h += '<td><code style="font-size:11px;">'+QL.esc(k)+'</code></td>';
            h += '<td>'+QL.esc(sg.name)+'</td>';
            h += '<td>'+(sg.popup?'<span class="ql-badge ql-badge-orange">popup</span>':'<span class="ql-badge ql-badge-gray">inline</span>')+'</td>';
            h += '<td>'+(sg.parent_subgroup?'<span class="ql-badge ql-badge-blue">'+QL.esc(sg.parent_subgroup)+'</span>':'-')+'</td>';
            h += '<td>'+(catNames.length>0?catNames.map(function(n){return '<span class="ql-badge ql-badge-blue">'+QL.esc(n)+'</span>';}).join(' '):'<span style="color:#999;">unused</span>')+'</td>';
            h += '<td>'+(iconCounts[k]||0)+'</td>';
            h += '<td>';
            h += '<button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.editSubgroup(&#39;'+QL.esc(k)+'&#39;)"><i class="fa fa-edit"></i></button> ';
            h += '<button class="ql-btn ql-btn-danger ql-btn-sm" onclick="QL.deleteSubgroup(&#39;'+QL.esc(k)+'&#39;)"><i class="fa fa-trash"></i></button>';
            h += '</td></tr>';
        }
        h += '</tbody></table>';
        area.innerHTML = h;
    },

    _populateSgParentDropdown: function(currentKey, currentParent) {
        var sel = document.getElementById('sg-parent');
        var h = '<option value="">None (top-level)</option>';
        var sgs = QL.config.subgroups || {};
        for (var k in sgs) {
            if (!sgs.hasOwnProperty(k)) continue;
            if (k === currentKey) continue; // Can't parent to self
            if (sgs[k].parent_subgroup === currentKey) continue; // Can't parent to own child
            var selected = (currentParent === k) ? ' selected' : '';
            h += '<option value="'+QL.esc(k)+'"'+selected+'>'+QL.esc(sgs[k].name || k)+'</option>';
        }
        sel.innerHTML = h;
    },

    addSubgroup: function() {
        document.getElementById('modal-sg-title').textContent = 'Add Subgroup';
        document.getElementById('sg-edit-key').value = '';
        document.getElementById('sg-id').value = '';
        document.getElementById('sg-name').value = '';
        document.getElementById('sg-icon-val').value = 'fa-folder';
        document.getElementById('sg-icon-preview').className = 'fa fa-folder';
        document.getElementById('sg-icon-label').textContent = 'fa-folder';
        document.getElementById('sg-popup').value = 'false';
        QL._populateSgParentDropdown('', '');
        QL.openModal('modal-subgroup');
    },

    editSubgroup: function(key) {
        var sg = QL.config.subgroups[key];
        document.getElementById('modal-sg-title').textContent = 'Edit Subgroup';
        document.getElementById('sg-edit-key').value = key;
        document.getElementById('sg-id').value = key;
        document.getElementById('sg-name').value = sg.name || '';
        document.getElementById('sg-icon-val').value = sg.icon || 'fa-folder';
        document.getElementById('sg-icon-preview').className = 'fa ' + (sg.icon||'fa-folder');
        document.getElementById('sg-icon-label').textContent = sg.icon || 'fa-folder';
        document.getElementById('sg-popup').value = sg.popup ? 'true' : 'false';
        QL._populateSgParentDropdown(key, sg.parent_subgroup || '');
        QL.openModal('modal-subgroup');
    },

    saveSubgroup: function() {
        var oldKey = document.getElementById('sg-edit-key').value;
        var newKey = document.getElementById('sg-id').value.trim();
        var name = document.getElementById('sg-name').value.trim();
        if (!newKey || !name) { QL.toast('ID and Name required','error'); return; }
        // Preserve existing sort_order if editing
        var existingOrder = (oldKey && QL.config.subgroups[oldKey] && QL.config.subgroups[oldKey].sort_order != null)
            ? QL.config.subgroups[oldKey].sort_order
            : Object.keys(QL.config.subgroups).length;
        var parentVal = document.getElementById('sg-parent').value || null;
        var obj = {
            name: name,
            icon: document.getElementById('sg-icon-val').value,
            popup: document.getElementById('sg-popup').value === 'true',
            sort_order: existingOrder,
            parent_subgroup: parentVal
        };
        if (oldKey && oldKey !== newKey) {
            delete QL.config.subgroups[oldKey];
        }
        QL.config.subgroups[newKey] = obj;
        QL.markDirty();
        QL.closeModal('modal-subgroup');
        QL.renderStructure();
        QL.renderSubgroupList();
    },

    deleteSubgroup: function(key) {
        if (!confirm('Delete subgroup "' + key + '"?')) return;
        delete QL.config.subgroups[key];
        QL.markDirty();
        QL.renderStructure();
        QL.renderSubgroupList();
    },

    // ===================== QUERIES =====================
    renderQueries: function() {
        var tbody = document.getElementById('ql-q-tbody');
        var cq = QL.config.count_queries || {};
        var keys = Object.keys(cq);
        if (keys.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4"><div class="ql-empty"><i class="fa fa-database"></i>No count queries.</div></td></tr>';
            return;
        }
        var h = '';
        for (var i=0;i<keys.length;i++) {
            var k = keys[i];
            var qq = cq[k];
            h += '<tr>';
            h += '<td><code>'+QL.esc(k)+'</code></td>';
            h += '<td>'+QL.esc(qq.description||'')+'</td>';
            h += '<td>'+(qq.uses_current_user?'<span class="ql-badge ql-badge-orange">Yes</span>':'No')+'</td>';
            h += '<td>';
            h += '<button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.editQuery(&#39;'+QL.esc(k)+'&#39;)"><i class="fa fa-edit"></i></button> ';
            h += '<button class="ql-btn ql-btn-sm" style="background:#00BCD4;color:white;" onclick="QL.testQuery(&#39;'+QL.esc(k)+'&#39;)"><i class="fa fa-play"></i> Test</button> ';
            h += '<button class="ql-btn ql-btn-danger ql-btn-sm" onclick="QL.deleteQuery(&#39;'+QL.esc(k)+'&#39;)"><i class="fa fa-trash"></i></button>';
            h += '</td></tr>';
        }
        tbody.innerHTML = h;
    },

    addQuery: function() {
        document.getElementById('modal-q-title').textContent = 'Add Count Query';
        document.getElementById('q-edit-key').value = '';
        document.getElementById('q-name').value = '';
        document.getElementById('q-desc').value = '';
        document.getElementById('q-sql').value = '';
        document.getElementById('q-uses-user').checked = false;
        QL.openModal('modal-query');
    },

    editQuery: function(key) {
        var qq = QL.config.count_queries[key];
        document.getElementById('modal-q-title').textContent = 'Edit Count Query';
        document.getElementById('q-edit-key').value = key;
        document.getElementById('q-name').value = key;
        document.getElementById('q-desc').value = qq.description || '';
        document.getElementById('q-sql').value = qq.sql || '';
        document.getElementById('q-uses-user').checked = !!qq.uses_current_user;
        QL.openModal('modal-query');
    },

    saveQuery: function() {
        var oldKey = document.getElementById('q-edit-key').value;
        var newKey = document.getElementById('q-name').value.trim();
        if (!newKey) { QL.toast('Query name required','error'); return; }
        var obj = {
            sql: document.getElementById('q-sql').value,
            description: document.getElementById('q-desc').value.trim(),
            uses_current_user: document.getElementById('q-uses-user').checked
        };
        if (oldKey && oldKey !== newKey) {
            delete QL.config.count_queries[oldKey];
        }
        QL.config.count_queries[newKey] = obj;
        QL.markDirty();
        QL.closeModal('modal-query');
        QL.renderQueries();
    },

    deleteQuery: function(key) {
        if (!confirm('Delete query "' + key + '"?')) return;
        delete QL.config.count_queries[key];
        QL.markDirty();
        QL.renderQueries();
    },

    testQuery: function(key) {
        var qq = QL.config.count_queries[key];
        if (!qq || !qq.sql) { QL.toast('No SQL for this query','error'); return; }
        var encoded = btoa(unescape(encodeURIComponent(qq.sql)));
        QL.ajax('test_query', {tpl_sql_b64: encoded, tpl_org: '0'}, function(r) {
            if (r.success) {
                QL.toast('Query "' + key + '" returned: ' + r.count, 'success');
            } else {
                QL.toast('Test failed: ' + (r.message||''), 'error');
            }
        });
    },

    // ===================== SETTINGS =====================
    renderSettings: function() {
        var s = QL.config.settings || {};
        document.getElementById('ql-setting-threshold').value = s.auto_subgroup_threshold || 8;
        document.getElementById('ql-setting-hide-small').value = s.hide_small_categories || 0;
    },

    updateSetting: function(key, val) {
        if (!QL.config.settings) QL.config.settings = {};
        QL.config.settings[key] = val;
        QL.markDirty();
    },

    // ===================== JSON EDITOR =====================
    refreshJsonEditor: function() {
        try {
            document.getElementById('ql-json-editor').value = JSON.stringify(QL.config, null, 2);
        } catch(e) {
            document.getElementById('ql-json-editor').value = 'Error: ' + e.message;
        }
    },

    applyJson: function() {
        try {
            var parsed = JSON.parse(document.getElementById('ql-json-editor').value);
            QL.config = parsed;
            QL.markDirty();
            QL.renderAll();
            QL.toast('JSON applied', 'success');
        } catch(e) {
            QL.toast('Invalid JSON: ' + e.message, 'error');
        }
    },

    // ===================== IMPORT / EXPORT =====================
    importHardcoded: function() {
        if (!confirm('This will REPLACE your current config with the hardcoded defaults. Continue?')) return;
        QL.config = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
        QL.markDirty();
        QL.renderAll();
        QL.toast('Imported hardcoded defaults. Click Save to persist.', 'success');
    },

    exportJson: function() {
        var blob = new Blob([JSON.stringify(QL.config, null, 2)], {type:'application/json'});
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'QuickLinksConfig_' + new Date().toISOString().slice(0,10) + '.json';
        a.click();
        URL.revokeObjectURL(url);
    },

    // ===================== BACKUPS =====================
    loadBackups: function() {
        QL.ajax('list_backups', {}, function(r) {
            var div = document.getElementById('ql-backup-list');
            if (!r.success) { div.innerHTML = '<p style="color:red;">'+QL.esc(r.message)+'</p>'; return; }
            var bk = r.backups || [];
            if (bk.length === 0) { div.innerHTML = '<p style="color:#999;">No backups found.</p>'; return; }
            var h = '<table class="ql-table"><thead><tr><th>Timestamp</th><th>User</th><th>Actions</th></tr></thead><tbody>';
            for (var i=0;i<bk.length;i++) {
                var b = bk[i];
                h += '<tr><td>'+QL.esc(b.timestamp)+'</td><td>'+QL.esc(b.user)+'</td><td>';
                h += '<button class="ql-btn ql-btn-outline ql-btn-sm" onclick="QL.previewBackup(&#39;'+QL.esc(b.backup_name)+'&#39;)"><i class="fa fa-eye"></i> Preview</button> ';
                h += '<button class="ql-btn ql-btn-sm" style="background:var(--ql-warning);color:white;" onclick="QL.restoreBackup(&#39;'+QL.esc(b.backup_name)+'&#39;)"><i class="fa fa-undo"></i> Restore</button>';
                h += '</td></tr>';
            }
            h += '</tbody></table>';
            div.innerHTML = h;
        });
    },

    previewBackup: function(name) {
        QL.ajax('preview_backup', {tpl_backup_name: name}, function(r) {
            if (r.success) {
                try {
                    var pretty = JSON.stringify(JSON.parse(r.content), null, 2);
                    document.getElementById('backup-preview-content').value = pretty;
                } catch(e) {
                    document.getElementById('backup-preview-content').value = r.content;
                }
                QL.openModal('modal-backup-preview');
            } else {
                QL.toast('Preview failed: '+(r.message||''), 'error');
            }
        });
    },

    restoreBackup: function(name) {
        if (!confirm('Restore this backup? A safety backup will be created first.')) return;
        QL.ajax('restore_backup', {tpl_backup_name: name}, function(r) {
            if (r.success) {
                QL.toast('Restored! Reloading...', 'success');
                setTimeout(function(){ QL.loadConfig(); }, 500);
            } else {
                QL.toast('Restore failed: '+(r.message||''), 'error');
            }
        });
    },

    // ===================== TREE DRAG AND DROP =====================
    _dragType: null,
    _dragIdx: null,

    treeDragStart: function(e, type, idx) {
        QL._dragType = type;
        QL._dragIdx = idx;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', type + ':' + idx);
        var el = e.target.closest('.ql-tree-cat,.ql-tree-icon');
        if (el) el.classList.add('ql-tree-dragging');
    },

    treeDragOver: function(e, type, idx) {
        if (QL._dragType === type || (QL._dragType === 'icon' && type === 'sg')) {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            var el = e.target.closest('.ql-tree-cat-header,.ql-tree-icon,.ql-tree-sg-header,.ql-tree-ungrouped-label');
            if (el) el.classList.add('ql-tree-drop-above');
        }
    },

    treeDragLeave: function(e) {
        var el = e.target.closest('.ql-tree-cat-header,.ql-tree-icon,.ql-tree-sg-header,.ql-tree-ungrouped-label');
        if (el) el.classList.remove('ql-tree-drop-above');
    },

    treeDrop: function(e, type, targetIdx) {
        e.preventDefault();
        // Clean up highlights
        var els = document.querySelectorAll('.ql-tree-drop-above');
        for (var i=0;i<els.length;i++) els[i].classList.remove('ql-tree-drop-above');

        if (QL._dragType !== type) return;
        var srcIdx = QL._dragIdx;
        if (srcIdx === targetIdx) return;

        var arr = type === 'cat' ? QL.config.categories : QL.config.icons;
        var item = arr.splice(srcIdx, 1)[0];
        // Adjust target index if source was before target
        if (srcIdx < targetIdx) targetIdx--;
        arr.splice(targetIdx, 0, item);

        QL.recomputePriorities();
        QL.markDirty();
        QL.renderStructure();
    },

    // ===================== ICON PICKER =====================
    openIconPicker: function(targetInputId) {
        QL._iconPickerTarget = targetInputId;
        QL._iconPickerSelected = document.getElementById(targetInputId).value;
        QL._iconPickerCat = 'All';
        document.getElementById('ip-search').value = '';
        document.getElementById('ip-custom').value = '';
        QL.renderIconPickerCats();
        QL.filterIconPicker();
        QL.openModal('modal-iconpicker');
    },

    renderIconPickerCats: function() {
        var catDiv = document.getElementById('ip-cats');
        var cats = ['All'].concat(Object.keys(FA_ICONS));
        var h = '';
        for (var i=0;i<cats.length;i++) {
            h += '<span class="ql-icon-picker-cat'+(QL._iconPickerCat===cats[i]?' active':'')+'" onclick="QL._iconPickerCat=&#39;'+cats[i]+'&#39;;QL.renderIconPickerCats();QL.filterIconPicker();">'+cats[i]+'</span>';
        }
        catDiv.innerHTML = h;
    },

    filterIconPicker: function() {
        var search = (document.getElementById('ip-search').value || '').toLowerCase();
        var grid = document.getElementById('ip-grid');
        var h = '';
        var cats = QL._iconPickerCat === 'All' ? Object.keys(FA_ICONS) : [QL._iconPickerCat];
        for (var ci=0;ci<cats.length;ci++) {
            var icons = FA_ICONS[cats[ci]] || [];
            for (var j=0;j<icons.length;j++) {
                var ic = icons[j];
                if (search && ic.l.toLowerCase().indexOf(search)===-1 && ic.i.toLowerCase().indexOf(search)===-1) continue;
                var sel = (QL._iconPickerSelected === ic.i) ? ' selected' : '';
                h += '<div class="ql-icon-picker-tile'+sel+'" onclick="QL._pickIcon(&#39;'+ic.i+'&#39;)" title="'+ic.i+'">';
                h += '<i class="fa '+ic.i+'"></i>';
                h += '<span>'+ic.l+'</span></div>';
            }
        }
        if (!h) h = '<div style="padding:20px;color:#999;grid-column:1/-1;text-align:center;">No icons match your search</div>';
        grid.innerHTML = h;
    },

    _pickIcon: function(icon) {
        QL._iconPickerSelected = icon;
        document.getElementById('ip-custom').value = icon;
        QL.filterIconPicker();
    },

    selectIconFromPicker: function() {
        var custom = document.getElementById('ip-custom').value.trim();
        var icon = custom || QL._iconPickerSelected || 'fa-link';
        var target = QL._iconPickerTarget;
        if (target) {
            document.getElementById(target).value = icon;
            // update preview
            var previewId = target.replace('-val','') + '-preview';
            var labelId = target.replace('-val','') + '-label';
            var prev = document.getElementById(previewId);
            var lab = document.getElementById(labelId);
            if (prev) prev.className = 'fa ' + icon;
            if (lab) lab.textContent = icon;
        }
        QL.closeModal('modal-iconpicker');
    },

    // ===================== ROLES SELECTOR =====================
    initRolesSelector: function(containerId, currentRoles) {
        var container = document.getElementById(containerId);
        currentRoles = currentRoles || [];
        var h = '<div class="ql-roles-tags" id="'+containerId+'-tags" onclick="document.getElementById(&#39;'+containerId+'-input&#39;).focus();">';
        for (var i=0;i<currentRoles.length;i++) {
            h += '<span class="ql-role-tag" data-role="'+QL.esc(currentRoles[i])+'">'+QL.esc(currentRoles[i]);
            h += '<span class="ql-role-tag-remove" onclick="QL.removeRole(&#39;'+containerId+'&#39;,&#39;'+QL.esc(currentRoles[i])+'&#39;)">&times;</span></span>';
        }
        h += '<input type="text" class="ql-roles-input" id="'+containerId+'-input" placeholder="Type role..." onfocus="QL.showRolesDropdown(&#39;'+containerId+'&#39;)" oninput="QL.filterRolesDropdown(&#39;'+containerId+'&#39;)">';
        h += '</div>';
        h += '<div class="ql-roles-dropdown" id="'+containerId+'-dropdown"></div>';
        container.innerHTML = h;
        // Close dropdown on outside click
        document.addEventListener('click', function handler(e) {
            if (!container.contains(e.target)) {
                var dd = document.getElementById(containerId+'-dropdown');
                if (dd) dd.className = 'ql-roles-dropdown';
            }
        });
    },

    showRolesDropdown: function(containerId) {
        QL.filterRolesDropdown(containerId);
        var dd = document.getElementById(containerId+'-dropdown');
        if (dd) dd.className = 'ql-roles-dropdown show';
    },

    filterRolesDropdown: function(containerId) {
        var input = document.getElementById(containerId+'-input');
        var dd = document.getElementById(containerId+'-dropdown');
        var filter = (input.value||'').toLowerCase();
        var current = QL.getRolesFromSelector(containerId);
        var h = '';
        var roles = QL.allRoles.length > 0 ? QL.allRoles : ['Access','Edit','Admin','SuperAdmin','Developer','Finance','FinanceAdmin','ManageTransactions','PastoralCare','IT-Team','Security-Doors','Beta','ManageGroups','ManageOrgMembers','BackgroundCheck','Staff'];
        for (var i=0;i<roles.length;i++) {
            if (!roles[i]) continue;
            if (current.indexOf(roles[i]) !== -1) continue;
            if (filter && roles[i].toLowerCase().indexOf(filter) === -1) continue;
            h += '<div class="ql-roles-option" onclick="QL.addRoleToSelector(&#39;'+containerId+'&#39;,&#39;'+QL.esc(roles[i])+'&#39;)">'+QL.esc(roles[i])+'</div>';
        }
        if (!h) h = '<div class="ql-roles-option" style="color:#999;">No matching roles</div>';
        dd.innerHTML = h;
        dd.className = 'ql-roles-dropdown show';
    },

    addRoleToSelector: function(containerId, role) {
        var tags = document.getElementById(containerId+'-tags');
        var input = document.getElementById(containerId+'-input');
        var span = document.createElement('span');
        span.className = 'ql-role-tag';
        span.setAttribute('data-role', role);
        span.innerHTML = QL.esc(role) + '<span class="ql-role-tag-remove" onclick="QL.removeRole(&#39;'+containerId+'&#39;,&#39;'+QL.esc(role)+'&#39;)">&times;</span>';
        tags.insertBefore(span, input);
        input.value = '';
        var dd = document.getElementById(containerId+'-dropdown');
        if (dd) dd.className = 'ql-roles-dropdown';
    },

    removeRole: function(containerId, role) {
        var tags = document.getElementById(containerId+'-tags');
        var spans = tags.querySelectorAll('.ql-role-tag');
        for (var i=0;i<spans.length;i++) {
            if (spans[i].getAttribute('data-role') === role) {
                spans[i].parentNode.removeChild(spans[i]);
                break;
            }
        }
    },

    getRolesFromSelector: function(containerId) {
        var tags = document.getElementById(containerId+'-tags');
        if (!tags) return [];
        var spans = tags.querySelectorAll('.ql-role-tag');
        var roles = [];
        for (var i=0;i<spans.length;i++) {
            var r = spans[i].getAttribute('data-role');
            if (r) roles.push(r);
        }
        return roles;
    },

    // ===================== GUIDED TOUR =====================
    _tourStep: 0,
    _tourSteps: [
        {
            target: '.ql-header',
            title: 'Welcome to QuickLinks Admin!',
            text: 'This tool lets you manage the QuickLinks widget that appears on your TouchPoint homepage. Use the Tour button anytime to see this guide again.'
        },
        {
            target: '#tab-structure',
            title: 'Structure Tab',
            text: 'This is your main workspace. It shows all your categories, subgroups, and links in a tree view. Everything is organized hierarchically.'
        },
        {
            target: '.ql-tree-cat:first-child .ql-tree-cat-header',
            title: 'Categories',
            text: 'Categories are the top-level groups in the widget (e.g. General, Finance). Click the chevron to expand/collapse. Drag the grip handle to reorder.'
        },
        {
            target: '.ql-tree-cat:first-child .ql-tree-cat-actions',
            title: 'Category Actions',
            text: '<b>+</b> adds a new link, <b>layers icon</b> adds a subgroup, <b>pencil</b> edits, <b>trash</b> deletes. Role badges show who can see the category.'
        },
        {
            target: '.ql-tree-icon:first-child',
            title: 'Links / Icons',
            text: 'Each link shows its icon, label, URL, role badges, and metadata. Drag the grip handle to reorder within its group. Click pencil to edit.'
        },
        {
            target: '.ql-tree-sg:first-child .ql-tree-sg-header',
            title: 'Subgroups',
            text: 'Subgroups organize links within a category. They can be inline (always visible) or popup (clickable tile that opens a menu). Manage all subgroups in the Settings tab.',
            fallback: '.ql-tabs'
        },
        {
            target: '[data-tab="queries"]',
            title: 'Count Queries',
            text: 'Define SQL queries that show count badges on icons (e.g. open tasks, member counts). Queries can use {0} for org ID or {} for the current user.'
        },
        {
            target: '[data-tab="settings"]',
            title: 'Settings & Subgroups',
            text: 'Manage all subgroup definitions, configure thresholds, import/export config, edit raw JSON, and manage backups.'
        },
        {
            target: '.ql-header-actions',
            title: 'Save Your Changes!',
            text: 'Changes are not saved automatically. Click <b>Save</b> to persist your config. The orange dot appears when you have unsaved changes. Click <b>Reload</b> to discard changes.'
        }
    ],

    startTour: function() {
        QL._tourStep = 0;
        QL._showTourStep();
        // Mark as seen
        try {
            localStorage.setItem('ql_tour_seen_' + (window._qlUserId || 'default'), '1');
        } catch(e) {}
    },

    endTour: function() {
        document.getElementById('ql-tour-overlay').className = 'ql-tour-overlay';
    },

    tourStep: function(delta) {
        QL._tourStep += delta;
        if (QL._tourStep < 0) QL._tourStep = 0;
        if (QL._tourStep >= QL._tourSteps.length) {
            QL.endTour();
            return;
        }
        QL._showTourStep();
    },

    _showTourStep: function() {
        var step = QL._tourSteps[QL._tourStep];
        var overlay = document.getElementById('ql-tour-overlay');
        var spotlight = document.getElementById('ql-tour-spotlight');
        var card = document.getElementById('ql-tour-card');

        // Find target element
        var target = document.querySelector(step.target);
        if (!target && step.fallback) target = document.querySelector(step.fallback);

        // Update card content
        document.getElementById('ql-tour-step').textContent = 'Step ' + (QL._tourStep + 1) + ' of ' + QL._tourSteps.length;
        document.getElementById('ql-tour-title').textContent = step.title;
        document.getElementById('ql-tour-text').innerHTML = step.text;

        // Show/hide prev button
        document.getElementById('ql-tour-prev').style.display = QL._tourStep === 0 ? 'none' : '';
        var nextBtn = document.getElementById('ql-tour-next');
        nextBtn.innerHTML = QL._tourStep === QL._tourSteps.length - 1 ? 'Done <i class="fa fa-check"></i>' : 'Next <i class="fa fa-arrow-right"></i>';

        // Position spotlight and card
        if (target) {
            target.scrollIntoView({behavior: 'smooth', block: 'center'});
            setTimeout(function() {
                var rect = target.getBoundingClientRect();
                var pad = 6;
                spotlight.style.left = (rect.left - pad + window.scrollX) + 'px';
                spotlight.style.top = (rect.top - pad + window.scrollY) + 'px';
                spotlight.style.width = (rect.width + pad * 2) + 'px';
                spotlight.style.height = (rect.height + pad * 2) + 'px';
                spotlight.style.display = 'block';

                // Position card below or above target
                var cardTop = rect.bottom + 12 + window.scrollY;
                var cardLeft = Math.max(16, Math.min(rect.left + window.scrollX, window.innerWidth - 400));
                if (cardTop + 200 > window.innerHeight + window.scrollY) {
                    cardTop = rect.top - 200 + window.scrollY;
                }
                card.style.top = cardTop + 'px';
                card.style.left = cardLeft + 'px';
            }, 150);
        } else {
            spotlight.style.display = 'none';
            card.style.top = '50%';
            card.style.left = '50%';
            card.style.transform = 'translate(-50%, -50%)';
        }

        overlay.className = 'ql-tour-overlay active';
    },

    checkFirstVisit: function() {
        try {
            var key = 'ql_tour_seen_' + (window._qlUserId || 'default');
            if (!localStorage.getItem(key)) {
                setTimeout(function() { QL.startTour(); }, 800);
            }
        } catch(e) {}
    }
};

// --- Unsaved changes warning ---
window.addEventListener('beforeunload', function(e) {
    if (QL.dirty) {
        e.preventDefault();
        e.returnValue = 'You have unsaved changes.';
    }
});

// --- Initialize ---
window.QL = QL;
window._qlUserId = '##USERID##';
QL.loadConfig();
QL.checkFirstVisit();

// Help tooltip positioning
document.addEventListener('mouseenter', function(e) {
    var help = e.target.closest('.ql-help');
    if (!help) return;
    var tip = help.querySelector('.ql-help-tip');
    if (!tip) return;
    var rect = help.getBoundingClientRect();
    tip.style.display = 'block';
    // Position below the ? badge
    var tipTop = rect.bottom + 6;
    var tipLeft = rect.left + rect.width / 2 - 120;
    // Keep within viewport
    if (tipLeft < 8) tipLeft = 8;
    if (tipLeft + 240 > window.innerWidth - 8) tipLeft = window.innerWidth - 248;
    // If below would go off screen, show above
    if (tipTop + 100 > window.innerHeight) tipTop = rect.top - 6;
    tip.style.top = tipTop + 'px';
    tip.style.left = tipLeft + 'px';
}, true);

document.addEventListener('mouseleave', function(e) {
    var help = e.target.closest('.ql-help');
    if (!help) return;
    var tip = help.querySelector('.ql-help-tip');
    if (tip) tip.style.display = 'none';
}, true);

})();
</script>
</body>
</html>
"""
        _uid = str(model.UserPeopleId) if model.UserPeopleId else '0'
        html = html.replace('##USERID##', _uid)
        print html
        model.Form = html
