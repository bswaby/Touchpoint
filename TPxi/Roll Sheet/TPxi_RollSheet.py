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
#
# Architecture:
#   Single .py file SPA - Python AJAX handlers for POST, HTML SPA for GET.
#   Data stored via model.WriteContentText / model.TextContent.
#
# Storage Keys:
#   RollSheet_Configs - All rollsheet configs (JSON)
#
# CSS Prefix: rs-
# Root Class: .rs-root
#
# Reference:
#   TPxi_DayOfRegistration.py - SPA architecture, CRUD, search
#   TPxi_RegistrationReportBuilder.py - print popup, HTML gen
#----------------------------------------------------------------------

import json
import datetime
import random
import re

model.Header = 'Rollsheet Generator'

# =====================================================================
# HELPER FUNCTIONS
# =====================================================================

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
    try:
        return int(val)
    except:
        return default

def safe_str(val):
    if val is None:
        return ''
    try:
        return str(val)
    except:
        return ''

ExcludeAnswers = [
    'Ma', 'Mon', 'Mone', 'No ', 'Non', 'Non ', 'Nine',
    'None', 'None.', 'None ', 'None. ', 'None know', 'None know ', 'no allergies', 'No alleriges',
    'No concerns', 'None Known', 'None Known ', 'None \\nNone', 'Nonr', 'Nome',
    'No food allergies ', 'null', 'N/A', 'N/', 'N_A',
    'NKDA', 'KNA', 'NA', 'NA ', 'NKA', 'N-A', 'NS', 'N/S', 'No',
    'no food allergies', '5', ''
]

# =====================================================================
# HELPER: SQL LIKE pattern matcher (Python-side)
# =====================================================================
def _sql_like_match(value, pattern):
    """Match a value against a SQL LIKE pattern using regex."""
    if not value or not pattern:
        return False
    regex_pattern = re.escape(pattern)
    regex_pattern = regex_pattern.replace(r'\%', '.*').replace(r'\_', '.')
    regex_pattern = '^' + regex_pattern + '$'
    try:
        return bool(re.match(regex_pattern, str(value), re.IGNORECASE))
    except:
        return False

# =====================================================================
# HELPER: Migrate old config format to dataSources
# =====================================================================
def _migrate_config(config):
    """Convert old config (includeMedical/recregFields) to dataSources format."""
    if 'dataSources' in config:
        return config

    data_sources = []

    # Migrate medical question labels
    if config.get('includeMedical', False):
        for idx, label in enumerate(config.get('medicalQuestionLabels', [])):
            data_sources.append({
                'id': 'ds_mig_med_' + str(idx),
                'sourceType': 'regQuestion',
                'label': 'Medical',
                'colorBg': '#fff3cd',
                'colorText': '#e65100',
                'questionPattern': label,
                'excludeNone': True,
                'matchAnswer': '',
                'showAsWarning': False,
                'enabled': True
            })

    # Migrate photo question label
    if config.get('includePhotos', False) and config.get('photoQuestionLabel', ''):
        data_sources.append({
            'id': 'ds_mig_photo_0',
            'sourceType': 'regQuestion',
            'label': 'NO PHOTOS',
            'colorBg': '#ffebee',
            'colorText': '#c62828',
            'questionPattern': config.get('photoQuestionLabel', ''),
            'matchAnswer': '%No%',
            'showAsWarning': True,
            'excludeNone': False,
            'enabled': True
        })

    # Migrate recreg fields
    recreg_fields = config.get('recregFields', {})
    color_map = [
        ('emergencyContact', 'EC', '#e3f2fd', '#1565c0'),
        ('medicalAllergies', 'Allergies', '#fff3cd', '#e65100'),
        ('doctor', 'Dr', '#e8f5e9', '#2e7d32'),
        ('insurance', 'Ins', '#f3e5f5', '#7b1fa2')
    ]
    for group, label, bg, text_color in color_map:
        if recreg_fields.get(group, False):
            data_sources.append({
                'id': 'ds_mig_rr_' + group,
                'sourceType': 'recreg',
                'label': label,
                'colorBg': bg,
                'colorText': text_color,
                'recregGroup': group,
                'enabled': True
            })

    config['dataSources'] = data_sources
    return config

# =====================================================================
# HELPER: Query all data sources efficiently
# =====================================================================
def _query_data_sources(data_sources, org_ids_str):
    """Process all data sources, returning {ds_id: {pid: value_text}}."""
    results = {}
    for ds in data_sources:
        results[ds['id']] = {}

    reg_sources = [ds for ds in data_sources if ds.get('sourceType') == 'regQuestion' and ds.get('enabled', True)]
    ev_sources = [ds for ds in data_sources if ds.get('sourceType') == 'extraValue' and ds.get('enabled', True)]
    rr_sources = [ds for ds in data_sources if ds.get('sourceType') == 'recreg' and ds.get('enabled', True)]

    # --- Registration Questions (single query, distribute in Python) ---
    if reg_sources:
        conditions = []
        for ds in reg_sources:
            pattern = ds.get('questionPattern', '').replace("'", "''")
            if pattern:
                conditions.append("rq.[Label] LIKE '{0}'".format(pattern))

        if conditions:
            rq_sql = """
                SELECT DISTINCT
                    rp.PeopleId,
                    rq.[Label] AS Question,
                    ra.AnswerValue AS Answer
                FROM Registration r
                    LEFT JOIN RegPeople rp ON rp.RegistrationId = r.RegistrationId
                    LEFT JOIN RegAnswer ra ON ra.RegPeopleId = rp.RegPeopleId
                    INNER JOIN RegQuestion rq ON rq.RegQuestionId = ra.RegQuestionId
                    INNER JOIN OrganizationMembers om ON om.PeopleId = rp.PeopleId
                WHERE om.OrganizationId IN ({0})
                    AND ({1})
                ORDER BY rp.PeopleId, rq.[Label]
            """.format(org_ids_str, " OR ".join(conditions))

            try:
                for row in q.QuerySql(rq_sql):
                    pid = row.PeopleId
                    question = safe_str(row.Question)
                    answer = safe_str(row.Answer).strip().strip('"')

                    for ds in reg_sources:
                        pattern = ds.get('questionPattern', '')
                        if not _sql_like_match(question, pattern):
                            continue

                        if ds.get('excludeNone', False):
                            should_exclude = False
                            answer_upper = answer.upper().strip()
                            for ea in ExcludeAnswers:
                                if answer_upper == ea.upper().strip():
                                    should_exclude = True
                                    break
                            if should_exclude or not answer:
                                continue

                        match_answer = ds.get('matchAnswer', '')
                        if match_answer:
                            if not _sql_like_match(answer, match_answer):
                                continue

                        ds_id = ds['id']
                        if ds.get('showAsWarning', False):
                            results[ds_id][pid] = ds.get('label', 'WARNING')
                        else:
                            if pid not in results[ds_id]:
                                results[ds_id][pid] = answer
                            else:
                                results[ds_id][pid] = results[ds_id][pid] + ' | ' + answer
            except:
                pass

    # --- Extra Values (single query with LEFT JOINs) ---
    if ev_sources:
        ev_joins = []
        ev_selects = ["p.PeopleId"]
        for idx, ds in enumerate(ev_sources):
            alias = "pe{0}".format(idx)
            field = ds.get('evField', '').replace("'", "''")
            ev_type = ds.get('evType', 'text')
            ev_joins.append("LEFT JOIN PeopleExtra {0} ON {0}.PeopleId = p.PeopleId AND {0}.Field = '{1}'".format(alias, field))
            if ev_type == 'bit':
                ev_selects.append("{0}.BitValue AS ev{1}".format(alias, idx))
            elif ev_type == 'int':
                ev_selects.append("{0}.IntValue AS ev{1}".format(alias, idx))
            elif ev_type == 'date':
                ev_selects.append("CONVERT(VARCHAR, {0}.DateValue, 101) AS ev{1}".format(alias, idx))
            elif ev_type == 'code':
                ev_selects.append("COALESCE({0}.CodeValue, {0}.StrValue) AS ev{1}".format(alias, idx))
            else:
                ev_selects.append("COALESCE({0}.Data, {0}.StrValue) AS ev{1}".format(alias, idx))

        ev_sql = """
            SELECT DISTINCT {0}
            FROM People p
            INNER JOIN OrganizationMembers om ON om.PeopleId = p.PeopleId
            {1}
            WHERE om.OrganizationId IN ({2})
        """.format(', '.join(ev_selects), '\n'.join(ev_joins), org_ids_str)

        try:
            for row in q.QuerySql(ev_sql):
                pid = row.PeopleId
                for idx, ds in enumerate(ev_sources):
                    val = getattr(row, 'ev' + str(idx), None)
                    if val is not None and safe_str(val).strip():
                        ev_type = ds.get('evType', 'text')
                        if ev_type == 'bit':
                            if val:
                                results[ds['id']][pid] = ds.get('label', 'Yes')
                        else:
                            results[ds['id']][pid] = safe_str(val).strip()
        except:
            pass

    # --- RecReg (single query, distribute by group) ---
    if rr_sources:
        try:
            recreg_sql = """
                SELECT DISTINCT
                    rr.PeopleId,
                    ISNULL(rr.MedAllergy, '') AS MedAllergy,
                    ISNULL(rr.MedicalDescription, '') AS MedicalDescription,
                    ISNULL(rr.emcontact, '') AS emcontact,
                    ISNULL(rr.emphone, '') AS emphone,
                    ISNULL(rr.doctor, '') AS doctor,
                    ISNULL(rr.docphone, '') AS docphone,
                    ISNULL(rr.insurance, '') AS insurance,
                    ISNULL(rr.policy, '') AS policy
                FROM RecReg rr
                INNER JOIN OrganizationMembers om ON om.PeopleId = rr.PeopleId
                WHERE om.OrganizationId IN ({0})
            """.format(org_ids_str)

            for row in q.QuerySql(recreg_sql):
                pid = row.PeopleId
                for ds in rr_sources:
                    group = ds.get('recregGroup', '')
                    text = ''
                    if group == 'emergencyContact':
                        ec_name = safe_str(row.emcontact)
                        ec_phone = safe_str(row.emphone)
                        if ec_name or ec_phone:
                            text = ec_name
                            if ec_phone:
                                text = text + ' (' + ec_phone + ')' if text else ec_phone
                    elif group == 'medicalAllergies':
                        allergies = safe_str(row.MedAllergy)
                        medical = safe_str(row.MedicalDescription)
                        med_parts = []
                        if allergies and allergies.lower() not in ['true', 'false', '0', '1']:
                            skip = False
                            for ea in ExcludeAnswers:
                                if allergies.upper().strip() == ea.upper().strip():
                                    skip = True
                                    break
                            if not skip:
                                med_parts.append(allergies)
                        if medical:
                            skip_med = False
                            for ea in ExcludeAnswers:
                                if medical.upper().strip() == ea.upper().strip():
                                    skip_med = True
                                    break
                            if not skip_med:
                                med_parts.append(medical)
                        text = ' | '.join(med_parts)
                    elif group == 'doctor':
                        doc = safe_str(row.doctor)
                        doc_ph = safe_str(row.docphone)
                        if doc or doc_ph:
                            text = doc
                            if doc_ph:
                                text = text + ' (' + doc_ph + ')' if text else doc_ph
                    elif group == 'insurance':
                        ins = safe_str(row.insurance)
                        pol = safe_str(row.policy)
                        if ins or pol:
                            text = ins
                            if pol:
                                text = text + ' #' + pol if text else '#' + pol

                    if text:
                        results[ds['id']][pid] = text
        except:
            pass

    return results

# =====================================================================
# HELPER: Build table row for a member
# =====================================================================
def _build_table_row(student, columns, ds_results, data_sources, name_line_cols=None):
    if name_line_cols is None:
        name_line_cols = {}
    parts = []
    pid = student['PeopleId']
    name = html_escape(student['Name'])

    # Collect name-line items (displayed inline after the name)
    nl_items = []
    # Map of column keys to their display values
    col_vals = [
        ('age', 'Age: ' + html_escape(student.get('Age', '')) if student.get('Age', '') else ''),
        ('gender', html_escape(student.get('Gender', '')) if student.get('Gender', '') else ''),
        ('phone', html_escape(student.get('Phone', '')) if student.get('Phone', '') else ''),
        ('email', html_escape(student.get('Email', '')) if student.get('Email', '') else ''),
        ('subgroup', 'Group: ' + html_escape(student.get('SubGroups', '')) if student.get('SubGroups', '') else ''),
        ('memberType', html_escape(student.get('MemberType', '')) if student.get('MemberType', '') else '')
    ]
    for key, val in col_vals:
        if columns.get(key, False) and val and name_line_cols.get(key, False):
            nl_items.append(val)

    # Build name display with optional inline items
    if nl_items:
        name_display = '{0} <span class="rs-info-inline">({1})</span>'.format(name, ', '.join(nl_items))
    else:
        name_display = name

    # Data source badges (replaces medical, photo, recreg)
    ds_html = ''
    pending_inline = []  # accumulates badges for inline grouping
    for ds in data_sources:
        if not ds.get('enabled', True):
            continue
        ds_id = ds.get('id', '')
        if ds_id in ds_results and pid in ds_results[ds_id]:
            value = html_escape(ds_results[ds_id][pid])
            bg = ds.get('colorBg', '#f0f0f0')
            color = ds.get('colorText', '#333')
            label = html_escape(ds.get('label', ''))
            if ds.get('showAsWarning', False):
                badge = '<span style="background:{0};color:{1};padding:1px 4px;border-radius:2px;font-weight:700;font-size:10px;">{2}</span>'.format(bg, color, label)
            else:
                display_text = label + ': ' + value if label else value
                badge = '<span style="background:{0};color:{1};padding:1px 4px;font-size:10px;border-radius:2px;">{2}</span>'.format(bg, color, display_text)

            mode = ds.get('displayMode', 'own_line')
            if mode == 'append' and pending_inline:
                pending_inline.append(badge)
            else:
                if pending_inline:
                    ds_html += '<div style="margin-top:2px;display:flex;flex-wrap:wrap;gap:4px;">' + ''.join(pending_inline) + '</div>'
                pending_inline = [badge]
    if pending_inline:
        ds_html += '<div style="margin-top:2px;display:flex;flex-wrap:wrap;gap:4px;">' + ''.join(pending_inline) + '</div>'

    # Build demographics info line (items NOT on name line)
    demo_items = []
    if columns.get('age', True) and student.get('Age', '') and not name_line_cols.get('age', False):
        demo_items.append('Age: ' + html_escape(student['Age']))
    if columns.get('gender', True) and student.get('Gender', '') and not name_line_cols.get('gender', False):
        demo_items.append(html_escape(student['Gender']))
    if columns.get('phone', False) and student.get('Phone', '') and not name_line_cols.get('phone', False):
        demo_items.append(html_escape(student['Phone']))
    if columns.get('email', False) and student.get('Email', '') and not name_line_cols.get('email', False):
        demo_items.append(html_escape(student['Email']))
    demo_html = ''
    if demo_items:
        demo_html = '<div class="rs-info-line">{0}</div>'.format(' &bull; '.join(demo_items))

    # Build org info line (items NOT on name line)
    org_items = []
    if columns.get('subgroup', False) and student.get('SubGroups', '') and not name_line_cols.get('subgroup', False):
        org_items.append('Group: ' + html_escape(student['SubGroups']))
    if columns.get('memberType', False) and student.get('MemberType', '') and not name_line_cols.get('memberType', False):
        org_items.append(html_escape(student['MemberType']))
    org_html = ''
    if org_items:
        org_html = '<div class="rs-info-line">{0}</div>'.format(' &bull; '.join(org_items))

    parts.append('<tr class="rs-row">')
    parts.append('<td class="rs-td rs-td-check"><div class="rs-checkbox"></div></td>')
    parts.append('<td class="rs-td rs-td-name"><div class="rs-name">{0}</div>{1}{2}{3}</td>'.format(name_display, demo_html, org_html, ds_html))
    parts.append('</tr>')
    return ''.join(parts)

# =====================================================================
# AJAX HANDLERS (POST)
# =====================================================================
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -----------------------------------------------------------------
    # Load all configs
    # -----------------------------------------------------------------
    if action == 'load_configs':
        try:
            raw = model.TextContent("RollSheet_Configs") or ''
            data = json.loads(raw) if raw else {'configs': []}
            print json.dumps({'success': True, 'configs': data.get('configs', [])})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Save config (create or update)
    # -----------------------------------------------------------------
    elif action == 'save_config':
        try:
            config_json = str(Data.config_json) if hasattr(Data, 'config_json') else ''
            config = json.loads(config_json)

            raw = model.TextContent("RollSheet_Configs") or ''
            data = json.loads(raw) if raw else {'configs': []}
            configs = data.get('configs', [])

            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

            existing_idx = -1
            for i, c in enumerate(configs):
                if c.get('id') == config.get('id'):
                    existing_idx = i
                    break

            if existing_idx >= 0:
                config['updatedAt'] = now
                config['createdAt'] = configs[existing_idx].get('createdAt', now)
                configs[existing_idx] = config
            else:
                if not config.get('id'):
                    config['id'] = 'rs_' + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + str(random.randint(100, 999))
                config['createdAt'] = now
                config['updatedAt'] = now
                configs.append(config)

            data['configs'] = configs
            model.WriteContentText("RollSheet_Configs", json.dumps(data), "")
            print json.dumps({'success': True, 'config': config, 'message': 'Config saved'})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Delete config
    # -----------------------------------------------------------------
    elif action == 'delete_config':
        try:
            config_id = str(Data.config_id) if hasattr(Data, 'config_id') else ''
            raw = model.TextContent("RollSheet_Configs") or ''
            data = json.loads(raw) if raw else {'configs': []}
            configs = data.get('configs', [])

            new_configs = [c for c in configs if c.get('id') != config_id]
            if len(new_configs) < len(configs):
                data['configs'] = new_configs
                model.WriteContentText("RollSheet_Configs", json.dumps(data), "")
                print json.dumps({'success': True, 'message': 'Config deleted'})
            else:
                print json.dumps({'success': False, 'message': 'Config not found'})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Get filter dropdowns (programs/divisions)
    # -----------------------------------------------------------------
    elif action == 'get_filters':
        try:
            prog_sql = "SELECT Id, Name FROM Program ORDER BY Name"
            div_sql = "SELECT Id, Name, ProgId FROM Division ORDER BY Name"

            programs = []
            for r in q.QuerySql(prog_sql):
                programs.append({'id': r.Id, 'name': r.Name})

            divisions = []
            for r in q.QuerySql(div_sql):
                divisions.append({'id': r.Id, 'name': r.Name, 'progId': r.ProgId})

            print json.dumps({'success': True, 'programs': programs, 'divisions': divisions})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Search involvements (orgs) by name or ID
    # -----------------------------------------------------------------
    elif action == 'search_involvements':
        try:
            search_term = getattr(Data, 'search_term', '')
            if search_term:
                search_term = str(search_term)
            division_id = getattr(Data, 'division_id', '')
            program_id = getattr(Data, 'program_id', '')

            where_clauses = ["o.OrganizationStatusId = 30"]

            if search_term:
                safe_term = search_term.replace("'", "''")
                where_clauses.append("""(o.OrganizationName LIKE '%{0}%'
                    OR CAST(o.OrganizationId AS VARCHAR) = '{0}')""".format(safe_term))

            if division_id:
                where_clauses.append("o.DivisionId = {0}".format(int(division_id)))

            if program_id:
                where_clauses.append("d.ProgId = {0}".format(int(program_id)))

            sql = """
                SELECT TOP 50
                    o.OrganizationId,
                    o.OrganizationName,
                    d.Name as DivisionName,
                    p.Name as ProgramName,
                    (SELECT COUNT(*) FROM OrganizationMembers om
                     WHERE om.OrganizationId = o.OrganizationId) as MemberCount
                FROM Organizations o
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE {0}
                ORDER BY o.OrganizationName
            """.format(" AND ".join(where_clauses))

            results = q.QuerySql(sql)
            involvements = []
            for r in results:
                involvements.append({
                    'orgId': r.OrganizationId,
                    'orgName': r.OrganizationName,
                    'divisionName': r.DivisionName or '',
                    'programName': r.ProgramName or '',
                    'memberCount': r.MemberCount or 0
                })
            print json.dumps({'success': True, 'involvements': involvements})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Get org member counts (quick count for selected orgs)
    # -----------------------------------------------------------------
    elif action == 'get_org_members_count':
        try:
            org_ids_str = str(Data.org_ids) if hasattr(Data, 'org_ids') else ''
            org_ids = [int(x.strip()) for x in org_ids_str.split(',') if x.strip()]
            counts = []
            for oid in org_ids:
                cnt_sql = "SELECT COUNT(*) as cnt FROM OrganizationMembers WHERE OrganizationId = {0}".format(oid)
                cnt_row = q.QuerySqlTop1(cnt_sql)
                counts.append({
                    'orgId': oid,
                    'count': cnt_row.cnt if cnt_row and hasattr(cnt_row, 'cnt') else 0
                })
            print json.dumps({'success': True, 'counts': counts})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Generate rollsheet HTML
    # -----------------------------------------------------------------
    elif action == 'generate_rollsheet':
        try:
            config_json = str(Data.config_json) if hasattr(Data, 'config_json') else ''
            config = json.loads(config_json)

            # Migrate old config format on-the-fly
            config = _migrate_config(config)

            source_type = config.get('sourceType', 'program_division')
            program_id = safe_int(config.get('programId', 0))
            division_id = safe_int(config.get('divisionId', 0))
            selected_orgs = config.get('selectedOrgs', [])
            exclude_org_ids = config.get('excludeOrgIds', [])
            columns = config.get('columns', {})
            name_line_cols = config.get('nameLineColumns', {})
            data_sources = config.get('dataSources', [])
            sort_by = config.get('sortBy', 'name')
            layout = config.get('layout', 'two_column')
            title = config.get('title', 'Rollsheet')
            show_footer = config.get('showFooter', True)
            report_date = str(Data.report_date) if hasattr(Data, 'report_date') and Data.report_date else ''
            report_date_iso = str(Data.report_date_iso) if hasattr(Data, 'report_date_iso') and Data.report_date_iso else ''
            only_with_meeting = config.get('onlyWithMeeting', False)

            # Determine org IDs to include
            org_ids = []
            if source_type == 'specific_orgs':
                org_ids = [safe_int(o.get('orgId', 0)) for o in selected_orgs if safe_int(o.get('orgId', 0)) > 0]
            else:
                org_where = []
                if program_id > 0:
                    org_where.append("os.ProgId = {0}".format(program_id))
                if division_id > 0:
                    org_where.append("os.DivId = {0}".format(division_id))
                if not org_where:
                    print json.dumps({'success': False, 'message': 'Program or Division required'})
                else:
                    org_sql = """
                        SELECT DISTINCT os.OrgId
                        FROM OrganizationStructure os
                        JOIN Organizations o ON os.OrgId = o.OrganizationId
                        WHERE {0} AND o.OrganizationStatusId = 30
                    """.format(" AND ".join(org_where))
                    for r in q.QuerySql(org_sql):
                        org_ids.append(r.OrgId)

            # Apply exclusions
            if exclude_org_ids:
                exclude_set = set(exclude_org_ids)
                org_ids = [oid for oid in org_ids if oid not in exclude_set]

            # Filter to only orgs with a meeting scheduled on the report date's day of week
            if only_with_meeting and report_date_iso and org_ids:
                import datetime
                safe_iso = report_date_iso.replace("'", "")
                try:
                    dt = datetime.datetime.strptime(safe_iso, '%Y-%m-%d')
                    # Python weekday(): Mon=0..Sun=6  ->  SchedDay: Sun=0..Sat=6
                    sched_day = (dt.weekday() + 1) % 7
                except:
                    sched_day = -1
                if sched_day >= 0:
                    schedule_sql = """
                        SELECT DISTINCT OrganizationId
                        FROM OrgSchedule
                        WHERE OrganizationId IN ({0})
                            AND SchedDay = {1}
                    """.format(','.join(str(oid) for oid in org_ids), sched_day)
                    meeting_org_ids = set()
                    for r in q.QuerySql(schedule_sql):
                        meeting_org_ids.add(r.OrganizationId)
                    org_ids = [oid for oid in org_ids if oid in meeting_org_ids]

            if not org_ids:
                no_org_msg = 'No organizations found matching your config.'
                if only_with_meeting and report_date_iso:
                    no_org_msg = 'No organizations have a meeting scheduled for ' + html_escape(report_date or report_date_iso) + '.'
                print json.dumps({'success': True, 'html': '<p style="text-align:center;color:#666;padding:40px;">' + no_org_msg + '</p>', 'orgCount': 0, 'totalMembers': 0})
            else:
                org_ids_str = ','.join(str(oid) for oid in org_ids)

                # Build column selections
                col_parts = ["p.PeopleId", "p.Name2 AS [Name]", "om.OrganizationId", "os.Organization"]
                if columns.get('age', True):
                    col_parts.append("p.Age")
                if columns.get('gender', True):
                    col_parts.append("g.[Description] AS Gender")
                if columns.get('phone', False):
                    col_parts.append("ISNULL(p.CellPhone, p.HomePhone) AS Phone")
                if columns.get('email', False):
                    col_parts.append("p.EmailAddress")
                if columns.get('memberType', False):
                    col_parts.append("mt.Description AS MemberType")
                if columns.get('subgroup', False):
                    col_parts.append("STUFF((SELECT ', ' + mt2.Name FROM OrgMemMemTags ommt2 JOIN MemberTags mt2 ON ommt2.MemberTagId = mt2.Id WHERE ommt2.PeopleId = p.PeopleId AND ommt2.OrgId = om.OrganizationId FOR XML PATH('')), 1, 2, '') AS SubGroups")

                join_parts = [
                    "INNER JOIN OrganizationMembers om ON om.PeopleId = p.PeopleId",
                    "INNER JOIN OrganizationStructure os ON os.OrgId = om.OrganizationId"
                ]
                if columns.get('gender', True):
                    join_parts.append("LEFT JOIN lookup.Gender g ON g.Id = p.GenderId")
                if columns.get('memberType', False):
                    join_parts.append("LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId")

                sort_clause = "os.Organization, p.Name2"
                if sort_by == 'age':
                    sort_clause = "os.Organization, p.Age, p.Name2"
                elif sort_by == 'gender':
                    sort_clause = "os.Organization, p.GenderId, p.Name2"

                members_sql = """
                    SELECT DISTINCT {0}
                    FROM People p
                    {1}
                    WHERE om.OrganizationId IN ({2})
                    ORDER BY {3}
                """.format(', '.join(col_parts), '\n'.join(join_parts), org_ids_str, sort_clause)

                members = q.QuerySql(members_sql)

                # Organize by org
                orgs = {}
                for row in members:
                    oid = row.OrganizationId
                    if oid not in orgs:
                        orgs[oid] = {'name': row.Organization, 'members': []}
                    member = {'PeopleId': row.PeopleId, 'Name': row.Name}
                    if columns.get('age', True):
                        member['Age'] = safe_str(row.Age) if hasattr(row, 'Age') and row.Age else ''
                    if columns.get('gender', True):
                        member['Gender'] = safe_str(row.Gender) if hasattr(row, 'Gender') and row.Gender else ''
                    if columns.get('phone', False):
                        member['Phone'] = safe_str(row.Phone) if hasattr(row, 'Phone') and row.Phone else ''
                    if columns.get('email', False):
                        member['Email'] = safe_str(row.EmailAddress) if hasattr(row, 'EmailAddress') and row.EmailAddress else ''
                    if columns.get('memberType', False):
                        member['MemberType'] = safe_str(row.MemberType) if hasattr(row, 'MemberType') and row.MemberType else ''
                    if columns.get('subgroup', False):
                        member['SubGroups'] = safe_str(row.SubGroups) if hasattr(row, 'SubGroups') and row.SubGroups else ''
                    orgs[oid]['members'].append(member)

                # Query all data sources
                enabled_ds = [ds for ds in data_sources if ds.get('enabled', True)]
                ds_results = _query_data_sources(enabled_ds, org_ids_str) if enabled_ds else {}

                # Build HTML
                html_parts = []
                sorted_orgs = sorted(orgs.items(), key=lambda x: x[1]['name'])
                total_members = 0
                org_count = len(sorted_orgs)

                for idx, (oid, org_data) in enumerate(sorted_orgs):
                    org_name = html_escape(org_data['name'])
                    students = org_data['members']
                    total_members += len(students)

                    page_break = ' style="page-break-after: always;"' if idx < org_count - 1 else ''
                    html_parts.append('<div class="rs-org-page"{0}>'.format(page_break))

                    html_parts.append('<div class="rs-org-header">')
                    html_parts.append('<h2 class="rs-org-title">{0}</h2>'.format(html_escape(title)))
                    html_parts.append('<h3 class="rs-org-name">{0}</h3>'.format(org_name))
                    date_display = html_escape(report_date) if report_date else '__________'
                    html_parts.append('<div class="rs-org-meta">Classroom {0} of {1} | Teacher: __________________ | Date: {2}</div>'.format(idx + 1, org_count, date_display))
                    html_parts.append('</div>')

                    th_parts = ['<th class="rs-th rs-th-check">&#10003;</th>', '<th class="rs-th rs-th-name">Name</th>']

                    if layout == 'table':
                        html_parts.append('<table class="rs-table"><thead><tr>{0}</tr></thead><tbody>'.format(''.join(th_parts)))
                        for s in students:
                            html_parts.append(_build_table_row(s, columns, ds_results, enabled_ds, name_line_cols))
                        html_parts.append('</tbody></table>')

                    elif layout == 'single_column':
                        html_parts.append('<table class="rs-table"><thead><tr>{0}</tr></thead><tbody>'.format(''.join(th_parts)))
                        for s in students:
                            html_parts.append(_build_table_row(s, columns, ds_results, enabled_ds, name_line_cols))
                        html_parts.append('</tbody></table>')

                    elif layout == 'three_column':
                        n = len(students)
                        third = int((n + 2) / 3)
                        cols = [students[:third], students[third:third*2], students[third*2:]]
                        max_len = max(len(c) for c in cols)

                        html_parts.append('<div class="rs-two-col">')
                        for ci in range(3):
                            html_parts.append('<div class="rs-col">')
                            html_parts.append('<table class="rs-table"><thead><tr>{0}</tr></thead><tbody>'.format(''.join(th_parts)))
                            for s in cols[ci]:
                                html_parts.append(_build_table_row(s, columns, ds_results, enabled_ds, name_line_cols))
                            # pad shorter columns with empty rows
                            for _ in range(max_len - len(cols[ci])):
                                html_parts.append('<tr class="rs-row"><td class="rs-td rs-td-check"><div class="rs-checkbox"></div></td>')
                                html_parts.append('<td class="rs-td">&nbsp;</td></tr>')
                            html_parts.append('</tbody></table>')
                            html_parts.append('</div>')
                        html_parts.append('</div>')

                    else:
                        mid_point = int((len(students) + 1) / 2)
                        left_col = students[:mid_point]
                        right_col = students[mid_point:]

                        html_parts.append('<div class="rs-two-col">')
                        html_parts.append('<div class="rs-col">')
                        html_parts.append('<table class="rs-table"><thead><tr>{0}</tr></thead><tbody>'.format(''.join(th_parts)))
                        for s in left_col:
                            html_parts.append(_build_table_row(s, columns, ds_results, enabled_ds, name_line_cols))
                        html_parts.append('</tbody></table>')
                        html_parts.append('</div>')

                        html_parts.append('<div class="rs-col">')
                        html_parts.append('<table class="rs-table"><thead><tr>{0}</tr></thead><tbody>'.format(''.join(th_parts)))
                        for s in right_col:
                            html_parts.append(_build_table_row(s, columns, ds_results, enabled_ds, name_line_cols))
                        while len(right_col) < len(left_col):
                            html_parts.append('<tr class="rs-row"><td class="rs-td rs-td-check"><div class="rs-checkbox"></div></td>')
                            html_parts.append('<td class="rs-td">&nbsp;</td>')
                            html_parts.append('</tr>')
                            right_col.append(None)
                        html_parts.append('</tbody></table>')
                        html_parts.append('</div>')
                        html_parts.append('</div>')

                    # Footer with dynamic legend
                    if show_footer:
                        html_parts.append('<div class="rs-footer">')
                        html_parts.append('<div class="rs-footer-stats">')
                        html_parts.append('<strong>Total:</strong> {0} &nbsp;&nbsp; <strong>Present:</strong> _____ &nbsp;&nbsp; <strong>Visitors:</strong> _____'.format(len(students)))
                        html_parts.append('</div>')
                        html_parts.append('<div class="rs-footer-sig"><strong>Teacher Signature:</strong> _________________________</div>')
                        html_parts.append('</div>')
                        if enabled_ds:
                            html_parts.append('<div class="rs-legend">')
                            html_parts.append('<strong>Legend:</strong> ')
                            for ds in enabled_ds:
                                bg = ds.get('colorBg', '#f0f0f0')
                                color = ds.get('colorText', '#333')
                                label = html_escape(ds.get('label', ''))
                                html_parts.append('<span style="background:{0};color:{1};padding:1px 6px;border-radius:2px;font-weight:700;margin-right:8px;">{2}</span> '.format(bg, color, label))
                            html_parts.append('</div>')

                    html_parts.append('</div>')

                final_html = ''.join(html_parts)
                print json.dumps({'success': True, 'html': final_html, 'orgCount': org_count, 'totalMembers': total_members})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Discover available fields for data sources
    # -----------------------------------------------------------------
    elif action == 'discover_fields':
        try:
            org_ids_str = str(Data.org_ids) if hasattr(Data, 'org_ids') else ''
            org_ids = [int(x.strip()) for x in org_ids_str.split(',') if x.strip()]
            if not org_ids:
                print json.dumps({'success': False, 'message': 'No org IDs provided'})
            else:
                safe_ids = ','.join(str(oid) for oid in org_ids)

                # 1. Registration questions
                rq_sql = """
                    SELECT DISTINCT rq.[Label], MIN(rq.[Order]) AS SortOrder
                    FROM Registration r
                    JOIN RegPeople rp ON r.RegistrationId = rp.RegistrationId
                    JOIN RegAnswer ra ON rp.RegPeopleId = ra.RegPeopleId
                    JOIN RegQuestion rq ON ra.RegQuestionId = rq.RegQuestionId
                    WHERE r.OrganizationId IN ({0}) AND rq.[Label] IS NOT NULL
                    GROUP BY rq.[Label]
                    ORDER BY SortOrder
                """.format(safe_ids)

                reg_questions = []
                for row in q.QuerySql(rq_sql):
                    reg_questions.append({'label': safe_str(row.Label)})

                # 2. Extra values
                ev_sql = """
                    SELECT DISTINCT pe.Field,
                        CASE WHEN pe.BitValue IS NOT NULL THEN 'bit'
                             WHEN pe.IntValue IS NOT NULL THEN 'int'
                             WHEN pe.DateValue IS NOT NULL THEN 'date'
                             ELSE 'text' END AS FieldType
                    FROM PeopleExtra pe
                    INNER JOIN OrganizationMembers om ON om.PeopleId = pe.PeopleId
                    WHERE om.OrganizationId IN ({0})
                        AND pe.Field IS NOT NULL AND pe.Field != ''
                    ORDER BY pe.Field
                """.format(safe_ids)

                extra_values = []
                for row in q.QuerySql(ev_sql):
                    extra_values.append({'field': safe_str(row.Field), 'fieldType': safe_str(row.FieldType)})

                # 3. RecReg (static list)
                recreg_options = [
                    {'group': 'emergencyContact', 'label': 'Emergency Contact (name + phone)'},
                    {'group': 'medicalAllergies', 'label': 'Allergies & Medical Conditions'},
                    {'group': 'doctor', 'label': 'Doctor (name + phone)'},
                    {'group': 'insurance', 'label': 'Insurance (company + policy #)'}
                ]

                print json.dumps({
                    'success': True,
                    'regQuestions': reg_questions,
                    'extraValues': extra_values,
                    'recregOptions': recreg_options
                })
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Unknown action
    # -----------------------------------------------------------------
    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + action})

# =====================================================================
# HTML OUTPUT (GET)
# =====================================================================
else:
    html = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script>
    if (window.location.pathname.indexOf('/PyScript/') > -1) {
        window.location.href = window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search;
    }
    </script>
    <style>
/* ================================================================
   Rollsheet Generator - Scoped CSS (rs- prefix)
   ================================================================ */
.rs-root {
    --rs-primary: #2563eb;
    --rs-primary-light: #dbeafe;
    --rs-secondary: #059669;
    --rs-secondary-light: #d1fae5;
    --rs-warning: #d97706;
    --rs-warning-light: #fef3c7;
    --rs-danger: #dc2626;
    --rs-danger-light: #fee2e2;
    --rs-dark: #1e293b;
    --rs-gray: #64748b;
    --rs-light-bg: #f8fafc;
    --rs-border: #e2e8f0;
    --rs-radius: 12px;
    --rs-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
    --rs-shadow-lg: 0 10px 15px rgba(0,0,0,0.1), 0 4px 6px rgba(0,0,0,0.05);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    color: var(--rs-dark);
    max-width: 1200px;
    margin: 0 auto;
    padding: 12px;
}

/* ---- Landing ---- */
.rs-landing {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 60vh;
    gap: 24px;
}
.rs-landing-title { font-size: 32px; font-weight: 700; text-align: center; margin-bottom: 8px; }
.rs-landing-subtitle { font-size: 18px; color: var(--rs-gray); text-align: center; margin-bottom: 32px; }
.rs-landing-buttons { display: flex; gap: 24px; flex-wrap: wrap; justify-content: center; }
.rs-landing-btn {
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    width: 260px; height: 200px; border-radius: var(--rs-radius); border: 2px solid var(--rs-border);
    background: white; cursor: pointer; transition: all 0.2s; text-decoration: none;
    color: var(--rs-dark); box-shadow: var(--rs-shadow); user-select: none;
}
.rs-landing-btn:active { transform: scale(0.97); box-shadow: var(--rs-shadow-lg); }
.rs-landing-btn i { font-size: 48px; margin-bottom: 16px; }
.rs-landing-btn span { font-size: 22px; font-weight: 600; }
.rs-landing-btn.rs-admin-btn { border-color: var(--rs-primary); }
.rs-landing-btn.rs-admin-btn i { color: var(--rs-primary); }
.rs-landing-btn.rs-gen-btn { border-color: var(--rs-secondary); }
.rs-landing-btn.rs-gen-btn i { color: var(--rs-secondary); }

/* ---- Common UI ---- */
.rs-btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 8px;
    padding: 10px 20px; border-radius: 8px; border: none; font-size: 15px; font-weight: 600;
    cursor: pointer; transition: all 0.15s; user-select: none;
}
.rs-btn:active { transform: scale(0.97); }
.rs-btn-primary { background: var(--rs-primary); color: white; }
.rs-btn-success { background: var(--rs-secondary); color: white; }
.rs-btn-warning { background: var(--rs-warning); color: white; }
.rs-btn-danger { background: var(--rs-danger); color: white; }
.rs-btn-outline { background: white; color: var(--rs-dark); border: 2px solid var(--rs-border); }
.rs-btn-sm { padding: 6px 14px; font-size: 13px; }
.rs-btn:disabled { opacity: 0.5; cursor: not-allowed; }

.rs-input {
    width: 100%; padding: 10px 14px; border: 2px solid var(--rs-border); border-radius: 8px;
    font-size: 15px; outline: none; transition: border-color 0.15s; box-sizing: border-box;
}
.rs-input:focus { border-color: var(--rs-primary); }

.rs-select {
    width: 100%; padding: 10px 14px; border: 2px solid var(--rs-border); border-radius: 8px;
    font-size: 15px; background: white; box-sizing: border-box;
}

.rs-textarea {
    width: 100%; padding: 10px 14px; border: 2px solid var(--rs-border); border-radius: 8px;
    font-size: 14px; outline: none; box-sizing: border-box; resize: vertical; font-family: monospace;
}
.rs-textarea:focus { border-color: var(--rs-primary); }

.rs-label { display: block; font-size: 13px; font-weight: 600; color: var(--rs-gray); margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px; }
.rs-help-text { font-size: 12px; color: var(--rs-gray); margin-top: 4px; }

.rs-panel {
    background: white; border-radius: var(--rs-radius); border: 1px solid var(--rs-border);
    box-shadow: var(--rs-shadow); margin-bottom: 16px; overflow: hidden;
}
.rs-panel-header {
    display: flex; align-items: center; justify-content: space-between; padding: 16px 20px;
    background: var(--rs-light-bg); border-bottom: 1px solid var(--rs-border); font-weight: 600; font-size: 16px;
}
.rs-panel-body { padding: 20px; }

/* ---- Back button ---- */
.rs-back-btn {
    display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px;
    color: var(--rs-primary); font-size: 16px; font-weight: 600; cursor: pointer;
    border: none; background: none;
}

/* ---- Config cards (admin list) ---- */
.rs-config-card {
    display: flex; align-items: center; justify-content: space-between; padding: 16px 20px;
    border: 1px solid var(--rs-border); border-radius: var(--rs-radius); margin-bottom: 12px;
    background: white; cursor: pointer; transition: all 0.15s;
}
.rs-config-card:hover { background: var(--rs-light-bg); }
.rs-config-name { font-size: 18px; font-weight: 600; }
.rs-config-meta { font-size: 13px; color: var(--rs-gray); margin-top: 4px; }
.rs-config-actions { display: flex; gap: 8px; }

/* ---- Form groups ---- */
.rs-form-group { margin-bottom: 20px; }
.rs-form-row { display: flex; gap: 16px; flex-wrap: wrap; }
.rs-form-row > * { flex: 1; min-width: 200px; }

/* ---- Toggle switch ---- */
.rs-toggle { position: relative; width: 44px; height: 24px; display: inline-block; vertical-align: middle; }
.rs-toggle input { display: none; }
.rs-toggle-slider {
    position: absolute; top: 0; left: 0; right: 0; bottom: 0;
    background: #cbd5e0; border-radius: 24px; cursor: pointer; transition: 0.2s;
}
.rs-toggle-slider:before {
    content: ""; position: absolute; width: 20px; height: 20px;
    left: 2px; bottom: 2px; background: white; border-radius: 50%; transition: 0.2s;
}
.rs-toggle input:checked + .rs-toggle-slider { background: var(--rs-secondary); }
.rs-toggle input:checked + .rs-toggle-slider:before { transform: translateX(20px); }

.rs-toggle-row {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 0; border-bottom: 1px solid #f0f0f0;
}
.rs-toggle-row:last-child { border-bottom: none; }
.rs-toggle-label { font-size: 14px; color: var(--rs-dark); }

/* ---- Checkbox grid ---- */
.rs-checkbox-grid { display: flex; flex-wrap: wrap; gap: 12px; }
.rs-checkbox-item { display: flex; align-items: center; gap: 6px; font-size: 14px; }
.rs-checkbox-item input[type="checkbox"] { width: 18px; height: 18px; }

/* ---- Search box ---- */
.rs-search-box { position: relative; }
.rs-search-results {
    position: absolute; top: 100%; left: 0; right: 0; background: white;
    border: 1px solid var(--rs-border); border-radius: 0 0 8px 8px;
    box-shadow: var(--rs-shadow-lg); max-height: 300px; overflow-y: auto; z-index: 100;
}
.rs-search-result-item {
    padding: 10px 14px; cursor: pointer; border-bottom: 1px solid var(--rs-border);
}
.rs-search-result-item:hover { background: var(--rs-primary-light); }
.rs-search-result-item:last-child { border-bottom: none; }
.rs-search-result-name { font-weight: 600; }
.rs-search-result-meta { font-size: 12px; color: var(--rs-gray); }

/* ---- Selected orgs list ---- */
.rs-selected-org-item {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px; background: var(--rs-primary-light); border-radius: 8px;
    margin-bottom: 8px; font-weight: 600;
}
.rs-remove-org { cursor: pointer; color: var(--rs-danger); font-size: 18px; padding: 4px 8px; }

/* ---- Rollsheet print styles ---- */
.rs-org-page { margin-bottom: 30px; }
.rs-org-header { text-align: center; margin-bottom: 15px; }
.rs-org-title { font-size: 22px; margin: 0 0 4px; color: #333; }
.rs-org-name { font-size: 18px; margin: 0 0 8px; color: #007bff; }
.rs-org-meta { font-size: 13px; color: #666; }
.rs-two-col { display: flex; gap: 16px; }
.rs-col { flex: 1; }
.rs-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.rs-th { padding: 6px 8px; text-align: left; border: 1px solid #ddd; background: #f8f9fa; font-weight: 700; }
.rs-th-check { width: 30px; text-align: center; }
.rs-td { padding: 6px 8px; border: 1px solid #ddd; vertical-align: top; }
.rs-td-check { text-align: center; }
.rs-td-name { min-width: 120px; }
.rs-checkbox { width: 18px; height: 18px; border: 2px solid #007bff; border-radius: 3px; margin: 0 auto; }
.rs-name { font-weight: 700; font-size: 13px; }
.rs-info-inline { font-weight: 400; font-size: 11px; color: #555; }
.rs-info-line { font-size: 11px; color: #555; margin-top: 1px; }
.rs-footer { margin-top: 16px; border-top: 1px solid #ddd; padding-top: 12px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }
.rs-footer-stats { font-size: 13px; }
.rs-footer-sig { font-size: 13px; }
.rs-legend { margin-top: 10px; padding: 8px; background: #f8f9fa; border-radius: 5px; font-size: 11px; }

/* ---- Data Source cards ---- */
.rs-ds-card {
    display: flex; align-items: center; justify-content: space-between; padding: 8px 12px;
    border: 1px solid var(--rs-border); border-radius: 8px; margin-bottom: 8px; background: white;
}
.rs-ds-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 11px; font-weight: 700; margin-right: 8px; }
.rs-ds-card-info { font-size: 13px; color: var(--rs-gray); }
.rs-ds-card-actions { display: flex; gap: 4px; }
.rs-ds-card-actions button { background: none; border: none; cursor: pointer; padding: 4px 6px; font-size: 14px; color: var(--rs-gray); border-radius: 4px; }
.rs-ds-card-actions button:hover { background: var(--rs-light-bg); }
.rs-ds-editor { padding: 16px; border: 2px solid var(--rs-primary); border-radius: 8px; margin-bottom: 8px; background: #fafbff; }
.rs-color-presets { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 8px; }
.rs-color-dot { width: 28px; height: 28px; border-radius: 50%; cursor: pointer; border: 2px solid transparent; display: flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 700; }
.rs-color-dot:hover, .rs-color-dot.rs-active { border-color: var(--rs-dark); }
.rs-color-inputs { display: flex; gap: 12px; align-items: center; }
.rs-color-inputs input[type="color"] { width: 40px; height: 32px; border: 1px solid var(--rs-border); border-radius: 4px; cursor: pointer; padding: 2px; }

/* ---- Preview area ---- */
.rs-preview-area {
    border: 1px solid var(--rs-border); border-radius: var(--rs-radius); padding: 20px;
    background: white; min-height: 400px;
}
.rs-preview-toolbar {
    display: flex; gap: 12px; align-items: center; margin-bottom: 16px;
    padding: 12px; background: var(--rs-light-bg); border-radius: 8px; flex-wrap: wrap;
}
.rs-stats { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.rs-stat-card { padding: 12px 20px; background: var(--rs-light-bg); border-radius: 8px; text-align: center; }
.rs-stat-num { font-size: 24px; font-weight: 700; color: var(--rs-primary); }
.rs-stat-label { font-size: 12px; color: var(--rs-gray); }

/* ---- Toast ---- */
.rs-toast-container {
    position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%);
    z-index: 9999; display: flex; flex-direction: column; align-items: center;
    gap: 8px; pointer-events: none;
}
.rs-toast {
    padding: 14px 28px; border-radius: 10px; font-size: 16px; font-weight: 600;
    color: white; opacity: 0; transform: translateY(20px); transition: all 0.3s;
    pointer-events: auto; max-width: 90vw; text-align: center;
}
.rs-toast.rs-show { opacity: 1; transform: translateY(0); }
.rs-toast-success { background: var(--rs-secondary); }
.rs-toast-danger { background: var(--rs-danger); }
.rs-toast-info { background: var(--rs-primary); }

/* ---- Loading spinner ---- */
.rs-loading { text-align: center; padding: 40px; color: var(--rs-gray); }
.rs-spinner {
    display: inline-block; width: 32px; height: 32px;
    border: 3px solid var(--rs-border); border-top: 3px solid var(--rs-primary);
    border-radius: 50%; animation: rsSpin 0.8s linear infinite;
}
@keyframes rsSpin { to { transform: rotate(360deg); } }

/* ---- Empty state ---- */
.rs-empty { text-align: center; padding: 40px 20px; color: #a0aec0; }

/* ---- Utilities ---- */
.rs-flex { display: flex; }
.rs-flex-wrap { flex-wrap: wrap; }
.rs-items-center { align-items: center; }
.rs-justify-between { justify-content: space-between; }
.rs-gap-8 { gap: 8px; }
.rs-gap-12 { gap: 12px; }
.rs-gap-16 { gap: 16px; }
.rs-mb-8 { margin-bottom: 8px; }
.rs-mb-12 { margin-bottom: 12px; }
.rs-mb-16 { margin-bottom: 16px; }
.rs-mb-24 { margin-bottom: 24px; }
.rs-mt-12 { margin-top: 12px; }
.rs-mt-16 { margin-top: 16px; }
.rs-text-center { text-align: center; }
.rs-text-muted { color: var(--rs-gray); }
.rs-text-sm { font-size: 14px; }
.rs-d-none { display: none; }

/* ---- Responsive ---- */
@media (max-width: 768px) {
    .rs-landing-buttons { flex-direction: column; align-items: center; }
    .rs-form-row { flex-direction: column; }
    .rs-two-col { flex-direction: column; }
    .rs-config-card { flex-direction: column; align-items: flex-start; gap: 12px; }
    .rs-config-actions { width: 100%; justify-content: flex-end; }
}

/* ---- Date Picker Modal ---- */
.rs-modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 9998; display: flex; align-items: center; justify-content: center; }
.rs-modal { background: #fff; border-radius: 12px; padding: 24px; max-width: 420px; width: 90%; box-shadow: 0 8px 32px rgba(0,0,0,0.2); }
.rs-modal h3 { margin: 0 0 16px; font-size: 18px; color: #333; }
.rs-date-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 16px; }
.rs-date-btn { padding: 12px 8px; border: 2px solid var(--rs-border); border-radius: 8px; background: #fff; cursor: pointer; text-align: center; font-size: 13px; transition: all 0.15s; }
.rs-date-btn:hover { border-color: var(--rs-primary); background: var(--rs-primary-light); }
.rs-date-btn strong { display: block; font-size: 15px; color: #333; margin-bottom: 2px; }
.rs-date-btn span { color: #666; font-size: 12px; }

/* ---- Print ---- */
@media print {
    * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
    .rs-table { border-collapse: collapse !important; }
    .rs-th, .rs-td { border: 1px solid #ddd !important; }
    .rs-org-page { page-break-inside: avoid; }
}
    </style>
</head>
<body>
<div class="rs-root" id="rsApp">
    <div class="rs-toast-container" id="rsToastContainer"></div>
    <div id="rsContent">
        <div class="rs-loading"><div class="rs-spinner"></div><p>Loading...</p></div>
    </div>
</div>

<script>
(function() {
    "use strict";

    // =====================================================================
    // STATE
    // =====================================================================
    var state = {
        mode: 'landing',
        configs: [],
        editConfig: null,
        programs: [],
        divisions: [],
        searchResults: [],
        searchTimeout: null,
        previewHtml: '',
        previewOrgCount: 0,
        previewTotalMembers: 0,
        selectedConfigId: null,
        filtersLoaded: false,
        loading: false,
        editingDsIndex: -1,
        discoveredFields: null,
        discoveringFields: false,
        showDatePicker: false,
        reportDate: '',
        reportDateIso: ''
    };

    var COLOR_PRESETS = [
        {bg: '#fff3cd', text: '#e65100'},
        {bg: '#ffebee', text: '#c62828'},
        {bg: '#e3f2fd', text: '#1565c0'},
        {bg: '#e8f5e9', text: '#2e7d32'},
        {bg: '#f3e5f5', text: '#7b1fa2'},
        {bg: '#fff8e1', text: '#f57f17'},
        {bg: '#e0f7fa', text: '#00695c'},
        {bg: '#fce4ec', text: '#880e4f'}
    ];

    // Script path for AJAX
    var scriptPath = (function() {
        var p = window.location.pathname;
        if (p.indexOf('/PyScriptForm/') > -1) return p;
        return p.replace('/PyScript/', '/PyScriptForm/');
    })();

    // =====================================================================
    // UTILITIES
    // =====================================================================
    function escHtml(s) {
        if (!s) return '';
        var d = document.createElement('div');
        d.appendChild(document.createTextNode(s));
        return d.innerHTML;
    }

    function escAttr(s) {
        return escHtml(s).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    function extractJson(text) {
        text = (text || '').trim();
        var start = text.indexOf('{');
        var end = text.lastIndexOf('}');
        if (start === -1 || end === -1) {
            start = text.indexOf('[');
            end = text.lastIndexOf(']');
        }
        if (start >= 0 && end > start) {
            return text.substring(start, end + 1);
        }
        return text;
    }

    function ajax(action, params, callback) {
        var data = 'action=' + encodeURIComponent(action);
        if (params) {
            for (var key in params) {
                if (params.hasOwnProperty(key)) {
                    data += '&' + encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
                }
            }
        }
        var xhr = new XMLHttpRequest();
        xhr.open('POST', scriptPath, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var json = JSON.parse(extractJson(xhr.responseText));
                        callback(null, json);
                    } catch (e) {
                        callback('Invalid response: ' + e.message, null);
                    }
                } else {
                    callback('HTTP ' + xhr.status, null);
                }
            }
        };
        xhr.send(data);
    }

    function showToast(msg, type) {
        type = type || 'info';
        var container = document.getElementById('rsToastContainer');
        var toast = document.createElement('div');
        toast.className = 'rs-toast rs-toast-' + type;
        toast.textContent = msg;
        container.appendChild(toast);
        setTimeout(function() { toast.classList.add('rs-show'); }, 10);
        setTimeout(function() {
            toast.classList.remove('rs-show');
            setTimeout(function() { container.removeChild(toast); }, 300);
        }, 3000);
    }

    function showLoading() {
        state.loading = true;
        var el = document.getElementById('rsContent');
        el.innerHTML = '<div class="rs-loading"><div class="rs-spinner"></div><p>Loading...</p></div>';
    }

    function defaultConfig() {
        return {
            id: '',
            name: '',
            sourceType: 'program_division',
            programId: 0,
            programName: '',
            divisionId: 0,
            divisionName: '',
            selectedOrgs: [],
            excludeOrgIds: [],
            columns: { age: true, gender: true, phone: false, email: false, subgroup: false, memberType: false },
            nameLineColumns: { age: true, gender: true },
            dataSources: [],
            sortBy: 'name',
            layout: 'two_column',
            title: '',
            showFooter: true,
            onlyWithMeeting: false
        };
    }

    function migrateConfig(c) {
        if (c.dataSources) return c;
        var ds = [];
        if (c.includeMedical) {
            var labels = c.medicalQuestionLabels || [];
            for (var i = 0; i < labels.length; i++) {
                ds.push({id: 'ds_mig_med_' + i, sourceType: 'regQuestion', label: 'Medical', colorBg: '#fff3cd', colorText: '#e65100', questionPattern: labels[i], excludeNone: true, matchAnswer: '', showAsWarning: false, enabled: true});
            }
        }
        if (c.includePhotos && c.photoQuestionLabel) {
            ds.push({id: 'ds_mig_photo_0', sourceType: 'regQuestion', label: 'NO PHOTOS', colorBg: '#ffebee', colorText: '#c62828', questionPattern: c.photoQuestionLabel, matchAnswer: '%No%', showAsWarning: true, excludeNone: false, enabled: true});
        }
        var rr = c.recregFields || {};
        var rrMap = [
            ['emergencyContact', 'EC', '#e3f2fd', '#1565c0'],
            ['medicalAllergies', 'Allergies', '#fff3cd', '#e65100'],
            ['doctor', 'Dr', '#e8f5e9', '#2e7d32'],
            ['insurance', 'Ins', '#f3e5f5', '#7b1fa2']
        ];
        for (var j = 0; j < rrMap.length; j++) {
            if (rr[rrMap[j][0]]) {
                ds.push({id: 'ds_mig_rr_' + rrMap[j][0], sourceType: 'recreg', label: rrMap[j][1], colorBg: rrMap[j][2], colorText: rrMap[j][3], recregGroup: rrMap[j][0], enabled: true});
            }
        }
        c.dataSources = ds;
        return c;
    }

    // =====================================================================
    // RENDER ROUTER
    // =====================================================================
    function render() {
        var el = document.getElementById('rsContent');
        state.loading = false;
        switch (state.mode) {
            case 'landing': el.innerHTML = renderLanding(); break;
            case 'admin': el.innerHTML = renderAdmin(); break;
            case 'admin_edit': el.innerHTML = renderAdminEdit(); break;
            case 'generate_pick': el.innerHTML = renderGeneratePick(); break;
            case 'generate_preview': el.innerHTML = renderGeneratePreview(); break;
            default: el.innerHTML = renderLanding();
        }
        // Date picker modal (overlay, outside main content)
        var existing = document.getElementById('rsDateModal');
        if (existing) existing.parentNode.removeChild(existing);
        if (state.showDatePicker) {
            var m = document.createElement('div');
            m.id = 'rsDateModal';
            m.innerHTML = renderDatePickerModal();
            document.querySelector('.rs-root').appendChild(m);
        }
    }

    // =====================================================================
    // LANDING
    // =====================================================================
    function renderLanding() {
        return '<div class="rs-landing">' +
            '<div class="rs-landing-title">Rollsheet Generator</div>' +
            '<div class="rs-landing-subtitle">Create configurable rollsheets for any ministry area</div>' +
            '<div class="rs-landing-buttons">' +
                '<div class="rs-landing-btn rs-admin-btn" onclick="RSApp.goAdmin()">' +
                    '<i>&#9881;</i><span>Admin Setup</span>' +
                '</div>' +
                '<div class="rs-landing-btn rs-gen-btn" onclick="RSApp.goGenerate()">' +
                    '<i>&#9997;</i><span>Generate Rollsheets</span>' +
                '</div>' +
            '</div>' +
        '</div>';
    }

    // =====================================================================
    // ADMIN LIST
    // =====================================================================
    function renderAdmin() {
        var h = '<button class="rs-back-btn" onclick="RSApp.goLanding()">&#8592; Back</button>';
        h += '<div class="rs-flex rs-items-center rs-justify-between rs-mb-16">';
        h += '<h2>Rollsheet Configs</h2>';
        h += '<button class="rs-btn rs-btn-primary" onclick="RSApp.newConfig()">+ New Config</button>';
        h += '</div>';

        if (!state.configs.length) {
            h += '<div class="rs-empty"><p>No configs yet. Create one to get started.</p></div>';
        } else {
            for (var i = 0; i < state.configs.length; i++) {
                var c = state.configs[i];
                var sourceLabel = c.sourceType === 'specific_orgs'
                    ? (c.selectedOrgs ? c.selectedOrgs.length : 0) + ' specific org(s)'
                    : (c.programName || 'Program') + ' / ' + (c.divisionName || 'All Divisions');
                h += '<div class="rs-config-card">';
                h += '<div>';
                h += '<div class="rs-config-name">' + escHtml(c.name || 'Untitled') + '</div>';
                h += '<div class="rs-config-meta">' + escHtml(sourceLabel) + ' &bull; Layout: ' + escHtml(c.layout || 'two_column') + ' &bull; Updated: ' + escHtml(c.updatedAt || '') + '</div>';
                h += '</div>';
                h += '<div class="rs-config-actions">';
                h += '<button class="rs-btn rs-btn-sm rs-btn-outline" onclick="RSApp.editConfig(\\'' + escAttr(c.id) + '\\')">Edit</button>';
                h += '<button class="rs-btn rs-btn-sm rs-btn-outline" onclick="RSApp.duplicateConfig(\\'' + escAttr(c.id) + '\\')">Duplicate</button>';
                h += '<button class="rs-btn rs-btn-sm rs-btn-danger" onclick="RSApp.deleteConfig(\\'' + escAttr(c.id) + '\\')">Delete</button>';
                h += '</div>';
                h += '</div>';
            }
        }
        return h;
    }

    // =====================================================================
    // ADMIN EDIT
    // =====================================================================
    function renderAdminEdit() {
        var c = state.editConfig;
        if (!c) return '<p>No config to edit.</p>';

        var h = '<button class="rs-back-btn" onclick="RSApp.goAdminList()">&#8592; Back to Configs</button>';
        h += '<h2 class="rs-mb-16">' + (c.id ? 'Edit Config' : 'New Config') + '</h2>';

        // Config Name
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Config Name</label>';
        h += '<input class="rs-input" id="rsConfigName" value="' + escAttr(c.name) + '" placeholder="e.g., VBS 2026 AM Rollsheet">';
        h += '</div>';

        // Print Title
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Print Title (shown on each page)</label>';
        h += '<input class="rs-input" id="rsConfigTitle" value="' + escAttr(c.title) + '" placeholder="e.g., VBS 2026 - AM Session">';
        h += '</div>';

        // Source Type
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Data Source</label>';
        h += '<div class="rs-flex rs-gap-16">';
        h += '<label class="rs-checkbox-item"><input type="radio" name="rsSourceType" value="program_division" ' + (c.sourceType !== 'specific_orgs' ? 'checked' : '') + ' onchange="RSApp.setSourceType(\\'program_division\\')"> Program / Division</label>';
        h += '<label class="rs-checkbox-item"><input type="radio" name="rsSourceType" value="specific_orgs" ' + (c.sourceType === 'specific_orgs' ? 'checked' : '') + ' onchange="RSApp.setSourceType(\\'specific_orgs\\')"> Specific Organizations</label>';
        h += '</div>';
        h += '</div>';

        // Program/Division dropdowns (when source = program_division)
        h += '<div id="rsProgramDivSection" style="' + (c.sourceType === 'specific_orgs' ? 'display:none' : '') + '">';
        h += '<div class="rs-form-row rs-form-group">';
        h += '<div>';
        h += '<label class="rs-label">Program</label>';
        h += '<select class="rs-select" id="rsProgramSelect" onchange="RSApp.onProgramChange()">';
        h += '<option value="">-- All Programs --</option>';
        for (var pi = 0; pi < state.programs.length; pi++) {
            var prog = state.programs[pi];
            var sel = (prog.id == c.programId) ? ' selected' : '';
            h += '<option value="' + prog.id + '"' + sel + '>' + escHtml(prog.name) + '</option>';
        }
        h += '</select>';
        h += '</div>';
        h += '<div>';
        h += '<label class="rs-label">Division</label>';
        h += '<select class="rs-select" id="rsDivisionSelect" onchange="RSApp.onDivisionChange()">';
        h += '<option value="">-- All Divisions --</option>';
        var filteredDivs = state.divisions;
        if (c.programId) {
            filteredDivs = state.divisions.filter(function(d) { return d.progId == c.programId; });
        }
        for (var di = 0; di < filteredDivs.length; di++) {
            var div = filteredDivs[di];
            var dsel = (div.id == c.divisionId) ? ' selected' : '';
            h += '<option value="' + div.id + '"' + dsel + '>' + escHtml(div.name) + '</option>';
        }
        h += '</select>';
        h += '</div>';
        h += '</div>';

        // Exclude org IDs
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Exclude Organization IDs (comma-separated, optional)</label>';
        h += '<input class="rs-input" id="rsExcludeOrgs" value="' + escAttr((c.excludeOrgIds || []).join(', ')) + '" placeholder="e.g., 2950, 2967, 2968">';
        h += '</div>';
        h += '</div>';

        // Specific orgs section
        h += '<div id="rsSpecificOrgsSection" style="' + (c.sourceType !== 'specific_orgs' ? 'display:none' : '') + '">';
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Search and Add Organizations</label>';
        h += '<div class="rs-search-box">';
        h += '<input class="rs-input" id="rsOrgSearch" placeholder="Search by name or ID..." oninput="RSApp.onOrgSearch()">';
        h += '<div id="rsOrgSearchResults" class="rs-search-results" style="display:none"></div>';
        h += '</div>';
        h += '</div>';

        // Selected orgs list
        h += '<div id="rsSelectedOrgsList">';
        if (c.selectedOrgs && c.selectedOrgs.length) {
            for (var si = 0; si < c.selectedOrgs.length; si++) {
                var so = c.selectedOrgs[si];
                h += '<div class="rs-selected-org-item">';
                h += '<span>' + escHtml(so.orgName) + ' (ID: ' + so.orgId + ', Members: ' + (so.memberCount || '?') + ')</span>';
                h += '<span class="rs-remove-org" onclick="RSApp.removeOrg(' + si + ')">&#10005;</span>';
                h += '</div>';
            }
        } else {
            h += '<div class="rs-text-muted rs-text-sm">No organizations selected yet.</div>';
        }
        h += '</div>';
        h += '</div>';

        // Column toggles
        h += '<div class="rs-panel rs-mt-16">';
        h += '<div class="rs-panel-header">Info to Display</div>';
        h += '<div class="rs-panel-body">';
        h += '<div class="rs-text-sm rs-text-muted rs-mb-12">Select which info to show. "On name line" places it next to the name instead of on a separate line below.</div>';
        var nlc = c.nameLineColumns || {};
        var colDefs = [
            {key: 'age', label: 'Age'}, {key: 'gender', label: 'Gender'},
            {key: 'phone', label: 'Phone'}, {key: 'email', label: 'Email'},
            {key: 'subgroup', label: 'Sub-Group'}, {key: 'memberType', label: 'Member Type'}
        ];
        h += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
        h += '<tr style="border-bottom:1px solid var(--rs-border);"><th style="text-align:left;padding:4px 8px;">Field</th><th style="text-align:center;padding:4px 8px;">Show</th><th style="text-align:center;padding:4px 8px;">On name line</th></tr>';
        for (var ci = 0; ci < colDefs.length; ci++) {
            var col = colDefs[ci];
            var checked = c.columns[col.key] ? ' checked' : '';
            var nlChecked = nlc[col.key] ? ' checked' : '';
            h += '<tr style="border-bottom:1px solid var(--rs-border);">';
            h += '<td style="padding:4px 8px;">' + col.label + '</td>';
            h += '<td style="text-align:center;padding:4px 8px;"><input type="checkbox" id="rsCol_' + col.key + '"' + checked + '></td>';
            h += '<td style="text-align:center;padding:4px 8px;"><input type="checkbox" id="rsNL_' + col.key + '"' + nlChecked + '></td>';
            h += '</tr>';
        }
        h += '</table>';
        h += '</div></div>';

        // Custom Data Sources panel (replaces RecReg + Medical/Photo panels)
        h += renderDsPanel(c);

        // Layout & Sort
        h += '<div class="rs-panel rs-mt-16">';
        h += '<div class="rs-panel-header">Layout & Sorting</div>';
        h += '<div class="rs-panel-body">';

        h += '<div class="rs-form-row">';
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Layout</label>';
        h += '<select class="rs-select" id="rsLayout">';
        h += '<option value="two_column"' + (c.layout === 'two_column' ? ' selected' : '') + '>Two Column</option>';
        h += '<option value="three_column"' + (c.layout === 'three_column' ? ' selected' : '') + '>Three Column</option>';
        h += '<option value="single_column"' + (c.layout === 'single_column' ? ' selected' : '') + '>Single Column</option>';
        h += '<option value="table"' + (c.layout === 'table' ? ' selected' : '') + '>Full-Width Table</option>';
        h += '</select>';
        h += '</div>';

        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Sort By</label>';
        h += '<select class="rs-select" id="rsSortBy">';
        h += '<option value="name"' + (c.sortBy === 'name' ? ' selected' : '') + '>Name</option>';
        h += '<option value="age"' + (c.sortBy === 'age' ? ' selected' : '') + '>Age</option>';
        h += '<option value="gender"' + (c.sortBy === 'gender' ? ' selected' : '') + '>Gender</option>';
        h += '</select>';
        h += '</div>';
        h += '</div>';

        h += '<div class="rs-toggle-row">';
        h += '<span class="rs-toggle-label">Show Footer (totals, signature line)</span>';
        h += '<label class="rs-toggle"><input type="checkbox" id="rsShowFooter"' + (c.showFooter ? ' checked' : '') + '><span class="rs-toggle-slider"></span></label>';
        h += '</div>';

        h += '<div class="rs-toggle-row">';
        h += '<span class="rs-toggle-label">Only include orgs with a meeting on the report date</span>';
        h += '<label class="rs-toggle"><input type="checkbox" id="rsOnlyWithMeeting"' + (c.onlyWithMeeting ? ' checked' : '') + '><span class="rs-toggle-slider"></span></label>';
        h += '</div>';

        h += '</div></div>';

        // Save button
        h += '<div class="rs-flex rs-gap-12 rs-mt-16">';
        h += '<button class="rs-btn rs-btn-primary" onclick="RSApp.saveConfig()">Save Config</button>';
        h += '<button class="rs-btn rs-btn-outline" onclick="RSApp.goAdminList()">Cancel</button>';
        h += '</div>';

        return h;
    }

    // =====================================================================
    // GENERATE PICK
    // =====================================================================
    function renderGeneratePick() {
        var h = '<button class="rs-back-btn" onclick="RSApp.goLanding()">&#8592; Back</button>';
        h += '<h2 class="rs-mb-16">Select a Config to Generate</h2>';

        if (!state.configs.length) {
            h += '<div class="rs-empty"><p>No configs available. Go to Admin Setup to create one first.</p></div>';
        } else {
            for (var i = 0; i < state.configs.length; i++) {
                var c = state.configs[i];
                var sourceLabel = c.sourceType === 'specific_orgs'
                    ? (c.selectedOrgs ? c.selectedOrgs.length : 0) + ' specific org(s)'
                    : (c.programName || 'Program') + ' / ' + (c.divisionName || 'All Divisions');
                var isSelected = (state.selectedConfigId === c.id);
                h += '<div class="rs-config-card' + (isSelected ? ' selected' : '') + '" onclick="RSApp.selectConfig(\\'' + escAttr(c.id) + '\\')" style="' + (isSelected ? 'border-color:var(--rs-primary);background:var(--rs-primary-light);' : '') + '">';
                h += '<div>';
                h += '<div class="rs-config-name">' + escHtml(c.name || 'Untitled') + '</div>';
                h += '<div class="rs-config-meta">' + escHtml(sourceLabel) + ' &bull; ' + escHtml(c.layout || 'two_column') + '</div>';
                h += '</div>';
                if (isSelected) {
                    h += '<button class="rs-btn rs-btn-success" onclick="event.stopPropagation();RSApp.generateRollsheet()">Generate</button>';
                }
                h += '</div>';
            }
        }
        return h;
    }

    // =====================================================================
    // GENERATE PREVIEW
    // =====================================================================
    function renderGeneratePreview() {
        var h = '<button class="rs-back-btn" onclick="RSApp.goGenerate()">&#8592; Back to Config Selection</button>';
        h += '<div class="rs-flex rs-items-center rs-justify-between rs-mb-16">';
        h += '<h2>Rollsheet Preview</h2>';
        h += '<button class="rs-btn rs-btn-primary" onclick="RSApp.printRollsheets()">Print Rollsheets</button>';
        h += '</div>';

        // Stats
        h += '<div class="rs-stats">';
        h += '<div class="rs-stat-card"><div class="rs-stat-num">' + state.previewOrgCount + '</div><div class="rs-stat-label">Organizations</div></div>';
        h += '<div class="rs-stat-card"><div class="rs-stat-num">' + state.previewTotalMembers + '</div><div class="rs-stat-label">Total Members</div></div>';
        h += '</div>';

        // Preview
        h += '<div class="rs-preview-area">' + state.previewHtml + '</div>';

        return h;
    }

    // =====================================================================
    // DATA SOURCES PANEL
    // =====================================================================
    function renderDsPanel(c) {
        var dsList = c.dataSources || [];
        var h = '<div class="rs-panel rs-mt-16">';
        h += '<div class="rs-panel-header"><span>Custom Data Sources</span>';
        h += '<button class="rs-btn rs-btn-sm rs-btn-outline" onclick="RSApp.discoverFields()" ' + (state.discoveringFields ? 'disabled' : '') + '>' + (state.discoveringFields ? 'Discovering...' : 'Discover Fields') + '</button>';
        h += '</div>';
        h += '<div class="rs-panel-body">';
        h += '<div class="rs-text-sm rs-text-muted rs-mb-12">Add registration questions, extra values, or built-in registration data to display below each person\\'s name as colored badges.</div>';

        // Field picker dropdown
        h += '<div class="rs-form-group">';
        h += renderDsFieldPicker();
        h += '</div>';

        // Existing data source cards
        if (!dsList.length) {
            h += '<div class="rs-text-muted rs-text-sm rs-text-center" style="padding:16px;">No data sources configured. Use the dropdown above or click "Discover Fields" to add sources.</div>';
        } else {
            for (var i = 0; i < dsList.length; i++) {
                if (state.editingDsIndex === i) {
                    h += renderDsEditor(dsList[i], i);
                } else {
                    h += renderDsCard(dsList[i], i);
                }
            }
        }

        // Show editor for new item
        if (state.editingDsIndex === -2) {
            h += renderDsEditor(dsList[dsList.length - 1], dsList.length - 1);
        }

        h += '</div></div>';
        return h;
    }

    function renderDsCard(ds, index) {
        var bg = ds.colorBg || '#f0f0f0';
        var color = ds.colorText || '#333';
        var label = escHtml(ds.label || '');
        var typeLabel = ds.sourceType === 'regQuestion' ? 'Reg Question' : ds.sourceType === 'extraValue' ? 'Extra Value' : 'RecReg';
        var detail = '';
        if (ds.sourceType === 'regQuestion') detail = 'Pattern: ' + escHtml(ds.questionPattern || '');
        else if (ds.sourceType === 'extraValue') detail = 'Field: ' + escHtml(ds.evField || '') + ' (' + escHtml(ds.evType || 'text') + ')';
        else if (ds.sourceType === 'recreg') detail = 'Group: ' + escHtml(ds.recregGroup || '');

        var modeLabel = ds.displayMode === 'append' ? ' [+append]' : '';

        var h = '<div class="rs-ds-card">';
        h += '<div style="display:flex;align-items:center;gap:8px;flex:1;min-width:0;">';
        h += '<span class="rs-ds-badge" style="background:' + bg + ';color:' + color + ';">' + label + '</span>';
        h += '<span class="rs-ds-card-info">' + typeLabel + ' - ' + detail + modeLabel + '</span>';
        h += '</div>';
        h += '<div class="rs-ds-card-actions">';
        if (index > 0) h += '<button onclick="RSApp.moveDsUp(' + index + ')" title="Move up">&#9650;</button>';
        if (index < (state.editConfig.dataSources || []).length - 1) h += '<button onclick="RSApp.moveDsDown(' + index + ')" title="Move down">&#9660;</button>';
        h += '<button onclick="RSApp.editDs(' + index + ')" title="Edit">&#9998;</button>';
        h += '<button onclick="RSApp.removeDs(' + index + ')" title="Remove" style="color:var(--rs-danger);">&#10005;</button>';
        h += '</div>';
        h += '</div>';
        return h;
    }

    function renderDsEditor(ds, index) {
        var h = '<div class="rs-ds-editor">';
        h += '<div class="rs-form-row rs-mb-12">';

        // Label
        h += '<div class="rs-form-group" style="flex:1;">';
        h += '<label class="rs-label">Short Label</label>';
        h += '<input class="rs-input" id="rsDsLabel" value="' + escAttr(ds.label || '') + '" placeholder="e.g., Allergies, NO PHOTOS, EC">';
        h += '</div>';

        // Source type (read-only display)
        h += '<div class="rs-form-group" style="flex:1;">';
        h += '<label class="rs-label">Source Type</label>';
        h += '<select class="rs-select" id="rsDsSourceType" onchange="RSApp.onDsTypeChange()">';
        h += '<option value="regQuestion"' + (ds.sourceType === 'regQuestion' ? ' selected' : '') + '>Registration Question</option>';
        h += '<option value="extraValue"' + (ds.sourceType === 'extraValue' ? ' selected' : '') + '>Extra Value</option>';
        h += '<option value="recreg"' + (ds.sourceType === 'recreg' ? ' selected' : '') + '>RecReg Built-in</option>';
        h += '</select>';
        h += '</div>';
        h += '</div>';

        // Colors
        h += '<div class="rs-form-group rs-mb-12">';
        h += '<label class="rs-label">Colors</label>';
        h += '<div class="rs-color-presets">';
        for (var ci = 0; ci < COLOR_PRESETS.length; ci++) {
            var cp = COLOR_PRESETS[ci];
            var isActive = (cp.bg === ds.colorBg && cp.text === ds.colorText) ? ' rs-active' : '';
            h += '<div class="rs-color-dot' + isActive + '" style="background:' + cp.bg + ';color:' + cp.text + ';" onclick="RSApp.pickPresetColor(' + ci + ')" title="' + cp.bg + '">Aa</div>';
        }
        h += '</div>';
        h += '<div class="rs-color-inputs">';
        h += '<label style="font-size:12px;">BG:</label><input type="color" id="rsDsColorBg" value="' + (ds.colorBg || '#f0f0f0') + '">';
        h += '<label style="font-size:12px;">Text:</label><input type="color" id="rsDsColorText" value="' + (ds.colorText || '#333333') + '">';
        h += '<span class="rs-ds-badge" style="background:' + (ds.colorBg || '#f0f0f0') + ';color:' + (ds.colorText || '#333') + ';">Preview</span>';
        h += '</div>';
        h += '</div>';

        // Type-specific fields
        h += '<div id="rsDsTypeFields">';
        h += renderDsTypeFields(ds);
        h += '</div>';

        // Display mode
        h += '<div class="rs-form-group rs-mb-12">';
        h += '<label class="rs-label">Display Mode</label>';
        h += '<select class="rs-select" id="rsDsDisplayMode">';
        h += '<option value="own_line"' + (ds.displayMode !== 'append' ? ' selected' : '') + '>Own line</option>';
        h += '<option value="append"' + (ds.displayMode === 'append' ? ' selected' : '') + '>Append to previous</option>';
        h += '</select>';
        h += '<div class="rs-help-text">"Own line" puts this badge on a new row. "Append to previous" places it inline next to the badge above.</div>';
        h += '</div>';

        // Action buttons
        h += '<div class="rs-flex rs-gap-8 rs-mt-12">';
        h += '<button class="rs-btn rs-btn-sm rs-btn-primary" onclick="RSApp.saveDsEdit(' + index + ')">Save</button>';
        h += '<button class="rs-btn rs-btn-sm rs-btn-outline" onclick="RSApp.cancelDsEdit()">Cancel</button>';
        h += '</div>';
        h += '</div>';
        return h;
    }

    function renderDsTypeFields(ds) {
        var h = '';
        if (ds.sourceType === 'regQuestion') {
            h += '<div class="rs-form-group rs-mb-8">';
            h += '<label class="rs-label">Question Pattern (SQL LIKE)</label>';
            h += '<input class="rs-input" id="rsDsQuestionPattern" value="' + escAttr(ds.questionPattern || '') + '" placeholder="%Allergies%">';
            h += '<div class="rs-help-text">Use % as wildcard. e.g., %Allergies% matches any question containing "Allergies".</div>';
            h += '</div>';
            h += '<div class="rs-form-group rs-mb-8">';
            h += '<label class="rs-label">Match Answer (optional, SQL LIKE)</label>';
            h += '<input class="rs-input" id="rsDsMatchAnswer" value="' + escAttr(ds.matchAnswer || '') + '" placeholder="Leave blank to show all answers">';
            h += '<div class="rs-help-text">Filter to only show when answer matches this pattern. e.g., %No% for photo opt-outs.</div>';
            h += '</div>';
            h += '<div class="rs-checkbox-grid">';
            h += '<label class="rs-checkbox-item"><input type="checkbox" id="rsDsExcludeNone"' + (ds.excludeNone ? ' checked' : '') + '> Exclude "None" type answers</label>';
            h += '<label class="rs-checkbox-item"><input type="checkbox" id="rsDsShowAsWarning"' + (ds.showAsWarning ? ' checked' : '') + '> Show as warning badge (label only, no answer text)</label>';
            h += '</div>';
        } else if (ds.sourceType === 'extraValue') {
            h += '<div class="rs-form-row rs-mb-8">';
            h += '<div class="rs-form-group" style="flex:2;">';
            h += '<label class="rs-label">Extra Value Field Name</label>';
            h += '<input class="rs-input" id="rsDsEvField" value="' + escAttr(ds.evField || '') + '" placeholder="e.g., VBSShirtSize">';
            h += '</div>';
            h += '<div class="rs-form-group" style="flex:1;">';
            h += '<label class="rs-label">Field Type</label>';
            h += '<select class="rs-select" id="rsDsEvType">';
            var evTypes = ['text', 'bit', 'int', 'date', 'code'];
            for (var ei = 0; ei < evTypes.length; ei++) {
                h += '<option value="' + evTypes[ei] + '"' + (ds.evType === evTypes[ei] ? ' selected' : '') + '>' + evTypes[ei] + '</option>';
            }
            h += '</select>';
            h += '</div>';
            h += '</div>';
        } else if (ds.sourceType === 'recreg') {
            h += '<div class="rs-form-group rs-mb-8">';
            h += '<label class="rs-label">RecReg Group</label>';
            h += '<select class="rs-select" id="rsDsRecregGroup">';
            var rrGroups = [
                ['emergencyContact', 'Emergency Contact (name + phone)'],
                ['medicalAllergies', 'Allergies & Medical Conditions'],
                ['doctor', 'Doctor (name + phone)'],
                ['insurance', 'Insurance (company + policy #)']
            ];
            for (var ri = 0; ri < rrGroups.length; ri++) {
                h += '<option value="' + rrGroups[ri][0] + '"' + (ds.recregGroup === rrGroups[ri][0] ? ' selected' : '') + '>' + rrGroups[ri][1] + '</option>';
            }
            h += '</select>';
            h += '</div>';
        }
        return h;
    }

    function renderDsFieldPicker() {
        var h = '<select class="rs-select" id="rsDsFieldPicker" onchange="RSApp.addDsFromPicker()">';
        h += '<option value="">+ Add a data source...</option>';

        var df = state.discoveredFields;
        if (df) {
            if (df.regQuestions && df.regQuestions.length) {
                h += '<optgroup label="Registration Questions">';
                for (var i = 0; i < df.regQuestions.length; i++) {
                    var rq = df.regQuestions[i];
                    var pat = '%' + rq.label + '%';
                    var shortLabel = rq.label.length > 20 ? rq.label.substring(0, 20) + '...' : rq.label;
                    h += '<option value="regQuestion|' + escAttr(pat) + '|' + escAttr(shortLabel) + '">' + escHtml(rq.label) + '</option>';
                }
                h += '</optgroup>';
            }
            if (df.extraValues && df.extraValues.length) {
                h += '<optgroup label="Extra Values">';
                for (var j = 0; j < df.extraValues.length; j++) {
                    var ev = df.extraValues[j];
                    h += '<option value="extraValue|' + escAttr(ev.field) + '|' + escAttr(ev.fieldType) + '|' + escAttr(ev.field) + '">' + escHtml(ev.field) + ' (' + ev.fieldType + ')</option>';
                }
                h += '</optgroup>';
            }
        }

        // RecReg always available (static)
        h += '<optgroup label="Built-in Registration (RecReg)">';
        h += '<option value="recreg|emergencyContact|EC">Emergency Contact</option>';
        h += '<option value="recreg|medicalAllergies|Allergies">Allergies & Medical</option>';
        h += '<option value="recreg|doctor|Dr">Doctor</option>';
        h += '<option value="recreg|insurance|Ins">Insurance</option>';
        h += '</optgroup>';

        // Custom pattern options
        h += '<optgroup label="Custom Pattern">';
        h += '<option value="custom|regQuestion">Custom Registration Question Pattern...</option>';
        h += '<option value="custom|extraValue">Custom Extra Value Field...</option>';
        h += '</optgroup>';

        h += '</select>';
        return h;
    }

    function discoverFields() {
        // Determine org IDs from current config
        var c = state.editConfig;
        if (!c) return;

        var orgIds = '';
        if (c.sourceType === 'specific_orgs' && c.selectedOrgs && c.selectedOrgs.length) {
            orgIds = c.selectedOrgs.map(function(o) { return o.orgId; }).join(',');
        } else if (c.programId || c.divisionId) {
            // Need to collect from form first
            var progEl = document.getElementById('rsProgramSelect');
            var divEl = document.getElementById('rsDivisionSelect');
            var pid = progEl ? parseInt(progEl.value) || 0 : c.programId || 0;
            var did = divEl ? parseInt(divEl.value) || 0 : c.divisionId || 0;
            if (!pid && !did) {
                showToast('Select a program/division or orgs first', 'danger');
                return;
            }
            // Query org IDs via search, then discover
            var searchParams = {};
            if (pid) searchParams.program_id = pid;
            if (did) searchParams.division_id = did;
            state.discoveringFields = true;
            render();
            ajax('search_involvements', searchParams, function(err, data) {
                if (!err && data && data.success && data.involvements) {
                    orgIds = data.involvements.map(function(inv) { return inv.orgId; }).join(',');
                    if (orgIds) {
                        ajax('discover_fields', {org_ids: orgIds}, function(err2, data2) {
                            state.discoveringFields = false;
                            if (!err2 && data2 && data2.success) {
                                state.discoveredFields = {regQuestions: data2.regQuestions, extraValues: data2.extraValues, recregOptions: data2.recregOptions};
                                showToast('Found ' + (data2.regQuestions || []).length + ' questions, ' + (data2.extraValues || []).length + ' extra values', 'success');
                            } else {
                                showToast('Error discovering fields', 'danger');
                            }
                            render();
                        });
                    } else {
                        state.discoveringFields = false;
                        showToast('No organizations found', 'danger');
                        render();
                    }
                } else {
                    state.discoveringFields = false;
                    showToast('Error finding orgs', 'danger');
                    render();
                }
            });
            return;
        } else {
            showToast('Select a program/division or orgs first', 'danger');
            return;
        }

        state.discoveringFields = true;
        render();
        ajax('discover_fields', {org_ids: orgIds}, function(err, data) {
            state.discoveringFields = false;
            if (!err && data && data.success) {
                state.discoveredFields = {regQuestions: data.regQuestions, extraValues: data.extraValues, recregOptions: data.recregOptions};
                showToast('Found ' + (data.regQuestions || []).length + ' questions, ' + (data.extraValues || []).length + ' extra values', 'success');
            } else {
                showToast('Error discovering fields: ' + (data ? data.message : err), 'danger');
            }
            render();
        });
    }

    function addDsFromPicker() {
        var picker = document.getElementById('rsDsFieldPicker');
        if (!picker || !picker.value) return;
        var val = picker.value;
        var parts = val.split('|');
        var c = state.editConfig;
        if (!c.dataSources) c.dataSources = [];

        var newDs = {
            id: 'ds_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
            enabled: true,
            colorBg: COLOR_PRESETS[c.dataSources.length % COLOR_PRESETS.length].bg,
            colorText: COLOR_PRESETS[c.dataSources.length % COLOR_PRESETS.length].text
        };

        if (parts[0] === 'regQuestion') {
            newDs.sourceType = 'regQuestion';
            newDs.questionPattern = parts[1] || '';
            newDs.label = parts[2] || 'Question';
            newDs.excludeNone = true;
            newDs.matchAnswer = '';
            newDs.showAsWarning = false;
        } else if (parts[0] === 'extraValue') {
            newDs.sourceType = 'extraValue';
            newDs.evField = parts[1] || '';
            newDs.evType = parts[2] || 'text';
            newDs.label = parts[3] || parts[1] || 'EV';
        } else if (parts[0] === 'recreg') {
            newDs.sourceType = 'recreg';
            newDs.recregGroup = parts[1] || 'emergencyContact';
            newDs.label = parts[2] || 'RR';
        } else if (parts[0] === 'custom') {
            newDs.sourceType = parts[1] || 'regQuestion';
            newDs.label = '';
            if (newDs.sourceType === 'regQuestion') {
                newDs.questionPattern = '';
                newDs.excludeNone = true;
                newDs.matchAnswer = '';
                newDs.showAsWarning = false;
            } else {
                newDs.evField = '';
                newDs.evType = 'text';
            }
        }

        c.dataSources.push(newDs);
        state.editingDsIndex = c.dataSources.length - 1;
        picker.value = '';
        render();
    }

    function editDs(index) {
        state.editingDsIndex = index;
        render();
    }

    function saveDsEdit(index) {
        var c = state.editConfig;
        var ds = c.dataSources[index];
        ds.label = (document.getElementById('rsDsLabel') || {}).value || ds.label || '';
        ds.sourceType = (document.getElementById('rsDsSourceType') || {}).value || ds.sourceType;
        ds.colorBg = (document.getElementById('rsDsColorBg') || {}).value || ds.colorBg;
        ds.colorText = (document.getElementById('rsDsColorText') || {}).value || ds.colorText;
        ds.displayMode = (document.getElementById('rsDsDisplayMode') || {}).value || 'own_line';

        if (ds.sourceType === 'regQuestion') {
            ds.questionPattern = (document.getElementById('rsDsQuestionPattern') || {}).value || '';
            ds.matchAnswer = (document.getElementById('rsDsMatchAnswer') || {}).value || '';
            ds.excludeNone = document.getElementById('rsDsExcludeNone') ? document.getElementById('rsDsExcludeNone').checked : false;
            ds.showAsWarning = document.getElementById('rsDsShowAsWarning') ? document.getElementById('rsDsShowAsWarning').checked : false;
        } else if (ds.sourceType === 'extraValue') {
            ds.evField = (document.getElementById('rsDsEvField') || {}).value || '';
            ds.evType = (document.getElementById('rsDsEvType') || {}).value || 'text';
        } else if (ds.sourceType === 'recreg') {
            ds.recregGroup = (document.getElementById('rsDsRecregGroup') || {}).value || 'emergencyContact';
        }

        state.editingDsIndex = -1;
        render();
    }

    function cancelDsEdit() {
        var c = state.editConfig;
        // If was adding new (last item with no label), remove it
        if (state.editingDsIndex === -2 || (state.editingDsIndex >= 0 && c.dataSources[state.editingDsIndex] && !c.dataSources[state.editingDsIndex].label)) {
            // Only remove if label is empty (was just added)
            var idx = state.editingDsIndex === -2 ? c.dataSources.length - 1 : state.editingDsIndex;
            if (idx >= 0 && c.dataSources[idx] && !c.dataSources[idx].label) {
                c.dataSources.splice(idx, 1);
            }
        }
        state.editingDsIndex = -1;
        render();
    }

    function removeDs(index) {
        if (!confirm('Remove this data source?')) return;
        state.editConfig.dataSources.splice(index, 1);
        state.editingDsIndex = -1;
        render();
    }

    function moveDsUp(index) {
        var arr = state.editConfig.dataSources;
        if (index > 0) {
            var tmp = arr[index];
            arr[index] = arr[index - 1];
            arr[index - 1] = tmp;
            render();
        }
    }

    function moveDsDown(index) {
        var arr = state.editConfig.dataSources;
        if (index < arr.length - 1) {
            var tmp = arr[index];
            arr[index] = arr[index + 1];
            arr[index + 1] = tmp;
            render();
        }
    }

    function onDsTypeChange() {
        var st = (document.getElementById('rsDsSourceType') || {}).value || 'regQuestion';
        var ds = {sourceType: st, questionPattern: '', matchAnswer: '', excludeNone: true, showAsWarning: false, evField: '', evType: 'text', recregGroup: 'emergencyContact'};
        var el = document.getElementById('rsDsTypeFields');
        if (el) el.innerHTML = renderDsTypeFields(ds);
    }

    function pickPresetColor(presetIndex) {
        var cp = COLOR_PRESETS[presetIndex];
        var bgEl = document.getElementById('rsDsColorBg');
        var textEl = document.getElementById('rsDsColorText');
        if (bgEl) bgEl.value = cp.bg;
        if (textEl) textEl.value = cp.text;
        // Re-render editor to update preview and active dot
        var idx = state.editingDsIndex;
        if (idx >= 0 && state.editConfig.dataSources[idx]) {
            state.editConfig.dataSources[idx].colorBg = cp.bg;
            state.editConfig.dataSources[idx].colorText = cp.text;
            render();
        }
    }

    // =====================================================================
    // PRINT POPUP
    // =====================================================================
    function printRollsheets() {
        var printHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Rollsheets</title><style>';
        printHtml += '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important;}';
        printHtml += 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:20px;color:#333;}';
        printHtml += '.rs-org-page{margin-bottom:30px;}';
        printHtml += '.rs-org-header{text-align:center;margin-bottom:15px;}';
        printHtml += '.rs-org-title{font-size:22px;margin:0 0 4px;color:#333;}';
        printHtml += '.rs-org-name{font-size:18px;margin:0 0 8px;color:#007bff;}';
        printHtml += '.rs-org-meta{font-size:13px;color:#666;}';
        printHtml += '.rs-two-col{display:flex;gap:16px;}';
        printHtml += '.rs-col{flex:1;}';
        printHtml += '.rs-table{width:100%;border-collapse:collapse;font-size:12px;}';
        printHtml += '.rs-th{padding:6px 8px;text-align:left;border:1px solid #ddd;background:#f8f9fa;font-weight:700;}';
        printHtml += '.rs-th-check{width:30px;text-align:center;}';
        printHtml += '.rs-td{padding:6px 8px;border:1px solid #ddd;vertical-align:top;}';
        printHtml += '.rs-td-check{text-align:center;}';
        printHtml += '.rs-checkbox{width:18px;height:18px;border:2px solid #007bff;border-radius:3px;margin:0 auto;}';
        printHtml += '.rs-name{font-weight:700;font-size:13px;}';
        printHtml += '.rs-info-inline{font-weight:400;font-size:11px;color:#555;}';
        printHtml += '.rs-info-line{font-size:11px;color:#555;margin-top:1px;}';
        printHtml += '.rs-footer{margin-top:16px;border-top:1px solid #ddd;padding-top:12px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;}';
        printHtml += '.rs-footer-stats{font-size:13px;}';
        printHtml += '.rs-footer-sig{font-size:13px;}';
        printHtml += '.rs-legend{margin-top:10px;padding:8px;background:#f8f9fa;border-radius:5px;font-size:11px;}';
        printHtml += '</style></head><body>';
        printHtml += state.previewHtml;
        printHtml += '<script>window.onload=function(){window.print();}<\/script>';
        printHtml += '</body></html>';

        var win = window.open('', '_blank', 'width=900,height=700');
        if (win) {
            win.document.write(printHtml);
            win.document.close();
        } else {
            showToast('Pop-up blocked. Please allow pop-ups for this site.', 'danger');
        }
    }

    // =====================================================================
    // EVENT HANDLERS
    // =====================================================================
    function goLanding() { state.mode = 'landing'; render(); }

    function goAdmin() {
        showLoading();
        loadConfigs(function() {
            state.mode = 'admin';
            render();
        });
    }

    function goAdminList() {
        showLoading();
        loadConfigs(function() {
            state.mode = 'admin';
            state.editConfig = null;
            render();
        });
    }

    function goGenerate() {
        showLoading();
        loadConfigs(function() {
            state.mode = 'generate_pick';
            state.selectedConfigId = null;
            render();
        });
    }

    function loadConfigs(callback) {
        ajax('load_configs', {}, function(err, data) {
            if (err) {
                showToast('Error loading configs: ' + err, 'danger');
                state.configs = [];
            } else if (data && data.success) {
                state.configs = data.configs || [];
            } else {
                showToast(data ? data.message : 'Unknown error', 'danger');
                state.configs = [];
            }
            if (callback) callback();
        });
    }

    function loadFilters(callback) {
        if (state.filtersLoaded) {
            if (callback) callback();
            return;
        }
        ajax('get_filters', {}, function(err, data) {
            if (!err && data && data.success) {
                state.programs = data.programs || [];
                state.divisions = data.divisions || [];
                state.filtersLoaded = true;
            }
            if (callback) callback();
        });
    }

    function newConfig() {
        loadFilters(function() {
            state.editConfig = defaultConfig();
            state.editingDsIndex = -1;
            state.mode = 'admin_edit';
            render();
        });
    }

    function editConfig(id) {
        loadFilters(function() {
            for (var i = 0; i < state.configs.length; i++) {
                if (state.configs[i].id === id) {
                    state.editConfig = migrateConfig(JSON.parse(JSON.stringify(state.configs[i])));
                    break;
                }
            }
            state.editingDsIndex = -1;
            state.mode = 'admin_edit';
            render();
        });
    }

    function duplicateConfig(id) {
        for (var i = 0; i < state.configs.length; i++) {
            if (state.configs[i].id === id) {
                var dup = migrateConfig(JSON.parse(JSON.stringify(state.configs[i])));
                dup.id = '';
                dup.name = (dup.name || 'Untitled') + ' (Copy)';
                loadFilters(function() {
                    state.editConfig = dup;
                    state.editingDsIndex = -1;
                    state.mode = 'admin_edit';
                    render();
                });
                return;
            }
        }
    }

    function deleteConfig(id) {
        if (!confirm('Are you sure you want to delete this config?')) return;
        ajax('delete_config', {config_id: id}, function(err, data) {
            if (!err && data && data.success) {
                showToast('Config deleted', 'success');
                goAdmin();
            } else {
                showToast('Error: ' + (data ? data.message : err), 'danger');
            }
        });
    }

    function collectConfigFromForm() {
        var c = state.editConfig;
        c.name = document.getElementById('rsConfigName').value.trim();
        c.title = document.getElementById('rsConfigTitle').value.trim();

        // Source type
        var radios = document.getElementsByName('rsSourceType');
        for (var r = 0; r < radios.length; r++) {
            if (radios[r].checked) { c.sourceType = radios[r].value; break; }
        }

        // Program/Division
        if (c.sourceType === 'program_division') {
            var progEl = document.getElementById('rsProgramSelect');
            var divEl = document.getElementById('rsDivisionSelect');
            c.programId = progEl ? parseInt(progEl.value) || 0 : 0;
            c.divisionId = divEl ? parseInt(divEl.value) || 0 : 0;
            c.programName = progEl && progEl.selectedIndex > 0 ? progEl.options[progEl.selectedIndex].text : '';
            c.divisionName = divEl && divEl.selectedIndex > 0 ? divEl.options[divEl.selectedIndex].text : '';

            var exEl = document.getElementById('rsExcludeOrgs');
            if (exEl && exEl.value.trim()) {
                c.excludeOrgIds = exEl.value.split(',').map(function(s) { return parseInt(s.trim()); }).filter(function(n) { return !isNaN(n) && n > 0; });
            } else {
                c.excludeOrgIds = [];
            }
        }

        // Columns and name line placement
        var colKeys = ['age', 'gender', 'phone', 'email', 'subgroup', 'memberType'];
        if (!c.nameLineColumns) c.nameLineColumns = {};
        for (var ci = 0; ci < colKeys.length; ci++) {
            var el = document.getElementById('rsCol_' + colKeys[ci]);
            c.columns[colKeys[ci]] = el ? el.checked : false;
            var nlEl = document.getElementById('rsNL_' + colKeys[ci]);
            c.nameLineColumns[colKeys[ci]] = nlEl ? nlEl.checked : false;
        }

        // dataSources are managed in-place on state.editConfig, no form collection needed

        // Remove old fields if present (clean up after migration)
        delete c.recregFields;
        delete c.includeMedical;
        delete c.includePhotos;
        delete c.medicalQuestionLabels;
        delete c.photoQuestionLabel;

        // Layout & Sort
        c.layout = document.getElementById('rsLayout').value;
        c.sortBy = document.getElementById('rsSortBy').value;
        c.showFooter = document.getElementById('rsShowFooter').checked;
        c.onlyWithMeeting = document.getElementById('rsOnlyWithMeeting') ? document.getElementById('rsOnlyWithMeeting').checked : false;

        return c;
    }

    function saveConfig() {
        var c = collectConfigFromForm();
        if (!c.name) {
            showToast('Please enter a config name', 'danger');
            return;
        }
        if (c.sourceType === 'program_division' && !c.programId && !c.divisionId) {
            showToast('Please select a program or division', 'danger');
            return;
        }
        if (c.sourceType === 'specific_orgs' && (!c.selectedOrgs || !c.selectedOrgs.length)) {
            showToast('Please add at least one organization', 'danger');
            return;
        }

        ajax('save_config', {config_json: JSON.stringify(c)}, function(err, data) {
            if (!err && data && data.success) {
                showToast('Config saved!', 'success');
                state.editConfig = data.config;
                goAdmin();
            } else {
                showToast('Error: ' + (data ? data.message : err), 'danger');
            }
        });
    }

    function setSourceType(type) {
        state.editConfig.sourceType = type;
        var progDiv = document.getElementById('rsProgramDivSection');
        var specOrgs = document.getElementById('rsSpecificOrgsSection');
        if (type === 'specific_orgs') {
            progDiv.style.display = 'none';
            specOrgs.style.display = '';
        } else {
            progDiv.style.display = '';
            specOrgs.style.display = 'none';
        }
    }

    function onProgramChange() {
        var progEl = document.getElementById('rsProgramSelect');
        var progId = parseInt(progEl.value) || 0;
        state.editConfig.programId = progId;

        // Filter divisions
        var divEl = document.getElementById('rsDivisionSelect');
        var currentDiv = parseInt(divEl.value) || 0;
        var html = '<option value="">-- All Divisions --</option>';
        for (var i = 0; i < state.divisions.length; i++) {
            var d = state.divisions[i];
            if (!progId || d.progId == progId) {
                var sel = (d.id == currentDiv) ? ' selected' : '';
                html += '<option value="' + d.id + '"' + sel + '>' + escHtml(d.name) + '</option>';
            }
        }
        divEl.innerHTML = html;
    }

    function onDivisionChange() {
        var divEl = document.getElementById('rsDivisionSelect');
        state.editConfig.divisionId = parseInt(divEl.value) || 0;
    }

    // Org search for specific_orgs mode
    function onOrgSearch() {
        clearTimeout(state.searchTimeout);
        var input = document.getElementById('rsOrgSearch');
        var term = input.value.trim();
        if (term.length < 2) {
            document.getElementById('rsOrgSearchResults').style.display = 'none';
            return;
        }
        state.searchTimeout = setTimeout(function() {
            ajax('search_involvements', {search_term: term}, function(err, data) {
                if (!err && data && data.success) {
                    var results = data.involvements || [];
                    var resultsEl = document.getElementById('rsOrgSearchResults');
                    if (!results.length) {
                        resultsEl.innerHTML = '<div style="padding:12px;color:#999;">No results found</div>';
                    } else {
                        var rh = '';
                        for (var i = 0; i < results.length; i++) {
                            var r = results[i];
                            rh += '<div class="rs-search-result-item" onclick="RSApp.addOrg(' + r.orgId + ',\\'' + escAttr(r.orgName) + '\\',' + (r.memberCount || 0) + ')">';
                            rh += '<div class="rs-search-result-name">' + escHtml(r.orgName) + ' (ID: ' + r.orgId + ')</div>';
                            rh += '<div class="rs-search-result-meta">' + escHtml(r.programName) + ' / ' + escHtml(r.divisionName) + ' &bull; ' + r.memberCount + ' members</div>';
                            rh += '</div>';
                        }
                        resultsEl.innerHTML = rh;
                    }
                    resultsEl.style.display = '';
                }
            });
        }, 300);
    }

    function addOrg(orgId, orgName, memberCount) {
        var c = state.editConfig;
        if (!c.selectedOrgs) c.selectedOrgs = [];
        // Don't add duplicates
        for (var i = 0; i < c.selectedOrgs.length; i++) {
            if (c.selectedOrgs[i].orgId == orgId) {
                showToast('Organization already added', 'info');
                return;
            }
        }
        c.selectedOrgs.push({orgId: orgId, orgName: orgName, memberCount: memberCount});

        // Re-render selected list
        var listEl = document.getElementById('rsSelectedOrgsList');
        var lh = '';
        for (var j = 0; j < c.selectedOrgs.length; j++) {
            var so = c.selectedOrgs[j];
            lh += '<div class="rs-selected-org-item">';
            lh += '<span>' + escHtml(so.orgName) + ' (ID: ' + so.orgId + ', Members: ' + (so.memberCount || '?') + ')</span>';
            lh += '<span class="rs-remove-org" onclick="RSApp.removeOrg(' + j + ')">&#10005;</span>';
            lh += '</div>';
        }
        listEl.innerHTML = lh;

        // Clear search
        document.getElementById('rsOrgSearch').value = '';
        document.getElementById('rsOrgSearchResults').style.display = 'none';
    }

    function removeOrg(index) {
        var c = state.editConfig;
        c.selectedOrgs.splice(index, 1);
        // Re-render selected list
        var listEl = document.getElementById('rsSelectedOrgsList');
        if (!c.selectedOrgs.length) {
            listEl.innerHTML = '<div class="rs-text-muted rs-text-sm">No organizations selected yet.</div>';
            return;
        }
        var lh = '';
        for (var j = 0; j < c.selectedOrgs.length; j++) {
            var so = c.selectedOrgs[j];
            lh += '<div class="rs-selected-org-item">';
            lh += '<span>' + escHtml(so.orgName) + ' (ID: ' + so.orgId + ', Members: ' + (so.memberCount || '?') + ')</span>';
            lh += '<span class="rs-remove-org" onclick="RSApp.removeOrg(' + j + ')">&#10005;</span>';
            lh += '</div>';
        }
        listEl.innerHTML = lh;
    }

    // Generate
    function selectConfig(id) {
        state.selectedConfigId = id;
        render();
    }

    // Date helpers
    function _formatDate(d) {
        var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
        var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        return days[d.getDay()] + ', ' + months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
    }
    function _isoDate(d) {
        var mm = ('0' + (d.getMonth()+1)).slice(-2);
        var dd = ('0' + d.getDate()).slice(-2);
        return d.getFullYear() + '-' + mm + '-' + dd;
    }
    function _nextDayOfWeek(dayNum) {
        // dayNum: 0=Sun,3=Wed,etc. Returns the next occurrence (today if it matches)
        var d = new Date();
        var diff = (dayNum - d.getDay() + 7) % 7;
        d.setDate(d.getDate() + diff);
        return d;
    }
    function _nextNextDayOfWeek(dayNum) {
        var d = _nextDayOfWeek(dayNum);
        d.setDate(d.getDate() + 7);
        return d;
    }

    function renderDatePickerModal() {
        var today = new Date();
        var thisSun = _nextDayOfWeek(0);
        var nextSun = _nextNextDayOfWeek(0);
        var thisWed = _nextDayOfWeek(3);

        var h = '<div class="rs-modal-overlay" onclick="RSApp.closeDatePicker()">';
        h += '<div class="rs-modal" onclick="event.stopPropagation()">';
        h += '<h3>Select Report Date</h3>';
        h += '<div class="rs-date-grid">';
        h += '<div class="rs-date-btn" onclick="RSApp.pickDate(\\'today\\')"><strong>Today</strong><span>' + _formatDate(today) + '</span></div>';
        h += '<div class="rs-date-btn" onclick="RSApp.pickDate(\\'this_sun\\')"><strong>Coming Sunday</strong><span>' + _formatDate(thisSun) + '</span></div>';
        h += '<div class="rs-date-btn" onclick="RSApp.pickDate(\\'next_sun\\')"><strong>Following Sunday</strong><span>' + _formatDate(nextSun) + '</span></div>';
        h += '<div class="rs-date-btn" onclick="RSApp.pickDate(\\'this_wed\\')"><strong>Coming Wednesday</strong><span>' + _formatDate(thisWed) + '</span></div>';
        h += '</div>';
        h += '<div class="rs-form-group">';
        h += '<label class="rs-label">Custom Date</label>';
        h += '<div class="rs-flex rs-gap-8">';
        h += '<input type="date" class="rs-input" id="rsCustomDate" style="flex:1;">';
        h += '<button class="rs-btn rs-btn-primary" onclick="RSApp.pickDate(\\'custom\\')">Use Date</button>';
        h += '</div>';
        h += '</div>';
        h += '<div style="margin-top:12px;text-align:right;">';
        h += '<button class="rs-btn rs-btn-outline" onclick="RSApp.closeDatePicker()">Cancel</button>';
        h += '</div>';
        h += '</div></div>';
        return h;
    }

    function generateRollsheet() {
        if (!state.selectedConfigId) {
            showToast('Please select a config first', 'danger');
            return;
        }
        state.showDatePicker = true;
        render();
    }

    function closeDatePicker() {
        state.showDatePicker = false;
        render();
    }

    function pickDate(which) {
        var dateStr = '';
        var isoStr = '';
        var d;
        if (which === 'today') {
            d = new Date();
        } else if (which === 'this_sun') {
            d = _nextDayOfWeek(0);
        } else if (which === 'next_sun') {
            d = _nextNextDayOfWeek(0);
        } else if (which === 'this_wed') {
            d = _nextDayOfWeek(3);
        } else if (which === 'custom') {
            var input = document.getElementById('rsCustomDate');
            if (!input || !input.value) {
                showToast('Please select a date', 'danger');
                return;
            }
            var parts = input.value.split('-');
            d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        }
        dateStr = _formatDate(d);
        isoStr = _isoDate(d);
        state.reportDate = dateStr;
        state.reportDateIso = isoStr;
        state.showDatePicker = false;
        doGenerate();
    }

    function doGenerate() {
        var config = null;
        for (var i = 0; i < state.configs.length; i++) {
            if (state.configs[i].id === state.selectedConfigId) {
                config = migrateConfig(JSON.parse(JSON.stringify(state.configs[i])));
                break;
            }
        }
        if (!config) {
            showToast('Please select a config first', 'danger');
            return;
        }
        showLoading();
        ajax('generate_rollsheet', {config_json: JSON.stringify(config), report_date: state.reportDate, report_date_iso: state.reportDateIso}, function(err, data) {
            if (!err && data && data.success) {
                state.previewHtml = data.html || '';
                state.previewOrgCount = data.orgCount || 0;
                state.previewTotalMembers = data.totalMembers || 0;
                state.mode = 'generate_preview';
                render();
            } else {
                showToast('Error: ' + (data ? data.message : err), 'danger');
                state.mode = 'generate_pick';
                render();
            }
        });
    }

    // =====================================================================
    // PUBLIC API
    // =====================================================================
    window.RSApp = {
        goLanding: goLanding,
        goAdmin: goAdmin,
        goAdminList: goAdminList,
        goGenerate: goGenerate,
        newConfig: newConfig,
        editConfig: editConfig,
        duplicateConfig: duplicateConfig,
        deleteConfig: deleteConfig,
        saveConfig: saveConfig,
        setSourceType: setSourceType,
        onProgramChange: onProgramChange,
        onDivisionChange: onDivisionChange,
        onOrgSearch: onOrgSearch,
        addOrg: addOrg,
        removeOrg: removeOrg,
        selectConfig: selectConfig,
        generateRollsheet: generateRollsheet,
        closeDatePicker: closeDatePicker,
        pickDate: pickDate,
        printRollsheets: printRollsheets,
        discoverFields: discoverFields,
        addDsFromPicker: addDsFromPicker,
        editDs: editDs,
        saveDsEdit: saveDsEdit,
        cancelDsEdit: cancelDsEdit,
        removeDs: removeDs,
        moveDsUp: moveDsUp,
        moveDsDown: moveDsDown,
        onDsTypeChange: onDsTypeChange,
        pickPresetColor: pickPresetColor
    };

    // =====================================================================
    // INIT
    // =====================================================================
    render();

})();
</script>
</div>
</body>
</html>'''

    model.Form = html
