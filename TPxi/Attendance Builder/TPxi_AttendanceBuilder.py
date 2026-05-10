#roles=Edit
#----------------------------------------------------------------------
# TPxi_AttendanceBuilder.py
#
# Configurable Attendance Report Builder for TouchPoint
#
# Allows admins to create reusable attendance report configurations,
# then generate reports showing historical attendance data from the
# Attend table. Reports include daily summaries, program/division
# breakdowns, per-org detail, subgroup breakdowns, and enrollment ratios.
#
# Key difference from RollSheet: queries the Attend table for historical
# attendance (people dropped post-event still show up), not current
# OrganizationMembers.
#
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
#
# Architecture:
#   Single .py file SPA - Python AJAX handlers for POST, HTML SPA for GET.
#   Data stored via model.WriteContentText / model.TextContent.
#
# Storage Keys:
#   AttendanceBuilder_Configs - All attendance report configs (JSON)
#
# CSS Prefix: ab-
# Root Class: .ab-root
#
# Reference:
#   TPxi_RollSheet.py - SPA architecture, CRUD, search
#----------------------------------------------------------------------

import json
import datetime
import random

model.Header = 'Attendance Report Builder'

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

def format_number(num):
    try:
        s = str(int(num))
        groups = []
        while s:
            groups.append(s[-3:])
            s = s[:-3]
        return ','.join(reversed(groups))
    except:
        return '0'

def normalize_date_str(s):
    """Normalize SQL date strings to YYYY-MM-DD.
    IronPython renders System.DateTime as 'M/D/YYYY hh:mm:ss AM' (US locale);
    SQL Server CONVERT(date, ...) may also yield 'YYYY-MM-DD'. Handle both so
    sorting and date-parsing work consistently."""
    if not s:
        return ''
    s = str(s).strip()
    if ' ' in s:
        s = s.split(' ')[0]
    if '/' in s:
        try:
            p = s.split('/')
            return '%04d-%02d-%02d' % (int(p[2]), int(p[0]), int(p[1]))
        except:
            return s
    return s

def format_pct(num, denom):
    if not denom or denom == 0:
        return '-'
    pct = int(num * 100 / denom)
    return str(pct) + '%'

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
            raw = model.TextContent("AttendanceBuilder_Configs") or ''
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

            raw = model.TextContent("AttendanceBuilder_Configs") or ''
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
                    config['id'] = 'ab_' + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + str(random.randint(100, 999))
                config['createdAt'] = now
                config['updatedAt'] = now
                configs.append(config)

            data['configs'] = configs
            model.WriteContentText("AttendanceBuilder_Configs", json.dumps(data), "")
            print json.dumps({'success': True, 'config': config, 'message': 'Config saved'})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # Delete config
    # -----------------------------------------------------------------
    elif action == 'delete_config':
        try:
            config_id = str(Data.config_id) if hasattr(Data, 'config_id') else ''
            raw = model.TextContent("AttendanceBuilder_Configs") or ''
            data = json.loads(raw) if raw else {'configs': []}
            configs = data.get('configs', [])

            new_configs = [c for c in configs if c.get('id') != config_id]
            if len(new_configs) < len(configs):
                data['configs'] = new_configs
                model.WriteContentText("AttendanceBuilder_Configs", json.dumps(data), "")
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
            include_inactive = str(getattr(Data, 'include_inactive', '')).lower() in ('1', 'true', 'yes')

            # 1=1 placeholder so WHERE clause is never empty if no other filters are set
            where_clauses = ["1=1"] if include_inactive else ["o.OrganizationStatusId = 30"]

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
    # Generate attendance report
    # -----------------------------------------------------------------
    elif action == 'generate_report':
        try:
            config_json = str(Data.config_json) if hasattr(Data, 'config_json') else ''
            config = json.loads(config_json)

            source_type = config.get('sourceType', 'program_division')
            program_div_groups = config.get('programDivGroups', [])
            # Backward compat: migrate old single programId/divisionId to groups
            if not program_div_groups and source_type != 'specific_orgs':
                old_pid = safe_int(config.get('programId', 0))
                old_did = safe_int(config.get('divisionId', 0))
                if old_pid > 0:
                    program_div_groups = [{'programId': old_pid, 'divisionId': old_did,
                                           'programName': config.get('programName', ''),
                                           'divisionName': config.get('divisionName', '')}]
            specific_org_ids = config.get('specificOrgIds', [])
            specific_org_names = config.get('specificOrgNames', {})
            exclude_org_ids_str = config.get('excludeOrgIds', '')
            print_title = config.get('printTitle', config.get('name', 'Attendance Report'))
            sections = config.get('reportSections', {})
            leader_member_types_str = config.get('leaderMemberTypes', '140,160')
            sort_by = config.get('sortBy', 'name')

            # Dates come exclusively from override params
            start_date = str(Data.override_start) if hasattr(Data, 'override_start') and Data.override_start else ''
            end_date = str(Data.override_end) if hasattr(Data, 'override_end') and Data.override_end else ''

            if not start_date or not end_date:
                print json.dumps({'success': False, 'message': 'Start and end dates are required'})
            else:
                # Honor the report's actual year in the print title. Replace any 4-digit
                # year (19xx/20xx) AND a literal {year} placeholder with the start date's
                # year - so titles like "VBS 2026 Attendance Report" stay accurate when
                # run on a different year's data without needing to edit the config.
                try:
                    import re as _re_year
                    rpt_year = str(start_date)[:4]
                    if rpt_year.isdigit() and len(rpt_year) == 4:
                        print_title = print_title.replace('{year}', rpt_year)
                        print_title = _re_year.sub(r'\b(?:19|20)\d{2}\b', rpt_year, print_title)
                except:
                    pass
                # Parse leader member types
                leader_types = []
                for lt in leader_member_types_str.split(','):
                    lt = lt.strip()
                    if lt:
                        leader_types.append(safe_int(lt))
                leader_types_str = ','.join(str(x) for x in leader_types) if leader_types else '140,160'

                # Step 1: Resolve org IDs
                org_ids = []
                if source_type == 'specific_orgs':
                    org_ids = [safe_int(x) for x in specific_org_ids if safe_int(x) > 0]
                else:
                    # Build OR conditions from programDivGroups
                    group_conditions = []
                    for grp in program_div_groups:
                        pid = safe_int(grp.get('programId', 0))
                        did = safe_int(grp.get('divisionId', 0))
                        if pid > 0 and did > 0:
                            group_conditions.append("(os.ProgId = {0} AND os.DivId = {1})".format(pid, did))
                        elif pid > 0:
                            group_conditions.append("(os.ProgId = {0})".format(pid))
                    if not group_conditions:
                        print json.dumps({'success': False, 'message': 'At least one program/division group is required'})
                        org_ids = []
                    else:
                        # Active-only by default; the includeInactive section toggle drops
                        # the status filter so historical reports (where the orgs are now
                        # archived/inactive) can still resolve their org list.
                        status_filter = "" if sections.get('includeInactive', False) else " AND o.OrganizationStatusId = 30"
                        org_sql = """
                            SELECT DISTINCT os.OrgId
                            FROM OrganizationStructure os
                            JOIN Organizations o ON os.OrgId = o.OrganizationId
                            WHERE ({0}){1}
                        """.format(" OR ".join(group_conditions), status_filter)
                        for r in q.QuerySql(org_sql):
                            org_ids.append(r.OrgId)

                # Apply exclusions
                if exclude_org_ids_str:
                    exclude_set = set()
                    for x in str(exclude_org_ids_str).split(','):
                        x = x.strip()
                        if x:
                            exclude_set.add(safe_int(x))
                    org_ids = [oid for oid in org_ids if oid not in exclude_set]

                if not org_ids:
                    print json.dumps({'success': True, 'html': '<p style="text-align:center;color:#666;padding:40px;">No organizations found matching your config.</p>', 'stats': {}})
                else:
                    org_ids_str = ','.join(str(oid) for oid in org_ids)
                    safe_start = start_date.replace("'", "")
                    safe_end = end_date.replace("'", "")

                    # An org can appear in OrganizationStructure under MULTIPLE (Prog, Div)
                    # combos (e.g., a VBS K-5 classroom is also under "VBS DAILY GRAND TOTAL").
                    # If we just JOIN OrganizationStructure, those duplicate rows fan out the
                    # attendance counts AND surface programs the user never selected.
                    # Build a CTE (OrgPD) that:
                    #   1. Restricts to the user's selected (Prog, Div) combos when the source
                    #      is program/division mode (skipped for specific_orgs - any row OK).
                    #   2. Picks ONE structure row per OrgId via ROW_NUMBER (rn=1).
                    if source_type == 'specific_orgs':
                        pd_inner_filter = ''
                    else:
                        pd_conds = []
                        for grp in program_div_groups:
                            pid = safe_int(grp.get('programId', 0))
                            did = safe_int(grp.get('divisionId', 0))
                            if pid > 0 and did > 0:
                                pd_conds.append("(ProgId = {0} AND DivId = {1})".format(pid, did))
                            elif pid > 0:
                                pd_conds.append("(ProgId = {0})".format(pid))
                        pd_inner_filter = " AND (" + " OR ".join(pd_conds) + ")" if pd_conds else ""
                    org_pd_cte = (
                        "WITH OrgPD AS ("
                        " SELECT OrgId, ProgId, Program, DivId, Division,"
                        " ROW_NUMBER() OVER (PARTITION BY OrgId ORDER BY ProgId, DivId) AS rn"
                        " FROM OrganizationStructure"
                        " WHERE OrgId IN ({0}){1})"
                    ).format(org_ids_str, pd_inner_filter)

                    # Step 2: Core attendance query
                    # Pull all Attend rows in range (both attended and absent). The Attend table
                    # records a row per person-meeting; AttendanceFlag=1 means they showed up.
                    # MemberTypeId reflects the person's role at meeting time. We use this to:
                    #   - count attended (flag=1) split into leaders vs members by MemberTypeId
                    #   - count enrolled (any flag) for ratio denominators
                    attend_sql = org_pd_cte + """
                        SELECT a.OrganizationId, o.OrganizationName,
                               p.Program, p.ProgId, p.Division, p.DivId,
                               CONVERT(date, a.MeetingDate) AS AttendDate,
                               a.MemberTypeId,
                               a.AttendanceFlag,
                               COUNT(DISTINCT a.PeopleId) AS PeopleCount
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        JOIN OrgPD p ON p.OrgId = a.OrganizationId AND p.rn = 1
                        WHERE a.MeetingDate >= '{0}' AND a.MeetingDate < DATEADD(day,1,'{1}')
                        GROUP BY a.OrganizationId, o.OrganizationName,
                                 p.Program, p.ProgId, p.Division, p.DivId,
                                 CONVERT(date, a.MeetingDate), a.MemberTypeId, a.AttendanceFlag
                    """.format(safe_start, safe_end)

                    # Step 3: Aggregate into Python dicts
                    # 'enrolled' = any-flag count (was on roster); 'leaders'/'members' = attended (flag=1)
                    by_date = {}       # date_str -> {leaders, members, total, enrolled_sum}
                    by_org = {}        # org_id -> {name, dates: {date_str -> {leaders, members, enrolled}}}
                    org_program = {}   # org_id -> {progId, progName, divId, divName}
                    by_prog = {}       # progId -> {name, dates: {date_str -> {leaders, members, total}}}
                    by_div = {}        # divId -> {name, progId, progName, dates: {date_str -> {leaders, members, total}}}
                    all_dates = set()
                    total_attendance = 0
                    orgs_with_data = set()

                    for row in q.QuerySql(attend_sql):
                        oid = row.OrganizationId
                        org_name = safe_str(row.OrganizationName)
                        date_str = normalize_date_str(safe_str(row.AttendDate))
                        mt = safe_int(row.MemberTypeId)
                        flag = safe_int(row.AttendanceFlag)
                        count = safe_int(row.PeopleCount)
                        prog_id = safe_int(row.ProgId)
                        prog_name = safe_str(row.Program)
                        div_id = safe_int(row.DivId)
                        div_name = safe_str(row.Division)

                        is_leader = mt in leader_types
                        attended = (flag == 1)
                        all_dates.add(date_str)
                        if attended:
                            total_attendance += count
                        orgs_with_data.add(oid)

                        # by_date
                        if date_str not in by_date:
                            by_date[date_str] = {'leaders': 0, 'members': 0, 'total': 0, 'enrolled_sum': 0}
                        # enrolled_sum overcounts people in multiple orgs; replaced by distinct query below
                        by_date[date_str]['enrolled_sum'] += count
                        if attended:
                            if is_leader:
                                by_date[date_str]['leaders'] += count
                            else:
                                by_date[date_str]['members'] += count
                            by_date[date_str]['total'] += count

                        # by_org
                        if oid not in by_org:
                            by_org[oid] = {'name': org_name, 'dates': {}}
                        if date_str not in by_org[oid]['dates']:
                            by_org[oid]['dates'][date_str] = {'leaders': 0, 'members': 0, 'enrolled': 0}
                        # within (org, date) summing across mtype/flag = distinct people (1 record per person/meeting)
                        by_org[oid]['dates'][date_str]['enrolled'] += count
                        if attended:
                            if is_leader:
                                by_org[oid]['dates'][date_str]['leaders'] += count
                            else:
                                by_org[oid]['dates'][date_str]['members'] += count

                        # org_program
                        if oid not in org_program:
                            org_program[oid] = {'progId': prog_id, 'progName': prog_name, 'divId': div_id, 'divName': div_name}

                        # by_prog (attended only - matches existing behavior for breakdown table)
                        if prog_id not in by_prog:
                            by_prog[prog_id] = {'name': prog_name, 'dates': {}}
                        if date_str not in by_prog[prog_id]['dates']:
                            by_prog[prog_id]['dates'][date_str] = {'leaders': 0, 'members': 0, 'total': 0}
                        if attended:
                            if is_leader:
                                by_prog[prog_id]['dates'][date_str]['leaders'] += count
                            else:
                                by_prog[prog_id]['dates'][date_str]['members'] += count
                            by_prog[prog_id]['dates'][date_str]['total'] += count

                        # by_div (attended only)
                        if div_id not in by_div:
                            by_div[div_id] = {'name': div_name, 'progId': prog_id, 'progName': prog_name, 'dates': {}}
                        if date_str not in by_div[div_id]['dates']:
                            by_div[div_id]['dates'][date_str] = {'leaders': 0, 'members': 0, 'total': 0}
                        if attended:
                            if is_leader:
                                by_div[div_id]['dates'][date_str]['leaders'] += count
                            else:
                                by_div[div_id]['dates'][date_str]['members'] += count
                            by_div[div_id]['dates'][date_str]['total'] += count

                    # Step 3b: Headcount support (tracked as a SEPARATE column)
                    # Per-meeting headcount diff: max(0, HeadCount - distinct_individual_attended).
                    # Examples:
                    #   - Individual 25, HeadCount 30 -> diff = 5
                    #   - Individual 25, HeadCount 20 -> diff = 0
                    #   - Individual 0,  HeadCount 30 -> diff = 30 (headcount-only meeting)
                    #
                    # HC is tracked separately and rendered as its own Headcount column +
                    # a Grand Total column (Total + Headcount). It is intentionally NOT
                    # added to Total, Members, Leaders, or Enrolled - so the enrollment
                    # ratio stays accurate (it's a per-person ratio, and HC has no people).
                    #
                    # First: build complete org->prog/div metadata covering ALL in-scope orgs
                    # (including any that have ONLY headcount meetings - they wouldn't appear
                    # in by_org from the main loop). Used both to label headcount-only orgs
                    # and to roll headcounts up into the right program/division.
                    org_meta = {}  # oid -> {name, progId, progName, divId, divName}
                    meta_sql = org_pd_cte + """
                        SELECT p.OrgId, o.OrganizationName,
                               p.ProgId, p.Program, p.DivId, p.Division
                        FROM OrgPD p
                        JOIN Organizations o ON o.OrganizationId = p.OrgId
                        WHERE p.rn = 1
                    """
                    for r in q.QuerySql(meta_sql):
                        org_meta[r.OrgId] = {
                            'name': safe_str(r.OrganizationName),
                            'progId': safe_int(r.ProgId),
                            'progName': safe_str(r.Program),
                            'divId': safe_int(r.DivId),
                            'divName': safe_str(r.Division)
                        }

                    hc_by_org_date = {}  # (oid, ds) -> sum of HC diffs (HeadCount - individual, when positive) across meetings
                    hc_sql = """
                        SELECT m.OrganizationId, CONVERT(date, m.MeetingDate) AS dt,
                               SUM(CASE WHEN ISNULL(m.HeadCount, 0) > ISNULL(att.cnt, 0)
                                        THEN ISNULL(m.HeadCount, 0) - ISNULL(att.cnt, 0)
                                        ELSE 0 END) AS HC_Diff
                        FROM Meetings m
                        OUTER APPLY (
                            SELECT COUNT(DISTINCT a.PeopleId) AS cnt
                            FROM Attend a
                            WHERE a.MeetingId = m.MeetingId AND a.AttendanceFlag = 1
                        ) att
                        WHERE m.MeetingDate >= '{0}' AND m.MeetingDate < DATEADD(day,1,'{1}')
                          AND m.OrganizationId IN ({2})
                          AND ISNULL(m.DidNotMeet, 0) = 0
                          AND ISNULL(m.HeadCount, 0) > 0
                        GROUP BY m.OrganizationId, CONVERT(date, m.MeetingDate)
                        HAVING SUM(CASE WHEN ISNULL(m.HeadCount, 0) > ISNULL(att.cnt, 0)
                                        THEN ISNULL(m.HeadCount, 0) - ISNULL(att.cnt, 0)
                                        ELSE 0 END) > 0
                    """.format(safe_start, safe_end, org_ids_str)
                    for r in q.QuerySql(hc_sql):
                        oid = r.OrganizationId
                        ds = normalize_date_str(safe_str(r.dt))
                        hc = safe_int(r.HC_Diff)
                        if hc <= 0:
                            continue
                        hc_by_org_date[(oid, ds)] = hc
                        all_dates.add(ds)
                        orgs_with_data.add(oid)
                        # NOTE: total_attendance intentionally NOT incremented - HC stays
                        # out of the "Total Attendance" stat (it's a Grand Total contributor).

                    # Build HC aggregates at every grain used by the renderers. HC is NEVER
                    # added to attendance counts - the dicts here exist solely to populate
                    # the Headcount column and compute Grand Total (Total + HC).
                    hc_per_day = {}        # ds -> hc
                    hc_per_org = {}        # oid -> hc
                    hc_per_prog = {}       # pid -> hc
                    hc_per_div = {}        # (pid, did) -> hc
                    hc_per_day_prog = {}   # (ds, pid) -> hc
                    hc_per_day_div = {}    # (ds, (pid, did)) -> hc
                    hc_period = 0
                    for (oid, ds), hc in hc_by_org_date.items():
                        meta = org_meta.get(oid, {})
                        pid = meta.get('progId', 0)
                        did = meta.get('divId', 0)
                        hc_period += hc
                        hc_per_day[ds] = hc_per_day.get(ds, 0) + hc
                        hc_per_org[oid] = hc_per_org.get(oid, 0) + hc
                        hc_per_prog[pid] = hc_per_prog.get(pid, 0) + hc
                        hc_per_div[(pid, did)] = hc_per_div.get((pid, did), 0) + hc
                        hc_per_day_prog[(ds, pid)] = hc_per_day_prog.get((ds, pid), 0) + hc
                        hc_per_day_div[(ds, (pid, did))] = hc_per_day_div.get((ds, (pid, did)), 0) + hc

                    # Populate zero-value placeholders so HC-only orgs / HC-only dates still
                    # appear in the rendered report (they wouldn't otherwise, since by_date /
                    # by_org / by_prog / by_div are built from Attend rows only).
                    for (oid, ds) in hc_by_org_date:
                        meta = org_meta.get(oid, {})
                        pid = meta.get('progId', 0)
                        pname = meta.get('progName', '')
                        did = meta.get('divId', 0)
                        dname = meta.get('divName', '')

                        if ds not in by_date:
                            by_date[ds] = {'leaders': 0, 'members': 0, 'total': 0, 'enrolled_sum': 0}
                        if oid not in by_org:
                            by_org[oid] = {'name': meta.get('name', ''), 'dates': {}}
                        if ds not in by_org[oid]['dates']:
                            by_org[oid]['dates'][ds] = {'leaders': 0, 'members': 0, 'enrolled': 0}
                        if pid not in by_prog:
                            by_prog[pid] = {'name': pname, 'dates': {}}
                        if ds not in by_prog[pid]['dates']:
                            by_prog[pid]['dates'][ds] = {'leaders': 0, 'members': 0, 'total': 0}
                        if did not in by_div:
                            by_div[did] = {'name': dname, 'progId': pid, 'progName': pname, 'dates': {}}
                        if ds not in by_div[did]['dates']:
                            by_div[did]['dates'][ds] = {'leaders': 0, 'members': 0, 'total': 0}
                        if oid not in org_program:
                            org_program[oid] = {'progId': pid, 'progName': pname, 'divId': did, 'divName': dname}

                    sorted_dates = sorted(list(all_dates))
                    day_names = {0: 'Mon', 1: 'Tue', 2: 'Wed', 3: 'Thu', 4: 'Fri', 5: 'Sat', 6: 'Sun'}

                    def get_day_name(ds):
                        try:
                            parts = ds.split('-')
                            d = datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
                            return day_names.get(d.weekday(), '')
                        except:
                            return ''

                    def format_date_short(ds):
                        try:
                            parts = ds.split('-')
                            return str(int(parts[1])) + '/' + str(int(parts[2]))
                        except:
                            return ds

                    # Step 4: Build summary metrics for daily/period/program/division/org views.
                    # Two modes (controlled by sections.countDistinct):
                    #   distinct (opt-in): COUNT(DISTINCT PeopleId) at each grain - true
                    #     unique-people counts. Required if you want to know "how many
                    #     individual people attended" without double-counting people who
                    #     appear in multiple orgs (e.g., volunteers serving rotations).
                    #   sum (default): aggregate from the main per-(org,date,mtype,flag)
                    #     query - matches the original Attendance Report behavior. A person
                    #     in N orgs counts N times.
                    # Either mode populates the same dicts so rendering is unified.
                    count_distinct = sections.get('countDistinct', False)
                    daily_distinct = {}   # date_str -> {leaders, total, enrolled}
                    period_distinct = {'leaders': 0, 'total': 0, 'enrolled': 0}
                    div_distinct = {}    # (prog_id, div_id) -> {leaders, total, enrolled, prog_name, div_name}
                    prog_distinct = {}   # prog_id -> {leaders, total, enrolled, name}
                    org_distinct = {}    # org_id -> {leaders, total, enrolled}
                    # Per-day program/division metrics for the Daily Summary expand-rows
                    daily_prog = {}      # date_str -> {prog_id -> {leaders, total, enrolled, name}}
                    daily_div = {}       # date_str -> {(prog_id, div_id) -> {leaders, total, enrolled, prog_name, div_name}}

                    need_prog_div = sections.get('programDivisionBreakdown', True)
                    need_org_summary = sorted_dates and (
                        sections.get('orgDetail', False) or need_prog_div
                    )

                    if count_distinct and sorted_dates:
                        # --- Distinct mode: SQL-side COUNT(DISTINCT) at each grain ---
                        per_day_sql = """
                            SELECT CONVERT(date, MeetingDate) AS AttendDate,
                                   COUNT(DISTINCT PeopleId) AS Enrolled,
                                   COUNT(DISTINCT CASE WHEN AttendanceFlag=1 THEN PeopleId END) AS Total,
                                   COUNT(DISTINCT CASE WHEN AttendanceFlag=1 AND MemberTypeId IN ({3}) THEN PeopleId END) AS Leaders
                            FROM Attend
                            WHERE MeetingDate >= '{0}' AND MeetingDate < DATEADD(day,1,'{1}')
                              AND OrganizationId IN ({2})
                            GROUP BY CONVERT(date, MeetingDate)
                        """.format(safe_start, safe_end, org_ids_str, leader_types_str)
                        for r in q.QuerySql(per_day_sql):
                            ds = normalize_date_str(safe_str(r.AttendDate))
                            daily_distinct[ds] = {
                                'leaders': safe_int(r.Leaders),
                                'total': safe_int(r.Total),
                                'enrolled': safe_int(r.Enrolled)
                            }

                        period_sql = """
                            SELECT COUNT(DISTINCT PeopleId) AS Enrolled,
                                   COUNT(DISTINCT CASE WHEN AttendanceFlag=1 THEN PeopleId END) AS Total,
                                   COUNT(DISTINCT CASE WHEN AttendanceFlag=1 AND MemberTypeId IN ({3}) THEN PeopleId END) AS Leaders
                            FROM Attend
                            WHERE MeetingDate >= '{0}' AND MeetingDate < DATEADD(day,1,'{1}')
                              AND OrganizationId IN ({2})
                        """.format(safe_start, safe_end, org_ids_str, leader_types_str)
                        rows = list(q.QuerySql(period_sql))
                        if rows:
                            period_distinct = {
                                'leaders': safe_int(rows[0].Leaders),
                                'total': safe_int(rows[0].Total),
                                'enrolled': safe_int(rows[0].Enrolled)
                            }

                        if need_prog_div:
                            per_div_sql = org_pd_cte + """
                                SELECT p.ProgId, p.Program, p.DivId, p.Division,
                                       COUNT(DISTINCT a.PeopleId) AS Enrolled,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 THEN a.PeopleId END) AS Total,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 AND a.MemberTypeId IN ({2}) THEN a.PeopleId END) AS Leaders
                                FROM Attend a
                                JOIN OrgPD p ON p.OrgId = a.OrganizationId AND p.rn = 1
                                WHERE a.MeetingDate >= '{0}' AND a.MeetingDate < DATEADD(day,1,'{1}')
                                GROUP BY p.ProgId, p.Program, p.DivId, p.Division
                            """.format(safe_start, safe_end, leader_types_str)
                            for r in q.QuerySql(per_div_sql):
                                key = (safe_int(r.ProgId), safe_int(r.DivId))
                                div_distinct[key] = {
                                    'leaders': safe_int(r.Leaders),
                                    'total': safe_int(r.Total),
                                    'enrolled': safe_int(r.Enrolled),
                                    'prog_name': safe_str(r.Program),
                                    'div_name': safe_str(r.Division)
                                }

                            per_prog_sql = org_pd_cte + """
                                SELECT p.ProgId, p.Program,
                                       COUNT(DISTINCT a.PeopleId) AS Enrolled,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 THEN a.PeopleId END) AS Total,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 AND a.MemberTypeId IN ({2}) THEN a.PeopleId END) AS Leaders
                                FROM Attend a
                                JOIN OrgPD p ON p.OrgId = a.OrganizationId AND p.rn = 1
                                WHERE a.MeetingDate >= '{0}' AND a.MeetingDate < DATEADD(day,1,'{1}')
                                GROUP BY p.ProgId, p.Program
                            """.format(safe_start, safe_end, leader_types_str)
                            for r in q.QuerySql(per_prog_sql):
                                prog_distinct[safe_int(r.ProgId)] = {
                                    'leaders': safe_int(r.Leaders),
                                    'total': safe_int(r.Total),
                                    'enrolled': safe_int(r.Enrolled),
                                    'name': safe_str(r.Program)
                                }

                            # Per-(day, prog) and per-(day, prog, div) for the Daily Summary expand-rows.
                            # Distinct doesn't sum across days, so these need their own queries.
                            per_day_prog_sql = org_pd_cte + """
                                SELECT CONVERT(date, a.MeetingDate) AS AttendDate,
                                       p.ProgId, p.Program,
                                       COUNT(DISTINCT a.PeopleId) AS Enrolled,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 THEN a.PeopleId END) AS Total,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 AND a.MemberTypeId IN ({2}) THEN a.PeopleId END) AS Leaders
                                FROM Attend a
                                JOIN OrgPD p ON p.OrgId = a.OrganizationId AND p.rn = 1
                                WHERE a.MeetingDate >= '{0}' AND a.MeetingDate < DATEADD(day,1,'{1}')
                                GROUP BY CONVERT(date, a.MeetingDate), p.ProgId, p.Program
                            """.format(safe_start, safe_end, leader_types_str)
                            for r in q.QuerySql(per_day_prog_sql):
                                ds = normalize_date_str(safe_str(r.AttendDate))
                                pid = safe_int(r.ProgId)
                                daily_prog.setdefault(ds, {})[pid] = {
                                    'leaders': safe_int(r.Leaders),
                                    'total': safe_int(r.Total),
                                    'enrolled': safe_int(r.Enrolled),
                                    'name': safe_str(r.Program)
                                }

                            per_day_div_sql = org_pd_cte + """
                                SELECT CONVERT(date, a.MeetingDate) AS AttendDate,
                                       p.ProgId, p.Program, p.DivId, p.Division,
                                       COUNT(DISTINCT a.PeopleId) AS Enrolled,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 THEN a.PeopleId END) AS Total,
                                       COUNT(DISTINCT CASE WHEN a.AttendanceFlag=1 AND a.MemberTypeId IN ({2}) THEN a.PeopleId END) AS Leaders
                                FROM Attend a
                                JOIN OrgPD p ON p.OrgId = a.OrganizationId AND p.rn = 1
                                WHERE a.MeetingDate >= '{0}' AND a.MeetingDate < DATEADD(day,1,'{1}')
                                GROUP BY CONVERT(date, a.MeetingDate), p.ProgId, p.Program, p.DivId, p.Division
                            """.format(safe_start, safe_end, leader_types_str)
                            for r in q.QuerySql(per_day_div_sql):
                                ds = normalize_date_str(safe_str(r.AttendDate))
                                key = (safe_int(r.ProgId), safe_int(r.DivId))
                                daily_div.setdefault(ds, {})[key] = {
                                    'leaders': safe_int(r.Leaders),
                                    'total': safe_int(r.Total),
                                    'enrolled': safe_int(r.Enrolled),
                                    'prog_name': safe_str(r.Program),
                                    'div_name': safe_str(r.Division)
                                }

                        if need_org_summary:
                            per_org_sql = """
                                SELECT OrganizationId,
                                       COUNT(DISTINCT PeopleId) AS Enrolled,
                                       COUNT(DISTINCT CASE WHEN AttendanceFlag=1 THEN PeopleId END) AS Total,
                                       COUNT(DISTINCT CASE WHEN AttendanceFlag=1 AND MemberTypeId IN ({3}) THEN PeopleId END) AS Leaders
                                FROM Attend
                                WHERE MeetingDate >= '{0}' AND MeetingDate < DATEADD(day,1,'{1}')
                                  AND OrganizationId IN ({2})
                                GROUP BY OrganizationId
                            """.format(safe_start, safe_end, org_ids_str, leader_types_str)
                            for r in q.QuerySql(per_org_sql):
                                org_distinct[r.OrganizationId] = {
                                    'leaders': safe_int(r.Leaders),
                                    'total': safe_int(r.Total),
                                    'enrolled': safe_int(r.Enrolled)
                                }
                    elif sorted_dates:
                        # --- Sum mode (default): roll up from main query in Python ---
                        # Daily: sum across orgs of per-(org,date) distinct counts
                        for ds in sorted_dates:
                            dd = by_date.get(ds, {'leaders': 0, 'members': 0, 'total': 0, 'enrolled_sum': 0})
                            daily_distinct[ds] = {
                                'leaders': dd['leaders'],
                                'total': dd['leaders'] + dd['members'],
                                'enrolled': dd.get('enrolled_sum', 0)
                            }
                        # Period: sum of dailies
                        period_distinct = {
                            'leaders': sum(daily_distinct[d]['leaders'] for d in sorted_dates if d in daily_distinct),
                            'total': sum(daily_distinct[d]['total'] for d in sorted_dates if d in daily_distinct),
                            'enrolled': sum(daily_distinct[d]['enrolled'] for d in sorted_dates if d in daily_distinct)
                        }
                        # Per-org: sum across dates from by_org
                        if need_org_summary:
                            for oid, odata in by_org.items():
                                o_l = sum(d['leaders'] for d in odata['dates'].values())
                                o_m = sum(d['members'] for d in odata['dates'].values())
                                o_e = sum(d.get('enrolled', 0) for d in odata['dates'].values())
                                org_distinct[oid] = {
                                    'leaders': o_l,
                                    'total': o_l + o_m,
                                    'enrolled': o_e
                                }
                        # Per-program & per-(prog,div): roll up from by_prog/by_div + by_org for enrolled
                        if need_prog_div:
                            # Build prog -> set of org_ids and (prog,div) -> set of org_ids
                            prog_orgs = {}
                            div_orgs_map = {}
                            for oid, oprg in org_program.items():
                                pid = oprg.get('progId', 0)
                                did = oprg.get('divId', 0)
                                prog_orgs.setdefault(pid, []).append(oid)
                                div_orgs_map.setdefault((pid, did), []).append(oid)

                            for pid, pdata in by_prog.items():
                                p_l = sum(d['leaders'] for d in pdata['dates'].values())
                                p_m = sum(d['members'] for d in pdata['dates'].values())
                                p_e = sum(
                                    by_org[oid]['dates'].get(ds, {}).get('enrolled', 0)
                                    for oid in prog_orgs.get(pid, [])
                                    for ds in sorted_dates
                                )
                                prog_distinct[pid] = {
                                    'leaders': p_l,
                                    'total': p_l + p_m,
                                    'enrolled': p_e,
                                    'name': pdata['name']
                                }

                            for did, ddata in by_div.items():
                                pid = ddata['progId']
                                d_l = sum(d['leaders'] for d in ddata['dates'].values())
                                d_m = sum(d['members'] for d in ddata['dates'].values())
                                d_e = sum(
                                    by_org[oid]['dates'].get(ds, {}).get('enrolled', 0)
                                    for oid in div_orgs_map.get((pid, did), [])
                                    for ds in sorted_dates
                                )
                                div_distinct[(pid, did)] = {
                                    'leaders': d_l,
                                    'total': d_l + d_m,
                                    'enrolled': d_e,
                                    'prog_name': ddata['progName'],
                                    'div_name': ddata['name']
                                }

                            # Per-(day, prog) and per-(day, prog, div) sum-mode population
                            for ds in sorted_dates:
                                # Per-program for this day
                                for pid, pdata in by_prog.items():
                                    pdd = pdata['dates'].get(ds)
                                    if not pdd:
                                        continue
                                    p_e = sum(
                                        by_org[oid]['dates'].get(ds, {}).get('enrolled', 0)
                                        for oid in prog_orgs.get(pid, [])
                                    )
                                    daily_prog.setdefault(ds, {})[pid] = {
                                        'leaders': pdd['leaders'],
                                        'total': pdd['leaders'] + pdd['members'],
                                        'enrolled': p_e,
                                        'name': pdata['name']
                                    }
                                # Per-(prog, div) for this day
                                for did, ddata in by_div.items():
                                    pid = ddata['progId']
                                    ddd = ddata['dates'].get(ds)
                                    if not ddd:
                                        continue
                                    d_e = sum(
                                        by_org[oid]['dates'].get(ds, {}).get('enrolled', 0)
                                        for oid in div_orgs_map.get((pid, did), [])
                                    )
                                    daily_div.setdefault(ds, {})[(pid, did)] = {
                                        'leaders': ddd['leaders'],
                                        'total': ddd['leaders'] + ddd['members'],
                                        'enrolled': d_e,
                                        'prog_name': ddata['progName'],
                                        'div_name': ddata['name']
                                    }

                    # In distinct mode the per-grain dicts (daily_distinct, prog_distinct,
                    # etc.) intentionally do NOT include headcount. HC is rendered via the
                    # separate hc_per_* dicts as its own column + Grand Total column.

                    # Ensure distinct-mode dicts have entries for HC-only orgs / dates so
                    # they show up in the rendered tables (with attendance values of 0).
                    if count_distinct and hc_by_org_date:
                        for ds in hc_per_day:
                            if ds not in daily_distinct:
                                daily_distinct[ds] = {'leaders': 0, 'total': 0, 'enrolled': 0}
                        if need_org_summary:
                            for oid in hc_per_org:
                                if oid not in org_distinct:
                                    org_distinct[oid] = {'leaders': 0, 'total': 0, 'enrolled': 0}
                        if need_prog_div:
                            for pid in hc_per_prog:
                                if pid not in prog_distinct:
                                    pname = ''
                                    for m in org_meta.values():
                                        if m.get('progId') == pid:
                                            pname = m.get('progName', '')
                                            break
                                    prog_distinct[pid] = {'leaders': 0, 'total': 0, 'enrolled': 0, 'name': pname}
                            for key in hc_per_div:
                                if key not in div_distinct:
                                    pid_k, did_k = key
                                    pname = ''
                                    dname = ''
                                    for m in org_meta.values():
                                        if m.get('progId') == pid_k and m.get('divId') == did_k:
                                            pname = m.get('progName', '')
                                            dname = m.get('divName', '')
                                            break
                                    div_distinct[key] = {'leaders': 0, 'total': 0, 'enrolled': 0,
                                                         'prog_name': pname, 'div_name': dname}
                            for (ds, pid) in hc_per_day_prog:
                                if ds not in daily_prog:
                                    daily_prog[ds] = {}
                                if pid not in daily_prog[ds]:
                                    pname = ''
                                    for m in org_meta.values():
                                        if m.get('progId') == pid:
                                            pname = m.get('progName', '')
                                            break
                                    daily_prog[ds][pid] = {'leaders': 0, 'total': 0, 'enrolled': 0, 'name': pname}
                            for (ds, key) in hc_per_day_div:
                                if ds not in daily_div:
                                    daily_div[ds] = {}
                                if key not in daily_div[ds]:
                                    pid_k, did_k = key
                                    pname = ''
                                    dname = ''
                                    for m in org_meta.values():
                                        if m.get('progId') == pid_k and m.get('divId') == did_k:
                                            pname = m.get('progName', '')
                                            dname = m.get('divName', '')
                                            break
                                    daily_div[ds][key] = {'leaders': 0, 'total': 0, 'enrolled': 0,
                                                          'prog_name': pname, 'div_name': dname}

                    # Step 5: Subgroup query (conditional)
                    subgroups = {}  # org_id -> {subgroup_name -> {date_str -> count}}
                    if sections.get('subgroupBreakdown', False):
                        sg_sql = """
                            SELECT a.OrganizationId,
                                   ISNULL(mt.Name, '(No Subgroup)') AS SubGroupName,
                                   CONVERT(date, a.MeetingDate) AS AttendDate,
                                   COUNT(DISTINCT a.PeopleId) AS AttendCount
                            FROM Attend a
                            LEFT JOIN OrgMemMemTags ommt ON a.PeopleId = ommt.PeopleId AND a.OrganizationId = ommt.OrgId
                            LEFT JOIN MemberTags mt ON ommt.MemberTagId = mt.Id
                            WHERE a.AttendanceFlag = 1
                              AND a.MeetingDate >= '{0}' AND a.MeetingDate < DATEADD(day,1,'{1}')
                              AND a.OrganizationId IN ({2})
                            GROUP BY a.OrganizationId, mt.Name, CONVERT(date, a.MeetingDate)
                        """.format(safe_start, safe_end, org_ids_str)
                        for r in q.QuerySql(sg_sql):
                            oid = r.OrganizationId
                            sg_name = safe_str(r.SubGroupName)
                            ds = normalize_date_str(safe_str(r.AttendDate))
                            cnt = safe_int(r.AttendCount)
                            if oid not in subgroups:
                                subgroups[oid] = {}
                            if sg_name not in subgroups[oid]:
                                subgroups[oid][sg_name] = {}
                            subgroups[oid][sg_name][ds] = cnt

                    # Step 6: Build HTML report
                    html_parts = []
                    sep_leaders = sections.get('separateLeaders', True)

                    # Report Header
                    html_parts.append('<div class="ab-report">')
                    html_parts.append('<div class="ab-rpt-header">')
                    html_parts.append('<h2 class="ab-rpt-title">{0}</h2>'.format(html_escape(print_title)))
                    html_parts.append('<div class="ab-rpt-meta">{0} to {1} &bull; {2} organizations &bull; {3} days</div>'.format(
                        html_escape(format_date_short(safe_start)), html_escape(format_date_short(safe_end)),
                        len(orgs_with_data), len(sorted_dates)))
                    html_parts.append('</div>')

                    # Headcount columns (Headcount + Grand Total) only render when there's
                    # at least one headcount in scope. Grand Total = Total + Headcount.
                    has_hc = hc_period > 0

                    # Section 1: Daily Summary Table
                    if sections.get('dailySummary', True) and sorted_dates:
                        html_parts.append('<div class="ab-section">')
                        html_parts.append('<h3 class="ab-section-title">Daily Summary</h3>')
                        html_parts.append('<table class="ab-table">')
                        # Header row
                        html_parts.append('<thead><tr>')
                        html_parts.append('<th>Date</th><th>Day</th>')
                        if sep_leaders:
                            html_parts.append('<th>Leaders</th><th>Members</th>')
                        html_parts.append('<th>Total</th>')
                        if has_hc:
                            html_parts.append('<th>Headcount<sup>*</sup></th><th>Grand Total</th>')
                        if sections.get('enrollmentRatio', False):
                            html_parts.append('<th>Enrolled</th><th>Ratio</th>')
                        html_parts.append('</tr></thead>')

                        html_parts.append('<tbody>')
                        # Daily rows use distinct-people counts (no cross-org duplicates)
                        # Each row gets an expand toggle that reveals indented program/division
                        # rows for that specific day (uses daily_prog/daily_div).
                        ds_show_ratio = sections.get('enrollmentRatio', False)
                        ds_has_prog_div = bool(daily_prog) or bool(daily_div)
                        for ds_idx, ds in enumerate(sorted_dates):
                            dist = daily_distinct.get(ds, {'leaders': 0, 'total': 0, 'enrolled': 0})
                            day_total = dist['total']
                            day_leaders = dist['leaders']
                            day_members = day_total - day_leaders
                            day_enrolled = dist['enrolled']
                            day_hc = hc_per_day.get(ds, 0)
                            day_grand = day_total + day_hc
                            day_detail_id = 'abDayDetail_{0}'.format(ds_idx)
                            html_parts.append('<tr>')
                            # Date cell - clickable expand toggle when prog/div data exists
                            if ds_has_prog_div:
                                html_parts.append((
                                    '<td>'
                                    '<span onclick="var e=document.getElementById(\'{0}\');e.style.display=e.style.display===\'none\'?\'table-row-group\':\'none\';this.innerHTML=e.style.display===\'none\'?\'&#9654;\':\'&#9660;\'" '
                                    'style="color:#7c3aed;cursor:pointer;font-size:10px;display:inline-block;width:14px;text-align:center;">&#9654;</span> '
                                    '{1}'
                                    '</td>'
                                ).format(day_detail_id, html_escape(format_date_short(ds))))
                            else:
                                html_parts.append('<td>{0}</td>'.format(html_escape(format_date_short(ds))))
                            html_parts.append('<td>{0}</td>'.format(html_escape(get_day_name(ds))))
                            if sep_leaders:
                                html_parts.append('<td>{0}</td><td>{1}</td>'.format(format_number(day_leaders), format_number(day_members)))
                            html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(day_total)))
                            if has_hc:
                                html_parts.append('<td>{0}</td><td><strong>{1}</strong></td>'.format(
                                    format_number(day_hc) if day_hc else '-',
                                    format_number(day_grand)))
                            if ds_show_ratio:
                                html_parts.append('<td>{0}</td><td>{1}</td>'.format(format_number(day_enrolled), format_pct(day_total, day_enrolled)))
                            html_parts.append('</tr>')

                            # Hidden detail group: per-program then per-division for this day
                            if ds_has_prog_div:
                                html_parts.append('<tbody id="{0}" style="display:none;">'.format(day_detail_id))
                                day_progs = daily_prog.get(ds, {})
                                day_divs = daily_div.get(ds, {})
                                for pid in sorted(day_progs.keys(), key=lambda p: day_progs[p]['name']):
                                    pdd = day_progs[pid]
                                    p_total = pdd['total']
                                    p_leaders = pdd['leaders']
                                    p_members = p_total - p_leaders
                                    p_enr = pdd['enrolled']
                                    p_hc = hc_per_day_prog.get((ds, pid), 0)
                                    p_grand = p_total + p_hc
                                    # Program header sub-row
                                    html_parts.append('<tr style="background:#ede9fe;">')
                                    html_parts.append('<td colspan="2" style="padding-left:32px;"><strong>{0}</strong></td>'.format(html_escape(pdd['name'])))
                                    if sep_leaders:
                                        html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(p_leaders), format_number(p_members)))
                                    html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(p_total)))
                                    if has_hc:
                                        html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(
                                            format_number(p_hc) if p_hc else '-',
                                            format_number(p_grand)))
                                    if ds_show_ratio:
                                        html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(p_enr), format_pct(p_total, p_enr)))
                                    html_parts.append('</tr>')
                                    # Division rows under this program for this day
                                    div_keys_for_prog = [k for k in day_divs.keys() if k[0] == pid]
                                    div_keys_for_prog.sort(key=lambda k: day_divs[k]['div_name'])
                                    for k in div_keys_for_prog:
                                        ddd = day_divs[k]
                                        d_total = ddd['total']
                                        d_leaders = ddd['leaders']
                                        d_members = d_total - d_leaders
                                        d_enr = ddd['enrolled']
                                        d_hc = hc_per_day_div.get((ds, k), 0)
                                        d_grand = d_total + d_hc
                                        html_parts.append('<tr style="background:#f8fafc;">')
                                        html_parts.append('<td colspan="2" style="padding-left:64px;font-size:12px;color:#64748b;">{0}</td>'.format(html_escape(ddd['div_name'])))
                                        if sep_leaders:
                                            html_parts.append('<td style="font-size:12px;color:#64748b;">{0}</td><td style="font-size:12px;color:#64748b;">{1}</td>'.format(format_number(d_leaders), format_number(d_members)))
                                        html_parts.append('<td style="font-size:12px;"><strong>{0}</strong></td>'.format(format_number(d_total)))
                                        if has_hc:
                                            html_parts.append('<td style="font-size:12px;color:#64748b;">{0}</td><td style="font-size:12px;"><strong>{1}</strong></td>'.format(
                                                format_number(d_hc) if d_hc else '-',
                                                format_number(d_grand)))
                                        if ds_show_ratio:
                                            html_parts.append('<td style="font-size:12px;color:#64748b;">{0}</td><td style="font-size:12px;color:#64748b;">{1}</td>'.format(format_number(d_enr), format_pct(d_total, d_enr)))
                                        html_parts.append('</tr>')
                                html_parts.append('</tbody>')

                        # TOTAL row uses period-distinct counts (unique people across whole period)
                        period_total = period_distinct['total']
                        period_leaders = period_distinct['leaders']
                        period_members = period_total - period_leaders
                        period_enrolled = period_distinct['enrolled']
                        period_grand = period_total + hc_period
                        html_parts.append('<tr class="ab-total-row">')
                        html_parts.append('<td colspan="2"><strong>TOTAL</strong></td>')
                        if sep_leaders:
                            html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(period_leaders), format_number(period_members)))
                        html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(period_total)))
                        if has_hc:
                            html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(hc_period), format_number(period_grand)))
                        if sections.get('enrollmentRatio', False):
                            html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(period_enrolled), format_pct(period_total, period_enrolled)))
                        html_parts.append('</tr>')

                        # Averages row = mean of daily distinct counts
                        if len(sorted_dates) > 1:
                            n = len(sorted_dates)
                            sum_total = sum(daily_distinct[d]['total'] for d in sorted_dates if d in daily_distinct)
                            sum_leaders = sum(daily_distinct[d]['leaders'] for d in sorted_dates if d in daily_distinct)
                            sum_enrolled = sum(daily_distinct[d]['enrolled'] for d in sorted_dates if d in daily_distinct)
                            avg_total = int(sum_total / n)
                            avg_leaders = int(sum_leaders / n)
                            avg_members = int((sum_total - sum_leaders) / n)
                            avg_enrolled = int(sum_enrolled / n)
                            avg_hc = int(hc_period / n)
                            avg_grand = avg_total + avg_hc
                            html_parts.append('<tr class="ab-avg-row">')
                            html_parts.append('<td colspan="2"><em>Average/day</em></td>')
                            if sep_leaders:
                                html_parts.append('<td><em>{0}</em></td><td><em>{1}</em></td>'.format(format_number(avg_leaders), format_number(avg_members)))
                            html_parts.append('<td><em>{0}</em></td>'.format(format_number(avg_total)))
                            if has_hc:
                                html_parts.append('<td><em>{0}</em></td><td><em>{1}</em></td>'.format(format_number(avg_hc), format_number(avg_grand)))
                            if sections.get('enrollmentRatio', False):
                                html_parts.append('<td><em>{0}</em></td><td><em>{1}</em></td>'.format(format_number(avg_enrolled), format_pct(avg_total, avg_enrolled)))
                            html_parts.append('</tr>')

                        html_parts.append('</tbody></table>')
                        html_parts.append('</div>')

                    # Section 2: Program/Division Breakdown
                    # Distinct-people summary per program/division for the whole period.
                    # Same column layout as Daily Summary; Headcount + Grand Total appear
                    # only when there's at least one headcount in scope.
                    if sections.get('programDivisionBreakdown', True) and prog_distinct:
                        show_ratio = sections.get('enrollmentRatio', False)
                        html_parts.append('<div class="ab-section">')
                        html_parts.append('<h3 class="ab-section-title">Program / Division Breakdown</h3>')
                        html_parts.append('<table class="ab-table">')
                        html_parts.append('<thead><tr><th>Program / Division</th>')
                        if sep_leaders:
                            html_parts.append('<th>Leaders</th><th>Members</th>')
                        html_parts.append('<th>Total</th>')
                        if has_hc:
                            html_parts.append('<th>Headcount<sup>*</sup></th><th>Grand Total</th>')
                        if show_ratio:
                            html_parts.append('<th>Enrolled</th><th>Ratio</th>')
                        html_parts.append('</tr></thead><tbody>')

                        div_counter = 0
                        for prog_id_key in sorted(prog_distinct.keys(), key=lambda p: prog_distinct[p]['name']):
                            pd = prog_distinct[prog_id_key]
                            p_total = pd['total']
                            p_leaders = pd['leaders']
                            p_members = p_total - p_leaders
                            p_enrolled = pd['enrolled']
                            p_hc = hc_per_prog.get(prog_id_key, 0)
                            p_grand = p_total + p_hc

                            html_parts.append('<tr style="background:#ede9fe;">')
                            html_parts.append('<td><strong>{0}</strong></td>'.format(html_escape(pd['name'])))
                            if sep_leaders:
                                html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(p_leaders), format_number(p_members)))
                            html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(p_total)))
                            if has_hc:
                                html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(
                                    format_number(p_hc) if p_hc else '-',
                                    format_number(p_grand)))
                            if show_ratio:
                                html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(p_enrolled), format_pct(p_total, p_enrolled)))
                            html_parts.append('</tr>')

                            prog_divs = [(did, dd) for (pid, did), dd in div_distinct.items() if pid == prog_id_key]
                            prog_divs.sort(key=lambda x: x[1]['div_name'])
                            for did, dd in prog_divs:
                                d_total = dd['total']
                                d_leaders = dd['leaders']
                                d_members = d_total - d_leaders
                                d_enrolled = dd['enrolled']
                                d_hc = hc_per_div.get((prog_id_key, did), 0)
                                d_grand = d_total + d_hc
                                div_orgs = sorted(
                                    [(oid, odata['name']) for oid, odata in by_org.items()
                                     if org_program.get(oid, {}).get('divId') == did
                                     and org_program.get(oid, {}).get('progId') == prog_id_key],
                                    key=lambda x: x[1]
                                )
                                has_orgs = len(div_orgs) > 0
                                detail_id = 'abDivDetail_{0}'.format(div_counter)
                                div_counter += 1

                                html_parts.append('<tr>')
                                html_parts.append('<td style="padding-left:16px;">')
                                if has_orgs:
                                    html_parts.append('<span onclick="var e=document.getElementById(\'{0}\');e.style.display=e.style.display===\'none\'?\'table-row-group\':\'none\';this.innerHTML=e.style.display===\'none\'?\'&#9654;\':\'&#9660;\'" style="color:#7c3aed;cursor:pointer;font-size:10px;display:inline-block;width:14px;text-align:center;">&#9654;</span> '.format(detail_id))
                                else:
                                    html_parts.append('<span style="display:inline-block;width:14px;"></span> ')
                                html_parts.append(html_escape(dd['div_name']))
                                html_parts.append('</td>')
                                if sep_leaders:
                                    html_parts.append('<td>{0}</td><td>{1}</td>'.format(format_number(d_leaders), format_number(d_members)))
                                html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(d_total)))
                                if has_hc:
                                    html_parts.append('<td>{0}</td><td><strong>{1}</strong></td>'.format(
                                        format_number(d_hc) if d_hc else '-',
                                        format_number(d_grand)))
                                if show_ratio:
                                    html_parts.append('<td>{0}</td><td>{1}</td>'.format(format_number(d_enrolled), format_pct(d_total, d_enrolled)))
                                html_parts.append('</tr>')

                                if has_orgs:
                                    html_parts.append('<tbody id="{0}" style="display:none;">'.format(detail_id))
                                    for oid, oname in div_orgs:
                                        od = org_distinct.get(oid, {'leaders': 0, 'total': 0, 'enrolled': 0})
                                        o_total = od['total']
                                        o_leaders = od['leaders']
                                        o_members = o_total - o_leaders
                                        o_enrolled = od['enrolled']
                                        o_hc = hc_per_org.get(oid, 0)
                                        o_grand = o_total + o_hc
                                        html_parts.append('<tr style="background:#f8fafc;">')
                                        html_parts.append('<td style="padding-left:48px;font-size:12px;color:#64748b;">{0}</td>'.format(html_escape(oname)))
                                        if sep_leaders:
                                            html_parts.append('<td style="font-size:12px;color:#64748b;">{0}</td><td style="font-size:12px;color:#64748b;">{1}</td>'.format(format_number(o_leaders), format_number(o_members)))
                                        html_parts.append('<td style="font-size:12px;"><strong>{0}</strong></td>'.format(format_number(o_total)))
                                        if has_hc:
                                            html_parts.append('<td style="font-size:12px;color:#64748b;">{0}</td><td style="font-size:12px;"><strong>{1}</strong></td>'.format(
                                                format_number(o_hc) if o_hc else '-',
                                                format_number(o_grand)))
                                        if show_ratio:
                                            html_parts.append('<td style="font-size:12px;color:#64748b;">{0}</td><td style="font-size:12px;color:#64748b;">{1}</td>'.format(format_number(o_enrolled), format_pct(o_total, o_enrolled)))
                                        html_parts.append('</tr>')
                                    html_parts.append('</tbody>')

                        html_parts.append('</tbody></table>')
                        html_parts.append('</div>')

                    # Section 3: Org Detail Cards (collapsed by default; click header to expand)
                    if sections.get('orgDetail', False) and by_org:
                        html_parts.append('<div class="ab-section">')
                        html_parts.append('<h3 class="ab-section-title">Organization Details ')
                        html_parts.append('<span class="ab-section-actions">')
                        html_parts.append('<button type="button" class="ab-link-btn" onclick="ABApp.toggleAllOrgs(true)">Expand all</button> | ')
                        html_parts.append('<button type="button" class="ab-link-btn" onclick="ABApp.toggleAllOrgs(false)">Collapse all</button>')
                        html_parts.append('</span></h3>')

                        show_ratio = sections.get('enrollmentRatio', False)
                        sorted_org_keys = sorted(by_org.keys(), key=lambda x: by_org[x]['name'])
                        for oid in sorted_org_keys:
                            odata = by_org[oid]
                            oprg = org_program.get(oid, {})
                            # Header summary: distinct unique people across the period
                            od = org_distinct.get(oid, {'leaders': 0, 'total': 0, 'enrolled': 0})
                            org_unique_attended = od['total']
                            org_unique_enrolled = od['enrolled']
                            # Per-row table totals: sum across dates (matches per-day distinct math)
                            org_total_att = sum(d['leaders'] + d['members'] for d in odata['dates'].values())
                            org_enrolled = sum(d.get('enrolled', 0) for d in odata['dates'].values())

                            org_card_hc = hc_per_org.get(oid, 0)
                            html_parts.append('<div class="ab-org-card">')
                            # Clickable header toggles next-sibling body
                            html_parts.append('<div class="ab-org-card-header ab-org-card-toggle" onclick="ABApp.toggleOrgCard(this)">')
                            html_parts.append('<span class="ab-org-card-title"><span class="ab-org-toggle-icon">&#9654;</span> <strong>{0}</strong></span>'.format(html_escape(odata['name'])))
                            html_parts.append('<span class="ab-org-card-summary">{0} attended'.format(format_number(org_unique_attended)))
                            if show_ratio and org_unique_enrolled > 0:
                                html_parts.append(' / {0} enrolled'.format(format_number(org_unique_enrolled)))
                            if org_card_hc > 0:
                                html_parts.append(' / +{0} headcount'.format(format_number(org_card_hc)))
                            html_parts.append('</span>')
                            html_parts.append('<span class="ab-org-card-meta">{0} / {1}</span>'.format(
                                html_escape(oprg.get('progName', '')), html_escape(oprg.get('divName', ''))))
                            html_parts.append('</div>')

                            html_parts.append('<div class="ab-org-card-body" style="display:none;">')
                            html_parts.append('<table class="ab-table ab-table-sm">')
                            html_parts.append('<thead><tr><th>Date</th><th>Day</th>')
                            if sep_leaders:
                                html_parts.append('<th>Leaders</th><th>Members</th>')
                            html_parts.append('<th>Total</th>')
                            if has_hc:
                                html_parts.append('<th>Headcount<sup>*</sup></th><th>Grand Total</th>')
                            if show_ratio:
                                html_parts.append('<th>Enrolled</th><th>Ratio</th>')
                            html_parts.append('</tr></thead><tbody>')

                            org_sum_l = 0
                            org_sum_m = 0
                            org_sum_hc = 0
                            for ds in sorted_dates:
                                dd = odata['dates'].get(ds, None)
                                row_hc = hc_by_org_date.get((oid, ds), 0)
                                # Render row if there's any attend data OR a headcount for this (org, date)
                                if dd or row_hc > 0:
                                    dl = dd['leaders'] if dd else 0
                                    dm = dd['members'] if dd else 0
                                    dt = dl + dm
                                    de = (dd.get('enrolled', 0) if dd else 0)
                                    grand = dt + row_hc
                                    org_sum_l += dl
                                    org_sum_m += dm
                                    org_sum_hc += row_hc
                                    html_parts.append('<tr>')
                                    html_parts.append('<td>{0}</td><td>{1}</td>'.format(html_escape(format_date_short(ds)), html_escape(get_day_name(ds))))
                                    if sep_leaders:
                                        html_parts.append('<td>{0}</td><td>{1}</td>'.format(format_number(dl), format_number(dm)))
                                    html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(dt)))
                                    if has_hc:
                                        html_parts.append('<td>{0}</td><td><strong>{1}</strong></td>'.format(
                                            format_number(row_hc) if row_hc else '-',
                                            format_number(grand)))
                                    if show_ratio:
                                        html_parts.append('<td>{0}</td><td>{1}</td>'.format(format_number(de), format_pct(dt, de)))
                                    html_parts.append('</tr>')

                            # Org total row
                            org_grand = org_total_att + org_sum_hc
                            html_parts.append('<tr class="ab-total-row">')
                            html_parts.append('<td colspan="2"><strong>Total</strong></td>')
                            if sep_leaders:
                                html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(org_sum_l), format_number(org_sum_m)))
                            html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(org_total_att)))
                            if has_hc:
                                html_parts.append('<td><strong>{0}</strong></td><td><strong>{1}</strong></td>'.format(format_number(org_sum_hc), format_number(org_grand)))
                            if show_ratio:
                                html_parts.append('<td><strong>{0}</strong></td><td></td>'.format(format_number(org_enrolled)))
                            html_parts.append('</tr>')

                            html_parts.append('</tbody></table>')
                            html_parts.append('</div>')  # end ab-org-card-body
                            html_parts.append('</div>')  # end ab-org-card

                        html_parts.append('</div>')

                    # Section 4: Subgroup Breakdown
                    if sections.get('subgroupBreakdown', False) and subgroups:
                        html_parts.append('<div class="ab-section">')
                        html_parts.append('<h3 class="ab-section-title">Subgroup Breakdown</h3>')

                        for oid in sorted(subgroups.keys(), key=lambda x: by_org.get(x, {}).get('name', '')):
                            sg_data = subgroups[oid]
                            org_name = by_org.get(oid, {}).get('name', 'Org ' + str(oid))
                            html_parts.append('<div class="ab-org-card">')
                            html_parts.append('<div class="ab-org-card-header"><strong>{0}</strong></div>'.format(html_escape(org_name)))
                            html_parts.append('<table class="ab-table ab-table-sm">')
                            html_parts.append('<thead><tr><th>Subgroup</th>')
                            for ds in sorted_dates:
                                html_parts.append('<th>{0}</th>'.format(html_escape(format_date_short(ds))))
                            html_parts.append('<th>Total</th></tr></thead><tbody>')

                            for sg_name in sorted(sg_data.keys()):
                                sg_dates = sg_data[sg_name]
                                sg_total = sum(sg_dates.values())
                                html_parts.append('<tr>')
                                html_parts.append('<td>{0}</td>'.format(html_escape(sg_name)))
                                for ds in sorted_dates:
                                    cnt = sg_dates.get(ds, 0)
                                    html_parts.append('<td>{0}</td>'.format(format_number(cnt) if cnt else '-'))
                                html_parts.append('<td><strong>{0}</strong></td>'.format(format_number(sg_total)))
                                html_parts.append('</tr>')

                            html_parts.append('</tbody></table>')
                            html_parts.append('</div>')

                        html_parts.append('</div>')

                    # Footnote when headcount data is present
                    if has_hc:
                        html_parts.append(
                            '<div class="ab-footnote" style="margin-top:16px;padding:10px 14px;'
                            'background:#fef3c7;border-left:3px solid #d97706;font-size:12px;color:#78350f;">'
                            '<strong>*</strong> <strong>Headcount</strong> represents people counted in bulk on '
                            'meetings without individual roll (e.g., a class that recorded only a total head count). '
                            'When a meeting has individual attendance and the headcount is higher, only the difference '
                            '(headcount - individual) is counted here. <em>Headcount feeds the Grand Total but is '
                            'excluded from Total, Enrolled, and the Enrollment Ratio</em> - those reflect '
                            'individually-tracked people only.'
                            '</div>'
                        )

                    html_parts.append('</div>')  # close ab-report

                    stats = {
                        'orgCount': len(orgs_with_data),
                        'totalAttendance': period_distinct.get('total', 0),
                        'dayCount': len(sorted_dates),
                        'totalEnrolled': period_distinct.get('enrolled', 0)
                    }
                    print json.dumps({'success': True, 'html': ''.join(html_parts), 'stats': stats})
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
   Attendance Report Builder - Scoped CSS (ab- prefix)
   ================================================================ */
.ab-root {
    --ab-primary: #7c3aed;
    --ab-primary-light: #ede9fe;
    --ab-secondary: #0d9488;
    --ab-secondary-light: #ccfbf1;
    --ab-warning: #d97706;
    --ab-warning-light: #fef3c7;
    --ab-danger: #dc2626;
    --ab-danger-light: #fee2e2;
    --ab-dark: #1e293b;
    --ab-gray: #64748b;
    --ab-light-bg: #f8fafc;
    --ab-border: #e2e8f0;
    --ab-radius: 12px;
    --ab-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
    --ab-shadow-lg: 0 10px 15px rgba(0,0,0,0.1), 0 4px 6px rgba(0,0,0,0.05);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    color: var(--ab-dark);
    max-width: 1200px;
    margin: 0 auto;
    padding: 12px;
}

/* ---- Landing ---- */
.ab-landing {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 60vh;
    gap: 24px;
}
.ab-landing-title { font-size: 32px; font-weight: 700; text-align: center; margin-bottom: 8px; }
.ab-landing-subtitle { font-size: 18px; color: var(--ab-gray); text-align: center; margin-bottom: 32px; }
.ab-landing-buttons { display: flex; gap: 24px; flex-wrap: wrap; justify-content: center; }
.ab-landing-btn {
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    width: 260px; height: 200px; border-radius: var(--ab-radius); border: 2px solid var(--ab-border);
    background: white; cursor: pointer; transition: all 0.2s; text-decoration: none;
    color: var(--ab-dark); box-shadow: var(--ab-shadow); user-select: none;
}
.ab-landing-btn:active { transform: scale(0.97); box-shadow: var(--ab-shadow-lg); }
.ab-landing-btn i { font-size: 48px; margin-bottom: 16px; }
.ab-landing-btn span { font-size: 22px; font-weight: 600; }
.ab-landing-btn.ab-admin-btn { border-color: var(--ab-primary); }
.ab-landing-btn.ab-admin-btn i { color: var(--ab-primary); }
.ab-landing-btn.ab-gen-btn { border-color: var(--ab-secondary); }
.ab-landing-btn.ab-gen-btn i { color: var(--ab-secondary); }

/* ---- Common UI ---- */
.ab-btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 8px;
    padding: 10px 20px; border-radius: 8px; border: none; font-size: 15px; font-weight: 600;
    cursor: pointer; transition: all 0.15s; user-select: none;
}
.ab-btn:active { transform: scale(0.97); }
.ab-btn-primary { background: var(--ab-primary); color: white; }
.ab-btn-success { background: var(--ab-secondary); color: white; }
.ab-btn-warning { background: var(--ab-warning); color: white; }
.ab-btn-danger { background: var(--ab-danger); color: white; }
.ab-btn-outline { background: white; color: var(--ab-dark); border: 2px solid var(--ab-border); }
.ab-btn-sm { padding: 6px 14px; font-size: 13px; }
.ab-btn:disabled { opacity: 0.5; cursor: not-allowed; }

.ab-input {
    width: 100%; padding: 10px 14px; border: 2px solid var(--ab-border); border-radius: 8px;
    font-size: 15px; outline: none; transition: border-color 0.15s; box-sizing: border-box;
}
.ab-input:focus { border-color: var(--ab-primary); }

.ab-select {
    width: 100%; padding: 10px 14px; border: 2px solid var(--ab-border); border-radius: 8px;
    font-size: 15px; background: white; box-sizing: border-box;
}

.ab-label { display: block; font-size: 13px; font-weight: 600; color: var(--ab-gray); margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px; }
.ab-help-text { font-size: 12px; color: var(--ab-gray); margin-top: 4px; }

.ab-panel {
    background: white; border-radius: var(--ab-radius); border: 1px solid var(--ab-border);
    box-shadow: var(--ab-shadow); margin-bottom: 16px; overflow: hidden;
}
.ab-panel-header {
    display: flex; align-items: center; justify-content: space-between; padding: 16px 20px;
    background: var(--ab-light-bg); border-bottom: 1px solid var(--ab-border); font-weight: 600; font-size: 16px;
}
.ab-panel-body { padding: 20px; }

/* ---- Back button ---- */
.ab-back-btn {
    display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px;
    color: var(--ab-primary); font-size: 16px; font-weight: 600; cursor: pointer;
    border: none; background: none;
}

/* ---- Config cards (admin list) ---- */
.ab-config-card {
    display: flex; align-items: center; justify-content: space-between; padding: 16px 20px;
    border: 1px solid var(--ab-border); border-radius: var(--ab-radius); margin-bottom: 12px;
    background: white; cursor: pointer; transition: all 0.15s;
}
.ab-config-card:hover { background: var(--ab-light-bg); }
.ab-config-name { font-size: 18px; font-weight: 600; }
.ab-config-meta { font-size: 13px; color: var(--ab-gray); margin-top: 4px; }
.ab-config-actions { display: flex; gap: 8px; }

/* ---- Form groups ---- */
.ab-form-group { margin-bottom: 20px; }
.ab-form-row { display: flex; gap: 16px; flex-wrap: wrap; }
.ab-form-row > * { flex: 1; min-width: 200px; }

/* ---- Toggle switch ---- */
.ab-toggle { position: relative; width: 44px; height: 24px; display: inline-block; vertical-align: middle; }
.ab-toggle input { display: none; }
.ab-toggle-slider {
    position: absolute; top: 0; left: 0; right: 0; bottom: 0;
    background: #cbd5e0; border-radius: 24px; cursor: pointer; transition: 0.2s;
}
.ab-toggle-slider:before {
    content: ""; position: absolute; width: 20px; height: 20px;
    left: 2px; bottom: 2px; background: white; border-radius: 50%; transition: 0.2s;
}
.ab-toggle input:checked + .ab-toggle-slider { background: var(--ab-secondary); }
.ab-toggle input:checked + .ab-toggle-slider:before { transform: translateX(20px); }

.ab-toggle-row {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 0; border-bottom: 1px solid #f0f0f0;
}
.ab-toggle-row:last-child { border-bottom: none; }
.ab-toggle-label { font-size: 14px; color: var(--ab-dark); }

/* ---- Search box ---- */
.ab-search-box { position: relative; }
.ab-search-results {
    position: absolute; top: 100%; left: 0; right: 0; background: white;
    border: 1px solid var(--ab-border); border-radius: 0 0 8px 8px;
    box-shadow: var(--ab-shadow-lg); max-height: 300px; overflow-y: auto; z-index: 100;
}
.ab-search-result-item {
    padding: 10px 14px; cursor: pointer; border-bottom: 1px solid var(--ab-border);
}
.ab-search-result-item:hover { background: var(--ab-primary-light); }
.ab-search-result-item:last-child { border-bottom: none; }
.ab-search-result-name { font-weight: 600; }
.ab-search-result-meta { font-size: 12px; color: var(--ab-gray); }

/* ---- Selected orgs list ---- */
.ab-selected-org-item {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px; background: var(--ab-primary-light); border-radius: 8px;
    margin-bottom: 8px; font-weight: 600;
}
.ab-remove-org { cursor: pointer; color: var(--ab-danger); font-size: 18px; padding: 4px 8px; }

/* ---- Report styles ---- */
.ab-report { font-size: 14px; }
.ab-rpt-header { text-align: center; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 2px solid var(--ab-primary); }
.ab-rpt-title { font-size: 24px; margin: 0 0 6px; color: var(--ab-primary); }
.ab-rpt-meta { font-size: 14px; color: var(--ab-gray); }

.ab-section { margin-bottom: 24px; }
.ab-section-title { font-size: 18px; color: var(--ab-primary); border-bottom: 2px solid var(--ab-primary-light); padding-bottom: 6px; margin: 0 0 12px; }

.ab-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 12px; }
.ab-table th { padding: 8px 10px; text-align: left; background: var(--ab-primary); color: white; font-weight: 600; font-size: 12px; }
.ab-table td { padding: 6px 10px; border-bottom: 1px solid #e2e8f0; }
.ab-table tr:hover { background: #f8fafc; }
.ab-table-sm th { padding: 5px 8px; font-size: 11px; }
.ab-table-sm td { padding: 4px 8px; font-size: 12px; }
.ab-total-row { background: var(--ab-warning-light) !important; }
.ab-total-row td { border-top: 2px solid var(--ab-warning); }
.ab-avg-row { background: #f0f9ff !important; font-style: italic; }
.ab-subtotal-row { background: #f1f5f9 !important; }
.ab-subtotal-row td { border-top: 1px solid #94a3b8; }

.ab-prog-block { margin-bottom: 20px; }
.ab-prog-name { font-size: 16px; color: var(--ab-dark); margin: 0 0 8px; padding: 8px 12px; background: var(--ab-primary-light); border-radius: 6px; }
.ab-prog-total { font-weight: 400; font-size: 14px; color: var(--ab-gray); }
.ab-div-block { margin: 0 0 12px 16px; }
.ab-div-name { font-size: 14px; color: var(--ab-secondary); margin: 0 0 6px; }
.ab-div-total { font-weight: 400; font-size: 13px; color: var(--ab-gray); }

.ab-org-card { border: 1px solid var(--ab-border); border-radius: 8px; margin-bottom: 12px; overflow: hidden; }
.ab-org-card-header { padding: 10px 14px; background: var(--ab-light-bg); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 4px; }
.ab-org-card-toggle { cursor: pointer; user-select: none; }
.ab-org-card-toggle:hover { background: var(--ab-primary-light); }
.ab-org-card-title { display: inline-flex; align-items: center; gap: 6px; }
.ab-org-toggle-icon { color: var(--ab-primary); font-size: 11px; display: inline-block; width: 12px; text-align: center; }
.ab-org-card-summary { font-size: 12px; color: var(--ab-secondary); font-weight: 600; }
.ab-org-card-meta { font-size: 12px; color: var(--ab-gray); }
.ab-org-card-body { padding: 0; border-top: 1px solid var(--ab-border); }
.ab-org-card-body .ab-table { margin: 0; }
.ab-org-card-footer { padding: 6px 14px; background: var(--ab-secondary-light); font-size: 12px; color: var(--ab-secondary); }
.ab-section-actions { float: right; font-size: 12px; font-weight: normal; }
.ab-link-btn { background: none; border: none; color: var(--ab-primary); cursor: pointer; font-size: 12px; padding: 0 4px; text-decoration: underline; }
.ab-link-btn:hover { color: var(--ab-secondary); }

/* ---- Preview area ---- */
.ab-preview-area {
    border: 1px solid var(--ab-border); border-radius: var(--ab-radius); padding: 20px;
    background: white; min-height: 400px;
}
.ab-preview-toolbar {
    display: flex; gap: 12px; align-items: center; margin-bottom: 16px;
    padding: 12px; background: var(--ab-light-bg); border-radius: 8px; flex-wrap: wrap;
}
.ab-stats { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.ab-stat-card { padding: 12px 20px; background: var(--ab-light-bg); border-radius: 8px; text-align: center; }
.ab-stat-num { font-size: 24px; font-weight: 700; color: var(--ab-primary); }
.ab-stat-label { font-size: 12px; color: var(--ab-gray); }

/* ---- Toast ---- */
.ab-toast-container {
    position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%);
    z-index: 9999; display: flex; flex-direction: column; align-items: center;
    gap: 8px; pointer-events: none;
}
.ab-toast {
    padding: 14px 28px; border-radius: 10px; font-size: 16px; font-weight: 600;
    color: white; opacity: 0; transform: translateY(20px); transition: all 0.3s;
    pointer-events: auto; max-width: 90vw; text-align: center;
}
.ab-toast.ab-show { opacity: 1; transform: translateY(0); }
.ab-toast-success { background: var(--ab-secondary); }
.ab-toast-danger { background: var(--ab-danger); }
.ab-toast-info { background: var(--ab-primary); }

/* ---- Loading spinner ---- */
.ab-loading { text-align: center; padding: 40px; color: var(--ab-gray); }
.ab-spinner {
    display: inline-block; width: 32px; height: 32px;
    border: 3px solid var(--ab-border); border-top: 3px solid var(--ab-primary);
    border-radius: 50%; animation: abSpin 0.8s linear infinite;
}
@keyframes abSpin { to { transform: rotate(360deg); } }

/* ---- Group detail toggle ---- */
.ab-group-summary { font-size: 13px; color: var(--ab-gray); }
.ab-group-toggle {
    display: inline-flex; align-items: center; gap: 4px; cursor: pointer;
    color: var(--ab-primary); font-size: 12px; font-weight: 600;
    border: none; background: none; padding: 2px 0; margin-left: 6px;
}
.ab-group-toggle:hover { text-decoration: underline; }
.ab-group-detail {
    display: none; margin-top: 6px; padding: 8px 12px;
    background: var(--ab-light-bg); border-radius: 6px; font-size: 12px; color: var(--ab-dark);
}
.ab-group-detail.ab-show { display: block; }
.ab-group-detail-item { padding: 2px 0; }

/* ---- Empty state ---- */
.ab-empty { text-align: center; padding: 40px 20px; color: #a0aec0; }

/* ---- Utilities ---- */
.ab-flex { display: flex; }
.ab-flex-wrap { flex-wrap: wrap; }
.ab-items-center { align-items: center; }
.ab-justify-between { justify-content: space-between; }
.ab-gap-8 { gap: 8px; }
.ab-gap-12 { gap: 12px; }
.ab-gap-16 { gap: 16px; }
.ab-mb-8 { margin-bottom: 8px; }
.ab-mb-12 { margin-bottom: 12px; }
.ab-mb-16 { margin-bottom: 16px; }
.ab-mb-24 { margin-bottom: 24px; }
.ab-mt-12 { margin-top: 12px; }
.ab-mt-16 { margin-top: 16px; }
.ab-text-center { text-align: center; }
.ab-text-muted { color: var(--ab-gray); }
.ab-text-sm { font-size: 14px; }
.ab-d-none { display: none; }

/* ---- Responsive ---- */
@media (max-width: 768px) {
    .ab-landing-buttons { flex-direction: column; align-items: center; }
    .ab-form-row { flex-direction: column; }
    .ab-config-card { flex-direction: column; align-items: flex-start; gap: 12px; }
    .ab-config-actions { width: 100%; justify-content: flex-end; }
}

/* ---- Print ---- */
@media print {
    * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
    .ab-table { border-collapse: collapse !important; }
    .ab-table th, .ab-table td { border: 1px solid #ddd !important; }
}
    </style>
</head>
<body>
<div class="ab-root" id="abApp">
    <div class="ab-toast-container" id="abToastContainer"></div>
    <div id="abContent">
        <div class="ab-loading"><div class="ab-spinner"></div><p>Loading...</p></div>
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
        previewStats: {},
        selectedConfigId: null,
        filtersLoaded: false,
        loading: false,
        overrideStart: '',
        overrideEnd: ''
    };

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
        var container = document.getElementById('abToastContainer');
        var toast = document.createElement('div');
        toast.className = 'ab-toast ab-toast-' + type;
        toast.textContent = msg;
        container.appendChild(toast);
        setTimeout(function() { toast.classList.add('ab-show'); }, 10);
        setTimeout(function() {
            toast.classList.remove('ab-show');
            setTimeout(function() { container.removeChild(toast); }, 300);
        }, 3000);
    }

    function showLoading() {
        state.loading = true;
        var el = document.getElementById('abContent');
        el.innerHTML = '<div class="ab-loading"><div class="ab-spinner"></div><p>Loading...</p></div>';
    }

    function defaultConfig() {
        return {
            id: '',
            name: '',
            printTitle: '',
            sourceType: 'program_division',
            programDivGroups: [],
            specificOrgIds: [],
            specificOrgNames: {},
            excludeOrgIds: '',
            reportSections: {
                dailySummary: true,
                programDivisionBreakdown: true,
                orgDetail: false,
                subgroupBreakdown: false,
                enrollmentRatio: true,
                separateLeaders: true,
                countDistinct: false,
                includeInactive: false
            },
            leaderMemberTypes: '140,160',
            sortBy: 'name',
            _schemaVersion: 2
        };
    }

    function fmtNum(n) {
        if (!n && n !== 0) return '0';
        return n.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',');
    }

    // Build a source summary with optional expandable detail for program/division groups
    // uid should be unique per card to scope the toggle
    function buildSourceSummaryHtml(c, uid) {
        if (c.sourceType === 'specific_orgs') {
            return '<span class="ab-group-summary">' + (c.specificOrgIds ? c.specificOrgIds.length : 0) + ' specific org(s)</span>';
        }
        var groups = c.programDivGroups || [];
        if (!groups.length) {
            // backward compat: old single programId/divisionId
            if (c.programId) {
                var lbl = escHtml(c.programName || 'Program') + ' / ' + escHtml(c.divisionName || 'All Divisions');
                return '<span class="ab-group-summary">' + lbl + '</span>';
            }
            return '<span class="ab-group-summary">No groups configured</span>';
        }
        // Build short summary: first group name(s), plus count
        var first = escHtml(groups[0].programName || 'Program');
        if (groups[0].divisionId) first += ' / ' + escHtml(groups[0].divisionName || 'Division');
        var summary = first;
        if (groups.length > 1) summary += ' + ' + (groups.length - 1) + ' more';

        var h = '<span class="ab-group-summary">' + summary + '</span>';
        if (groups.length > 1) {
            h += '<button class="ab-group-toggle" onclick="event.stopPropagation();ABApp.toggleGroupDetail(\\'' + uid + '\\')">&#9660; Details</button>';
        }
        h += '<div class="ab-group-detail" id="abGrpDetail_' + uid + '">';
        for (var i = 0; i < groups.length; i++) {
            var g = groups[i];
            var glbl = escHtml(g.programName || 'Program ' + g.programId);
            glbl += ' / ' + (g.divisionId ? escHtml(g.divisionName || 'Division ' + g.divisionId) : 'All Divisions');
            h += '<div class="ab-group-detail-item">' + glbl + '</div>';
        }
        h += '</div>';
        return h;
    }

    function toggleGroupDetail(uid) {
        var el = document.getElementById('abGrpDetail_' + uid);
        if (el) el.classList.toggle('ab-show');
    }

    // =====================================================================
    // RENDER ROUTER
    // =====================================================================
    function render() {
        var el = document.getElementById('abContent');
        state.loading = false;
        switch (state.mode) {
            case 'landing': el.innerHTML = renderLanding(); break;
            case 'admin': el.innerHTML = renderAdmin(); break;
            case 'admin_edit': el.innerHTML = renderAdminEdit(); break;
            case 'generate_pick': el.innerHTML = renderGeneratePick(); break;
            case 'generate_preview': el.innerHTML = renderGeneratePreview(); break;
            default: el.innerHTML = renderLanding();
        }
    }

    // =====================================================================
    // LANDING
    // =====================================================================
    function renderLanding() {
        return '<div class="ab-landing">' +
            '<div class="ab-landing-title">Attendance Report Builder</div>' +
            '<div class="ab-landing-subtitle">Create configurable attendance reports for any ministry area</div>' +
            '<div class="ab-landing-buttons">' +
                '<div class="ab-landing-btn ab-admin-btn" onclick="ABApp.goAdmin()">' +
                    '<i>&#9881;</i><span>Admin Setup</span>' +
                '</div>' +
                '<div class="ab-landing-btn ab-gen-btn" onclick="ABApp.goGenerate()">' +
                    '<i>&#128202;</i><span>Generate Report</span>' +
                '</div>' +
            '</div>' +
        '</div>';
    }

    // =====================================================================
    // ADMIN LIST
    // =====================================================================
    function renderAdmin() {
        var h = '<button class="ab-back-btn" onclick="ABApp.goLanding()">&#8592; Back</button>';
        h += '<div class="ab-flex ab-items-center ab-justify-between ab-mb-16">';
        h += '<h2>Attendance Report Configs</h2>';
        h += '<button class="ab-btn ab-btn-primary" onclick="ABApp.newConfig()">+ New Config</button>';
        h += '</div>';

        if (!state.configs.length) {
            h += '<div class="ab-empty"><p>No configs yet. Create one to get started.</p></div>';
        } else {
            for (var i = 0; i < state.configs.length; i++) {
                var c = state.configs[i];
                h += '<div class="ab-config-card">';
                h += '<div>';
                h += '<div class="ab-config-name">' + escHtml(c.name || 'Untitled') + '</div>';
                h += '<div class="ab-config-meta">' + buildSourceSummaryHtml(c, 'admin_' + i) + ' &bull; Updated: ' + escHtml(c.updatedAt || '') + '</div>';
                h += '</div>';
                h += '<div class="ab-config-actions">';
                h += '<button class="ab-btn ab-btn-sm ab-btn-outline" onclick="ABApp.editConfig(\\'' + escAttr(c.id) + '\\')">Edit</button>';
                h += '<button class="ab-btn ab-btn-sm ab-btn-outline" onclick="ABApp.duplicateConfig(\\'' + escAttr(c.id) + '\\')">Duplicate</button>';
                h += '<button class="ab-btn ab-btn-sm ab-btn-danger" onclick="ABApp.deleteConfig(\\'' + escAttr(c.id) + '\\')">Delete</button>';
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

        var h = '<button class="ab-back-btn" onclick="ABApp.goAdminList()">&#8592; Back to Configs</button>';
        h += '<h2 class="ab-mb-16">' + (c.id ? 'Edit Config' : 'New Config') + '</h2>';

        // Config Name
        h += '<div class="ab-form-group">';
        h += '<label class="ab-label">Config Name</label>';
        h += '<input class="ab-input" id="abConfigName" value="' + escAttr(c.name) + '" placeholder="e.g., VBS 2026 Attendance">';
        h += '</div>';

        // Print Title
        h += '<div class="ab-form-group">';
        h += '<label class="ab-label">Print Title (shown on report header)</label>';
        h += '<input class="ab-input" id="abPrintTitle" value="' + escAttr(c.printTitle || '') + '" placeholder="e.g., VBS {year} Attendance Report">';
        h += '<div class="ab-help-text">Any 4-digit year in the title (or the literal <code>{year}</code> placeholder) is replaced with the year of the report\\'s start date - so historical reports show the correct year automatically.</div>';
        h += '</div>';

        // Source Type
        h += '<div class="ab-form-group">';
        h += '<label class="ab-label">Data Source</label>';
        h += '<div class="ab-flex ab-gap-16">';
        h += '<label style="display:flex;align-items:center;gap:6px;font-size:14px;cursor:pointer;"><input type="radio" name="abSourceType" value="program_division" ' + (c.sourceType !== 'specific_orgs' ? 'checked' : '') + ' onchange="ABApp.setSourceType(\\'program_division\\')"> Program / Division</label>';
        h += '<label style="display:flex;align-items:center;gap:6px;font-size:14px;cursor:pointer;"><input type="radio" name="abSourceType" value="specific_orgs" ' + (c.sourceType === 'specific_orgs' ? 'checked' : '') + ' onchange="ABApp.setSourceType(\\'specific_orgs\\')"> Specific Organizations</label>';
        h += '</div>';
        h += '</div>';

        // Program/Division Groups
        h += '<div id="abProgramDivSection" style="' + (c.sourceType === 'specific_orgs' ? 'display:none' : '') + '">';
        h += '<div class="ab-form-group">';
        h += '<label class="ab-label">Program / Division Groups</label>';
        h += '<div class="ab-help-text ab-mb-8">Add one or more program/division pairs. Leave division blank to include all divisions in a program.</div>';
        h += '<div class="ab-form-row" style="align-items:flex-end;">';
        h += '<div style="flex:2;">';
        h += '<label class="ab-label" style="font-size:12px;">Program</label>';
        h += '<select class="ab-select" id="abAddProgram" onchange="ABApp.onAddProgramChange()">';
        h += '<option value="">-- Select Program --</option>';
        for (var pi = 0; pi < state.programs.length; pi++) {
            var prog = state.programs[pi];
            h += '<option value="' + prog.id + '">' + escHtml(prog.name) + '</option>';
        }
        h += '</select>';
        h += '</div>';
        h += '<div style="flex:2;">';
        h += '<label class="ab-label" style="font-size:12px;">Division (optional)</label>';
        h += '<select class="ab-select" id="abAddDivision">';
        h += '<option value="">-- All Divisions --</option>';
        h += '</select>';
        h += '</div>';
        h += '<div style="flex:0 0 auto;">';
        h += '<button class="ab-btn ab-btn-primary ab-btn-sm" onclick="ABApp.addProgDivGroup()" style="white-space:nowrap;">+ Add</button>';
        h += '</div>';
        h += '</div>';
        h += '</div>';

        // Render existing groups
        h += '<div id="abProgDivGroupList" class="ab-mt-12">';
        if (!c.programDivGroups) c.programDivGroups = [];
        h += renderProgDivGroupListHtml(c.programDivGroups);
        h += '</div>';

        // Exclude org IDs
        h += '<div class="ab-form-group ab-mt-12">';
        h += '<label class="ab-label">Exclude Organization IDs (comma-separated, optional)</label>';
        h += '<input class="ab-input" id="abExcludeOrgs" value="' + escAttr(c.excludeOrgIds || '') + '" placeholder="e.g., 2950, 2967">';
        h += '</div>';
        h += '</div>';

        // Specific orgs section
        h += '<div id="abSpecificOrgsSection" style="' + (c.sourceType !== 'specific_orgs' ? 'display:none' : '') + '">';
        h += '<div class="ab-form-group">';
        h += '<label class="ab-label">Search and Add Organizations</label>';
        h += '<div class="ab-search-box">';
        h += '<input class="ab-input" id="abOrgSearch" placeholder="Search by name or ID..." oninput="ABApp.onOrgSearch()">';
        h += '<div id="abOrgSearchResults" class="ab-search-results" style="display:none"></div>';
        h += '</div>';
        h += '</div>';

        // Selected orgs list
        h += '<div id="abSelectedOrgsList">';
        if (c.specificOrgIds && c.specificOrgIds.length) {
            for (var si = 0; si < c.specificOrgIds.length; si++) {
                var soId = c.specificOrgIds[si];
                var soName = (c.specificOrgNames || {})[String(soId)] || 'Org ' + soId;
                h += '<div class="ab-selected-org-item">';
                h += '<span>' + escHtml(soName) + ' (ID: ' + soId + ')</span>';
                h += '<span class="ab-remove-org" onclick="ABApp.removeOrg(' + si + ')">&#10005;</span>';
                h += '</div>';
            }
        } else {
            h += '<div class="ab-text-muted ab-text-sm">No organizations selected yet.</div>';
        }
        h += '</div>';
        h += '</div>';

        // Report Sections toggles
        h += '<div class="ab-panel ab-mt-16">';
        h += '<div class="ab-panel-header">Report Sections</div>';
        h += '<div class="ab-panel-body">';
        var secs = c.reportSections || {};
        var secDefs = [
            {key: 'dailySummary', label: 'Daily Summary Table'},
            {key: 'programDivisionBreakdown', label: 'Program / Division Breakdown'},
            {key: 'orgDetail', label: 'Organization Detail Cards'},
            {key: 'subgroupBreakdown', label: 'Subgroup Breakdown'},
            {key: 'enrollmentRatio', label: 'Enrollment Ratio'},
            {key: 'separateLeaders', label: 'Separate Leaders vs Members'},
            {key: 'countDistinct', label: 'Count Distinct People (vs sum across orgs)'},
            {key: 'includeInactive', label: 'Include Inactive Involvements (for historical reports)'}
        ];
        for (var si2 = 0; si2 < secDefs.length; si2++) {
            var sec = secDefs[si2];
            var isOn = secs[sec.key] !== false && secs[sec.key] !== undefined ? secs[sec.key] : false;
            // dailySummary, programDivisionBreakdown, orgDetail, separateLeaders default to true
            if (secs[sec.key] === undefined && (sec.key === 'dailySummary' || sec.key === 'programDivisionBreakdown' || sec.key === 'orgDetail' || sec.key === 'separateLeaders')) {
                isOn = true;
            }
            h += '<div class="ab-toggle-row">';
            h += '<span class="ab-toggle-label">' + sec.label + '</span>';
            h += '<label class="ab-toggle"><input type="checkbox" id="abSec_' + sec.key + '"' + (isOn ? ' checked' : '') + '><span class="ab-toggle-slider"></span></label>';
            h += '</div>';
        }
        h += '</div></div>';

        // Leader Member Types
        h += '<div class="ab-form-group ab-mt-16">';
        h += '<label class="ab-label">Leader Member Type IDs (comma-separated)</label>';
        h += '<input class="ab-input" id="abLeaderTypes" value="' + escAttr(c.leaderMemberTypes || '140,160') + '" placeholder="140,160">';
        h += '<div class="ab-help-text">MemberTypeId values to classify as leaders (e.g., 140=Leader, 160=Teacher). Used when "Separate Leaders" is enabled.</div>';
        h += '</div>';

        // Sort
        h += '<div class="ab-form-group">';
        h += '<label class="ab-label">Sort Organizations By</label>';
        h += '<select class="ab-select" id="abSortBy">';
        h += '<option value="name"' + (c.sortBy === 'name' ? ' selected' : '') + '>Name</option>';
        h += '<option value="attendance"' + (c.sortBy === 'attendance' ? ' selected' : '') + '>Total Attendance (desc)</option>';
        h += '</select>';
        h += '</div>';

        // Save button
        h += '<div class="ab-flex ab-gap-12 ab-mt-16">';
        h += '<button class="ab-btn ab-btn-primary" onclick="ABApp.saveConfig()">Save Config</button>';
        h += '<button class="ab-btn ab-btn-outline" onclick="ABApp.goAdminList()">Cancel</button>';
        h += '</div>';

        return h;
    }

    // =====================================================================
    // GENERATE PICK
    // =====================================================================
    function renderGeneratePick() {
        var h = '<button class="ab-back-btn" onclick="ABApp.goLanding()">&#8592; Back</button>';
        h += '<h2 class="ab-mb-16">Select a Config to Generate</h2>';

        if (!state.configs.length) {
            h += '<div class="ab-empty"><p>No configs available. Go to Admin Setup to create one first.</p></div>';
        } else {
            for (var i = 0; i < state.configs.length; i++) {
                var c = state.configs[i];
                var isSelected = (state.selectedConfigId === c.id);
                h += '<div class="ab-config-card" onclick="ABApp.selectConfig(\\'' + escAttr(c.id) + '\\')" style="' + (isSelected ? 'border-color:var(--ab-primary);background:var(--ab-primary-light);' : '') + '">';
                h += '<div>';
                h += '<div class="ab-config-name">' + escHtml(c.name || 'Untitled') + '</div>';
                h += '<div class="ab-config-meta">' + buildSourceSummaryHtml(c, 'gen_' + i) + '</div>';
                h += '</div>';
                if (isSelected) {
                    h += '<div style="display:flex;flex-direction:column;gap:8px;align-items:flex-end;" onclick="event.stopPropagation()">';
                    h += '<div class="ab-flex ab-gap-8 ab-items-center">';
                    h += '<label class="ab-label" style="margin:0;white-space:nowrap;">Start Date:</label>';
                    h += '<input type="date" class="ab-input" id="abOverrideStart" value="' + escAttr(state.overrideStart || '') + '" style="width:170px;" onchange="ABApp.onOverrideChange()">';
                    h += '</div>';
                    h += '<div class="ab-flex ab-gap-8 ab-items-center">';
                    h += '<label class="ab-label" style="margin:0;white-space:nowrap;">End Date:</label>';
                    h += '<input type="date" class="ab-input" id="abOverrideEnd" value="' + escAttr(state.overrideEnd || '') + '" style="width:170px;" onchange="ABApp.onOverrideChange()">';
                    h += '</div>';
                    h += '<button class="ab-btn ab-btn-success" onclick="ABApp.generateReport()">Generate Report</button>';
                    h += '</div>';
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
        var h = '<button class="ab-back-btn" onclick="ABApp.goGenerate()">&#8592; Back to Config Selection</button>';
        h += '<div class="ab-flex ab-items-center ab-justify-between ab-mb-16">';
        h += '<h2>Attendance Report Preview</h2>';
        h += '<button class="ab-btn ab-btn-primary" onclick="ABApp.printReport()">Print Report</button>';
        h += '</div>';

        // Stats
        var st = state.previewStats || {};
        h += '<div class="ab-stats">';
        h += '<div class="ab-stat-card"><div class="ab-stat-num">' + fmtNum(st.orgCount || 0) + '</div><div class="ab-stat-label">Organizations</div></div>';
        h += '<div class="ab-stat-card"><div class="ab-stat-num">' + fmtNum(st.totalAttendance || 0) + '</div><div class="ab-stat-label">Total Attendance</div></div>';
        h += '<div class="ab-stat-card"><div class="ab-stat-num">' + fmtNum(st.dayCount || 0) + '</div><div class="ab-stat-label">Days</div></div>';
        if (st.totalEnrolled) {
            h += '<div class="ab-stat-card"><div class="ab-stat-num">' + fmtNum(st.totalEnrolled) + '</div><div class="ab-stat-label">Enrolled</div></div>';
        }
        h += '</div>';

        // Preview
        h += '<div class="ab-preview-area">' + state.previewHtml + '</div>';

        return h;
    }

    // =====================================================================
    // ORG CARD COLLAPSE/EXPAND
    // =====================================================================
    function toggleOrgCard(headerEl) {
        var body = headerEl.nextElementSibling;
        if (!body) return;
        var isHidden = (body.style.display === 'none' || body.style.display === '');
        body.style.display = isHidden ? 'block' : 'none';
        var icon = headerEl.querySelector('.ab-org-toggle-icon');
        if (icon) icon.innerHTML = isHidden ? '&#9660;' : '&#9654;';
    }

    function toggleAllOrgs(expand) {
        var bodies = document.querySelectorAll('.ab-org-card-body');
        for (var i = 0; i < bodies.length; i++) {
            bodies[i].style.display = expand ? 'block' : 'none';
        }
        var icons = document.querySelectorAll('.ab-org-toggle-icon');
        for (var j = 0; j < icons.length; j++) {
            icons[j].innerHTML = expand ? '&#9660;' : '&#9654;';
        }
    }

    // =====================================================================
    // PRINT POPUP
    // =====================================================================
    function printReport() {
        // Read live DOM so on-screen expand/collapse state is preserved in print.
        // Inline display:none on collapsed sections (org card bodies, division detail
        // groups) carries through naturally; only expanded sections will print.
        var previewEl = document.querySelector('.ab-preview-area');
        var html = previewEl ? previewEl.innerHTML : state.previewHtml;

        var printHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Attendance Report</title><style>';
        printHtml += '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important;}';
        printHtml += 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:20px;color:#333;font-size:13px;}';
        printHtml += '.ab-report{max-width:100%;}';
        printHtml += '.ab-rpt-header{text-align:center;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid #7c3aed;}';
        printHtml += '.ab-rpt-title{font-size:22px;margin:0 0 4px;color:#7c3aed;}';
        printHtml += '.ab-rpt-meta{font-size:13px;color:#666;}';
        printHtml += '.ab-section{margin-bottom:20px;}';
        printHtml += '.ab-section-title{font-size:16px;color:#7c3aed;border-bottom:2px solid #ede9fe;padding-bottom:4px;margin:0 0 8px;}';
        printHtml += '.ab-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:10px;}';
        printHtml += '.ab-table th{padding:6px 8px;text-align:left;background:#7c3aed;color:#fff;font-weight:600;font-size:11px;}';
        printHtml += '.ab-table td{padding:4px 8px;border-bottom:1px solid #e2e8f0;}';
        printHtml += '.ab-table-sm th{padding:4px 6px;font-size:10px;}';
        printHtml += '.ab-table-sm td{padding:3px 6px;font-size:11px;}';
        printHtml += '.ab-total-row{background:#fef3c7!important;}';
        printHtml += '.ab-total-row td{border-top:2px solid #d97706;}';
        printHtml += '.ab-avg-row{background:#f0f9ff!important;font-style:italic;}';
        printHtml += '.ab-subtotal-row{background:#f1f5f9!important;}';
        printHtml += '.ab-subtotal-row td{border-top:1px solid #94a3b8;}';
        printHtml += '.ab-org-card{border:1px solid #e2e8f0;border-radius:6px;margin-bottom:10px;overflow:hidden;page-break-inside:avoid;}';
        printHtml += '.ab-org-card-header{padding:6px 10px;background:#f8fafc;display:flex;justify-content:space-between;align-items:center;}';
        printHtml += '.ab-org-card-body{border-top:1px solid #e2e8f0;}';
        printHtml += '.ab-org-card-summary{font-size:11px;color:#0d9488;font-weight:600;}';
        printHtml += '.ab-org-card-meta{font-size:11px;color:#64748b;}';
        printHtml += '.ab-org-card-footer{padding:4px 10px;background:#ccfbf1;font-size:11px;color:#0d9488;}';
        // Hide UI affordances that don't make sense in print
        printHtml += '.ab-org-toggle-icon,.ab-section-actions,.ab-link-btn{display:none!important;}';
        // Division detail expand caret in Program/Division Breakdown
        printHtml += 'span[onclick]{display:none!important;}';
        printHtml += '</style></head><body>';
        printHtml += html;
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
            state.overrideStart = '';
            state.overrideEnd = '';
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
            state.mode = 'admin_edit';
            render();
        });
    }

    // Migrate old config format (single programId/divisionId) to new (programDivGroups)
    function migrateConfig(c) {
        if (!c.programDivGroups) c.programDivGroups = [];
        if (!c.programDivGroups.length && c.sourceType !== 'specific_orgs' && c.programId) {
            c.programDivGroups.push({
                programId: c.programId,
                programName: c.programName || '',
                divisionId: c.divisionId || 0,
                divisionName: c.divisionName || ''
            });
        }
        // Clean up old fields
        delete c.programId;
        delete c.programName;
        delete c.divisionId;
        delete c.divisionName;
        delete c.startDate;
        delete c.endDate;
        // v2: orgDetail defaults to false for pre-existing configs
        if (!c._schemaVersion || c._schemaVersion < 2) {
            if (c.reportSections) c.reportSections.orgDetail = false;
        }
        c._schemaVersion = 2;
        return c;
    }

    function editConfig(id) {
        loadFilters(function() {
            for (var i = 0; i < state.configs.length; i++) {
                if (state.configs[i].id === id) {
                    state.editConfig = migrateConfig(JSON.parse(JSON.stringify(state.configs[i])));
                    break;
                }
            }
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
        c.name = document.getElementById('abConfigName').value.trim();
        c.printTitle = document.getElementById('abPrintTitle').value.trim();

        // Source type
        var radios = document.getElementsByName('abSourceType');
        for (var r = 0; r < radios.length; r++) {
            if (radios[r].checked) { c.sourceType = radios[r].value; break; }
        }

        // programDivGroups are managed in-place on state.editConfig
        if (c.sourceType === 'program_division') {
            c.excludeOrgIds = (document.getElementById('abExcludeOrgs') || {}).value || '';
        }

        // Report sections
        var secKeys = ['dailySummary', 'programDivisionBreakdown', 'orgDetail', 'subgroupBreakdown', 'enrollmentRatio', 'separateLeaders', 'countDistinct', 'includeInactive'];
        if (!c.reportSections) c.reportSections = {};
        for (var si = 0; si < secKeys.length; si++) {
            var el = document.getElementById('abSec_' + secKeys[si]);
            c.reportSections[secKeys[si]] = el ? el.checked : false;
        }

        c.leaderMemberTypes = (document.getElementById('abLeaderTypes') || {}).value || '140,160';
        c.sortBy = (document.getElementById('abSortBy') || {}).value || 'name';

        return c;
    }

    function saveConfig() {
        var c = collectConfigFromForm();
        if (!c.name) {
            showToast('Please enter a config name', 'danger');
            return;
        }
        if (c.sourceType === 'program_division' && (!c.programDivGroups || !c.programDivGroups.length)) {
            showToast('Please add at least one program/division group', 'danger');
            return;
        }
        if (c.sourceType === 'specific_orgs' && (!c.specificOrgIds || !c.specificOrgIds.length)) {
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
        var progDiv = document.getElementById('abProgramDivSection');
        var specOrgs = document.getElementById('abSpecificOrgsSection');
        if (type === 'specific_orgs') {
            progDiv.style.display = 'none';
            specOrgs.style.display = '';
        } else {
            progDiv.style.display = '';
            specOrgs.style.display = 'none';
        }
    }

    // Program/Division Group helpers
    function renderProgDivGroupListHtml(groups) {
        if (!groups || !groups.length) {
            return '<div class="ab-text-muted ab-text-sm">No program/division groups added yet.</div>';
        }
        var lh = '';
        for (var i = 0; i < groups.length; i++) {
            var g = groups[i];
            var label = escHtml(g.programName || 'Program ' + g.programId);
            if (g.divisionId) {
                label += ' / ' + escHtml(g.divisionName || 'Division ' + g.divisionId);
            } else {
                label += ' / All Divisions';
            }
            lh += '<div class="ab-selected-org-item">';
            lh += '<span>' + label + '</span>';
            lh += '<span class="ab-remove-org" onclick="ABApp.removeProgDivGroup(' + i + ')">&#10005;</span>';
            lh += '</div>';
        }
        return lh;
    }

    function onAddProgramChange() {
        var progEl = document.getElementById('abAddProgram');
        var progId = parseInt(progEl.value) || 0;
        var divEl = document.getElementById('abAddDivision');
        var html = '<option value="">-- All Divisions --</option>';
        if (progId) {
            for (var i = 0; i < state.divisions.length; i++) {
                var d = state.divisions[i];
                if (d.progId == progId) {
                    html += '<option value="' + d.id + '">' + escHtml(d.name) + '</option>';
                }
            }
        }
        divEl.innerHTML = html;
    }

    function addProgDivGroup() {
        var progEl = document.getElementById('abAddProgram');
        var divEl = document.getElementById('abAddDivision');
        var progId = parseInt(progEl.value) || 0;
        var divId = parseInt(divEl.value) || 0;
        if (!progId) {
            showToast('Please select a program', 'danger');
            return;
        }
        var progName = progEl.options[progEl.selectedIndex].text;
        var divName = divId && divEl.selectedIndex > 0 ? divEl.options[divEl.selectedIndex].text : '';

        var c = state.editConfig;
        if (!c.programDivGroups) c.programDivGroups = [];
        // Check for duplicates
        for (var i = 0; i < c.programDivGroups.length; i++) {
            var g = c.programDivGroups[i];
            if (g.programId == progId && g.divisionId == divId) {
                showToast('This group is already added', 'info');
                return;
            }
        }
        c.programDivGroups.push({
            programId: progId,
            programName: progName,
            divisionId: divId,
            divisionName: divName
        });

        // Re-render list
        document.getElementById('abProgDivGroupList').innerHTML = renderProgDivGroupListHtml(c.programDivGroups);
        // Reset dropdowns
        progEl.selectedIndex = 0;
        onAddProgramChange();
    }

    function removeProgDivGroup(index) {
        var c = state.editConfig;
        c.programDivGroups.splice(index, 1);
        document.getElementById('abProgDivGroupList').innerHTML = renderProgDivGroupListHtml(c.programDivGroups);
    }

    // Org search for specific_orgs mode
    function onOrgSearch() {
        clearTimeout(state.searchTimeout);
        var input = document.getElementById('abOrgSearch');
        var term = input.value.trim();
        if (term.length < 2) {
            document.getElementById('abOrgSearchResults').style.display = 'none';
            return;
        }
        // Pass through the current edit-config's includeInactive toggle so users can
        // find archived orgs when adding to specific_orgs configs
        var inc = '';
        if (state.editConfig && state.editConfig.reportSections && state.editConfig.reportSections.includeInactive) {
            inc = '1';
        }
        state.searchTimeout = setTimeout(function() {
            ajax('search_involvements', {search_term: term, include_inactive: inc}, function(err, data) {
                if (!err && data && data.success) {
                    var results = data.involvements || [];
                    var resultsEl = document.getElementById('abOrgSearchResults');
                    if (!results.length) {
                        resultsEl.innerHTML = '<div style="padding:12px;color:#999;">No results found</div>';
                    } else {
                        var rh = '';
                        for (var i = 0; i < results.length; i++) {
                            var r = results[i];
                            rh += '<div class="ab-search-result-item" onclick="ABApp.addOrg(' + r.orgId + ',\\'' + escAttr(r.orgName) + '\\')">';
                            rh += '<div class="ab-search-result-name">' + escHtml(r.orgName) + ' (ID: ' + r.orgId + ')</div>';
                            rh += '<div class="ab-search-result-meta">' + escHtml(r.programName) + ' / ' + escHtml(r.divisionName) + ' &bull; ' + r.memberCount + ' members</div>';
                            rh += '</div>';
                        }
                        resultsEl.innerHTML = rh;
                    }
                    resultsEl.style.display = '';
                }
            });
        }, 300);
    }

    function addOrg(orgId, orgName) {
        var c = state.editConfig;
        if (!c.specificOrgIds) c.specificOrgIds = [];
        if (!c.specificOrgNames) c.specificOrgNames = {};
        // Don't add duplicates
        for (var i = 0; i < c.specificOrgIds.length; i++) {
            if (c.specificOrgIds[i] == orgId) {
                showToast('Organization already added', 'info');
                return;
            }
        }
        c.specificOrgIds.push(orgId);
        c.specificOrgNames[String(orgId)] = orgName;

        // Re-render selected list
        var listEl = document.getElementById('abSelectedOrgsList');
        var lh = '';
        for (var j = 0; j < c.specificOrgIds.length; j++) {
            var soId = c.specificOrgIds[j];
            var soName = c.specificOrgNames[String(soId)] || 'Org ' + soId;
            lh += '<div class="ab-selected-org-item">';
            lh += '<span>' + escHtml(soName) + ' (ID: ' + soId + ')</span>';
            lh += '<span class="ab-remove-org" onclick="ABApp.removeOrg(' + j + ')">&#10005;</span>';
            lh += '</div>';
        }
        listEl.innerHTML = lh;

        // Clear search
        document.getElementById('abOrgSearch').value = '';
        document.getElementById('abOrgSearchResults').style.display = 'none';
    }

    function removeOrg(index) {
        var c = state.editConfig;
        var removedId = c.specificOrgIds[index];
        c.specificOrgIds.splice(index, 1);
        delete c.specificOrgNames[String(removedId)];

        var listEl = document.getElementById('abSelectedOrgsList');
        if (!c.specificOrgIds.length) {
            listEl.innerHTML = '<div class="ab-text-muted ab-text-sm">No organizations selected yet.</div>';
            return;
        }
        var lh = '';
        for (var j = 0; j < c.specificOrgIds.length; j++) {
            var soId = c.specificOrgIds[j];
            var soName = c.specificOrgNames[String(soId)] || 'Org ' + soId;
            lh += '<div class="ab-selected-org-item">';
            lh += '<span>' + escHtml(soName) + ' (ID: ' + soId + ')</span>';
            lh += '<span class="ab-remove-org" onclick="ABApp.removeOrg(' + j + ')">&#10005;</span>';
            lh += '</div>';
        }
        listEl.innerHTML = lh;
    }

    // Generate
    function selectConfig(id) {
        if (state.selectedConfigId === id) return;
        state.selectedConfigId = id;
        state.overrideStart = '';
        state.overrideEnd = '';
        render();
    }

    function onOverrideChange() {
        var startEl = document.getElementById('abOverrideStart');
        var endEl = document.getElementById('abOverrideEnd');
        if (startEl) state.overrideStart = startEl.value;
        if (endEl) state.overrideEnd = endEl.value;
    }

    function generateReport() {
        if (!state.selectedConfigId) {
            showToast('Please select a config first', 'danger');
            return;
        }
        // Get dates from form
        onOverrideChange();
        if (!state.overrideStart || !state.overrideEnd) {
            showToast('Please enter both start and end dates', 'danger');
            return;
        }

        var config = null;
        for (var i = 0; i < state.configs.length; i++) {
            if (state.configs[i].id === state.selectedConfigId) {
                config = migrateConfig(JSON.parse(JSON.stringify(state.configs[i])));
                break;
            }
        }
        if (!config) {
            showToast('Config not found', 'danger');
            return;
        }

        showLoading();
        var params = {
            config_json: JSON.stringify(config),
            override_start: state.overrideStart || '',
            override_end: state.overrideEnd || ''
        };
        ajax('generate_report', params, function(err, data) {
            if (!err && data && data.success) {
                state.previewHtml = data.html || '';
                state.previewStats = data.stats || {};
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
    window.ABApp = {
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
        onAddProgramChange: onAddProgramChange,
        addProgDivGroup: addProgDivGroup,
        removeProgDivGroup: removeProgDivGroup,
        toggleGroupDetail: toggleGroupDetail,
        onOrgSearch: onOrgSearch,
        addOrg: addOrg,
        removeOrg: removeOrg,
        selectConfig: selectConfig,
        onOverrideChange: onOverrideChange,
        generateReport: generateReport,
        printReport: printReport,
        toggleOrgCard: toggleOrgCard,
        toggleAllOrgs: toggleAllOrgs
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
