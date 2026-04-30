#roles=Edit
#----------------------------------------------------------------------
# TPxi_AttendanceMarkings.py
#
# Team-friendly attendance entry for large events (VBS, camps, etc.)
#
# Built around two ideas:
#   1) Inverted attendance: when most people attend, you click only those
#      who didn't (default Present, click to mark Absent). Configurable
#      per saved config to also support default Absent (click Present)
#      and default Unmarked (must explicitly mark each).
#   2) Team workflow: multiple staff work the involvement list together,
#      claiming rows as they go and watching live counts. The dashboard
#      polls every 10 seconds. Working "down to zero" = clearing the
#      remaining-involvements counter.
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
#   Per-click writes via model.EditPersonAttendance for live team sync.
#   Meeting auto-created on first roster open via GetMeetingIdByDateTime.
#
# Storage Keys:
#   AttendanceMarkings_Configs                            - Saved configs
#   AttendanceMarkings_Session_<configId>_<dateIso>       - Live session state
#   AttendanceMarkings_Log_<configId>_<dateIso>           - Action audit log
#
# CSS Prefix: va-
# Root Class: .va-root
#
# Reference:
#   TPxi_RollSheet.py          - SPA + config CRUD + schedule filter
#   TPxi_AttendanceBuilder.py  - SPA dispatch + filters
#   TPxi_DayOfRegistration.py  - team coordination, polling, optimistic UI
#----------------------------------------------------------------------

import json
import datetime
import re

model.Header = 'Attendance Markings'

# =====================================================================
# HELPERS
# =====================================================================

def html_escape(text):
    if text is None:
        return ''
    s = str(text)
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    s = s.replace('"', '&quot;').replace("'", '&#39;')
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

def safe_bool(val):
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    s = str(val).lower().strip()
    return s in ('1', 'true', 'yes', 'on')

def now_iso():
    return datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')

def now_display():
    return datetime.datetime.now().strftime('%I:%M %p').lstrip('0')

def parse_iso_datetime(s):
    """Parse ISO datetime back to datetime object. Returns None on failure."""
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s[:19], '%Y-%m-%dT%H:%M:%S')
    except:
        try:
            return datetime.datetime.strptime(s[:10], '%Y-%m-%d')
        except:
            return None

def minutes_ago(iso_str):
    """Return integer minutes between iso_str and now, or None."""
    dt = parse_iso_datetime(iso_str)
    if not dt:
        return None
    delta = datetime.datetime.now() - dt
    return int(delta.total_seconds() / 60)

def has_data(name):
    return hasattr(Data, name) and getattr(Data, name) not in (None, '')

def get_data(name, default=''):
    if has_data(name):
        return getattr(Data, name)
    return default

def safe_id(s):
    """Sanitize a string for use as a content-key suffix."""
    if not s:
        return ''
    return re.sub(r'[^a-zA-Z0-9_-]', '_', str(s))[:80]

def sql_in_list(int_list):
    """Build a SQL-safe comma list from a list of ints. Returns '0' if empty."""
    if not int_list:
        return '0'
    return ','.join(str(safe_int(v, 0)) for v in int_list)

# =====================================================================
# STORAGE KEYS
# =====================================================================

CONFIGS_KEY = "AttendanceMarkings_Configs"

def session_key(cfg_id, date_iso):
    return "AttendanceMarkings_Session_" + safe_id(cfg_id) + "_" + safe_id(date_iso)

def log_key(cfg_id, date_iso):
    return "AttendanceMarkings_Log_" + safe_id(cfg_id) + "_" + safe_id(date_iso)

def load_json(key, default):
    """Safely load JSON content, returning default on any error."""
    try:
        raw = model.TextContent(key)
        if raw:
            return json.loads(raw)
    except:
        pass
    return default

def save_json(key, data):
    model.WriteContentText(key, json.dumps(data), "")

# =====================================================================
# CONFIG CRUD
# =====================================================================

CONFIG_SCHEMA_VERSION = 2
DEFAULT_CONFIG = {
    '_schemaVersion': CONFIG_SCHEMA_VERSION,
    'name': '',
    'sourceType': 'program_division',     # 'program_division' | 'specific_orgs'
    'programDivGroups': [],                # [{programId, divisionId, programName, divisionName}]
    'specificOrgIds': [],
    'specificOrgNames': {},
    'excludeOrgIds': [],
    'defaultState': 'present',             # 'present' | 'absent' | 'unmarked'
    'onlyWithMeeting': True,
    'excludeMemberTypes': '',              # comma-separated (e.g. leaders to skip)
    'walkInMemberType': 'Visitor',         # 'Member' | 'Visitor' | 'Prospect' | 'Guest' (or any custom MemberType description)
    'allowWalkIns': True,                  # show "+ Add Person" button on roster
}

def normalize_config(c):
    out = dict(DEFAULT_CONFIG)
    if not isinstance(c, dict):
        return out
    for k, v in c.items():
        out[k] = v
    out['_schemaVersion'] = CONFIG_SCHEMA_VERSION
    if out.get('defaultState') not in ('present', 'absent', 'unmarked'):
        out['defaultState'] = 'present'
    if out.get('sourceType') not in ('program_division', 'specific_orgs'):
        out['sourceType'] = 'program_division'
    if not out.get('walkInMemberType'):
        out['walkInMemberType'] = 'Visitor'
    if 'allowWalkIns' not in out or out['allowWalkIns'] is None:
        out['allowWalkIns'] = True
    return out

def handle_load_configs():
    data = load_json(CONFIGS_KEY, {'configs': []})
    configs = data.get('configs', []) if isinstance(data, dict) else []
    configs = [normalize_config(c) for c in configs]
    print json.dumps({'success': True, 'configs': configs})

def handle_save_config():
    try:
        cfg_json = get_data('va_config_json')
        if not cfg_json:
            print json.dumps({'success': False, 'message': 'Missing config_json'})
            return
        config = json.loads(cfg_json)
        config = normalize_config(config)
        if not config.get('name', '').strip():
            print json.dumps({'success': False, 'message': 'Config name is required'})
            return

        data = load_json(CONFIGS_KEY, {'configs': []})
        configs = data.get('configs', []) if isinstance(data, dict) else []

        cfg_id = config.get('id', '').strip()
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

        if cfg_id:
            replaced = False
            for i, existing in enumerate(configs):
                if existing.get('id') == cfg_id:
                    config['createdAt'] = existing.get('createdAt', ts)
                    config['updatedAt'] = ts
                    configs[i] = config
                    replaced = True
                    break
            if not replaced:
                config['createdAt'] = ts
                config['updatedAt'] = ts
                configs.append(config)
        else:
            cfg_id = 'va_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            config['id'] = cfg_id
            config['createdAt'] = ts
            config['updatedAt'] = ts
            configs.append(config)

        save_json(CONFIGS_KEY, {'configs': configs})
        print json.dumps({'success': True, 'config': config})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save failed: ' + str(e)})

def handle_delete_config():
    cfg_id = get_data('va_config_id')
    if not cfg_id:
        print json.dumps({'success': False, 'message': 'Missing config_id'})
        return
    data = load_json(CONFIGS_KEY, {'configs': []})
    configs = data.get('configs', []) if isinstance(data, dict) else []
    configs = [c for c in configs if c.get('id') != cfg_id]
    save_json(CONFIGS_KEY, {'configs': configs})
    print json.dumps({'success': True})

def find_config(cfg_id):
    data = load_json(CONFIGS_KEY, {'configs': []})
    configs = data.get('configs', []) if isinstance(data, dict) else []
    for c in configs:
        if c.get('id') == cfg_id:
            return normalize_config(c)
    return None

# =====================================================================
# LOOKUPS (programs, divisions, involvement search)
# =====================================================================

def handle_get_filters():
    """Return programs and divisions for the config form dropdowns."""
    try:
        progs = []
        for r in q.QuerySql("SELECT Id, Name FROM Program WHERE Name IS NOT NULL ORDER BY Name"):
            progs.append({'id': r.Id, 'name': safe_str(r.Name)})
        divs = []
        for r in q.QuerySql("SELECT Id, Name, ProgId FROM Division WHERE Name IS NOT NULL ORDER BY Name"):
            divs.append({'id': r.Id, 'name': safe_str(r.Name), 'programId': r.ProgId})
        print json.dumps({'success': True, 'programs': progs, 'divisions': divs})
    except Exception as e:
        print json.dumps({'success': False, 'message': str(e)})

def handle_search_involvements():
    """Search organizations by name with optional program/division filter.
    Uses OrganizationStructure so orgs assigned to multiple programs surface
    under any of their assignments (not just the primary)."""
    try:
        term = get_data('va_term', '').strip()
        program_id = safe_int(get_data('va_program_id', 0))
        division_id = safe_int(get_data('va_division_id', 0))
        include_inactive = safe_bool(get_data('va_include_inactive', '0'))

        wh = []
        if not include_inactive:
            wh.append("o.OrganizationStatusId = 30")
        if term:
            safe_term = term.replace("'", "''")
            wh.append("o.OrganizationName LIKE '%" + safe_term + "%'")
        if program_id > 0:
            wh.append("os.ProgId = " + str(program_id))
        if division_id > 0:
            wh.append("os.DivId = " + str(division_id))
        where_sql = (" WHERE " + " AND ".join(wh)) if wh else ""

        # When a program/division filter is on, scope via OrganizationStructure
        # so we catch ALL program/division assignments, not just primary.
        # Otherwise just search by name.
        use_structure = program_id > 0 or division_id > 0
        if use_structure:
            sql = """
                SELECT TOP 50
                    o.OrganizationId, o.OrganizationName,
                    ISNULL(MAX(os.Division), '') AS DivisionName,
                    ISNULL(MAX(os.Program), '') AS ProgramName,
                    (SELECT COUNT(*) FROM OrganizationMembers om
                     WHERE om.OrganizationId = o.OrganizationId) AS MemberCount
                FROM Organizations o
                JOIN OrganizationStructure os ON o.OrganizationId = os.OrgId
            """ + where_sql + " GROUP BY o.OrganizationId, o.OrganizationName ORDER BY o.OrganizationName"
        else:
            sql = """
                SELECT TOP 50
                    o.OrganizationId, o.OrganizationName,
                    ISNULL(d.Name, '') AS DivisionName,
                    ISNULL(p.Name, '') AS ProgramName,
                    (SELECT COUNT(*) FROM OrganizationMembers om
                     WHERE om.OrganizationId = o.OrganizationId) AS MemberCount
                FROM Organizations o
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
            """ + where_sql + " ORDER BY o.OrganizationName"

        results = []
        for r in q.QuerySql(sql):
            results.append({
                'orgId': r.OrganizationId,
                'orgName': safe_str(r.OrganizationName),
                'divisionName': safe_str(r.DivisionName),
                'programName': safe_str(r.ProgramName),
                'memberCount': r.MemberCount,
            })
        print json.dumps({'success': True, 'results': results})
    except Exception as e:
        print json.dumps({'success': False, 'message': str(e)})

# =====================================================================
# RESOLVE ORGS IN SCOPE
# =====================================================================

def resolve_orgs_for_session(config, date_iso):
    """Returns list of org dicts in scope: [{orgId, orgName, programName,
    divisionName, memberCount, schedTime (HH:MM or '')}]."""
    src = config.get('sourceType', 'program_division')
    only_with_meeting = bool(config.get('onlyWithMeeting', True))
    exclude_set = set(safe_int(o, 0) for o in (config.get('excludeOrgIds') or []))

    # Step 1: candidate orgs
    candidate_ids = []
    if src == 'specific_orgs':
        for oid in (config.get('specificOrgIds') or []):
            try:
                oid_int = int(oid)
                if oid_int > 0 and oid_int not in exclude_set:
                    candidate_ids.append(oid_int)
            except:
                pass
    else:
        # Use OrganizationStructure — orgs can be assigned to MULTIPLE
        # Program/Division pairs (e.g., a VBS room rolls up under both
        # "VBS K-5 / Classrooms" and "VBS DAILY GRAND TOTAL"). Joining
        # via Organizations.DivisionId only catches the "primary" mapping
        # and misses orgs whose secondary mapping matches the config.
        groups = config.get('programDivGroups') or []
        for grp in groups:
            pid = safe_int(grp.get('programId', 0))
            did = safe_int(grp.get('divisionId', 0))
            if pid <= 0 and did <= 0:
                continue
            wh = ["o.OrganizationStatusId = 30"]
            if pid > 0:
                wh.append("os.ProgId = " + str(pid))
            if did > 0:
                wh.append("os.DivId = " + str(did))
            sql = """
                SELECT DISTINCT os.OrgId AS OrganizationId
                FROM OrganizationStructure os
                JOIN Organizations o ON o.OrganizationId = os.OrgId
                WHERE """ + " AND ".join(wh)
            for r in q.QuerySql(sql):
                if r.OrganizationId not in exclude_set:
                    candidate_ids.append(r.OrganizationId)

    candidate_ids = list(set(candidate_ids))
    if not candidate_ids:
        return []

    # Step 2: optional schedule filter (only orgs with a meeting on this day-of-week)
    sched_day = -1
    if date_iso:
        dt = parse_iso_datetime(date_iso)
        if dt:
            # Python weekday(): Mon=0..Sun=6  ->  SchedDay: Sun=0..Sat=6
            sched_day = (dt.weekday() + 1) % 7

    # ---- Signal 1: legacy OrgSchedule (recurring weekly schedule) ----
    sched_time_by_org = {}
    if sched_day >= 0:
        sql = """
            SELECT OrganizationId, MIN(SchedTime) AS SchedTime
            FROM OrgSchedule
            WHERE OrganizationId IN ({0}) AND SchedDay = {1}
            GROUP BY OrganizationId
        """.format(sql_in_list(candidate_ids), sched_day)
        for r in q.QuerySql(sql):
            sched_time_by_org[r.OrganizationId] = r.SchedTime

    # First/Last meeting date window (legacy gate — NULL = no boundary)
    in_window = set()
    if date_iso:
        safe_iso = date_iso.replace("'", "")
        window_sql = """
            SELECT OrganizationId
            FROM Organizations
            WHERE OrganizationId IN ({0})
                AND (FirstMeetingDate IS NULL OR FirstMeetingDate <= '{1}')
                AND (LastMeetingDate IS NULL OR LastMeetingDate >= '{1}')
        """.format(sql_in_list(candidate_ids), safe_iso)
        for r in q.QuerySql(window_sql):
            in_window.add(r.OrganizationId)

    # ---- Signal 2: pre-created Meetings rows for the selected date ----
    # When an org uses TouchPoint's new Scheduler, Meeting rows are created
    # ahead of time. Catches them even without an OrgSchedule entry.
    meeting_time_by_org = {}
    if date_iso:
        safe_iso = date_iso.replace("'", "")
        meet_sql = """
            SELECT OrganizationId, MIN(MeetingDate) AS MeetingDate
            FROM Meetings
            WHERE OrganizationId IN ({0})
                AND CONVERT(date, MeetingDate) = '{1}'
                AND (Canceled IS NULL OR Canceled = 0)
            GROUP BY OrganizationId
        """.format(sql_in_list(candidate_ids), safe_iso)
        for r in q.QuerySql(meet_sql):
            meeting_time_by_org[r.OrganizationId] = r.MeetingDate

    # ---- Signal 3: MeetingSeries (Scheduler-defined recurring schedules) ----
    # RRULE parsing — based on observed live data, every series in the wild is
    # FREQ=WEEKLY;INTERVAL=1;BYDAY=XX with a DTSTART. Anything more complex
    # falls through and we rely on signals 1+2.
    series_time_by_org = {}
    if date_iso:
        try:
            target_dt = parse_iso_datetime(date_iso)
            if target_dt:
                target_date_only = target_dt.date()
                target_weekday = target_dt.weekday()  # Mon=0..Sun=6
                day_map = {'MO': 0, 'TU': 1, 'WE': 2, 'TH': 3, 'FR': 4, 'SA': 5, 'SU': 6}
                series_sql = """
                    SELECT OrganizationId, MeetingStart, RruleString
                    FROM MeetingSeries
                    WHERE OrganizationId IN ({0})
                """.format(sql_in_list(candidate_ids))
                for r in q.QuerySql(series_sql):
                    rrule = r.RruleString or ''
                    # DTSTART: must be on or before target date
                    dt_match = re.search(r'DTSTART:(\d{8})', rrule)
                    if dt_match:
                        try:
                            dtstart_date = datetime.datetime.strptime(dt_match.group(1), '%Y%m%d').date()
                            if target_date_only < dtstart_date:
                                continue
                        except:
                            pass
                    # FREQ — only handle WEEKLY for now
                    if 'FREQ=WEEKLY' not in rrule:
                        continue
                    # BYDAY — at least one day must match target weekday
                    bd = re.search(r'BYDAY=([A-Z,]+)', rrule)
                    if not bd:
                        continue
                    matched = False
                    for d in bd.group(1).split(','):
                        d = d.strip()
                        if d in day_map and day_map[d] == target_weekday:
                            matched = True
                            break
                    if not matched:
                        continue
                    # Use this series's MeetingStart as the time of day
                    series_time_by_org[r.OrganizationId] = r.MeetingStart
        except:
            pass

    if only_with_meeting:
        kept = []
        for oid in candidate_ids:
            # Pre-created Meeting wins — it's the most authoritative
            if oid in meeting_time_by_org:
                kept.append(oid)
                continue
            # MeetingSeries match (Scheduler-defined, future event)
            if oid in series_time_by_org:
                kept.append(oid)
                continue
            # Legacy: OrgSchedule day match AND in [First, Last] meeting window
            if oid in sched_time_by_org and (not in_window or oid in in_window):
                kept.append(oid)
                continue
        candidate_ids = kept
        if not candidate_ids:
            return []

    # Build a unified "schedule time" map preferring most-specific source
    for oid, t in meeting_time_by_org.items():
        sched_time_by_org[oid] = t
    for oid, t in series_time_by_org.items():
        if oid not in meeting_time_by_org:
            sched_time_by_org[oid] = t

    # Step 3: hydrate org details
    sql = """
        SELECT o.OrganizationId, o.OrganizationName,
               ISNULL(d.Name, '') AS DivisionName,
               ISNULL(p.Name, '') AS ProgramName,
               (SELECT COUNT(*) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId) AS MemberCount
        FROM Organizations o
        LEFT JOIN Division d ON o.DivisionId = d.Id
        LEFT JOIN Program p ON d.ProgId = p.Id
        WHERE o.OrganizationId IN ({0})
        ORDER BY p.Name, d.Name, o.OrganizationName
    """.format(sql_in_list(candidate_ids))

    out = []
    for r in q.QuerySql(sql):
        st = sched_time_by_org.get(r.OrganizationId)
        sched_str = ''
        if st is not None:
            try:
                sched_str = st.strftime('%H:%M')
            except:
                sched_str = ''
        out.append({
            'orgId': r.OrganizationId,
            'orgName': safe_str(r.OrganizationName),
            'divisionName': safe_str(r.DivisionName),
            'programName': safe_str(r.ProgramName),
            'memberCount': r.MemberCount,
            'schedTime': sched_str,
        })
    return out

# =====================================================================
# SESSION STATE
# =====================================================================

def get_session(cfg_id, date_iso):
    """Load session state, initializing structure if missing."""
    raw = load_json(session_key(cfg_id, date_iso), None)
    if not isinstance(raw, dict):
        raw = {}
    raw.setdefault('configId', cfg_id)
    raw.setdefault('dateIso', date_iso)
    raw.setdefault('globalTime', '09:00')
    raw.setdefault('orgs', {})        # { orgId(str): {claimedBy, claimedAt, lastTouchAt, lastTouchBy, completed, completedAt, completedBy, meetingId} }
    raw.setdefault('volunteers', [])
    return raw

def write_session(sess):
    save_json(session_key(sess.get('configId'), sess.get('dateIso')), sess)

def append_log(cfg_id, date_iso, entry):
    key = log_key(cfg_id, date_iso)
    data = load_json(key, {'entries': []})
    if not isinstance(data, dict):
        data = {'entries': []}
    data.setdefault('entries', [])
    entry['at'] = now_iso()
    data['entries'].append(entry)
    # Trim oldest if very large
    if len(data['entries']) > 5000:
        data['entries'] = data['entries'][-5000:]
    save_json(key, data)

# =====================================================================
# DASHBOARD: live involvement list with counts
# =====================================================================

def get_attend_counts_for_meetings(meeting_ids):
    """Return dict {meetingId: {'present': N, 'absent': N}}."""
    if not meeting_ids:
        return {}
    sql = """
        SELECT MeetingId,
               SUM(CASE WHEN AttendanceFlag = 1 THEN 1 ELSE 0 END) AS PresentCount,
               SUM(CASE WHEN AttendanceFlag = 0 THEN 1 ELSE 0 END) AS AbsentCount
        FROM Attend
        WHERE MeetingId IN ({0})
        GROUP BY MeetingId
    """.format(sql_in_list(meeting_ids))
    out = {}
    for r in q.QuerySql(sql):
        out[r.MeetingId] = {
            'present': r.PresentCount or 0,
            'absent': r.AbsentCount or 0,
        }
    return out

def get_enrolled_count_for_orgs(org_ids, exclude_member_types):
    """Return dict {orgId: enrolledCount} respecting exclude_member_types."""
    if not org_ids:
        return {}
    excl = ''
    if exclude_member_types:
        # Sanitize comma list of ints
        ids = []
        for tok in str(exclude_member_types).split(','):
            n = safe_int(tok.strip(), 0)
            if n > 0:
                ids.append(n)
        if ids:
            excl = " AND om.MemberTypeId NOT IN (" + ','.join(str(i) for i in ids) + ")"
    sql = """
        SELECT om.OrganizationId, COUNT(*) AS Cnt
        FROM OrganizationMembers om
        JOIN People p ON p.PeopleId = om.PeopleId
        WHERE om.OrganizationId IN ({0})
            AND p.IsDeceased = 0
            AND p.ArchivedFlag = 0
            {1}
        GROUP BY om.OrganizationId
    """.format(sql_in_list(org_ids), excl)
    out = {}
    for r in q.QuerySql(sql):
        out[r.OrganizationId] = r.Cnt
    return out

def handle_init_session():
    """Called when user picks config/date/volunteer/global-time. Resolves orgs in scope, registers volunteer."""
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        global_time = get_data('va_global_time', '09:00')
        volunteer = get_data('va_volunteer', '').strip()

        if not cfg_id or not date_iso:
            print json.dumps({'success': False, 'message': 'Missing config_id or date_iso'})
            return
        if not volunteer:
            print json.dumps({'success': False, 'message': 'Volunteer name is required'})
            return

        config = find_config(cfg_id)
        if not config:
            print json.dumps({'success': False, 'message': 'Config not found'})
            return

        sess = get_session(cfg_id, date_iso)
        if global_time:
            sess['globalTime'] = global_time
        if volunteer not in sess.get('volunteers', []):
            sess.setdefault('volunteers', []).append(volunteer)
        write_session(sess)

        orgs = resolve_orgs_for_session(config, date_iso)
        print json.dumps({
            'success': True,
            'config': config,
            'session': sess,
            'orgs': orgs,
            'orgCount': len(orgs),
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Init failed: ' + str(e)})

def handle_dashboard():
    """Return live dashboard data: orgs in scope + counts + claim/completion state."""
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        if not cfg_id or not date_iso:
            print json.dumps({'success': False, 'message': 'Missing args'})
            return
        config = find_config(cfg_id)
        if not config:
            print json.dumps({'success': False, 'message': 'Config not found'})
            return

        orgs = resolve_orgs_for_session(config, date_iso)
        org_ids = [o['orgId'] for o in orgs]
        enrolled = get_enrolled_count_for_orgs(org_ids, config.get('excludeMemberTypes', ''))

        sess = get_session(cfg_id, date_iso)
        sess_orgs = sess.get('orgs', {}) or {}

        # Build map of meetingId -> orgId for orgs that have a meeting recorded
        meeting_ids = []
        meeting_to_org = {}
        for o in orgs:
            st = sess_orgs.get(str(o['orgId']), {})
            mid = st.get('meetingId')
            if mid:
                meeting_ids.append(mid)
                meeting_to_org[mid] = o['orgId']
        counts_by_mid = get_attend_counts_for_meetings(meeting_ids)

        rows = []
        remaining = 0
        for o in orgs:
            oid = o['orgId']
            sst = sess_orgs.get(str(oid), {}) or {}
            enrolled_n = enrolled.get(oid, 0)
            mid = sst.get('meetingId')
            counts = counts_by_mid.get(mid, {'present': 0, 'absent': 0}) if mid else {'present': 0, 'absent': 0}
            marked = counts['present'] + counts['absent']
            unmarked_n = max(0, enrolled_n - marked)
            completed = bool(sst.get('completed'))
            if not completed:
                remaining += 1
            rows.append({
                'orgId': oid,
                'orgName': o['orgName'],
                'programName': o['programName'],
                'divisionName': o['divisionName'],
                'enrolled': enrolled_n,
                'present': counts['present'],
                'absent': counts['absent'],
                'unmarked': unmarked_n,
                'schedTime': o['schedTime'],
                'claimedBy': sst.get('claimedBy', ''),
                'claimedAt': sst.get('claimedAt', ''),
                'lastTouchBy': sst.get('lastTouchBy', ''),
                'lastTouchAt': sst.get('lastTouchAt', ''),
                'lastTouchAgo': minutes_ago(sst.get('lastTouchAt', '')) if sst.get('lastTouchAt') else None,
                'claimedAgo': minutes_ago(sst.get('claimedAt', '')) if sst.get('claimedAt') else None,
                'completed': completed,
                'completedBy': sst.get('completedBy', ''),
                'completedAt': sst.get('completedAt', ''),
            })

        print json.dumps({
            'success': True,
            'rows': rows,
            'remaining': remaining,
            'total': len(rows),
            'volunteers': sess.get('volunteers', []),
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Dashboard failed: ' + str(e)})

def handle_claim_org():
    cfg_id = get_data('va_config_id')
    date_iso = get_data('va_date_iso')
    org_id = safe_int(get_data('va_org_id', 0))
    volunteer = get_data('va_volunteer', '').strip()
    if not cfg_id or not date_iso or org_id <= 0 or not volunteer:
        print json.dumps({'success': False, 'message': 'Missing args'})
        return
    sess = get_session(cfg_id, date_iso)
    sess.setdefault('orgs', {})
    sst = sess['orgs'].setdefault(str(org_id), {})
    # Soft-claim: just stamp who's currently working it
    sst['claimedBy'] = volunteer
    sst['claimedAt'] = now_iso()
    sst['lastTouchBy'] = volunteer
    sst['lastTouchAt'] = now_iso()
    write_session(sess)
    append_log(cfg_id, date_iso, {'action': 'claim', 'orgId': org_id, 'by': volunteer})
    print json.dumps({'success': True})

def handle_release_org():
    """User backed out without finalizing. Clear claim but keep counts."""
    cfg_id = get_data('va_config_id')
    date_iso = get_data('va_date_iso')
    org_id = safe_int(get_data('va_org_id', 0))
    if not cfg_id or not date_iso or org_id <= 0:
        print json.dumps({'success': False, 'message': 'Missing args'})
        return
    sess = get_session(cfg_id, date_iso)
    sst = sess.get('orgs', {}).get(str(org_id))
    if sst:
        sst['claimedBy'] = ''
        sst['claimedAt'] = ''
        write_session(sess)
    print json.dumps({'success': True})

# =====================================================================
# ROSTER + ATTENDANCE WRITES
# =====================================================================

def resolve_meeting_for_org(cfg_id, date_iso, org_id, sess, config):
    """Get or create meeting for this org/date. Caches meetingId in session.
    Returns (meeting_id, error_message).

    Time-resolution priority:
      1) Existing Meeting row for the date (Scheduler pre-created or any past
         meeting we'd reuse). Use its time exactly so GetMeetingIdByDateTime
         finds the same one rather than creating a duplicate.
      2) MeetingSeries.MeetingStart time-of-day (Scheduler RRULE)
      3) OrgSchedule.SchedTime (legacy)
      4) Session globalTime fallback
    """
    sess.setdefault('orgs', {})
    sst = sess['orgs'].setdefault(str(org_id), {})
    if sst.get('meetingId'):
        return sst['meetingId'], None

    dt_date = parse_iso_datetime(date_iso)
    if not dt_date:
        return None, 'Invalid date'

    safe_iso = date_iso.replace("'", "")[:10]
    resolved_time = None

    # 1) Existing Meeting row for this date
    sql = """
        SELECT TOP 1 MeetingDate FROM Meetings
        WHERE OrganizationId = {0}
            AND CONVERT(date, MeetingDate) = '{1}'
            AND (Canceled IS NULL OR Canceled = 0)
        ORDER BY MeetingDate
    """.format(int(org_id), safe_iso)
    for r in q.QuerySql(sql):
        resolved_time = r.MeetingDate
        break

    # 2) MeetingSeries with matching weekly RRULE
    if resolved_time is None:
        try:
            target_weekday = dt_date.weekday()
            day_map = {'MO': 0, 'TU': 1, 'WE': 2, 'TH': 3, 'FR': 4, 'SA': 5, 'SU': 6}
            sql2 = """
                SELECT TOP 5 MeetingStart, RruleString
                FROM MeetingSeries
                WHERE OrganizationId = {0}
            """.format(int(org_id))
            for r in q.QuerySql(sql2):
                rrule = r.RruleString or ''
                if 'FREQ=WEEKLY' not in rrule:
                    continue
                dt_match = re.search(r'DTSTART:(\d{8})', rrule)
                if dt_match:
                    try:
                        dtstart_date = datetime.datetime.strptime(dt_match.group(1), '%Y%m%d').date()
                        if dt_date.date() < dtstart_date:
                            continue
                    except:
                        pass
                bd = re.search(r'BYDAY=([A-Z,]+)', rrule)
                if not bd:
                    continue
                matched = False
                for d in bd.group(1).split(','):
                    d = d.strip()
                    if d in day_map and day_map[d] == target_weekday:
                        matched = True
                        break
                if matched:
                    resolved_time = r.MeetingStart
                    break
        except:
            pass

    # 3) Legacy OrgSchedule
    if resolved_time is None:
        sched_day = (dt_date.weekday() + 1) % 7
        sql3 = """
            SELECT TOP 1 SchedTime FROM OrgSchedule
            WHERE OrganizationId = {0} AND SchedDay = {1}
            ORDER BY SchedTime
        """.format(int(org_id), sched_day)
        for r in q.QuerySql(sql3):
            resolved_time = r.SchedTime
            break

    # 4) Session globalTime fallback
    hour, minute = 9, 0
    if resolved_time is not None:
        try:
            hour = resolved_time.hour
            minute = resolved_time.minute
        except:
            pass
    else:
        gt = (sess.get('globalTime') or '09:00').strip()
        if re.match(r'^\d{1,2}:\d{2}$', gt):
            parts = gt.split(':')
            hour = safe_int(parts[0], 9) % 24
            minute = safe_int(parts[1], 0) % 60

    meeting_dt = datetime.datetime(dt_date.year, dt_date.month, dt_date.day, hour, minute, 0)

    try:
        mid = model.GetMeetingIdByDateTime(int(org_id), meeting_dt, True)
    except Exception as e:
        return None, 'Meeting create failed: ' + str(e)
    if not mid:
        return None, 'Meeting create returned no id'

    sst['meetingId'] = mid
    write_session(sess)
    return mid, None

def get_roster(org_id, meeting_id, exclude_member_types):
    """Return list of [{peopleId, name, age, memberType, attendanceFlag (or None)}]."""
    excl = ''
    if exclude_member_types:
        ids = []
        for tok in str(exclude_member_types).split(','):
            n = safe_int(tok.strip(), 0)
            if n > 0:
                ids.append(n)
        if ids:
            excl = " AND om.MemberTypeId NOT IN (" + ','.join(str(i) for i in ids) + ")"
    sql = """
        SELECT p.PeopleId, p.Name2 AS Name, p.Age,
               ISNULL(mt.Description, '') AS MemberType,
               a.AttendanceFlag
        FROM OrganizationMembers om
        JOIN People p ON p.PeopleId = om.PeopleId
        LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
        LEFT JOIN Attend a ON a.PeopleId = p.PeopleId AND a.MeetingId = {1}
        WHERE om.OrganizationId = {0}
            AND p.IsDeceased = 0
            AND p.ArchivedFlag = 0
            {2}
        ORDER BY p.Name2
    """.format(int(org_id), int(meeting_id), excl)
    rows = []
    for r in q.QuerySql(sql):
        flag = r.AttendanceFlag
        if flag is None:
            state = 'unmarked'
        elif flag is True or flag == 1:
            state = 'present'
        else:
            state = 'absent'
        rows.append({
            'peopleId': r.PeopleId,
            'name': safe_str(r.Name),
            'age': r.Age if r.Age is not None else '',
            'memberType': safe_str(r.MemberType),
            'state': state,
        })
    return rows

def handle_open_roster():
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        org_id = safe_int(get_data('va_org_id', 0))
        volunteer = get_data('va_volunteer', '').strip()
        if not cfg_id or not date_iso or org_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing args'})
            return
        config = find_config(cfg_id)
        if not config:
            print json.dumps({'success': False, 'message': 'Config not found'})
            return
        sess = get_session(cfg_id, date_iso)

        mid, err = resolve_meeting_for_org(cfg_id, date_iso, org_id, sess, config)
        if err:
            print json.dumps({'success': False, 'message': err})
            return

        # Soft claim
        sst = sess['orgs'].setdefault(str(org_id), {})
        sst['claimedBy'] = volunteer or sst.get('claimedBy', '')
        sst['claimedAt'] = now_iso()
        sst['lastTouchBy'] = volunteer or sst.get('lastTouchBy', '')
        sst['lastTouchAt'] = now_iso()
        write_session(sess)

        roster = get_roster(org_id, mid, config.get('excludeMemberTypes', ''))
        # Org name
        org_name = ''
        for r in q.QuerySql("SELECT OrganizationName FROM Organizations WHERE OrganizationId = " + str(int(org_id))):
            org_name = safe_str(r.OrganizationName)
            break

        print json.dumps({
            'success': True,
            'orgId': org_id,
            'orgName': org_name,
            'meetingId': mid,
            'defaultState': config.get('defaultState', 'present'),
            'roster': roster,
            'completed': bool(sst.get('completed')),
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Open roster failed: ' + str(e)})

def handle_mark_attendance():
    """Mark a single person's attendance. Writes to DB via EditPersonAttendance."""
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        org_id = safe_int(get_data('va_org_id', 0))
        meeting_id = safe_int(get_data('va_meeting_id', 0))
        people_id = safe_int(get_data('va_people_id', 0))
        new_state = get_data('va_new_state', '').lower()  # 'present' | 'absent'
        volunteer = get_data('va_volunteer', '').strip()

        if not cfg_id or not date_iso or org_id <= 0 or meeting_id <= 0 or people_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing args'})
            return
        if new_state not in ('present', 'absent'):
            print json.dumps({'success': False, 'message': 'Invalid state'})
            return

        attended = (new_state == 'present')
        try:
            model.EditPersonAttendance(int(meeting_id), int(people_id), attended)
        except Exception as e:
            print json.dumps({'success': False, 'message': 'Write failed: ' + str(e)})
            return

        # Update session touch stamp
        sess = get_session(cfg_id, date_iso)
        sst = sess.get('orgs', {}).setdefault(str(org_id), {})
        sst['lastTouchBy'] = volunteer or sst.get('lastTouchBy', '')
        sst['lastTouchAt'] = now_iso()
        write_session(sess)
        append_log(cfg_id, date_iso, {
            'action': 'mark', 'orgId': org_id, 'peopleId': people_id,
            'state': new_state, 'by': volunteer
        })

        # Return fresh counts for the dashboard widget
        counts = get_attend_counts_for_meetings([meeting_id]).get(meeting_id, {'present': 0, 'absent': 0})
        print json.dumps({'success': True, 'present': counts['present'], 'absent': counts['absent']})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Mark failed: ' + str(e)})

def handle_finalize_org():
    """Apply default state to all unmarked people for this meeting, then mark org complete."""
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        org_id = safe_int(get_data('va_org_id', 0))
        meeting_id = safe_int(get_data('va_meeting_id', 0))
        volunteer = get_data('va_volunteer', '').strip()

        if not cfg_id or not date_iso or org_id <= 0 or meeting_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing args'})
            return

        config = find_config(cfg_id)
        if not config:
            print json.dumps({'success': False, 'message': 'Config not found'})
            return

        default_state = config.get('defaultState', 'present')
        # 'unmarked' default means we DON'T auto-write; user must have marked everyone explicitly.
        # In that case, we just mark complete without touching DB.

        roster = get_roster(org_id, meeting_id, config.get('excludeMemberTypes', ''))
        applied = 0
        skipped = 0
        if default_state in ('present', 'absent'):
            attended = (default_state == 'present')
            for row in roster:
                if row['state'] == 'unmarked':
                    try:
                        model.EditPersonAttendance(int(meeting_id), int(row['peopleId']), attended)
                        applied += 1
                    except:
                        skipped += 1

        sess = get_session(cfg_id, date_iso)
        sst = sess.get('orgs', {}).setdefault(str(org_id), {})
        sst['completed'] = True
        sst['completedBy'] = volunteer or sst.get('completedBy', '')
        sst['completedAt'] = now_iso()
        sst['lastTouchBy'] = volunteer or sst.get('lastTouchBy', '')
        sst['lastTouchAt'] = now_iso()
        # Clear soft-claim now that it's done
        sst['claimedBy'] = ''
        sst['claimedAt'] = ''
        write_session(sess)
        append_log(cfg_id, date_iso, {
            'action': 'finalize', 'orgId': org_id, 'applied': applied, 'skipped': skipped, 'by': volunteer
        })

        counts = get_attend_counts_for_meetings([meeting_id]).get(meeting_id, {'present': 0, 'absent': 0})
        print json.dumps({
            'success': True, 'applied': applied, 'skipped': skipped,
            'present': counts['present'], 'absent': counts['absent']
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Finalize failed: ' + str(e)})

def handle_reopen_org():
    """Allow re-opening a finalized org (in case of mistakes)."""
    cfg_id = get_data('va_config_id')
    date_iso = get_data('va_date_iso')
    org_id = safe_int(get_data('va_org_id', 0))
    volunteer = get_data('va_volunteer', '').strip()
    if not cfg_id or not date_iso or org_id <= 0:
        print json.dumps({'success': False, 'message': 'Missing args'})
        return
    sess = get_session(cfg_id, date_iso)
    sst = sess.get('orgs', {}).setdefault(str(org_id), {})
    sst['completed'] = False
    sst['completedBy'] = ''
    sst['completedAt'] = ''
    sst['lastTouchBy'] = volunteer or sst.get('lastTouchBy', '')
    sst['lastTouchAt'] = now_iso()
    write_session(sess)
    append_log(cfg_id, date_iso, {'action': 'reopen', 'orgId': org_id, 'by': volunteer})
    print json.dumps({'success': True})

# =====================================================================
# WALK-IN PERSON: search, add existing, create new
# Uses MemberType description (e.g. "Visitor") configured per session.
# =====================================================================

def handle_search_people():
    """Find people by name / email / phone with disambiguating fields.

    Smart matching:
      - Splits the search term on whitespace/commas into tokens.
      - For each token, requires a match in at least one name field (FirstName,
        LastName, NickName, AltName, MaidenName, Name, Name2, EmailAddress).
      - All tokens must match (AND between tokens) — so "Be Swab" finds Ben Swaby,
        and "Swa, Be" also finds Ben Swaby.
      - If the term contains 4+ digits, additionally matches against phone numbers.

    Returns up to 20 results; does NOT scope to an org so leaders can pull in
    anyone across the database.
    """
    try:
        term = get_data('va_term', '').strip()
        if len(term) < 2:
            print json.dumps({'success': True, 'results': []})
            return

        # Tokenize on whitespace/commas; strip empties
        tokens = [t for t in re.split(r'[\s,]+', term) if t and len(t) >= 1]
        if not tokens:
            print json.dumps({'success': True, 'results': []})
            return

        digits = re.sub(r'[^\d]', '', term)

        def sql_like_safe(s):
            # Escape SQL LIKE special chars and quote
            return (s.replace("'", "''")
                     .replace('[', '[[]')
                     .replace('%', '[%]')
                     .replace('_', '[_]'))

        name_fields = ['p.FirstName', 'p.LastName', 'p.NickName', 'p.AltName',
                       'p.MaidenName', 'p.Name', 'p.Name2', 'p.EmailAddress']

        # Each token: must match at least one name/email field. Combined with AND across tokens.
        token_clauses = []
        for tok in tokens:
            safe_tok = sql_like_safe(tok)
            field_or = " OR ".join([f + " LIKE '%" + safe_tok + "%'" for f in name_fields])
            token_clauses.append("(" + field_or + ")")
        name_match = "(" + " AND ".join(token_clauses) + ")"

        # Phone fallback — matches against any phone column with non-digits stripped
        phone_clause = ""
        if len(digits) >= 4:
            phone_clause = (
                " OR REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(p.CellPhone, ''),'-',''),' ',''),'(',''),')','') LIKE '%" + digits + "%'"
                " OR REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(p.HomePhone, ''),'-',''),' ',''),'(',''),')','') LIKE '%" + digits + "%'"
                " OR REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(p.WorkPhone, ''),'-',''),' ',''),'(',''),')','') LIKE '%" + digits + "%'"
            )

        wh = ["p.IsDeceased = 0", "p.ArchivedFlag = 0"]
        wh.append("(" + name_match + phone_clause + ")")
        sql = """
            SELECT TOP 20
                p.PeopleId, p.Name2, p.Age,
                LTRIM(RTRIM(ISNULL(p.CityName, '') +
                    CASE WHEN p.StateCode IS NOT NULL AND LEN(LTRIM(p.StateCode)) > 0 THEN ', ' + p.StateCode ELSE '' END +
                    CASE WHEN p.ZipCode IS NOT NULL AND LEN(LTRIM(p.ZipCode)) > 0 THEN ' ' + LEFT(p.ZipCode, 5) ELSE '' END
                )) AS CityStateZip,
                ISNULL(ms.Description, '') AS MemberStatus,
                ISNULL((SELECT TOP 1 sp.Name2 FROM People sp WHERE sp.PeopleId = p.SpouseId), '') AS SpouseName,
                ISNULL(p.EmailAddress, '') AS EmailAddress,
                ISNULL(p.CellPhone, '') AS CellPhone
            FROM People p
            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
            WHERE """ + " AND ".join(wh) + """
            ORDER BY p.Name2
        """
        results = []
        for r in q.QuerySql(sql):
            results.append({
                'peopleId': r.PeopleId,
                'name': safe_str(r.Name2),
                'age': r.Age if r.Age is not None else '',
                'cityStateZip': safe_str(r.CityStateZip),
                'memberStatus': safe_str(r.MemberStatus),
                'spouseName': safe_str(r.SpouseName),
                'email': safe_str(r.EmailAddress),
                'cell': safe_str(r.CellPhone),
            })
        print json.dumps({'success': True, 'results': results})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Search failed: ' + str(e)})

def _add_existing_to_org(people_id, org_id, member_type_desc):
    """JoinOrg + SetMemberType for an existing person. Returns (ok, error)."""
    try:
        person = model.GetPerson(int(people_id))
        if not person or not person.PeopleId:
            return False, 'Person not found'
        try:
            model.JoinOrg(int(org_id), person)
        except Exception as e:
            # JoinOrg can throw "already a member" — that's fine
            if 'already' not in str(e).lower():
                return False, 'JoinOrg failed: ' + str(e)
        if member_type_desc:
            try:
                model.SetMemberType(int(people_id), int(org_id), member_type_desc)
            except Exception as e:
                # Non-fatal — they're added to the org, just maybe wrong type
                pass
        return True, None
    except Exception as e:
        return False, str(e)

def handle_add_walkin_existing():
    """Add an existing person to the org as walk-in and mark them present."""
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        org_id = safe_int(get_data('va_org_id', 0))
        meeting_id = safe_int(get_data('va_meeting_id', 0))
        people_id = safe_int(get_data('va_people_id', 0))
        volunteer = get_data('va_volunteer', '').strip()
        if not cfg_id or not date_iso or org_id <= 0 or meeting_id <= 0 or people_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing args'})
            return
        config = find_config(cfg_id)
        if not config:
            print json.dumps({'success': False, 'message': 'Config not found'})
            return

        member_type_desc = (config.get('walkInMemberType') or 'Visitor').strip()
        ok, err = _add_existing_to_org(people_id, org_id, member_type_desc)
        if not ok:
            print json.dumps({'success': False, 'message': err})
            return

        try:
            model.EditPersonAttendance(int(meeting_id), int(people_id), True)
        except Exception as e:
            print json.dumps({'success': False, 'message': 'Mark attended failed: ' + str(e)})
            return

        sess = get_session(cfg_id, date_iso)
        sst = sess.get('orgs', {}).setdefault(str(org_id), {})
        sst['lastTouchBy'] = volunteer or sst.get('lastTouchBy', '')
        sst['lastTouchAt'] = now_iso()
        write_session(sess)
        append_log(cfg_id, date_iso, {
            'action': 'walkin_existing', 'orgId': org_id,
            'peopleId': people_id, 'memberType': member_type_desc, 'by': volunteer
        })

        # Return the fresh roster + counts so client can refresh
        roster = get_roster(org_id, meeting_id, config.get('excludeMemberTypes', ''))
        counts = get_attend_counts_for_meetings([meeting_id]).get(meeting_id, {'present': 0, 'absent': 0})
        print json.dumps({
            'success': True, 'peopleId': people_id, 'memberType': member_type_desc,
            'roster': roster, 'present': counts['present'], 'absent': counts['absent']
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Add walk-in failed: ' + str(e)})

def handle_create_walkin():
    """Create a brand-new person via AddPerson + add to org + mark present."""
    try:
        cfg_id = get_data('va_config_id')
        date_iso = get_data('va_date_iso')
        org_id = safe_int(get_data('va_org_id', 0))
        meeting_id = safe_int(get_data('va_meeting_id', 0))
        first_name = get_data('va_first_name', '').strip()
        last_name = get_data('va_last_name', '').strip()
        email = get_data('va_email', '').strip()
        cell = get_data('va_cell', '').strip()
        volunteer = get_data('va_volunteer', '').strip()
        if not cfg_id or not date_iso or org_id <= 0 or meeting_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing args'})
            return
        if not first_name or not last_name:
            print json.dumps({'success': False, 'message': 'First and last name are required'})
            return
        config = find_config(cfg_id)
        if not config:
            print json.dumps({'success': False, 'message': 'Config not found'})
            return

        # Try FindAddPerson first — avoids creating duplicates if email matches
        people_id = 0
        if email:
            try:
                people_id = safe_int(model.FindAddPerson(first_name, last_name, email), 0)
            except:
                people_id = 0

        if people_id <= 0:
            # Create from scratch. AddPerson signature:
            # AddPerson(familyId, position, title, firstName, nickName, lastName, dob,
            #          married, gender, address, address2, city, state, zip, homePhone,
            #          email, cellPhone)
            try:
                people_id = safe_int(model.AddPerson(
                    0, 10, '', first_name, '', last_name, '',
                    False, 0, '', '', '', '', '', '',
                    email, cell
                ), 0)
            except Exception as e:
                print json.dumps({'success': False, 'message': 'Create person failed: ' + str(e)})
                return

        if people_id <= 0:
            print json.dumps({'success': False, 'message': 'Could not create or find person'})
            return

        # If they didn't already have a cell phone and one was provided, set it
        if cell:
            try:
                model.UpdatePerson(int(people_id), 'CellPhone', cell)
            except:
                pass

        member_type_desc = (config.get('walkInMemberType') or 'Visitor').strip()
        ok, err = _add_existing_to_org(people_id, org_id, member_type_desc)
        if not ok:
            print json.dumps({'success': False, 'message': err, 'peopleId': people_id})
            return

        try:
            model.EditPersonAttendance(int(meeting_id), int(people_id), True)
        except Exception as e:
            print json.dumps({'success': False, 'message': 'Mark attended failed: ' + str(e)})
            return

        sess = get_session(cfg_id, date_iso)
        sst = sess.get('orgs', {}).setdefault(str(org_id), {})
        sst['lastTouchBy'] = volunteer or sst.get('lastTouchBy', '')
        sst['lastTouchAt'] = now_iso()
        write_session(sess)
        append_log(cfg_id, date_iso, {
            'action': 'walkin_create', 'orgId': org_id,
            'peopleId': people_id, 'name': first_name + ' ' + last_name,
            'memberType': member_type_desc, 'by': volunteer
        })

        roster = get_roster(org_id, meeting_id, config.get('excludeMemberTypes', ''))
        counts = get_attend_counts_for_meetings([meeting_id]).get(meeting_id, {'present': 0, 'absent': 0})
        print json.dumps({
            'success': True, 'peopleId': people_id, 'memberType': member_type_desc,
            'roster': roster, 'present': counts['present'], 'absent': counts['absent']
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Create walk-in failed: ' + str(e)})

# =====================================================================
# DISPATCH
# =====================================================================

if model.HttpMethod == "post":
    action = get_data('va_action', '')
    if action == 'load_configs':
        handle_load_configs()
    elif action == 'save_config':
        handle_save_config()
    elif action == 'delete_config':
        handle_delete_config()
    elif action == 'get_filters':
        handle_get_filters()
    elif action == 'search_involvements':
        handle_search_involvements()
    elif action == 'init_session':
        handle_init_session()
    elif action == 'dashboard':
        handle_dashboard()
    elif action == 'claim_org':
        handle_claim_org()
    elif action == 'release_org':
        handle_release_org()
    elif action == 'open_roster':
        handle_open_roster()
    elif action == 'mark_attendance':
        handle_mark_attendance()
    elif action == 'finalize_org':
        handle_finalize_org()
    elif action == 'reopen_org':
        handle_reopen_org()
    elif action == 'search_people':
        handle_search_people()
    elif action == 'add_walkin_existing':
        handle_add_walkin_existing()
    elif action == 'create_walkin':
        handle_create_walkin()
    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + safe_str(action)})

else:
    # =================================================================
    # GET: render the SPA
    # =================================================================

    css = """
    <style>
    .va-root { font-family: 'Segoe UI', Arial, sans-serif; color: #222; max-width: 1400px; margin: 0 auto; padding: 12px; }
    .va-root * { box-sizing: border-box; }
    .va-h1 { font-size: 22px; font-weight: 700; margin: 0 0 4px 0; color: #1f4e79; }
    .va-sub { font-size: 13px; color: #666; margin: 0 0 16px 0; }
    .va-card { background: #fff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
    .va-row { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; }
    .va-row + .va-row { margin-top: 8px; }
    .va-grow { flex: 1 1 auto; }
    .va-label { display: block; font-size: 12px; color: #555; margin-bottom: 4px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
    .va-input, .va-select { width: 100%; padding: 8px 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; background: #fff; }
    .va-input:focus, .va-select:focus { outline: none; border-color: #1f4e79; box-shadow: 0 0 0 2px rgba(31,78,121,0.15); }
    .va-btn { display: inline-block; padding: 8px 14px; font-size: 14px; font-weight: 600; border: 1px solid #1f4e79; background: #1f4e79; color: #fff; border-radius: 4px; cursor: pointer; text-decoration: none; }
    .va-btn:hover { background: #2a5e8e; border-color: #2a5e8e; }
    .va-btn.va-secondary { background: #fff; color: #1f4e79; }
    .va-btn.va-secondary:hover { background: #f0f4f8; }
    .va-btn.va-danger { background: #c0392b; border-color: #c0392b; }
    .va-btn.va-danger:hover { background: #d04639; border-color: #d04639; }
    .va-btn.va-success { background: #27ae60; border-color: #27ae60; }
    .va-btn.va-success:hover { background: #2ecc71; border-color: #2ecc71; }
    .va-btn.va-sm { padding: 4px 10px; font-size: 13px; }
    .va-btn.va-lg { padding: 12px 24px; font-size: 16px; }
    .va-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .va-pill { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
    .va-pill.va-grey { background: #eef0f2; color: #555; }
    .va-pill.va-blue { background: #d6e6f5; color: #1f4e79; }
    .va-pill.va-green { background: #d4f0db; color: #1f6b3a; }
    .va-pill.va-red { background: #f9d6d6; color: #8a2020; }
    .va-pill.va-amber { background: #fce7c2; color: #7a4a00; }
    .va-muted { color: #888; font-size: 12px; }
    .va-empty { text-align: center; color: #888; padding: 30px 0; }
    .va-config-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px,1fr)); gap: 10px; }
    .va-config-item { border: 1px solid #e1e4e8; border-radius: 6px; padding: 12px; background: #fafbfc; cursor: pointer; transition: all 0.1s; }
    .va-config-item:hover { border-color: #1f4e79; background: #fff; }
    .va-config-item h4 { margin: 0 0 4px 0; font-size: 15px; color: #1f4e79; }
    .va-config-item .va-meta { font-size: 12px; color: #666; }
    .va-config-actions { margin-top: 8px; display: flex; gap: 6px; }
    .va-counter { display: inline-block; padding: 6px 14px; background: #1f4e79; color: #fff; font-size: 14px; font-weight: 700; border-radius: 999px; margin-right: 8px; }
    .va-counter.va-good { background: #27ae60; }
    .va-counter.va-warn { background: #e67e22; }
    .va-dashboard-row { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; padding: 10px 12px; border: 1px solid #e1e4e8; border-radius: 6px; background: #fff; margin-bottom: 6px; cursor: pointer; transition: all 0.1s; }
    .va-dashboard-row:hover { border-color: #1f4e79; background: #f6f9fc; }
    .va-dashboard-row.va-completed { background: #f3faf5; border-color: #c8e6c9; opacity: 0.85; }
    .va-dashboard-row.va-completed:hover { opacity: 1; }
    .va-dashboard-row.va-claimed-other { border-left: 4px solid #e67e22; }
    .va-dashboard-row.va-claimed-self { border-left: 4px solid #1f4e79; }
    .va-org-name { font-weight: 700; font-size: 15px; flex: 1 1 240px; }
    .va-org-meta { font-size: 12px; color: #666; flex: 0 0 auto; }
    .va-counts { display: flex; gap: 4px; flex-wrap: wrap; }
    .va-count-chip { padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; }
    .va-count-chip.va-p { background: #d4f0db; color: #1f6b3a; }
    .va-count-chip.va-a { background: #f9d6d6; color: #8a2020; }
    .va-count-chip.va-u { background: #eef0f2; color: #555; }
    .va-roster-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px,1fr)); gap: 8px; margin-top: 12px; }
    .va-roster-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; border: 2px solid #e1e4e8; border-radius: 6px; cursor: pointer; user-select: none; transition: all 0.08s; background: #fff; min-height: 56px; }
    .va-roster-row .va-name { font-weight: 600; flex: 1; }
    .va-roster-row .va-info { font-size: 11px; color: #888; margin-left: 8px; }
    .va-roster-row.va-state-present { background: #e6f5ec; border-color: #b8e0c4; }
    .va-roster-row.va-state-absent { background: #fbeaea; border-color: #ecbcbc; }
    .va-roster-row.va-state-unmarked { background: #fff; border-color: #ccc; border-style: dashed; }
    .va-roster-row.va-pending { opacity: 0.6; }
    .va-tag { font-size: 10px; padding: 1px 6px; border-radius: 3px; margin-left: 4px; vertical-align: middle; }
    .va-tag.va-p { background: #27ae60; color: #fff; }
    .va-tag.va-a { background: #c0392b; color: #fff; }
    .va-tag.va-u { background: #999; color: #fff; }
    .va-search { width: 100%; padding: 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; }
    .va-toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); padding: 10px 18px; border-radius: 6px; color: #fff; font-weight: 600; z-index: 9999; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    .va-toast.va-success { background: #27ae60; }
    .va-toast.va-error { background: #c0392b; }
    .va-toast.va-info { background: #1f4e79; }
    .va-divider { height: 1px; background: #e1e4e8; margin: 12px 0; }
    .va-toolbar { display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
    .va-toolbar .va-actions { display: flex; gap: 8px; flex-wrap: wrap; }
    .va-toggle-block { display: flex; gap: 8px; flex-wrap: wrap; }
    .va-toggle-block label { display: inline-flex; align-items: center; gap: 6px; padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px; cursor: pointer; background: #fff; font-size: 13px; }
    .va-toggle-block input[type=radio]:checked + span { font-weight: 700; }
    .va-toggle-block label:has(input:checked) { background: #d6e6f5; border-color: #1f4e79; }
    .va-org-search-results { max-height: 260px; overflow-y: auto; border: 1px solid #e1e4e8; border-radius: 4px; margin-top: 6px; }
    .va-org-search-results .va-org-result { padding: 8px 10px; border-bottom: 1px solid #eee; cursor: pointer; }
    .va-org-search-results .va-org-result:hover { background: #f0f4f8; }
    .va-selected-orgs { margin-top: 10px; }
    .va-selected-orgs .va-chip { display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; background: #d6e6f5; color: #1f4e79; border-radius: 4px; margin: 2px; font-size: 13px; }
    .va-selected-orgs .va-chip button { background: transparent; border: 0; color: #1f4e79; cursor: pointer; font-size: 14px; padding: 0 0 0 4px; line-height: 1; }
    .va-progdiv-rows { display: flex; flex-direction: column; gap: 8px; }
    .va-progdiv-rows .va-progdiv-row { display: flex; gap: 8px; align-items: center; }
    .va-progdiv-rows select { flex: 1; }
    .va-completed-section { margin-top: 16px; }
    .va-completed-section h3 { font-size: 14px; color: #555; margin: 0 0 6px 0; }
    @media (max-width: 700px) {
      .va-org-name { flex: 1 1 100%; }
      .va-roster-grid { grid-template-columns: 1fr 1fr; }
      .va-counter { display: block; margin-bottom: 6px; }
    }
    </style>
    """

    js = r"""
    <script>
    (function(){
      var state = {
        view: 'landing',     // landing | edit_config | session_entry | dashboard | roster
        configs: [],
        editingConfig: null,
        programs: [],
        divisions: [],
        currentConfigId: '',
        currentDateIso: '',
        currentDateDisplay: '',
        currentGlobalTime: '09:00',
        currentVolunteer: '',
        dashboardRows: [],
        currentOrg: null,    // { orgId, orgName, meetingId, defaultState, roster }
        rosterFilter: '',
        pollTimer: null,
        walkInOpen: false,    // walk-in panel toggle on roster screen
        walkInSearch: '',     // current search term in walk-in panel
        walkInCreateOpen: false, // "create new" subform toggle
        dashboardFilter: 'all', // 'all' | 'not_started' | 'in_progress' | 'done'
      };

      var root = document.getElementById('vaRoot');

      // ----------------- helpers -----------------
      function el(tag, attrs, children) {
        var e = document.createElement(tag);
        if (attrs) {
          for (var k in attrs) {
            if (k === 'class') e.className = attrs[k];
            else if (k === 'html') e.innerHTML = attrs[k];
            else if (k.indexOf('on') === 0) e.addEventListener(k.substr(2), attrs[k]);
            else e.setAttribute(k, attrs[k]);
          }
        }
        if (children) {
          (children instanceof Array ? children : [children]).forEach(function(c){
            if (c == null) return;
            if (typeof c === 'string') e.appendChild(document.createTextNode(c));
            else e.appendChild(c);
          });
        }
        return e;
      }

      function htmlEscape(s) {
        if (s == null) return '';
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                        .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
      }

      function ajax(action, params, cb) {
        var data = new FormData();
        data.append('va_action', action);
        for (var k in params) {
          if (params[k] != null) data.append(k, params[k]);
        }
        var xhr = new XMLHttpRequest();
        xhr.open('POST', window.location.pathname + window.location.search, true);
        xhr.onreadystatechange = function() {
          if (xhr.readyState !== 4) return;
          if (xhr.status >= 200 && xhr.status < 300) {
            try {
              var resp = JSON.parse(xhr.responseText);
              cb(null, resp);
            } catch(e) {
              cb('Bad response: ' + xhr.responseText.substring(0,200), null);
            }
          } else {
            cb('HTTP ' + xhr.status, null);
          }
        };
        xhr.send(data);
      }

      function toast(msg, kind) {
        var t = el('div', {class: 'va-toast va-' + (kind || 'info')}, msg);
        document.body.appendChild(t);
        setTimeout(function(){ if (t.parentNode) t.parentNode.removeChild(t); }, 2500);
      }

      function todayIso() {
        var d = new Date();
        return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
      }

      function isoToDisplay(iso) {
        if (!iso) return '';
        var parts = iso.split('-');
        if (parts.length !== 3) return iso;
        var d = new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]));
        var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
        var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return days[d.getDay()] + ', ' + months[d.getMonth()] + ' ' + d.getDate();
      }

      function clearPolling() {
        if (state.pollTimer) { clearInterval(state.pollTimer); state.pollTimer = null; }
      }

      // ----------------- render dispatcher -----------------
      function render() {
        clearPolling();
        root.innerHTML = '';
        if (state.view === 'landing') renderLanding();
        else if (state.view === 'edit_config') renderEditConfig();
        else if (state.view === 'session_entry') renderSessionEntry();
        else if (state.view === 'dashboard') renderDashboard();
        else if (state.view === 'roster') renderRoster();
      }

      // ----------------- landing -----------------
      function renderLanding() {
        var head = el('div', {class:'va-toolbar'}, [
          el('div', null, [
            el('div', {class:'va-h1'}, 'Attendance Markings'),
            el('div', {class:'va-sub'}, 'Pick a saved config to start a session, or create a new one.'),
          ]),
          el('div', {class:'va-actions'}, [
            el('button', {class:'va-btn', onclick: function(){ state.editingConfig = null; loadFiltersThen(showEditConfig); }}, '+ New Config'),
          ])
        ]);
        root.appendChild(head);

        var card = el('div', {class:'va-card'});
        if (!state.configs.length) {
          card.appendChild(el('div', {class:'va-empty'}, 'No configs yet. Click "+ New Config" to make one.'));
        } else {
          var grid = el('div', {class:'va-config-list'});
          state.configs.forEach(function(c){
            var summary = c.sourceType === 'specific_orgs'
              ? ((c.specificOrgIds || []).length + ' involvement(s)')
              : ((c.programDivGroups || []).length + ' program/division group(s)');
            var defLabel = c.defaultState === 'absent' ? 'Default Absent' :
                           c.defaultState === 'unmarked' ? 'Default Unmarked' : 'Default Present';
            var item = el('div', {class:'va-config-item'}, [
              el('h4', null, c.name || '(no name)'),
              el('div', {class:'va-meta'}, summary),
              el('div', {class:'va-meta'}, defLabel + ' · ' + (c.onlyWithMeeting ? 'Schedule-only' : 'All in scope')),
              el('div', {class:'va-config-actions'}, [
                el('button', {class:'va-btn va-sm', onclick: function(ev){ ev.stopPropagation(); startSession(c); }}, 'Start Session'),
                el('button', {class:'va-btn va-sm va-secondary', onclick: function(ev){ ev.stopPropagation(); state.editingConfig = JSON.parse(JSON.stringify(c)); loadFiltersThen(showEditConfig); }}, 'Edit'),
                el('button', {class:'va-btn va-sm va-danger', onclick: function(ev){ ev.stopPropagation(); deleteConfig(c); }}, 'Delete'),
              ])
            ]);
            grid.appendChild(item);
          });
          card.appendChild(grid);
        }
        root.appendChild(card);
      }

      function loadConfigs(after) {
        ajax('load_configs', {}, function(err, resp){
          if (err || !resp || !resp.success) { toast('Failed to load configs', 'error'); return; }
          state.configs = resp.configs || [];
          if (after) after();
        });
      }

      function deleteConfig(c) {
        if (!confirm('Delete config "' + c.name + '"?')) return;
        ajax('delete_config', {va_config_id: c.id}, function(err, resp){
          if (err || !resp.success) { toast('Delete failed', 'error'); return; }
          loadConfigs(render);
        });
      }

      // ----------------- edit config -----------------
      function showEditConfig() {
        state.view = 'edit_config';
        if (!state.editingConfig) {
          state.editingConfig = {
            id: '',
            name: '',
            sourceType: 'program_division',
            programDivGroups: [{programId:0, divisionId:0}],
            specificOrgIds: [],
            specificOrgNames: {},
            excludeOrgIds: [],
            defaultState: 'present',
            onlyWithMeeting: true,
            excludeMemberTypes: '',
            walkInMemberType: 'Visitor',
            allowWalkIns: true,
          };
        }
        if (!state.editingConfig.programDivGroups || !state.editingConfig.programDivGroups.length) {
          state.editingConfig.programDivGroups = [{programId:0, divisionId:0}];
        }
        render();
      }

      function loadFiltersThen(after) {
        if (state.programs.length) { after(); return; }
        ajax('get_filters', {}, function(err, resp){
          if (err || !resp.success) { toast('Failed to load filters', 'error'); return; }
          state.programs = resp.programs || [];
          state.divisions = resp.divisions || [];
          after();
        });
      }

      function renderEditConfig() {
        var c = state.editingConfig;
        var head = el('div', {class:'va-toolbar'}, [
          el('div', {class:'va-h1'}, c.id ? 'Edit Config' : 'New Config'),
          el('div', {class:'va-actions'}, [
            el('button', {class:'va-btn va-secondary', onclick: function(){ state.view='landing'; render(); }}, '← Back'),
          ])
        ]);
        root.appendChild(head);

        var card = el('div', {class:'va-card'});

        // Name
        var nameRow = el('div', {class:'va-row'}, [
          el('div', {class:'va-grow'}, [
            el('label', {class:'va-label'}, 'Name'),
            el('input', {class:'va-input', type:'text', value: c.name || '', placeholder:'e.g., VBS 2026 Mornings',
              oninput: function(e){ c.name = e.target.value; }})
          ]),
        ]);
        card.appendChild(nameRow);

        // Source type
        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('label', {class:'va-label'}, 'Source'));
        var srcBlock = el('div', {class:'va-toggle-block'}, [
          radioToggle('va_src', 'program_division', 'Program / Division', c.sourceType === 'program_division', function(){ c.sourceType='program_division'; render(); }),
          radioToggle('va_src', 'specific_orgs', 'Specific Involvements', c.sourceType === 'specific_orgs', function(){ c.sourceType='specific_orgs'; render(); }),
        ]);
        card.appendChild(srcBlock);

        if (c.sourceType === 'program_division') {
          card.appendChild(renderProgDivBlock(c));
        } else {
          card.appendChild(renderSpecificOrgBlock(c));
        }

        // Default state
        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('label', {class:'va-label'}, 'Default Attendance State'));
        card.appendChild(el('div', {class:'va-toggle-block'}, [
          radioToggle('va_def', 'present', 'Default Present (click to mark Absent)', c.defaultState === 'present', function(){ c.defaultState='present'; }),
          radioToggle('va_def', 'absent',  'Default Absent (click to mark Present)', c.defaultState === 'absent',  function(){ c.defaultState='absent'; }),
          radioToggle('va_def', 'unmarked','Default Unmarked (must mark each)',     c.defaultState === 'unmarked',function(){ c.defaultState='unmarked'; }),
        ]));
        card.appendChild(el('div', {class:'va-muted'}, 'On Finalize, any Unmarked person gets the default state — except in Default Unmarked mode, where Finalize requires every person to be explicitly marked.'));

        // Schedule filter
        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('label', {class:'va-label'}, 'Involvement Filter'));
        var schedRow = el('label', null, [
          el('input', {type:'checkbox', checked: c.onlyWithMeeting, onchange: function(e){ c.onlyWithMeeting = e.target.checked; }}),
          ' Only show involvements with a meeting scheduled on the selected day'
        ]);
        card.appendChild(schedRow);
        card.appendChild(el('div', {class:'va-muted'}, 'When OFF, all involvements in scope appear. You\'ll be asked for a single global meeting time when starting the session.'));

        // Exclude member types
        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('label', {class:'va-label'}, 'Exclude Member Type IDs (comma-separated, optional)'));
        card.appendChild(el('input', {class:'va-input', type:'text', value: c.excludeMemberTypes || '', placeholder:'e.g., 140,310,320 to exclude leaders',
          oninput: function(e){ c.excludeMemberTypes = e.target.value; }}));
        card.appendChild(el('div', {class:'va-muted'}, 'Leaders are typically 140, 310, 320. Leave blank to include everyone enrolled.'));

        // Walk-in defaults
        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('label', {class:'va-label'}, 'Walk-Ins'));
        card.appendChild(el('label', null, [
          el('input', {type:'checkbox', checked: c.allowWalkIns !== false, onchange: function(e){ c.allowWalkIns = e.target.checked; }}),
          ' Allow leaders to add walk-in people from the roster screen'
        ]));
        card.appendChild(el('div', {style:'margin-top:8px'}, [
          el('label', {class:'va-label'}, 'Add walk-ins as'),
          el('select', {class:'va-select', onchange: function(e){ c.walkInMemberType = e.target.value || 'Visitor'; }},
            ['Visitor', 'Member', 'Prospect', 'Guest', 'Inactive Member'].map(function(opt){
              var o = el('option', {value: opt}, opt);
              if ((c.walkInMemberType || 'Visitor') === opt) o.selected = true;
              return o;
            }))
        ]));
        card.appendChild(el('div', {class:'va-muted'}, 'Member type applied automatically when a leader adds a walk-in. Use whichever description matches your church\'s lookup.MemberType table.'));

        // Save
        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('div', {class:'va-row'}, [
          el('button', {class:'va-btn va-success va-lg', onclick: saveCurrentConfig}, 'Save Config'),
          el('button', {class:'va-btn va-secondary', onclick: function(){ state.view='landing'; render(); }}, 'Cancel'),
        ]));

        root.appendChild(card);
      }

      function radioToggle(name, value, label, checked, onChange) {
        var wrap = el('label', null);
        var input = el('input', {type:'radio', name:name, value:value});
        if (checked) input.checked = true;
        input.addEventListener('change', function(){ if (input.checked) onChange(); });
        wrap.appendChild(input);
        wrap.appendChild(el('span', null, label));
        return wrap;
      }

      function renderProgDivBlock(c) {
        var wrap = el('div', {class:'va-progdiv-rows'});
        var hint = el('div', {class:'va-muted'}, 'Pick one or more Program / Division combinations. Leave Division as "(any)" to include all divisions in that program.');
        wrap.appendChild(hint);

        function refresh() {
          // Re-render block in place
          var newBlock = renderProgDivBlock(c);
          wrap.parentNode.replaceChild(newBlock, wrap);
        }

        (c.programDivGroups || []).forEach(function(g, idx){
          var row = el('div', {class:'va-progdiv-row'});
          var pSel = el('select', {class:'va-select', onchange: function(e){
            g.programId = parseInt(e.target.value || 0);
            g.divisionId = 0;
            refresh();
          }});
          pSel.appendChild(el('option', {value:0}, '(any program)'));
          state.programs.forEach(function(p){
            var opt = el('option', {value:p.id}, p.name);
            if (p.id === g.programId) opt.selected = true;
            pSel.appendChild(opt);
          });
          var dSel = el('select', {class:'va-select', onchange: function(e){
            g.divisionId = parseInt(e.target.value || 0);
          }});
          dSel.appendChild(el('option', {value:0}, '(any division)'));
          state.divisions.filter(function(d){ return !g.programId || d.programId === g.programId; }).forEach(function(d){
            var opt = el('option', {value:d.id}, d.name);
            if (d.id === g.divisionId) opt.selected = true;
            dSel.appendChild(opt);
          });
          var rmBtn = el('button', {class:'va-btn va-sm va-danger', onclick: function(){
            c.programDivGroups.splice(idx, 1);
            if (!c.programDivGroups.length) c.programDivGroups.push({programId:0, divisionId:0});
            refresh();
          }}, '×');
          row.appendChild(pSel);
          row.appendChild(dSel);
          row.appendChild(rmBtn);
          wrap.appendChild(row);
        });

        var addBtn = el('button', {class:'va-btn va-sm va-secondary', onclick: function(){
          c.programDivGroups.push({programId:0, divisionId:0});
          refresh();
        }}, '+ Add Program/Division');
        wrap.appendChild(addBtn);
        return wrap;
      }

      function renderSpecificOrgBlock(c) {
        var wrap = el('div');
        wrap.appendChild(el('div', {class:'va-muted'}, 'Search by name and click to add. Selected involvements appear below.'));
        var search = el('input', {class:'va-search', type:'text', placeholder:'Search involvements...',
          oninput: function(e){ doSearch(e.target.value); }});
        wrap.appendChild(search);
        var resultsBox = el('div', {class:'va-org-search-results', style:'display:none'});
        wrap.appendChild(resultsBox);

        function doSearch(term) {
          term = term.trim();
          if (term.length < 2) { resultsBox.style.display='none'; return; }
          ajax('search_involvements', {va_term: term}, function(err, resp){
            if (err || !resp.success) return;
            resultsBox.innerHTML = '';
            (resp.results || []).forEach(function(o){
              var r = el('div', {class:'va-org-result', onclick: function(){ addOrg(o); search.value=''; resultsBox.style.display='none'; }}, [
                el('strong', null, o.orgName),
                ' ',
                el('span', {class:'va-muted'}, '(' + (o.programName || '') + (o.divisionName ? ' / ' + o.divisionName : '') + ', ' + o.memberCount + ' enrolled)')
              ]);
              resultsBox.appendChild(r);
            });
            resultsBox.style.display = resp.results && resp.results.length ? 'block' : 'none';
          });
        }

        function addOrg(o) {
          if (!c.specificOrgIds) c.specificOrgIds = [];
          if (!c.specificOrgNames) c.specificOrgNames = {};
          if (c.specificOrgIds.indexOf(o.orgId) === -1) {
            c.specificOrgIds.push(o.orgId);
            c.specificOrgNames[String(o.orgId)] = o.orgName;
            redrawSelected();
          }
        }

        var selectedBox = el('div', {class:'va-selected-orgs'});
        function redrawSelected() {
          selectedBox.innerHTML = '';
          (c.specificOrgIds || []).forEach(function(oid){
            var nm = (c.specificOrgNames || {})[String(oid)] || ('Org ' + oid);
            var chip = el('span', {class:'va-chip'}, [
              nm,
              el('button', {onclick: function(){
                var idx = c.specificOrgIds.indexOf(oid);
                if (idx >= 0) c.specificOrgIds.splice(idx, 1);
                if (c.specificOrgNames) delete c.specificOrgNames[String(oid)];
                redrawSelected();
              }}, '×')
            ]);
            selectedBox.appendChild(chip);
          });
        }
        redrawSelected();
        wrap.appendChild(selectedBox);
        return wrap;
      }

      function saveCurrentConfig() {
        var c = state.editingConfig;
        if (!c.name || !c.name.trim()) { toast('Name is required', 'error'); return; }
        if (c.sourceType === 'program_division') {
          var hasAny = (c.programDivGroups || []).some(function(g){ return g.programId > 0 || g.divisionId > 0; });
          if (!hasAny) { toast('Add at least one Program or Division', 'error'); return; }
        } else {
          if (!(c.specificOrgIds || []).length) { toast('Add at least one involvement', 'error'); return; }
        }
        ajax('save_config', {va_config_json: JSON.stringify(c)}, function(err, resp){
          if (err || !resp.success) { toast('Save failed: ' + (resp ? resp.message : err), 'error'); return; }
          toast('Saved', 'success');
          loadConfigs(function(){ state.view='landing'; render(); });
        });
      }

      // ----------------- session entry -----------------
      function startSession(c) {
        state.currentConfigId = c.id;
        state.currentDateIso = todayIso();
        state.currentDateDisplay = isoToDisplay(state.currentDateIso);
        state.currentGlobalTime = '09:00';
        state.currentVolunteer = state.currentVolunteer || '';
        state.view = 'session_entry';
        render();
      }

      function renderSessionEntry() {
        var c = state.configs.filter(function(x){ return x.id === state.currentConfigId; })[0];
        if (!c) { state.view='landing'; render(); return; }

        var head = el('div', {class:'va-toolbar'}, [
          el('div', null, [
            el('div', {class:'va-h1'}, 'Start Session'),
            el('div', {class:'va-sub'}, c.name),
          ]),
          el('div', {class:'va-actions'}, [
            el('button', {class:'va-btn va-secondary', onclick: function(){ state.view='landing'; render(); }}, '← Back'),
          ])
        ]);
        root.appendChild(head);

        var card = el('div', {class:'va-card'});

        card.appendChild(el('label', {class:'va-label'}, 'Date'));
        var dateInput = el('input', {class:'va-input', type:'date', value: state.currentDateIso,
          oninput: function(e){ state.currentDateIso = e.target.value; state.currentDateDisplay = isoToDisplay(e.target.value); displayDate.textContent = state.currentDateDisplay; }});
        var displayDate = el('div', {class:'va-muted', style:'margin-top:4px'}, state.currentDateDisplay);
        card.appendChild(dateInput);
        card.appendChild(displayDate);

        if (!c.onlyWithMeeting) {
          card.appendChild(el('div', {class:'va-divider'}));
          card.appendChild(el('label', {class:'va-label'}, 'Global Meeting Time'));
          card.appendChild(el('input', {class:'va-input', type:'time', value: state.currentGlobalTime,
            oninput: function(e){ state.currentGlobalTime = e.target.value || '09:00'; }}));
          card.appendChild(el('div', {class:'va-muted'}, 'Used for any involvement without a scheduled meeting on this date.'));
        }

        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('label', {class:'va-label'}, 'Your Name'));
        card.appendChild(el('input', {class:'va-input', type:'text', value: state.currentVolunteer, placeholder:'e.g., Mary Smith',
          oninput: function(e){ state.currentVolunteer = e.target.value; }}));
        card.appendChild(el('div', {class:'va-muted'}, 'Shown to teammates so everyone can see who is working which involvement.'));

        card.appendChild(el('div', {class:'va-divider'}));
        card.appendChild(el('button', {class:'va-btn va-success va-lg', onclick: enterDashboard}, 'Begin Attendance'));

        root.appendChild(card);
      }

      function enterDashboard() {
        if (!state.currentDateIso) { toast('Pick a date', 'error'); return; }
        if (!state.currentVolunteer || !state.currentVolunteer.trim()) { toast('Enter your name', 'error'); return; }
        ajax('init_session', {
          va_config_id: state.currentConfigId,
          va_date_iso: state.currentDateIso,
          va_global_time: state.currentGlobalTime,
          va_volunteer: state.currentVolunteer.trim(),
        }, function(err, resp){
          if (err || !resp.success) { toast('Could not start: ' + (resp ? resp.message : err), 'error'); return; }
          state.view = 'dashboard';
          render();
        });
      }

      // ----------------- dashboard -----------------
      function renderDashboard() {
        var c = state.configs.filter(function(x){ return x.id === state.currentConfigId; })[0];
        var head = el('div', {class:'va-toolbar'}, [
          el('div', null, [
            el('div', {class:'va-h1'}, 'Attendance Dashboard'),
            el('div', {class:'va-sub'}, (c ? c.name : '') + ' · ' + state.currentDateDisplay + ' · ' + htmlEscape(state.currentVolunteer)),
          ]),
          el('div', {class:'va-actions'}, [
            el('button', {class:'va-btn va-secondary', onclick: function(){ state.view='session_entry'; render(); }}, 'Change'),
            el('button', {class:'va-btn va-secondary', onclick: function(){ state.view='landing'; render(); }}, '← Home'),
          ])
        ]);
        root.appendChild(head);

        var statsCard = el('div', {class:'va-card', id:'vaStatsCard'}, 'Loading…');
        root.appendChild(statsCard);

        var listCard = el('div', {class:'va-card', id:'vaListCard'}, '');
        root.appendChild(listCard);

        function loadDashboard() {
          ajax('dashboard', {va_config_id: state.currentConfigId, va_date_iso: state.currentDateIso}, function(err, resp){
            if (err || !resp.success) { return; }
            state.dashboardRows = resp.rows || [];
            paintDashboard(resp);
          });
        }

        function paintDashboard(resp) {
          var rows = resp.rows || [];
          var total = rows.length;

          // Bucket each row by status
          function statusOf(r) {
            if (r.completed) return 'done';
            if ((r.present || 0) + (r.absent || 0) > 0) return 'in_progress';
            return 'not_started';
          }
          var notStarted = rows.filter(function(r){ return statusOf(r) === 'not_started'; }).length;
          var inProgress = rows.filter(function(r){ return statusOf(r) === 'in_progress'; }).length;
          var done = rows.filter(function(r){ return statusOf(r) === 'done'; }).length;
          var remaining = notStarted + inProgress;
          var pct = total ? Math.round(done * 100 / total) : 0;

          // Clickable filter pills
          function filterPill(key, label, count, baseClass) {
            var active = state.dashboardFilter === key;
            return el('span', {
              class: 'va-counter ' + (baseClass || '') + (active ? ' va-counter-active' : ''),
              style: 'cursor:pointer;' + (active ? 'box-shadow:0 0 0 3px rgba(31,78,121,0.25);' : 'opacity:0.85;'),
              title: 'Click to filter to ' + label,
              onclick: function(){ state.dashboardFilter = key; paintDashboard(resp); }
            }, label + ': ' + count);
          }
          var remainingClass = remaining === 0 ? 'va-good' : (remaining < total/2 ? 'va-warn' : '');

          statsCard.innerHTML = '';
          statsCard.appendChild(el('div', {class:'va-row', style:'flex-wrap:wrap;gap:6px;'}, [
            filterPill('all', 'All', total, ''),
            filterPill('not_started', 'Not Started', notStarted, ''),
            filterPill('in_progress', 'In Progress', inProgress, 'va-warn'),
            filterPill('done', 'Done', done, 'va-good'),
          ]));
          statsCard.appendChild(el('div', {class:'va-muted', style:'margin-top:6px;font-size:11px;'},
            'Remaining: ' + remaining + ' / ' + total + ' (' + pct + '% done) · Auto-refreshes every 10 seconds · Click any pill above to filter'));

          // Apply filter to row list
          var filterKey = state.dashboardFilter || 'all';
          var visibleRows = (filterKey === 'all') ? rows : rows.filter(function(r){ return statusOf(r) === filterKey; });

          listCard.innerHTML = '';
          if (!total) {
            listCard.appendChild(el('div', {class:'va-empty'}, 'No involvements in scope for this date.'));
            return;
          }
          if (!visibleRows.length) {
            listCard.appendChild(el('div', {class:'va-empty'}, 'No involvements match the current filter.'));
            return;
          }

          // When viewing All, keep the visual split (Not Started → In Progress → Done).
          // For specific filters, just render the matching rows in order.
          if (filterKey === 'all') {
            var nsRows = visibleRows.filter(function(r){ return statusOf(r) === 'not_started'; });
            var ipRows = visibleRows.filter(function(r){ return statusOf(r) === 'in_progress'; });
            var dnRows = visibleRows.filter(function(r){ return statusOf(r) === 'done'; });
            if (nsRows.length) {
              listCard.appendChild(el('h3', {style:'margin:0 0 6px 0;font-size:14px;color:#555;'}, 'Not Started (' + nsRows.length + ')'));
              nsRows.forEach(function(r){ listCard.appendChild(buildOrgRow(r)); });
            }
            if (ipRows.length) {
              listCard.appendChild(el('h3', {style:'margin:14px 0 6px 0;font-size:14px;color:#7a4a00;'}, 'In Progress (' + ipRows.length + ')'));
              ipRows.forEach(function(r){ listCard.appendChild(buildOrgRow(r)); });
            }
            if (dnRows.length) {
              listCard.appendChild(el('div', {class:'va-completed-section'}, [
                el('h3', null, 'Done (' + dnRows.length + ')'),
              ]));
              dnRows.forEach(function(r){ listCard.appendChild(buildOrgRow(r)); });
            }
          } else {
            visibleRows.forEach(function(r){ listCard.appendChild(buildOrgRow(r)); });
          }
        }

        function buildOrgRow(r) {
          var classes = ['va-dashboard-row'];
          if (r.completed) classes.push('va-completed');
          if (r.claimedBy && r.claimedBy === state.currentVolunteer) classes.push('va-claimed-self');
          else if (r.claimedBy) classes.push('va-claimed-other');

          var statusPill;
          if (r.completed) {
            statusPill = el('span', {class:'va-pill va-green'}, 'Done · ' + (r.completedBy || '') + (r.completedAt ? ' ' + (minutesAgoLabel(r.completedAt)) : ''));
          } else if (r.claimedBy) {
            statusPill = el('span', {class:'va-pill va-amber'}, 'Working: ' + r.claimedBy + (r.claimedAt ? ' (' + minutesAgoLabel(r.claimedAt) + ')' : ''));
          } else if (r.present + r.absent > 0) {
            statusPill = el('span', {class:'va-pill va-blue'}, 'In Progress');
          } else {
            statusPill = el('span', {class:'va-pill va-grey'}, 'Not Started');
          }

          var counts = el('div', {class:'va-counts'}, [
            el('span', {class:'va-count-chip va-p'}, 'P ' + r.present),
            el('span', {class:'va-count-chip va-a'}, 'A ' + r.absent),
            el('span', {class:'va-count-chip va-u'}, 'U ' + r.unmarked),
          ]);

          var meta = [];
          if (r.programName) meta.push(r.programName);
          if (r.divisionName) meta.push(r.divisionName);
          if (r.schedTime) meta.push(r.schedTime);
          meta.push(r.enrolled + ' enrolled');

          var row = el('div', {class: classes.join(' '), onclick: function(){ openRoster(r); }}, [
            el('div', {class:'va-org-name'}, r.orgName),
            el('div', {class:'va-org-meta'}, meta.join(' · ')),
            statusPill,
            counts,
          ]);
          return row;
        }

        loadDashboard();
        state.pollTimer = setInterval(loadDashboard, 10000);
      }

      function minutesAgoLabel(iso) {
        if (!iso) return '';
        var dt;
        try { dt = new Date(iso); } catch(e) { return ''; }
        if (isNaN(dt.getTime())) return '';
        var mins = Math.floor((Date.now() - dt.getTime()) / 60000);
        if (mins < 1) return 'just now';
        if (mins < 60) return mins + 'm ago';
        var hrs = Math.floor(mins/60);
        if (hrs < 24) return hrs + 'h ago';
        return Math.floor(hrs/24) + 'd ago';
      }

      // ----------------- roster -----------------
      function openRoster(row) {
        ajax('open_roster', {
          va_config_id: state.currentConfigId,
          va_date_iso: state.currentDateIso,
          va_org_id: row.orgId,
          va_volunteer: state.currentVolunteer,
        }, function(err, resp){
          if (err || !resp.success) { toast('Failed to open: ' + (resp ? resp.message : err), 'error'); return; }
          state.currentOrg = {
            orgId: resp.orgId,
            orgName: resp.orgName,
            meetingId: resp.meetingId,
            defaultState: resp.defaultState,
            roster: resp.roster || [],
            completed: !!resp.completed,
          };
          state.rosterFilter = '';
          state.view = 'roster';
          render();
        });
      }

      function renderRoster() {
        var o = state.currentOrg;
        if (!o) { state.view='dashboard'; render(); return; }

        var head = el('div', {class:'va-toolbar'}, [
          el('div', null, [
            el('div', {class:'va-h1'}, o.orgName),
            el('div', {class:'va-sub'}, state.currentDateDisplay + ' · ' + htmlEscape(state.currentVolunteer) + (o.completed ? ' · COMPLETED' : '')),
          ]),
          el('div', {class:'va-actions'}, [
            el('button', {class:'va-btn va-secondary', onclick: backToDashboard}, '← Dashboard'),
          ])
        ]);
        root.appendChild(head);

        // Stats / actions card
        var statsCard = el('div', {class:'va-card', id:'vaRosterStats'});
        root.appendChild(statsCard);
        // List card
        var listCard = el('div', {class:'va-card', id:'vaRosterList'});
        root.appendChild(listCard);

        function paint() {
          var present = 0, absent = 0, unmarked = 0;
          o.roster.forEach(function(p){
            if (p.state === 'present') present++;
            else if (p.state === 'absent') absent++;
            else unmarked++;
          });

          statsCard.innerHTML = '';
          var defLabel = o.defaultState === 'absent' ? 'Default Absent' :
                         o.defaultState === 'unmarked' ? 'Default Unmarked' : 'Default Present';
          statsCard.appendChild(el('div', {class:'va-row'}, [
            el('span', {class:'va-counter va-good'}, 'Present: ' + present),
            el('span', {class:'va-counter va-warn'}, 'Absent: ' + absent),
            el('span', {class:'va-counter'}, 'Unmarked: ' + unmarked),
            el('span', {class:'va-muted'}, defLabel + ' · ' + o.roster.length + ' total'),
          ]));
          // Legend explaining the per-person P/A/U chips
          statsCard.appendChild(el('div', {class:'va-muted', style:'margin-top:6px;font-size:11px;'}, [
            'Per-person tags: ',
            el('span', {class:'va-tag va-p', style:'margin-right:2px'}, 'P'), ' = Present · ',
            el('span', {class:'va-tag va-a', style:'margin-right:2px'}, 'A'), ' = Absent · ',
            el('span', {class:'va-tag va-u', style:'margin-right:2px'}, 'U'), ' = Unmarked (gets the default on Finalize)'
          ]));

          var c = state.configs.filter(function(x){ return x.id === state.currentConfigId; })[0] || {};
          var allowWalkIns = c.allowWalkIns !== false;
          var walkInType = c.walkInMemberType || 'Visitor';

          var actions = el('div', {class:'va-row', style:'margin-top:10px;flex-wrap:wrap;gap:6px;'});
          actions.appendChild(el('input', {class:'va-input', type:'text', placeholder:'Filter by name…', value: state.rosterFilter,
            oninput: function(e){ state.rosterFilter = e.target.value; renderList(); }, style:'max-width:240px;'}));
          if (!o.completed && allowWalkIns) {
            actions.appendChild(el('button', {class:'va-btn va-secondary', onclick: function(){ state.walkInOpen = !state.walkInOpen; paint(); }}, state.walkInOpen ? 'Close Add Person' : '+ Add Person'));
          }
          // Headcount link → opens the TouchPoint meeting page in a new tab so leaders
          // can enter official HeadCount / NumNewVisit values via TouchPoint's native UI.
          // (TouchPoint's Python API doesn't expose a setter for those columns.)
          if (o.meetingId) {
            actions.appendChild(el('a', {
              class: 'va-btn va-secondary',
              href: '/Meeting/' + o.meetingId,
              target: '_blank',
              title: 'Set HeadCount, NumNewVisit, etc. in TouchPoint'
            }, 'Headcount in TouchPoint ↗'));
          }
          if (o.completed) {
            actions.appendChild(el('button', {class:'va-btn va-secondary', onclick: reopenOrg}, 'Re-open'));
          } else {
            actions.appendChild(el('button', {class:'va-btn va-success', onclick: finalizeOrg}, 'Finalize ' + (unmarked ? '(' + unmarked + ' will become ' + (o.defaultState === 'absent' ? 'Absent' : (o.defaultState === 'unmarked' ? 'left unmarked' : 'Present')) + ')' : '')));
          }
          statsCard.appendChild(actions);

          if (state.walkInOpen && !o.completed && allowWalkIns) {
            statsCard.appendChild(renderWalkInPanel(walkInType));
          }

          renderList();
        }

        function renderWalkInPanel(walkInType) {
          var panel = el('div', {class:'va-card', style:'margin-top:12px;background:#f6f9fc;border-color:#bcd5ee;'});
          panel.appendChild(el('div', {class:'va-row'}, [
            el('strong', null, 'Add walk-in as: '),
            el('span', {class:'va-pill va-blue'}, walkInType),
            el('span', {class:'va-muted', style:'margin-left:8px'}, '(set in config)')
          ]));

          var searchRow = el('div', {class:'va-row', style:'margin-top:8px'});
          var searchInput = el('input', {class:'va-input', type:'text', placeholder:'Start typing name, email, or phone…', value: state.walkInSearch || '', style:'flex:1', oninput: function(e){
            state.walkInSearch = e.target.value;
            if (state._walkInDebounce) clearTimeout(state._walkInDebounce);
            // 250ms debounce — feels responsive without flooding the server on every keystroke
            state._walkInDebounce = setTimeout(function(){
              doWalkInSearch();
            }, 250);
          }});
          var searchBtn = el('button', {class:'va-btn', onclick: function(){
            if (state._walkInDebounce) clearTimeout(state._walkInDebounce);
            doWalkInSearch();
          }}, 'Search');
          searchInput.addEventListener('keydown', function(e){
            if (e.key === 'Enter') {
              e.preventDefault();
              if (state._walkInDebounce) clearTimeout(state._walkInDebounce);
              doWalkInSearch();
            }
          });
          searchRow.appendChild(searchInput);
          searchRow.appendChild(searchBtn);
          panel.appendChild(searchRow);
          // Auto-focus the search input when the panel opens
          setTimeout(function(){ try { searchInput.focus(); } catch(e){} }, 50);

          var resultsBox = el('div', {id:'vaWalkInResults', style:'margin-top:8px'});
          panel.appendChild(resultsBox);

          // Create-new form (collapsed by default)
          var createBtn = el('button', {class:'va-btn va-secondary va-sm', style:'margin-top:8px', onclick: function(){ state.walkInCreateOpen = !state.walkInCreateOpen; paint(); }}, state.walkInCreateOpen ? 'Cancel new contact' : '+ Can\'t find them — create new');
          panel.appendChild(createBtn);

          if (state.walkInCreateOpen) {
            var formGrid = el('div', {style:'display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px'});
            formGrid.appendChild(el('input', {class:'va-input', type:'text', placeholder:'First name', id:'vaNewFirst'}));
            formGrid.appendChild(el('input', {class:'va-input', type:'text', placeholder:'Last name', id:'vaNewLast'}));
            formGrid.appendChild(el('input', {class:'va-input', type:'email', placeholder:'Email (optional)', id:'vaNewEmail'}));
            formGrid.appendChild(el('input', {class:'va-input', type:'tel', placeholder:'Cell (optional)', id:'vaNewCell'}));
            panel.appendChild(formGrid);
            panel.appendChild(el('button', {class:'va-btn va-success', style:'margin-top:8px', onclick: function(){ doCreateWalkIn(); }}, 'Create & Mark Present'));
          }

          return panel;
        }

        function doWalkInSearch() {
          var box = document.getElementById('vaWalkInResults');
          if (!box) return;
          var term = (state.walkInSearch || '').trim();
          if (term.length < 2) {
            box.innerHTML = '<div class="va-muted">Type at least 2 characters.</div>';
            return;
          }
          box.innerHTML = '<div class="va-muted">Searching…</div>';
          ajax('search_people', {va_term: term}, function(err, resp){
            if (err || !resp.success) { box.innerHTML = '<div class="va-pill va-red">Search failed</div>'; return; }
            box.innerHTML = '';
            if (!resp.results || !resp.results.length) {
              box.appendChild(el('div', {class:'va-muted'}, 'No matches.'));
              return;
            }
            resp.results.forEach(function(p){
              var bits = [];
              if (p.age !== '' && p.age != null) bits.push('Age ' + p.age);
              if (p.cityStateZip) bits.push(p.cityStateZip);
              if (p.spouseName) bits.push('spouse: ' + p.spouseName);
              if (p.memberStatus) bits.push(p.memberStatus);
              var contactBits = [];
              if (p.email) contactBits.push(p.email);
              if (p.cell) contactBits.push(p.cell);
              var row = el('div', {style:'display:flex;justify-content:space-between;align-items:center;padding:8px 10px;border:1px solid #e1e4e8;border-radius:6px;margin-bottom:4px;background:#fff;'}, [
                el('div', null, [
                  el('div', {style:'font-weight:600'}, p.name),
                  el('div', {class:'va-muted'}, bits.join(' · ')),
                  contactBits.length ? el('div', {class:'va-muted', style:'font-size:11px'}, contactBits.join(' · ')) : null,
                ]),
                el('button', {class:'va-btn va-sm va-success', onclick: function(){ doAddExisting(p.peopleId, p.name); }}, 'Add')
              ]);
              box.appendChild(row);
            });
          });
        }

        function doAddExisting(peopleId, name) {
          if (!peopleId) return;
          ajax('add_walkin_existing', {
            va_config_id: state.currentConfigId,
            va_date_iso: state.currentDateIso,
            va_org_id: o.orgId,
            va_meeting_id: o.meetingId,
            va_people_id: peopleId,
            va_volunteer: state.currentVolunteer,
          }, function(err, resp){
            if (err || !resp.success) { toast('Add failed: ' + (resp ? resp.message : err), 'error'); return; }
            o.roster = resp.roster || o.roster;
            toast('Added ' + (name || 'person') + ' as ' + (resp.memberType || 'walk-in'), 'success');
            // Close the panel after a successful add to prep for the next click
            state.walkInOpen = false;
            state.walkInSearch = '';
            state.walkInCreateOpen = false;
            paint();
          });
        }

        function doCreateWalkIn() {
          var first = (document.getElementById('vaNewFirst') || {}).value || '';
          var last = (document.getElementById('vaNewLast') || {}).value || '';
          var email = (document.getElementById('vaNewEmail') || {}).value || '';
          var cell = (document.getElementById('vaNewCell') || {}).value || '';
          if (!first.trim() || !last.trim()) { toast('First and last name required', 'error'); return; }
          ajax('create_walkin', {
            va_config_id: state.currentConfigId,
            va_date_iso: state.currentDateIso,
            va_org_id: o.orgId,
            va_meeting_id: o.meetingId,
            va_first_name: first,
            va_last_name: last,
            va_email: email,
            va_cell: cell,
            va_volunteer: state.currentVolunteer,
          }, function(err, resp){
            if (err || !resp.success) { toast('Create failed: ' + (resp ? resp.message : err), 'error'); return; }
            o.roster = resp.roster || o.roster;
            toast('Created ' + first + ' ' + last + ' as ' + (resp.memberType || 'walk-in'), 'success');
            state.walkInOpen = false;
            state.walkInSearch = '';
            state.walkInCreateOpen = false;
            paint();
          });
        }

        function renderList() {
          listCard.innerHTML = '';
          var filter = (state.rosterFilter || '').toLowerCase().trim();
          var rows = o.roster.filter(function(p){
            if (!filter) return true;
            return (p.name || '').toLowerCase().indexOf(filter) >= 0;
          });
          if (!rows.length) {
            listCard.appendChild(el('div', {class:'va-empty'}, filter ? 'No matches.' : 'No people enrolled.'));
            return;
          }
          var grid = el('div', {class:'va-roster-grid'});
          rows.forEach(function(p){
            grid.appendChild(buildPersonRow(p));
          });
          listCard.appendChild(grid);
        }

        function buildPersonRow(p) {
          var visState = p.state;
          // If unmarked, show as the default state visually so user sees what Finalize will produce.
          if (visState === 'unmarked' && o.defaultState !== 'unmarked') {
            visState = o.defaultState;
          }
          var clsList = ['va-roster-row', 'va-state-' + visState];
          if (p._pending) clsList.push('va-pending');

          var tag;
          if (p.state === 'present') tag = el('span', {class:'va-tag va-p', title:'Present — confirmed attending'}, 'P');
          else if (p.state === 'absent') tag = el('span', {class:'va-tag va-a', title:'Absent — confirmed not attending'}, 'A');
          else tag = el('span', {class:'va-tag va-u', title:'Unmarked — gets the default on Finalize'}, 'U');

          var info = [];
          if (p.age !== '' && p.age != null) info.push('Age ' + p.age);
          if (p.memberType) info.push(p.memberType);

          var row = el('div', {class: clsList.join(' '), onclick: function(){ togglePerson(p); }}, [
            el('div', {class:'va-name'}, [p.name, ' ', tag]),
            el('div', {class:'va-info'}, info.join(' · ')),
          ]);
          return row;
        }

        function togglePerson(p) {
          if (o.completed) { toast('Re-open before editing', 'error'); return; }
          if (p._pending) return;
          // Cycle: present -> absent -> present (or absent -> present). Unmarked starts as default-state visually but click moves to opposite.
          var newState;
          if (p.state === 'present') newState = 'absent';
          else if (p.state === 'absent') newState = 'present';
          else {
            // Unmarked: if default is present, click marks Absent. If default is absent, click marks Present. If default is unmarked, default to Present.
            if (o.defaultState === 'present') newState = 'absent';
            else if (o.defaultState === 'absent') newState = 'present';
            else newState = 'present';
          }
          var prevState = p.state;
          p.state = newState;
          p._pending = true;
          paint();
          ajax('mark_attendance', {
            va_config_id: state.currentConfigId,
            va_date_iso: state.currentDateIso,
            va_org_id: o.orgId,
            va_meeting_id: o.meetingId,
            va_people_id: p.peopleId,
            va_new_state: newState,
            va_volunteer: state.currentVolunteer,
          }, function(err, resp){
            p._pending = false;
            if (err || !resp.success) {
              toast('Save failed', 'error');
              p.state = prevState;
            }
            paint();
          });
        }

        function finalizeOrg() {
          if (o.defaultState === 'unmarked') {
            var unmarkedCount = o.roster.filter(function(p){ return p.state === 'unmarked'; }).length;
            if (unmarkedCount > 0) {
              if (!confirm(unmarkedCount + ' people are still Unmarked. With Default Unmarked, they will stay unmarked (no row written). Continue?')) return;
            }
          }
          ajax('finalize_org', {
            va_config_id: state.currentConfigId,
            va_date_iso: state.currentDateIso,
            va_org_id: o.orgId,
            va_meeting_id: o.meetingId,
            va_volunteer: state.currentVolunteer,
          }, function(err, resp){
            if (err || !resp.success) { toast('Finalize failed: ' + (resp ? resp.message : err), 'error'); return; }
            toast('Finalized · ' + resp.applied + ' applied', 'success');
            backToDashboard();
          });
        }

        function reopenOrg() {
          if (!confirm('Re-open this involvement? You can mark more people, then finalize again.')) return;
          ajax('reopen_org', {
            va_config_id: state.currentConfigId,
            va_date_iso: state.currentDateIso,
            va_org_id: o.orgId,
            va_volunteer: state.currentVolunteer,
          }, function(err, resp){
            if (err || !resp.success) { toast('Re-open failed', 'error'); return; }
            o.completed = false;
            paint();
          });
        }

        paint();
      }

      function backToDashboard() {
        // Soft-release claim if we never finalized
        if (state.currentOrg && !state.currentOrg.completed) {
          ajax('release_org', {
            va_config_id: state.currentConfigId,
            va_date_iso: state.currentDateIso,
            va_org_id: state.currentOrg.orgId,
          }, function(){});
        }
        state.currentOrg = null;
        state.view = 'dashboard';
        render();
      }

      // ----------------- bootstrap -----------------
      loadConfigs(render);
    })();
    </script>
    """

    body = """
    <div class="va-root">
      <div id="vaRoot"></div>
    </div>
    """

    model.Form = css + body + js
