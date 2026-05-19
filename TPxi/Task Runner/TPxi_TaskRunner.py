#roles=Edit
#----------------------------------------------------------------------
# TPxi_TaskRunner.py
#
# Individual-first task triage view. The pastor / IC opens this and
# sees what to do RIGHT NOW, not a workload chart. Tasks grouped by
# urgency (Overdue / Today / This Week / Later / Undated) with quick
# actions (complete, snooze, open person) per row.
#
# Phase 1a (this file): identity + load + minimal grouped render.
# Phase 1b: prettier render, mobile-friendly CSS, keyword chips.
# Phase 1c: complete / snooze / open-person AJAX actions.
# Phase 1d: Daily Focus card + capacity awareness.
#
# Storage Keys:
#   TaskRunner_UserPrefs_<peopleId>  - per-user UI prefs (JSON)
#   TaskRunner_OrgSettings           - install-wide config (JSON)
#
# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub: https://github.com/bswaby/Touchpoint
#----------------------------------------------------------------------

import json
import datetime
import re

model.Header = 'Task Runner'

# --- Version / Auto-update -------------------------------------------
APP_VERSION = '0.1.0'  # Phase 1a — pre-release skeleton
DC_SCRIPT_ID = 'TPxi_TaskRunner'
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'


def get_script_name():
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


# =====================================================================
# UNICODE / JSON SAFETY (IronPython)
# =====================================================================
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
    if s is None:
        return ''
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
        try:
            return _to_ascii(s.decode('cp1252'))
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
        return dict((sanitize_for_json(k), sanitize_for_json(v)) for k, v in obj.items())
    if isinstance(obj, (list, tuple)):
        return [sanitize_for_json(item) for item in obj]
    return safe_str(obj)


def jprint(obj):
    import sys
    sys.stdout.write(json.dumps(sanitize_for_json(obj)))


def get_data(name, default=''):
    try:
        if hasattr(Data, name):
            v = getattr(Data, name)
            return v if v is not None else default
    except:
        pass
    return default


def safe_int(v, default=0):
    try:
        return int(v)
    except:
        return default


def _iso_date(val):
    """Return YYYY-MM-DD for a SQL/.NET DateTime, or '' if None.

    IronPython's strftime() doesn't reliably handle .NET DateTime objects,
    and str() formats them as M/D/YYYY HH:MM AM/PM (which broke the bucket
    string compares). We try multiple paths so any value shape works.
    """
    if val is None:
        return ''
    # .NET DateTime.ToString('yyyy-MM-dd') is the most reliable path
    try:
        s = val.ToString('yyyy-MM-dd')
        if s:
            return safe_str(s)
    except:
        pass
    # CPython / Python datetime
    try:
        return val.strftime('%Y-%m-%d')
    except:
        pass
    # Last resort: pick the date apart by attributes
    try:
        return '{0:04d}-{1:02d}-{2:02d}'.format(int(val.Year), int(val.Month), int(val.Day))
    except:
        pass
    try:
        return '{0:04d}-{1:02d}-{2:02d}'.format(val.year, val.month, val.day)
    except:
        pass
    return ''


# =====================================================================
# CONFIG
# =====================================================================
PREFS_KEY_PREFIX = 'TaskRunner_UserPrefs_'
ORG_SETTINGS_KEY = 'TaskRunner_OrgSettings'
SNOOZED_KEY_PREFIX = 'TaskRunner_Snoozed_'  # local triage state per user
# Contact methods are stored under the same key ProspectBuilder uses, so
# admins configure them once and both tools see the same list.
SHARED_CONTACT_KEY = 'ProgramPulse_Settings'

# Suggested starter methods. Admins can click "Use these as starters" in
# Settings to populate the table, but they're never auto-applied -- a new
# church install sees an empty list and an explanation, so the deployer
# doesn't have to accept our naming conventions.
SUGGESTED_CONTACT_METHODS = [
    {'code': 'P', 'label': 'Phone', 'keyword': '', 'keywordId': None},
    {'code': 'E', 'label': 'Email', 'keyword': '', 'keywordId': None},
    {'code': 'T', 'label': 'Text',  'keyword': '', 'keywordId': None},
    {'code': 'V', 'label': 'Visit', 'keyword': '', 'keywordId': None},
    {'code': 'M', 'label': 'Mail',  'keyword': '', 'keywordId': None},
]

# How far back to count contact attempts. 26 weeks matches the historical
# ProspectBuilder default. Stored in the shared settings blob so ProspectBuilder
# can adopt the same value later.
DEFAULT_CONTACT_LOOKBACK_DAYS = 182
MIN_CONTACT_LOOKBACK_DAYS = 7
MAX_CONTACT_LOOKBACK_DAYS = 1825   # ~5 years

# Closed-task status filter -- Completed (40) and Declined (70) are closed.
# Everything else (including NULL) is "still on the assignee's plate".
# Inverting the whitelist this way is more forgiving for installs that have
# added custom statuses or have rows with NULL StatusId.
CLOSED_STATUS_IDS = '40,70'

# Page render cap so a 1000-task user doesn't blow the response. Tasks
# beyond this fall under "Later" anyway; user can drill in via filters.
MAX_TASKS_PER_LOAD = 500


def load_user_prefs(people_id):
    try:
        raw = model.TextContent(PREFS_KEY_PREFIX + str(people_id))
        if raw:
            return json.loads(raw)
    except:
        pass
    return {}


def save_user_prefs(people_id, prefs):
    try:
        model.WriteContentText(PREFS_KEY_PREFIX + str(people_id), json.dumps(prefs), '')
    except:
        pass


# --- Team / drill-in policy values (validated server-side) ---
DRILLIN_SCOPE_VALUES = ('off', 'subgroup', 'all_staff')
DRILLIN_ACTIONS_VALUES = ('none', 'reassign', 'reassign_complete')


def load_org_settings():
    """Loads TaskRunner_OrgSettings and migrates legacy fields on read so
    existing installs keep working as we evolve the schema.
    """
    try:
        raw = model.TextContent(ORG_SETTINGS_KEY)
        if raw:
            settings = json.loads(raw)
        else:
            settings = {}
    except:
        settings = {}

    # Legacy: use_subgroups was a single bool. Map it onto the new pair:
    #   true  -> scope=subgroup, group display by department on
    #   false -> scope=all_staff (everyone in the staff org, flat)
    if 'drillin_scope' not in settings:
        if 'use_subgroups' in settings:
            settings['drillin_scope'] = 'subgroup' if settings.get('use_subgroups') else 'all_staff'
            settings.setdefault('team_display_grouped', bool(settings.get('use_subgroups')))
        else:
            settings['drillin_scope'] = 'off'
    if 'drillin_actions' not in settings:
        settings['drillin_actions'] = 'reassign'  # conservative default
    if 'team_display_grouped' not in settings:
        settings['team_display_grouped'] = bool(settings.get('use_subgroups', False))

    # Sanity-clamp values
    if settings.get('drillin_scope') not in DRILLIN_SCOPE_VALUES:
        settings['drillin_scope'] = 'off'
    if settings.get('drillin_actions') not in DRILLIN_ACTIONS_VALUES:
        settings['drillin_actions'] = 'reassign'
    return settings


def save_org_settings(settings):
    try:
        model.WriteContentText(ORG_SETTINGS_KEY, json.dumps(settings), '')
        return True
    except:
        return False


def get_team_people_ids(my_pid, org_settings=None):
    """Return the list of PeopleIds the calling user counts as 'team'.

    Rules:
      - off: empty list (Team View hidden entirely)
      - subgroup: people in the staff org who share at least one subgroup
                  with the caller. If the caller has no subgroups, return
                  empty (avoids accidentally exposing the whole staff).
      - all_staff: every active member of the staff org (excluding self).
    """
    if not my_pid:
        return []
    if org_settings is None:
        org_settings = load_org_settings()
    scope = org_settings.get('drillin_scope', 'off')
    staff_org_id = safe_int(org_settings.get('staff_org_id', 0))
    if scope == 'off' or not staff_org_id:
        return []

    if scope == 'all_staff':
        sql = """
            SELECT DISTINCT om.PeopleId
            FROM OrganizationMembers om
            WHERE om.OrganizationId = {oid}
              AND om.PeopleId <> {me}
        """.format(oid=int(staff_org_id), me=int(my_pid))
    else:
        # subgroup: intersect my subgroups in the staff org with everyone else's
        sql = """
            SELECT DISTINCT om.PeopleId
            FROM OrganizationMembers om
            JOIN OrgMemMemTags ommt ON ommt.PeopleId = om.PeopleId AND ommt.OrgId = om.OrganizationId
            WHERE om.OrganizationId = {oid}
              AND om.PeopleId <> {me}
              AND ommt.MemberTagId IN (
                  SELECT mine.MemberTagId
                  FROM OrgMemMemTags mine
                  WHERE mine.OrgId = {oid} AND mine.PeopleId = {me}
              )
        """.format(oid=int(staff_org_id), me=int(my_pid))

    pids = []
    try:
        for r in q.QuerySql(sql):
            pids.append(int(r.PeopleId))
    except:
        pass
    return pids


def is_on_my_team(my_pid, other_pid, org_settings=None):
    """Cheap drill-in auth check. Used before exposing another user's tasks."""
    if not other_pid or other_pid == my_pid:
        return False
    return int(other_pid) in get_team_people_ids(my_pid, org_settings)


def can_reassign_for_others(org_settings=None):
    if org_settings is None:
        org_settings = load_org_settings()
    return org_settings.get('drillin_actions') in ('reassign', 'reassign_complete')


def can_complete_for_others(org_settings=None):
    if org_settings is None:
        org_settings = load_org_settings()
    return org_settings.get('drillin_actions') == 'reassign_complete'


def load_my_team(my_pid):
    """Return team members + per-member task stats. One main query for the
    person list, one batched stats query keyed by PeopleId.
    """
    settings = load_org_settings()
    team_pids = get_team_people_ids(my_pid, settings)
    if not team_pids:
        return {
            'enabled': settings.get('drillin_scope', 'off') != 'off',
            'members': [],
            'permissions': {
                'canReassign': can_reassign_for_others(settings),
                'canComplete': can_complete_for_others(settings),
            }
        }

    pid_csv = ','.join(str(p) for p in team_pids)
    staff_org_id = safe_int(settings.get('staff_org_id', 0))

    # Per-person identity + subgroups (for the department display)
    people_sql = """
        SELECT p.PeopleId, ISNULL(p.Name2, '') AS Name2,
               ISNULL(p.EmailAddress, '') AS EmailAddress,
               STUFF((
                   SELECT '|' + ISNULL(mt.Name, '')
                   FROM OrgMemMemTags ommt
                   JOIN MemberTags mt ON mt.Id = ommt.MemberTagId AND mt.OrgId = ommt.OrgId
                   WHERE ommt.PeopleId = p.PeopleId AND ommt.OrgId = {oid}
                   FOR XML PATH('')
               ), 1, 1, '') AS Subgroups
        FROM People p
        WHERE p.PeopleId IN ({pids})
        ORDER BY p.LastName, p.FirstName
    """.format(pids=pid_csv, oid=staff_org_id)

    # Per-person task stats. One query, conditional aggregates.
    stats_sql = """
        SELECT
            COALESCE(tn.AssigneeId, tn.OwnerId) AS PeopleId,
            SUM(CASE WHEN tn.CompletedDate IS NULL AND tn.IsArchived = 0 AND tn.IsNote = 0
                     AND (tn.StatusId IS NULL OR tn.StatusId NOT IN ({closed}))
                THEN 1 ELSE 0 END) AS OpenCount,
            SUM(CASE WHEN tn.CompletedDate IS NULL AND tn.IsArchived = 0 AND tn.IsNote = 0
                     AND (tn.StatusId IS NULL OR tn.StatusId NOT IN ({closed}))
                     AND tn.DueDate < CAST(GETDATE() AS DATE)
                THEN 1 ELSE 0 END) AS OverdueCount,
            SUM(CASE WHEN tn.CompletedDate IS NOT NULL
                     AND CAST(tn.CompletedDate AS DATE) = CAST(GETDATE() AS DATE)
                THEN 1 ELSE 0 END) AS CompletedToday,
            MAX(tn.CompletedDate) AS LastCompleted
        FROM TaskNote tn
        WHERE (
            tn.AssigneeId IN ({pids})
            OR (tn.OwnerId IN ({pids}) AND tn.AssigneeId IS NULL)
        )
          AND tn.IsNote = 0
        GROUP BY COALESCE(tn.AssigneeId, tn.OwnerId)
    """.format(pids=pid_csv, closed=CLOSED_STATUS_IDS)

    stats_by_pid = {}
    try:
        for r in q.QuerySql(stats_sql):
            stats_by_pid[int(r.PeopleId)] = {
                'open': int(r.OpenCount or 0),
                'overdue': int(r.OverdueCount or 0),
                'completedToday': int(r.CompletedToday or 0),
                'lastCompleted': _iso_date(r.LastCompleted),
            }
    except:
        pass

    members = []
    try:
        for r in q.QuerySql(people_sql):
            pid = int(r.PeopleId)
            subs = []
            if r.Subgroups:
                subs = [safe_str(s) for s in safe_str(r.Subgroups).split('|') if s]
            stats = stats_by_pid.get(pid, {'open': 0, 'overdue': 0, 'completedToday': 0, 'lastCompleted': ''})
            members.append({
                'peopleId': pid,
                'name': safe_str(r.Name2),
                'email': safe_str(r.EmailAddress),
                'subgroups': subs,
                'open': stats['open'],
                'overdue': stats['overdue'],
                'completedToday': stats['completedToday'],
                'lastCompleted': stats['lastCompleted'],
            })
    except:
        pass

    return {
        'enabled': True,
        'members': members,
        'permissions': {
            'canReassign': can_reassign_for_others(settings),
            'canComplete': can_complete_for_others(settings),
        },
        'displayGrouped': bool(settings.get('team_display_grouped', False)),
    }


def get_staff_org_subgroup_for_person(people_id, staff_org_id):
    """Return the subgroup name(s) on the staff involvement for the given
    person, if any. Used by Phase 2 Team View to group staff by department.
    Returns a list of subgroup names (since a person can be in multiple)."""
    if not people_id or not staff_org_id:
        return []
    sql = """
        SELECT ISNULL(mt.Name, '') AS Name
        FROM OrgMemMemTags ommt
        JOIN MemberTags mt ON mt.Id = ommt.MemberTagId AND mt.OrgId = ommt.OrgId
        WHERE ommt.PeopleId = {pid} AND ommt.OrgId = {oid}
    """.format(pid=int(people_id), oid=int(staff_org_id))
    out = []
    try:
        for r in q.QuerySql(sql):
            n = safe_str(r.Name)
            if n:
                out.append(n)
    except:
        pass
    return out


def load_contact_methods():
    """Load the shared contact-methods list (same storage ProspectBuilder
    uses). Returns [] when nothing is configured -- the UI handles the
    empty state explicitly so we don't impose code/label conventions on
    every church that installs this tool.
    """
    try:
        raw = model.TextContent(SHARED_CONTACT_KEY)
        if raw:
            settings = json.loads(raw)
            methods = settings.get('contact_methods', [])
            if methods:
                return methods
    except:
        pass
    return []


def save_contact_methods(methods):
    """Persist the contact-methods list back into the shared key. Other keys
    in the same blob (Program Pulse / ProspectBuilder settings) are preserved.
    """
    try:
        raw = model.TextContent(SHARED_CONTACT_KEY)
        settings = json.loads(raw) if raw else {}
    except:
        settings = {}
    settings['contact_methods'] = methods
    try:
        model.WriteContentText(SHARED_CONTACT_KEY, json.dumps(settings), '')
        return True
    except:
        return False


def find_contact_method_by_code(code):
    for m in load_contact_methods():
        if str(m.get('code', '')).upper() == str(code).upper():
            return m
    return None


def load_contact_lookback_days():
    """How many days back the contact-effort badges should count.
    Stored in the same shared settings blob as the methods themselves."""
    try:
        raw = model.TextContent(SHARED_CONTACT_KEY)
        if raw:
            settings = json.loads(raw)
            v = settings.get('contact_lookback_days')
            if v is not None:
                v = int(v)
                if v < MIN_CONTACT_LOOKBACK_DAYS: v = MIN_CONTACT_LOOKBACK_DAYS
                if v > MAX_CONTACT_LOOKBACK_DAYS: v = MAX_CONTACT_LOOKBACK_DAYS
                return v
    except:
        pass
    return DEFAULT_CONTACT_LOOKBACK_DAYS


def save_contact_lookback_days(days):
    try:
        raw = model.TextContent(SHARED_CONTACT_KEY)
        settings = json.loads(raw) if raw else {}
    except:
        settings = {}
    try:
        days_i = int(days)
    except:
        days_i = DEFAULT_CONTACT_LOOKBACK_DAYS
    if days_i < MIN_CONTACT_LOOKBACK_DAYS: days_i = MIN_CONTACT_LOOKBACK_DAYS
    if days_i > MAX_CONTACT_LOOKBACK_DAYS: days_i = MAX_CONTACT_LOOKBACK_DAYS
    settings['contact_lookback_days'] = days_i
    try:
        model.WriteContentText(SHARED_CONTACT_KEY, json.dumps(settings), '')
        return True
    except:
        return False


def load_snoozed(people_id):
    """{ '<taskId>': 'YYYY-MM-DD', ... } - tasks hidden from this user's
    view until the given date. SQL is read-only on most installs, so snooze
    lives in content storage as personal triage state rather than mutating
    TaskNote.DueDate. The real DueDate stays accurate for other tools.
    """
    try:
        raw = model.TextContent(SNOOZED_KEY_PREFIX + str(people_id))
        if raw:
            d = json.loads(raw)
            if isinstance(d, dict):
                return d
    except:
        pass
    return {}


def save_snoozed(people_id, snoozed):
    try:
        model.WriteContentText(SNOOZED_KEY_PREFIX + str(people_id),
                               json.dumps(snoozed), '')
        return True
    except:
        return False


def prune_snoozed(snoozed):
    """Drop entries whose snooze date has already passed. Keeps the storage
    small over time without needing a separate cleanup job.
    """
    today = datetime.datetime.now().strftime('%Y-%m-%d')
    return dict((k, v) for k, v in snoozed.items() if v and v >= today)


# =====================================================================
# SQL — load my tasks
# =====================================================================
def get_current_people_id():
    """The TouchPoint logged-in user's PeopleId. Falls back to 0 if unset.

    TouchPoint exposes this as either a property or a method depending on
    the install / version, so try both forms before giving up.
    """
    try:
        pid = model.UserPeopleId
        if callable(pid):
            pid = pid()
        if pid:
            return int(pid)
    except:
        pass
    try:
        pid = model.UserPeopleId()
        if pid:
            return int(pid)
    except:
        pass
    return 0


def get_current_user_info(people_id):
    """Return display info about the current user for the header."""
    if not people_id:
        return {'peopleId': 0, 'name': 'Unknown', 'username': ''}
    sql = """
        SELECT p.PeopleId, ISNULL(p.Name2, '') AS Name2,
               ISNULL(u.Username, '') AS Username
        FROM People p
        LEFT JOIN Users u ON u.PeopleId = p.PeopleId
        WHERE p.PeopleId = {0}
    """.format(int(people_id))
    try:
        for r in q.QuerySql(sql):
            return {
                'peopleId': r.PeopleId,
                'name': safe_str(r.Name2),
                'username': safe_str(r.Username)
            }
    except:
        pass
    return {'peopleId': people_id, 'name': '', 'username': ''}


def load_my_tasks(assignee_id, view_as_id=None):
    """Pull open TaskNotes for the given assignee. view_as_id allows a
    future leader-mode override; phase 1a always uses assignee_id.

    Returns a list of dicts shaped for the UI. No grouping done here —
    that happens client-side based on the date, so a user changing the
    week boundary doesn't need a server round trip.
    """
    aid = view_as_id or assignee_id
    if not aid:
        return []

    # "On my plate" includes:
    #   - Tasks delegated to me (AssigneeId = me)
    #   - Tasks I own that haven't been delegated to someone else
    # Tasks I created and handed off (Owner=me, Assignee=someone else) are
    # NOT on my plate -- they're on the assignee's plate.
    sql = """
        SELECT TOP {limit}
            tn.TaskNoteId,
            ISNULL(tn.Instructions, '') AS Instructions,
            ISNULL(tn.Notes, '') AS Notes,
            tn.DueDate,
            tn.CreatedDate,
            tn.StatusId,
            ts.Description AS StatusDescription,
            tn.AboutPersonId,
            ISNULL(about.Name2, '') AS AboutName,
            tn.OwnerId,
            tn.AssigneeId,
            ISNULL(owner.Name2, '') AS OwnerName,
            tn.OrgId,
            ISNULL(o.OrganizationName, '') AS OrgName,
            STUFF((
                SELECT '|' + ISNULL(k.Description, '')
                FROM TaskNoteKeyword tnk
                JOIN dbo.Keyword k ON k.KeywordId = tnk.KeywordId
                WHERE tnk.TaskNoteId = tn.TaskNoteId
                FOR XML PATH('')
            ), 1, 1, '') AS KeywordsPipe
        FROM TaskNote tn
        LEFT JOIN lookup.TaskStatus ts ON ts.Id = tn.StatusId
        LEFT JOIN People about ON about.PeopleId = tn.AboutPersonId
        LEFT JOIN People owner ON owner.PeopleId = tn.OwnerId
        LEFT JOIN Organizations o ON o.OrganizationId = tn.OrgId
        WHERE (
                tn.AssigneeId = {aid}
                OR (tn.OwnerId = {aid} AND tn.AssigneeId IS NULL)
              )
          AND tn.IsArchived = 0
          AND tn.IsNote = 0
          AND tn.CompletedDate IS NULL
          AND (tn.StatusId IS NULL OR tn.StatusId NOT IN ({closed_status}))
        ORDER BY
            CASE WHEN tn.DueDate IS NULL THEN 1 ELSE 0 END,
            tn.DueDate ASC,
            tn.CreatedDate DESC
    """.format(
        limit=MAX_TASKS_PER_LOAD,
        aid=int(aid),
        closed_status=CLOSED_STATUS_IDS
    )

    tasks = []
    try:
        for r in q.QuerySql(sql):
            kw = []
            if r.KeywordsPipe:
                kw = [safe_str(x) for x in safe_str(r.KeywordsPipe).split('|') if x]
            due_iso = _iso_date(r.DueDate)
            created_iso = _iso_date(r.CreatedDate)
            tasks.append({
                'id': int(r.TaskNoteId),
                'instructions': safe_str(r.Instructions),
                'notes': safe_str(r.Notes),
                'dueDate': due_iso,
                'createdDate': created_iso,
                'statusId': int(r.StatusId or 0),
                'status': safe_str(r.StatusDescription),
                'aboutPersonId': int(r.AboutPersonId) if r.AboutPersonId else 0,
                'aboutName': safe_str(r.AboutName),
                'ownerId': int(r.OwnerId) if r.OwnerId else 0,
                'ownerName': safe_str(r.OwnerName),
                'orgId': int(r.OrgId) if r.OrgId else 0,
                'orgName': safe_str(r.OrgName),
                'keywords': kw,
                'extras': [],
            })
    except Exception as e:
        # Surfaced via the error response shape so the UI can show it
        return {'error': safe_str(e), 'tasks': []}

    # Split into visible / hidden based on local snooze state. SQL is read-only
    # on this install, so snooze lives in content storage rather than mutating
    # TaskNote.DueDate. Prune expired entries opportunistically.
    snoozed_before = load_snoozed(aid)
    snoozed = prune_snoozed(snoozed_before)
    if snoozed != snoozed_before:
        save_snoozed(aid, snoozed)

    today = datetime.datetime.now().strftime('%Y-%m-%d')
    visible = []
    hidden = []
    for t in tasks:
        snz_until = snoozed.get(str(t['id']))
        if snz_until and snz_until > today:
            t['hiddenUntil'] = snz_until
            hidden.append(t)
        else:
            visible.append(t)

    # Batch fetch extras for both lists in one query (no N+1).
    all_ids = [t['id'] for t in visible] + [t['id'] for t in hidden]
    if all_ids:
        extras_by_task = load_task_extras(all_ids)
        for t in visible:
            t['extras'] = extras_by_task.get(t['id'], [])
        for t in hidden:
            t['extras'] = extras_by_task.get(t['id'], [])

    # Contact-effort badges per about-person (last 26 weeks). One batched
    # query keyed by AboutPersonId, then attached back to each task.
    about_pids = []
    seen_pids = set()
    for t in visible + hidden:
        ap = t.get('aboutPersonId')
        if ap and ap not in seen_pids:
            seen_pids.add(ap)
            about_pids.append(ap)
    if about_pids:
        contact_map = load_contact_efforts(about_pids)
        for t in visible + hidden:
            ap = t.get('aboutPersonId')
            if ap:
                t['contactEfforts'] = contact_map.get(str(ap), None)

    return {'tasks': visible, 'hiddenTasks': hidden,
            'snoozedCount': len(hidden)}


def load_contact_efforts(people_ids):
    """Tally TaskNote-based contact efforts for the given about-people, by
    configured contact-method code. Matches the ProspectBuilder pattern:

        { '<peopleId>': {
            'methods': {'P': 3, 'E': 1, 'V': 0, ...},
            'total':   8,
            'lastDate': '2026-04-15',
            'other':   2     # TaskNotes that didn't match any configured keyword
        } }

    Window: configurable via load_contact_lookback_days (defaults to 182d).
    Stored in the shared settings blob, so once ProspectBuilder reads it
    the same value drives badges in both tools.
    """
    if not people_ids:
        return {}
    methods = load_contact_methods()
    lookback_days = load_contact_lookback_days()
    keyword_id_to_code = {}
    keyword_ids = []
    for m in methods:
        kid = m.get('keywordId')
        if kid:
            try:
                kid_i = int(kid)
                keyword_id_to_code[kid_i] = safe_str(m.get('code', '?'))
                keyword_ids.append(kid_i)
            except:
                pass

    pid_list = ','.join(str(int(p)) for p in people_ids if p)
    if not pid_list:
        return {}

    # Step 1: total TaskNote count and last-contact date per person.
    total_sql = """
        SELECT tn.AboutPersonId AS PeopleId,
               COUNT(*) AS TotalCount,
               MAX(tn.CreatedDate) AS LastContactDate
        FROM TaskNote tn WITH (NOLOCK)
        WHERE tn.AboutPersonId IN ({pids})
          AND tn.CreatedDate >= DATEADD(day, -{days}, GETDATE())
        GROUP BY tn.AboutPersonId
    """.format(pids=pid_list, days=lookback_days)

    contact_map = {}
    try:
        for r in q.QuerySql(total_sql):
            contact_map[str(int(r.PeopleId))] = {
                'methods': {},
                'total': int(r.TotalCount or 0),
                'lastDate': _iso_date(r.LastContactDate),
                'other': 0,
            }
    except:
        pass

    # Step 2: per-code counts (only for the people we found in step 1).
    if keyword_ids and contact_map:
        contacted_pids = ','.join(contact_map.keys())
        kid_list = ','.join(str(k) for k in keyword_ids)
        method_sql = """
            SELECT tn.AboutPersonId AS PeopleId, tnk.KeywordId, COUNT(*) AS Cnt
            FROM TaskNote tn WITH (NOLOCK)
            JOIN TaskNoteKeyword tnk ON tnk.TaskNoteId = tn.TaskNoteId
            WHERE tn.AboutPersonId IN ({pids})
              AND tnk.KeywordId IN ({kids})
              AND tn.CreatedDate >= DATEADD(day, -{days}, GETDATE())
            GROUP BY tn.AboutPersonId, tnk.KeywordId
        """.format(pids=contacted_pids, kids=kid_list, days=lookback_days)
        try:
            for row in q.QuerySql(method_sql):
                pid_key = str(int(row.PeopleId))
                if pid_key in contact_map:
                    code = keyword_id_to_code.get(int(row.KeywordId), '')
                    if code:
                        contact_map[pid_key]['methods'][code] = int(row.Cnt or 0)
        except:
            pass

    # Step 3: derive "Other" -- TaskNotes that didn't match any configured keyword.
    for pid_key in contact_map:
        matched = sum(contact_map[pid_key]['methods'].values())
        other = contact_map[pid_key]['total'] - matched
        contact_map[pid_key]['other'] = max(other, 0)

    return contact_map


def load_task_extras(task_ids):
    """Pull keyword-extra-value fields for a batch of TaskNotes.

    Returns {taskId: [{keyword, name, value}]} so each task in the list can
    show its filled-in extras inline (e.g. "PC Stay Info: Facility=Summitt
    Medical Center, Admission=2026-05-05"). Skips extras whose response is
    NULL/empty.

    DataType conventions seen in the wild:
      3 = text / number (Response holds the literal value)
      5 = dropdown (Response holds the OptionId, look up via KeywordExtraValueOption)
      8 = date (Response holds the date string)
    The logic below handles all three uniformly: prefer the resolved option
    label when present, otherwise use Response as-is.
    """
    if not task_ids:
        return {}
    ids_csv = ','.join(str(int(t)) for t in task_ids)
    sql = """
        SELECT
            tnev.TaskNoteId,
            ISNULL(k.Description, '') AS Keyword,
            ISNULL(kev.Name, '') AS FieldName,
            kev.SortOrder,
            kev.DataType,
            tnev.Response AS RawResponse,
            kevo.Name AS OptionLabel
        FROM TaskNoteExtraValue tnev
        JOIN KeywordExtraValue kev ON kev.KeywordExtraValueId = tnev.KeywordExtraValueId
        LEFT JOIN dbo.Keyword k ON k.KeywordId = kev.KeywordId
        LEFT JOIN KeywordExtraValueOption kevo
               ON kevo.KeywordExtraValueId = kev.KeywordExtraValueId
              AND kevo.KeywordExtraValueOptionId = TRY_CAST(tnev.Response AS INT)
        WHERE tnev.TaskNoteId IN ({ids})
        ORDER BY tnev.TaskNoteId, k.Description, kev.SortOrder, kev.Name
    """.format(ids=ids_csv)

    by_task = {}
    try:
        for r in q.QuerySql(sql):
            raw = safe_str(r.RawResponse)
            opt = safe_str(r.OptionLabel)
            # Pick the human-readable value. Option label wins; otherwise the
            # raw response. Skip if neither has anything meaningful.
            value = opt if opt else raw
            if not value or not value.strip():
                continue
            tid = int(r.TaskNoteId)
            by_task.setdefault(tid, []).append({
                'keyword': safe_str(r.Keyword),
                'name': safe_str(r.FieldName),
                'value': value,
            })
    except:
        pass
    return by_task


def get_my_stats(people_id):
    """Capacity awareness: how much have I been completing recently?

    Returns counts the UI can use to compute averages and zero-projection
    without doing math server-side.
    """
    if not people_id:
        return {'completedToday': 0, 'completed7d': 0, 'completed30d': 0,
                'oldestOpenDays': 0}
    sql = """
        SELECT
            SUM(CASE WHEN tn.CompletedDate IS NOT NULL
                     AND CAST(tn.CompletedDate AS DATE) = CAST(GETDATE() AS DATE)
                THEN 1 ELSE 0 END) AS CompletedToday,
            SUM(CASE WHEN tn.CompletedDate IS NOT NULL
                     AND tn.CompletedDate >= DATEADD(day, -7, GETDATE())
                THEN 1 ELSE 0 END) AS Completed7d,
            SUM(CASE WHEN tn.CompletedDate IS NOT NULL
                     AND tn.CompletedDate >= DATEADD(day, -30, GETDATE())
                THEN 1 ELSE 0 END) AS Completed30d
        FROM TaskNote tn
        WHERE (
                tn.AssigneeId = {pid}
                OR tn.OwnerId = {pid}
                OR tn.CompletedBy = {pid}
              )
          AND tn.IsNote = 0
          AND tn.CompletedDate IS NOT NULL
    """.format(pid=int(people_id))
    out = {'completedToday': 0, 'completed7d': 0, 'completed30d': 0}
    try:
        r = q.QuerySqlTop1(sql)
        if r:
            out['completedToday'] = int(r.CompletedToday or 0)
            out['completed7d'] = int(r.Completed7d or 0)
            out['completed30d'] = int(r.Completed30d or 0)
    except:
        pass
    return out


# Keyword names that hint a task should rank higher in the focus picker.
# Lowercased. Per-install config can override later (Phase 2 settings UI).
URGENT_KEYWORDS = ('hospital', 'death', 'emergency', 'urgent', 'baptism',
                   'salvation', 'bereavement')


def score_task(task, today_iso):
    """Priority score for ranking the Daily Focus picks. Higher = more
    important. Pure function -- caller passes everything needed.

    Components:
      - Days overdue (10 pts per day, capped at 200)
      - Due today bonus (15 pts)
      - "Urgent" keyword hit (40 pts each, capped at 80)
      - No-due-date penalty (-5) so tasks with no date don't dominate
    """
    score = 0
    due = task.get('dueDate', '')
    if due:
        # Lex compare on YYYY-MM-DD works for ordering. Day count needs math.
        try:
            d_due = datetime.datetime.strptime(due, '%Y-%m-%d').date()
            d_today = datetime.datetime.strptime(today_iso, '%Y-%m-%d').date()
            delta = (d_today - d_due).days   # positive = overdue
            if delta > 0:
                score += min(delta * 10, 200)
            elif delta == 0:
                score += 15
        except:
            pass
    else:
        score -= 5
    # Keyword boost
    kw_lower = [str(k).lower() for k in task.get('keywords', [])]
    boost = 0
    for u in URGENT_KEYWORDS:
        for k in kw_lower:
            if u in k:
                boost += 40
                break
    score += min(boost, 80)
    return score


def _name_search_where_tasks(search_term, alias='p'):
    """Build a tokenized name-search WHERE expression. Shared pattern with
    TPxi_DayOfRegistration. Handles two input styles:

      "Last, First" / "Swa, B"  -> starts-with on LastName + FirstName/NickName
      "ben swa" / "Ben Swa"     -> each token must match somewhere in
                                   FirstName/LastName/NickName/Name2

    Returns a SQL fragment with no leading WHERE/AND. Returns '1=0' for
    empty input so the surrounding query stays valid.
    """
    term = (search_term or '').strip()
    if not term:
        return '1=0'
    a = alias

    if ',' in term:
        parts = term.split(',', 1)
        last_tok = parts[0].strip().replace("'", "''")
        first_tok = parts[1].strip().replace("'", "''")
        clauses = []
        if last_tok:
            clauses.append("{0}.LastName LIKE '{1}%'".format(a, last_tok))
        if first_tok:
            clauses.append("({0}.FirstName LIKE '{1}%' OR {0}.NickName LIKE '{1}%')".format(a, first_tok))
        if not clauses:
            return '1=0'
        return ' AND '.join(clauses)

    tokens = [t for t in term.split() if t]
    if not tokens:
        return '1=0'
    token_clauses = []
    for raw in tokens:
        tok = raw.replace("'", "''")
        token_clauses.append(
            "({0}.FirstName LIKE '%{1}%' "
            "OR {0}.LastName LIKE '%{1}%' "
            "OR {0}.NickName LIKE '%{1}%' "
            "OR {0}.Name2 LIKE '%{1}%')".format(a, tok)
        )
    return ' AND '.join(token_clauses)


def task_belongs_to_user(task_id, people_id, allow_team_pids=None):
    """Authorization gate: confirm the task is actually on this user's plate
    (or on a teammate's plate when allow_team_pids is supplied) before letting
    them act on it. Open status only -- can't re-complete a completed task.

    allow_team_pids: optional list of PeopleIds the caller is allowed to act
    on behalf of (e.g., direct reports when drill-in actions are enabled).
    """
    if not task_id or not people_id:
        return False
    eligible = [int(people_id)]
    if allow_team_pids:
        for p in allow_team_pids:
            try:
                eligible.append(int(p))
            except:
                pass
    pids_csv = ','.join(str(p) for p in eligible)
    sql = """
        SELECT TOP 1 1 AS Hit
        FROM TaskNote tn
        WHERE tn.TaskNoteId = {tid}
          AND tn.IsArchived = 0
          AND tn.IsNote = 0
          AND tn.CompletedDate IS NULL
          AND (tn.StatusId IS NULL OR tn.StatusId NOT IN ({closed}))
          AND (
                tn.AssigneeId IN ({pids})
                OR (tn.OwnerId IN ({pids}) AND tn.AssigneeId IS NULL)
              )
    """.format(tid=int(task_id), pids=pids_csv, closed=CLOSED_STATUS_IDS)
    try:
        r = q.QuerySqlTop1(sql)
        return bool(r and r.Hit == 1)
    except:
        return False


def diagnose_empty_tasks(people_id):
    """Run when load_my_tasks returns 0 tasks. Returns counts so we can see
    whether the user really has none, or whether the filter is wrong."""
    if not people_id:
        return {'reason': 'no-people-id'}
    sql = """
        SELECT
            SUM(CASE WHEN tn.AssigneeId = {pid} THEN 1 ELSE 0 END) AS AsAssignee_All,
            SUM(CASE WHEN tn.AssigneeId = {pid} AND tn.CompletedDate IS NULL
                     AND tn.IsArchived = 0 AND tn.IsNote = 0 THEN 1 ELSE 0 END) AS AsAssignee_Open,
            SUM(CASE WHEN tn.OwnerId = {pid} AND tn.AssigneeId IS NULL THEN 1 ELSE 0 END) AS AsSoloOwner_All,
            SUM(CASE WHEN tn.OwnerId = {pid} AND tn.AssigneeId IS NULL AND tn.CompletedDate IS NULL
                     AND tn.IsArchived = 0 AND tn.IsNote = 0 THEN 1 ELSE 0 END) AS AsSoloOwner_Open,
            SUM(CASE WHEN tn.OwnerId = {pid} AND tn.AssigneeId IS NOT NULL THEN 1 ELSE 0 END) AS DelegatedAway,
            COUNT(*) AS TotalRowsInTable
        FROM TaskNote tn
    """.format(pid=int(people_id))
    try:
        r = q.QuerySqlTop1(sql)
        if not r:
            return {'reason': 'query-empty'}
        result = {
            'reason': 'no-matching-tasks',
            'peopleId': int(people_id),
            'asAssignee_all': int(r.AsAssignee_All or 0),
            'asAssignee_open': int(r.AsAssignee_Open or 0),
            'asSoloOwner_all': int(r.AsSoloOwner_All or 0),
            'asSoloOwner_open': int(r.AsSoloOwner_Open or 0),
            'delegatedAway': int(r.DelegatedAway or 0),
            'taskNoteTableRows': int(r.TotalRowsInTable or 0),
        }
        # Also break down by StatusId so we can spot custom or NULL statuses.
        try:
            status_sql = """
                SELECT
                    ISNULL(CAST(tn.StatusId AS VARCHAR(10)), '(null)') AS StatusKey,
                    ISNULL(ts.Description, '(unknown)') AS StatusDesc,
                    COUNT(*) AS Cnt
                FROM TaskNote tn
                LEFT JOIN lookup.TaskStatus ts ON ts.Id = tn.StatusId
                WHERE (tn.AssigneeId = {pid}
                       OR (tn.OwnerId = {pid} AND tn.AssigneeId IS NULL))
                  AND tn.IsArchived = 0 AND tn.IsNote = 0 AND tn.CompletedDate IS NULL
                GROUP BY tn.StatusId, ts.Description
                ORDER BY COUNT(*) DESC
            """.format(pid=int(people_id))
            statuses = []
            for sr in q.QuerySql(status_sql):
                statuses.append('{0}={1} ({2})'.format(
                    safe_str(sr.StatusKey), safe_str(sr.StatusDesc), int(sr.Cnt or 0)))
            if statuses:
                result['openStatusBreakdown'] = ', '.join(statuses)
        except:
            pass
        return result
    except Exception as e:
        return {'reason': 'diagnose-error', 'error': safe_str(e)}


# =====================================================================
# AJAX DISPATCH
# =====================================================================
if model.HttpMethod == "post":
    action = safe_str(get_data('action', ''))
    try:
        if action == 'who_am_i':
            pid = get_current_people_id()
            info = get_current_user_info(pid)
            info['prefs'] = load_user_prefs(pid)
            info['orgSettings'] = load_org_settings()
            jprint({'success': True, 'user': info})

        elif action == 'load_my_tasks':
            pid = get_current_people_id()
            view_as = safe_int(get_data('view_as_id', 0))
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user',
                        'diagnosis': {'reason': 'no-people-id'}})
            elif view_as and not is_on_my_team(pid, view_as):
                # Drill-in auth: only allowed when the target is on the
                # caller's team (per drillin_scope settings).
                jprint({'success': False, 'message': "Not allowed to view that person's tasks"})
            else:
                target_pid = view_as if view_as else pid
                result = load_my_tasks(target_pid)
                if isinstance(result, dict) and 'error' in result:
                    jprint({'success': False, 'message': result['error']})
                else:
                    tasks = result.get('tasks', [])
                    resp = {
                        'success': True,
                        'tasks': tasks,
                        'peopleId': pid,
                        'viewAsId': view_as if view_as else 0,
                        'asOf': datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
                    }
                    resp['snoozedCount'] = result.get('snoozedCount', 0)
                    resp['hiddenTasks'] = result.get('hiddenTasks', [])
                    resp['contactMethods'] = load_contact_methods()
                    resp['contactLookbackDays'] = load_contact_lookback_days()
                    resp['orgSettings'] = load_org_settings()
                    resp['suggestedContactMethods'] = SUGGESTED_CONTACT_METHODS
                    # Capacity stats: how much I've been getting done. Cheap
                    # extra query so the Focus card can render in one trip.
                    resp['stats'] = get_my_stats(pid)
                    # Top-3 focus picks computed server-side from the same
                    # in-memory task list so the UI doesn't have to score
                    # everything client-side. Tied tasks broken by due date.
                    today_iso = datetime.datetime.now().strftime('%Y-%m-%d')
                    scored = sorted(tasks,
                                    key=lambda t: (-score_task(t, today_iso),
                                                   t.get('dueDate') or '9999-99-99'))
                    resp['focusIds'] = [t['id'] for t in scored[:3]]
                    if not tasks:
                        resp['diagnosis'] = diagnose_empty_tasks(pid)
                    jprint(resp)

        elif action == 'complete_task':
            pid = get_current_people_id()
            task_id = safe_int(get_data('task_id', 0))
            note = safe_str(get_data('note', '')).strip()
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            elif not task_id:
                jprint({'success': False, 'message': 'Missing task_id'})
            else:
                # If the user has team-complete perms, allow tasks owned by
                # their teammates. Otherwise restrict to their own plate.
                team_pids = get_team_people_ids(pid) if can_complete_for_others() else []
                if not task_belongs_to_user(task_id, pid, allow_team_pids=team_pids):
                    jprint({'success': False, 'message': 'That task is not on your list'})
                else:
                    try:
                        model.TaskNoteComplete(task_id, note or 'Completed via Task Runner', [])
                        jprint({'success': True, 'taskId': task_id})
                    except Exception as e:
                        jprint({'success': False, 'message': 'Complete failed: ' + safe_str(e)})

        elif action == 'snooze_task':
            pid = get_current_people_id()
            task_id = safe_int(get_data('task_id', 0))
            days = safe_int(get_data('days', 1))
            if days < 1:
                days = 1
            if days > 365:
                days = 365
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            elif not task_id:
                jprint({'success': False, 'message': 'Missing task_id'})
            elif not task_belongs_to_user(task_id, pid):
                jprint({'success': False, 'message': 'That task is not on your list'})
            else:
                # Snooze is local triage state -- does NOT mutate the TaskNote
                # row. Stored in content storage keyed by the user. The real
                # DueDate stays accurate for every other tool that reads it.
                snoozed = load_snoozed(pid)
                until = (datetime.datetime.now() + datetime.timedelta(days=days)).strftime('%Y-%m-%d')
                snoozed[str(task_id)] = until
                if save_snoozed(pid, snoozed):
                    jprint({'success': True, 'taskId': task_id, 'snoozedUntil': until})
                else:
                    jprint({'success': False, 'message': 'Could not save snooze state'})

        elif action == 'search_people':
            # Tokenized people search for the Reassign picker. Same pattern
            # used in TPxi_DayOfRegistration:
            #   "Last, First" / "swa, b" -> starts-with on LastName + FirstName/NickName
            #   "ben swa" / "Ben Swa"    -> every whitespace token must match
            #                               somewhere in FirstName / LastName /
            #                               NickName / Name2 (substring, AND
            #                               across tokens, OR across fields)
            term = safe_str(get_data('term', '')).strip()
            if len(term) < 2:
                jprint({'success': True, 'people': []})
            else:
                where = _name_search_where_tasks(term, 'p')
                sql = """
                    SELECT TOP 20 p.PeopleId, ISNULL(p.Name2,'') AS Name2
                    FROM People p
                    WHERE p.IsDeceased = 0
                      AND p.ArchivedFlag = 0
                      AND ({0})
                    ORDER BY p.LastName, p.FirstName
                """.format(where)
                try:
                    people = []
                    for r in q.QuerySql(sql):
                        people.append({'peopleId': int(r.PeopleId), 'name': safe_str(r.Name2)})
                    jprint({'success': True, 'people': people})
                except Exception as e:
                    jprint({'success': False, 'message': safe_str(e)})

        elif action == 'reassign_task':
            # Reassign via model.TaskNoteMassAssign (no direct SQL write).
            # Optionally add a note about the reassignment.
            pid = get_current_people_id()
            task_id = safe_int(get_data('task_id', 0))
            new_assignee = safe_int(get_data('new_assignee_id', 0))
            note_text = safe_str(get_data('note', '')).strip()
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            elif not task_id or not new_assignee:
                jprint({'success': False, 'message': 'Missing task_id or new_assignee_id'})
            else:
                team_pids = get_team_people_ids(pid) if can_reassign_for_others() else []
                if not task_belongs_to_user(task_id, pid, allow_team_pids=team_pids):
                    jprint({'success': False, 'message': 'That task is not on your list'})
                else:
                    try:
                        model.TaskNoteMassAssign([int(task_id)], int(new_assignee))
                        # Optional context note attached to the same TaskNote so
                        # the new assignee sees why it was handed off.
                        if note_text:
                            try:
                                me_name = ''
                                them_name = ''
                                try:
                                    rr = q.QuerySqlTop1(
                                        "SELECT (SELECT Name FROM People WHERE PeopleId={0}) AS Me, "
                                        "(SELECT Name FROM People WHERE PeopleId={1}) AS Them".format(int(pid), int(new_assignee))
                                    )
                                    if rr:
                                        me_name = safe_str(rr.Me)
                                        them_name = safe_str(rr.Them)
                                except:
                                    pass
                                preface = 'Reassigned'
                                if me_name and them_name:
                                    preface = 'Reassigned from ' + me_name + ' to ' + them_name
                                elif them_name:
                                    preface = 'Reassigned to ' + them_name
                                # No clean "append note to an existing TaskNote" API
                                # exists; pipeDashboard logs a comment but uses raw
                                # SQL which is blocked here. Best-effort: create a
                                # linked TaskNote on the about-person instead.
                                try:
                                    about_sql = "SELECT TOP 1 AboutPersonId FROM TaskNote WHERE TaskNoteId = {0}".format(int(task_id))
                                    rr = q.QuerySqlTop1(about_sql)
                                    about_pid = int(rr.AboutPersonId) if rr and rr.AboutPersonId else None
                                except:
                                    about_pid = None
                                if about_pid:
                                    model.CreateTaskNote(
                                        int(pid), about_pid, None, None, True,
                                        None, preface + ': ' + note_text, None, [], False
                                    )
                            except:
                                pass
                        jprint({'success': True, 'taskId': task_id, 'newAssigneeId': new_assignee})
                    except Exception as e:
                        jprint({'success': False, 'message': 'Reassign failed: ' + safe_str(e)})

        elif action == 'log_contact_attempt':
            # Log a contact effort against the task's about-person. Creates a
            # NEW TaskNote (so it shows in their contact history) attached to
            # the configured keyword for that method. Does NOT complete the
            # underlying task -- the user can still click Complete after.
            pid = get_current_people_id()
            task_id = safe_int(get_data('task_id', 0))
            code = safe_str(get_data('code', '')).strip()
            note_text = safe_str(get_data('note', '')).strip()
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            elif not task_id:
                jprint({'success': False, 'message': 'Missing task_id'})
            elif not task_belongs_to_user(task_id, pid):
                jprint({'success': False, 'message': 'That task is not on your list'})
            elif not code:
                jprint({'success': False, 'message': 'Pick a contact method'})
            else:
                method = find_contact_method_by_code(code)
                if not method:
                    jprint({'success': False, 'message': 'Unknown contact method code: ' + code})
                else:
                    # Look up the about-person from the task so we can attach
                    # the contact log to the right contact-history surface.
                    about_pid = 0
                    try:
                        about_sql = "SELECT TOP 1 AboutPersonId FROM TaskNote WHERE TaskNoteId = {0}".format(int(task_id))
                        r = q.QuerySqlTop1(about_sql)
                        if r and r.AboutPersonId:
                            about_pid = int(r.AboutPersonId)
                    except:
                        pass
                    if not about_pid:
                        jprint({'success': False, 'message': 'Task has no about-person to log against'})
                    else:
                        # Default note text shows the method label so the log
                        # entry is readable even without the user typing anything.
                        full_note = note_text or ('Contact attempt: ' + safe_str(method.get('label', code)))
                        if code:
                            full_note = '[' + code + '] ' + full_note
                        try:
                            kw_id = method.get('keywordId')
                            kw_list = [int(kw_id)] if kw_id else []
                            # CreateTaskNote signature varies by version; positional args
                            # match what pipeDashboard uses successfully on these installs.
                            # ownerId, peopleId, roleId, instructions, isNote, dueDate,
                            # notes, assigneeId, keywordIds, sendEmail
                            new_id = model.CreateTaskNote(
                                int(pid), int(about_pid), None, None, True,
                                None, full_note, None, kw_list, False
                            )
                            jprint({'success': True, 'newTaskNoteId': int(new_id) if new_id else 0,
                                    'code': code})
                        except Exception as e:
                            jprint({'success': False, 'message': 'CreateTaskNote failed: ' + safe_str(e)})

        elif action == 'save_contact_methods':
            # Settings save: methods list + lookback window. Both live in
            # the shared ProgramPulse_Settings blob.
            try:
                methods_json = safe_str(get_data('methods_json', '[]'))
                methods = json.loads(methods_json)
            except:
                methods = []
            cleaned = []
            for m in methods:
                code = safe_str(m.get('code', '')).strip()[:2].upper()
                if not code:
                    continue
                cleaned.append({
                    'code': code,
                    'label': safe_str(m.get('label', '')).strip(),
                    'keyword': safe_str(m.get('keyword', '')).strip(),
                    'keywordId': safe_int(m.get('keywordId', 0)) or None,
                })
            ok_methods = save_contact_methods(cleaned)

            # Lookback is optional in the payload; only update when provided.
            new_lookback = None
            try:
                raw = get_data('lookback_days', '')
                if raw is not None and safe_str(raw).strip() != '':
                    new_lookback = int(raw)
            except:
                new_lookback = None
            ok_lookback = True
            if new_lookback is not None:
                ok_lookback = save_contact_lookback_days(new_lookback)

            if ok_methods and ok_lookback:
                jprint({
                    'success': True,
                    'methods': cleaned,
                    'lookbackDays': load_contact_lookback_days()
                })
            else:
                jprint({'success': False, 'message': 'Could not save settings'})

        elif action == 'load_my_team':
            pid = get_current_people_id()
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            else:
                jprint({'success': True, 'team': load_my_team(pid)})

        elif action == 'save_org_settings':
            # Staff involvement + team policy. drillin_scope and drillin_actions
            # are validated against fixed value lists before persisting.
            settings = load_org_settings()
            staff_org_id = safe_int(get_data('staff_org_id', 0))
            staff_org_name = safe_str(get_data('staff_org_name', ''))
            scope = safe_str(get_data('drillin_scope', '')).strip().lower()
            actions = safe_str(get_data('drillin_actions', '')).strip().lower()
            group_disp = safe_str(get_data('team_display_grouped', '')).lower() in ('true', '1', 'yes', 'on')

            if staff_org_id:
                settings['staff_org_id'] = staff_org_id
                settings['staff_org_name'] = staff_org_name
            else:
                settings.pop('staff_org_id', None)
                settings.pop('staff_org_name', None)

            if scope in DRILLIN_SCOPE_VALUES:
                settings['drillin_scope'] = scope
            if actions in DRILLIN_ACTIONS_VALUES:
                settings['drillin_actions'] = actions
            settings['team_display_grouped'] = group_disp
            # Drop the legacy field once a write has happened with the new schema.
            settings.pop('use_subgroups', None)

            if save_org_settings(settings):
                jprint({'success': True, 'orgSettings': settings})
            else:
                jprint({'success': False, 'message': 'Could not save settings'})

        elif action == 'search_involvements':
            # Search for the staff involvement picker. Active orgs only.
            term = safe_str(get_data('term', '')).strip()
            if len(term) < 2:
                jprint({'success': True, 'involvements': []})
            else:
                tok = term.replace("'", "''")
                sql = """
                    SELECT TOP 20 o.OrganizationId, ISNULL(o.OrganizationName,'') AS Name,
                           ISNULL(d.Name,'') AS DivisionName,
                           ISNULL(p.Name,'') AS ProgramName
                    FROM Organizations o
                    LEFT JOIN Division d ON d.Id = o.DivisionId
                    LEFT JOIN Program p ON p.Id = d.ProgId
                    WHERE o.OrganizationStatusId = 30
                      AND o.OrganizationName LIKE '%{0}%'
                    ORDER BY o.OrganizationName
                """.format(tok)
                try:
                    orgs = []
                    for r in q.QuerySql(sql):
                        orgs.append({
                            'orgId': int(r.OrganizationId),
                            'name': safe_str(r.Name),
                            'division': safe_str(r.DivisionName),
                            'program': safe_str(r.ProgramName),
                        })
                    jprint({'success': True, 'involvements': orgs})
                except Exception as e:
                    jprint({'success': False, 'message': safe_str(e)})

        elif action == 'list_keywords':
            # Used by the settings UI to populate keyword dropdowns.
            try:
                rows = []
                for r in q.QuerySql("SELECT KeywordId, Description FROM dbo.Keyword WHERE IsActive = 1 ORDER BY Description"):
                    rows.append({'keywordId': int(r.KeywordId), 'description': safe_str(r.Description)})
                jprint({'success': True, 'keywords': rows})
            except Exception as e:
                jprint({'success': False, 'message': safe_str(e)})

        elif action == 'clear_all_snoozed':
            pid = get_current_people_id()
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            else:
                save_snoozed(pid, {})
                jprint({'success': True})

        elif action == 'unsnooze_task':
            pid = get_current_people_id()
            task_id = safe_int(get_data('task_id', 0))
            if not pid or not task_id:
                jprint({'success': False, 'message': 'Missing user or task_id'})
            else:
                snoozed = load_snoozed(pid)
                if str(task_id) in snoozed:
                    del snoozed[str(task_id)]
                    save_snoozed(pid, snoozed)
                jprint({'success': True, 'taskId': task_id, 'remainingHidden': len(snoozed)})

        elif action == 'save_user_pref':
            # Generic per-user pref store. Whitelist keys here so the client
            # can't write arbitrary JSON into the prefs blob.
            ALLOWED_PREF_KEYS = ('hideEmptyTeam',)
            pid = get_current_people_id()
            key = safe_str(get_data('key', '')).strip()
            raw_value = safe_str(get_data('value', ''))
            if not pid:
                jprint({'success': False, 'message': 'No logged-in user'})
            elif key not in ALLOWED_PREF_KEYS:
                jprint({'success': False, 'message': 'Unknown pref key: ' + key})
            else:
                prefs = load_user_prefs(pid)
                # Coerce known bool prefs to a real bool so client-side
                # truthy checks behave consistently.
                if key in ('hideEmptyTeam',):
                    prefs[key] = raw_value.lower() in ('true', '1', 'yes', 'on')
                else:
                    prefs[key] = raw_value
                save_user_prefs(pid, prefs)
                jprint({'success': True, 'key': key, 'value': prefs[key]})

        else:
            jprint({'success': False, 'message': 'Unknown action: ' + action})
    except Exception as e:
        import traceback
        jprint({
            'success': False,
            'message': 'Server error: ' + safe_str(e),
            'trace': traceback.format_exc(),
        })


# =====================================================================
# HTML / SPA
# =====================================================================
else:
    model.Form = '''
<style>
:root { --tr-primary:#1f4e79; --tr-accent:#0078d4; --tr-success:#107c10;
        --tr-warning:#f0ad4e; --tr-danger:#d13438; --tr-muted:#666;
        --tr-bg:#f8f9fa; --tr-border:#e1e4e8; --tr-card:#ffffff; }
* { box-sizing: border-box; }
.tr-root { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
           color: #1e293b; max-width: 1100px; margin: 0 auto; padding: 12px; }
.tr-h1 { font-size: 22px; font-weight: 700; margin: 0 0 4px; color: var(--tr-primary); }
.tr-sub { font-size: 13px; color: var(--tr-muted); margin-bottom: 12px; }
.tr-loading { padding: 30px; text-align: center; color: var(--tr-muted); }
.tr-spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid #ccc;
              border-top-color: var(--tr-accent); border-radius: 50%;
              animation: tr-spin 0.8s linear infinite; vertical-align: middle; margin-right: 6px; }
@keyframes tr-spin { to { transform: rotate(360deg); } }

/* No overflow:hidden -- the per-row Hide menu (position:absolute) needs to
   escape the bucket boundary. Header gets its own rounded top to keep the
   visual clean. */
.tr-bucket { background: var(--tr-card); border: 1px solid var(--tr-border);
             border-radius: 8px; margin-bottom: 12px; }
.tr-bucket-header { padding: 10px 14px; font-weight: 700; font-size: 14px;
                    display: flex; align-items: center; justify-content: space-between;
                    border-bottom: 1px solid var(--tr-border);
                    border-radius: 7px 7px 0 0; }
.tr-bucket-overdue .tr-bucket-header { background: #fde7e9; color: var(--tr-danger); }
.tr-bucket-today   .tr-bucket-header { background: #fff4ce; color: #7a5c00; }
.tr-bucket-week    .tr-bucket-header { background: #e1f5fe; color: var(--tr-accent); }
.tr-bucket-later   .tr-bucket-header { background: #f0f4f8; color: var(--tr-primary); }
.tr-bucket-undated .tr-bucket-header { background: #eee; color: var(--tr-muted); }
.tr-bucket-hidden  { border-style: dashed; border-color: #c8d4e3; opacity: 0.92; }
.tr-bucket-hidden  .tr-bucket-header { background: #f8f9fa; color: var(--tr-muted); border-style: dashed; }
.tr-bucket-hidden  .tr-row-title { color: #475569; }
.tr-hidden-subhead { display: flex; justify-content: space-between; align-items: center;
                     padding: 6px 14px; background: #fbfcfd; font-size: 12px;
                     border-top: 1px dashed #d6dee8; border-bottom: 1px solid #f0f0f0;
                     color: var(--tr-muted); }
.tr-hidden-subhead:first-child { border-top: 0; }
.tr-hidden-subhead-label { font-weight: 700; color: var(--tr-primary); text-transform: uppercase;
                           letter-spacing: 0.4px; font-size: 11px; }
.tr-hidden-subhead-count { display: inline-block; min-width: 18px; text-align: center;
                           background: #e1e8f0; color: var(--tr-primary);
                           border-radius: 10px; padding: 0 6px; font-weight: 700;
                           font-size: 11px; margin-right: 8px; }
.tr-hidden-subhead-action { color: var(--tr-accent); cursor: pointer; text-decoration: underline;
                            font-weight: 600; font-size: 11px; }
.tr-bucket-count { font-weight: 600; font-size: 12px; opacity: 0.85; }
.tr-hidden-until { display: inline-block; padding: 1px 7px; border-radius: 10px;
                   font-size: 11px; font-weight: 600; background: #eef2f7;
                   color: var(--tr-muted); margin-right: 4px; }
.tr-act-unhide { background: #fff; color: var(--tr-primary); border-color: var(--tr-primary); }
.tr-act-unhide:hover { background: #f0f4f8; }

.tr-row { padding: 10px 14px; border-bottom: 1px solid #f0f0f0;
          display: flex; gap: 10px; align-items: flex-start; }
.tr-row:last-child { border-bottom: none; }
.tr-row-main { flex: 1; min-width: 0; }
.tr-row-actions { display: flex; gap: 6px; flex-shrink: 0; align-items: center;
                  flex-wrap: wrap; justify-content: flex-end; }
@media (max-width: 600px) {
    .tr-row { flex-direction: column; }
    .tr-row-actions { justify-content: flex-start; }
}
.tr-act { padding: 4px 8px; font-size: 12px; font-weight: 600;
          border: 1px solid #c8d4e3; background: #fff; color: var(--tr-primary);
          border-radius: 4px; cursor: pointer; white-space: nowrap;
          text-decoration: none; display: inline-block; }
.tr-act:hover { background: #f0f4f8; }
.tr-act-complete { background: var(--tr-success); color: #fff; border-color: var(--tr-success); }
.tr-act-complete:hover { background: #0e6b0e; color: #fff; }
.tr-act-snooze { background: #fff; color: #7a5c00; border-color: #f4d35e; }
.tr-act-snooze:hover { background: #fff8e1; color: #7a5c00; }
.tr-act:disabled { opacity: 0.5; cursor: not-allowed; }

/* Snooze submenu */
.tr-snooze-wrap { position: relative; display: inline-block; }
.tr-snooze-menu { position: absolute; top: calc(100% + 4px); right: 0; z-index: 100;
                  background: #fff; border: 1px solid var(--tr-border);
                  border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.18);
                  min-width: 180px; padding: 4px 0; display: none; }
.tr-snooze-menu.open { display: block; }
.tr-snooze-menu button { display: block; width: 100%; text-align: left;
                         padding: 6px 12px; font-size: 12px; background: none;
                         border: 0; color: #1e293b; cursor: pointer; }
.tr-snooze-menu button:hover { background: #f0f4f8; }

/* View toggle (My / Team) */
.tr-view-toggle { display: inline-flex; border: 1px solid var(--tr-border);
                  border-radius: 5px; overflow: hidden; }
.tr-view-btn { padding: 5px 12px; font-size: 13px; background: #fff;
               color: var(--tr-primary); border: 0; cursor: pointer;
               font-weight: 600; }
.tr-view-btn:not(:last-child) { border-right: 1px solid var(--tr-border); }
.tr-view-btn.active { background: var(--tr-primary); color: #fff; }
.tr-view-btn:hover:not(.active) { background: #f0f4f8; }

/* Team list */
.tr-team-toolbar { display: flex; align-items: center; justify-content: space-between;
                   gap: 12px; padding: 8px 4px 12px; flex-wrap: wrap; }
.tr-team-toolbar-toggle { display: inline-flex; align-items: center; gap: 6px;
                          font-size: 13px; color: var(--tr-primary); cursor: pointer;
                          user-select: none; }
.tr-team-toolbar-toggle input { margin: 0; cursor: pointer; }
.tr-team-toolbar-meta { font-size: 12px; color: var(--tr-muted); }
.tr-team-section { background: var(--tr-card); border: 1px solid var(--tr-border);
                   border-radius: 8px; margin-bottom: 12px; }
.tr-team-section-header { padding: 10px 14px; font-weight: 700; font-size: 14px;
                          background: #f0f4f8; color: var(--tr-primary);
                          border-bottom: 1px solid var(--tr-border);
                          border-radius: 7px 7px 0 0; }
.tr-team-card { padding: 10px 14px; display: flex; align-items: center;
                gap: 12px; border-bottom: 1px solid #f0f0f0; cursor: pointer; }
.tr-team-card:last-child { border-bottom: none; }
.tr-team-card:hover { background: #fafbfc; }
.tr-team-card-main { flex: 1; min-width: 0; }
.tr-team-name { font-weight: 600; font-size: 14px; color: var(--tr-primary); }
.tr-team-meta { font-size: 12px; color: var(--tr-muted); margin-top: 2px; }
.tr-team-stat { display: inline-block; min-width: 70px; text-align: center; }
.tr-team-stat-num { font-weight: 700; font-size: 15px; }
.tr-team-stat-num.overdue { color: var(--tr-danger); }
.tr-team-stat-num.zero { color: var(--tr-muted); }
.tr-team-stat-label { font-size: 10px; color: var(--tr-muted);
                      text-transform: uppercase; letter-spacing: 0.4px; }
.tr-team-stats { display: flex; gap: 10px; flex-shrink: 0; }
.tr-team-empty { padding: 20px; color: var(--tr-muted); font-style: italic;
                 text-align: center; }
.tr-team-disabled { padding: 16px; background: #fff8e1; border: 1px solid #f4d35e;
                    border-radius: 8px; font-size: 13px; color: #7a5c00;
                    margin-bottom: 12px; }

/* Drill-in banner */
.tr-drillin-banner { background: var(--tr-accent); color: #fff;
                     padding: 10px 14px; border-radius: 8px;
                     display: flex; justify-content: space-between;
                     align-items: center; margin-bottom: 12px; }
.tr-drillin-banner button { background: rgba(255,255,255,0.2); color: #fff;
                            border: 1px solid rgba(255,255,255,0.4);
                            padding: 4px 10px; border-radius: 4px;
                            cursor: pointer; font-size: 12px; font-weight: 600; }
.tr-drillin-banner button:hover { background: rgba(255,255,255,0.3); }

/* Daily Focus card */
.tr-focus { background: linear-gradient(135deg, #1f4e79 0%, #2a5e8e 100%);
            color: #fff; border-radius: 10px; padding: 14px 18px;
            margin-bottom: 14px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
.tr-focus-title { font-size: 13px; font-weight: 700; opacity: 0.85;
                  text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
.tr-focus-help { font-size: 11px; opacity: 0.75; margin-bottom: 10px; font-style: italic; }
.tr-focus-empty { font-size: 14px; opacity: 0.85; padding: 4px 0; }
.tr-focus-list { display: flex; flex-direction: column; gap: 8px; }
.tr-focus-item { background: rgba(255,255,255,0.10); border-left: 3px solid #fff;
                 border-radius: 6px; padding: 8px 12px; display: flex;
                 justify-content: space-between; align-items: center; gap: 10px; }
.tr-focus-item-main { flex: 1; min-width: 0; }
.tr-focus-item-title { font-weight: 600; font-size: 14px; line-height: 1.3; }
.tr-focus-item-meta { font-size: 11px; opacity: 0.85; margin-top: 2px; }
.tr-focus-item button {
    padding: 4px 10px; font-size: 12px; font-weight: 600;
    background: #fff; color: var(--tr-primary); border: 0;
    border-radius: 4px; cursor: pointer; white-space: nowrap;
}
.tr-focus-item button:hover { background: #f0f4f8; }

/* Capacity line */
.tr-capacity { display: flex; gap: 16px; flex-wrap: wrap;
               padding: 8px 14px; background: var(--tr-card);
               border: 1px solid var(--tr-border); border-radius: 8px;
               margin-bottom: 12px; font-size: 12px; color: var(--tr-muted); }
.tr-capacity strong { color: #1e293b; font-weight: 700; }

/* Snoozed footer */
.tr-snoozed-footer { margin-top: 14px; padding: 10px 14px; background: #f8f9fa;
                     border: 1px dashed #c8d4e3; border-radius: 8px;
                     font-size: 12px; color: var(--tr-muted);
                     display: flex; justify-content: space-between; align-items: center; }
.tr-snoozed-footer a { color: var(--tr-accent); cursor: pointer; text-decoration: underline; }

/* Modals (Log Contact, Reassign, Settings) */
.tr-modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.45);
                    z-index: 8000; display: none; align-items: center;
                    justify-content: center; padding: 20px; }
.tr-modal-overlay.open { display: flex; }
.tr-modal { background: #fff; border-radius: 10px; max-width: 520px; width: 100%;
            max-height: 85vh; overflow-y: auto; padding: 18px 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.25); }
.tr-modal-header { display: flex; justify-content: space-between; align-items: center;
                   margin-bottom: 12px; padding-bottom: 8px;
                   border-bottom: 1px solid var(--tr-border); }
.tr-modal-title { font-size: 16px; font-weight: 700; color: var(--tr-primary); }
.tr-modal-close { background: none; border: 0; font-size: 22px; cursor: pointer;
                  color: var(--tr-muted); line-height: 1; padding: 0 4px; }
.tr-modal-body { font-size: 13px; color: #1e293b; }
.tr-modal-body label { display: block; font-size: 12px; font-weight: 600;
                       color: var(--tr-muted); margin: 10px 0 4px; }
.tr-modal-body textarea, .tr-modal-body input[type="text"] {
    width: 100%; box-sizing: border-box; padding: 6px 8px; font-size: 13px;
    border: 1px solid #c8ccd0; border-radius: 4px; font-family: inherit;
}
.tr-modal-body textarea { min-height: 60px; resize: vertical; }
.tr-modal-footer { margin-top: 14px; display: flex; gap: 8px; justify-content: flex-end; }

.tr-method-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(110px, 1fr));
                  gap: 8px; margin-top: 6px; }
.tr-method-btn { display: flex; align-items: center; justify-content: center; gap: 6px;
                 padding: 10px 12px; border: 1px solid var(--tr-border);
                 background: #f8f9fa; color: var(--tr-primary); font-weight: 600;
                 border-radius: 6px; cursor: pointer; font-size: 13px; }
.tr-method-btn:hover { background: #e1f5fe; border-color: var(--tr-accent); }
.tr-method-btn .tr-method-code {
    background: var(--tr-primary); color: #fff; width: 22px; height: 22px;
    border-radius: 50%; display: inline-flex; align-items: center;
    justify-content: center; font-size: 11px; font-weight: 700;
}
.tr-method-warn { font-size: 12px; color: #7a5c00; background: #fff4ce;
                  border: 1px solid #f4d35e; border-radius: 4px;
                  padding: 8px 10px; margin-top: 6px; }

.tr-people-results { max-height: 280px; overflow-y: auto;
                     border: 1px solid var(--tr-border); border-radius: 6px;
                     margin-top: 4px; background: #fff; }
.tr-people-result { padding: 7px 10px; cursor: pointer; font-size: 13px;
                    border-bottom: 1px solid #f0f0f0; }
.tr-people-result:hover { background: #f0f4f8; }
.tr-people-result.selected { background: #e1f5fe; }
.tr-people-result:last-child { border-bottom: none; }

/* Toast for action feedback */
.tr-toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%);
            padding: 9px 16px; border-radius: 6px; color: #fff; font-weight: 600;
            font-size: 13px; z-index: 9999; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            opacity: 0; transition: opacity 0.2s; }
.tr-toast.show { opacity: 1; }
.tr-toast-success { background: var(--tr-success); }
.tr-toast-error   { background: var(--tr-danger); }
.tr-toast-info    { background: var(--tr-primary); }
.tr-row-title { font-weight: 600; font-size: 14px; margin-bottom: 2px; line-height: 1.3; }
.tr-row-meta { font-size: 12px; color: var(--tr-muted); display: flex; gap: 8px; flex-wrap: wrap; }
.tr-row-meta a { color: var(--tr-accent); text-decoration: none; }
.tr-row-meta a:hover { text-decoration: underline; }
.tr-chip { display: inline-block; padding: 1px 7px; border-radius: 10px;
           font-size: 11px; font-weight: 600; background: #eef2f7; color: var(--tr-primary); }
.tr-chip-overdue { background: #fde7e9; color: var(--tr-danger); }
.tr-chip-today   { background: #fff4ce; color: #7a5c00; }
.tr-row-contact { display: inline-flex; gap: 3px; align-items: center;
                  margin-left: 6px; vertical-align: middle; }
.tr-contact-code { display: inline-block; padding: 1px 5px; border-radius: 3px;
                   font-size: 11px; font-weight: 700; background: #f0f4f8;
                   color: var(--tr-muted); line-height: 1.4; }
.tr-contact-code.has-count { background: rgba(52,152,219,0.15); color: var(--tr-accent); }
.tr-contact-code.other     { background: #fff4ce; color: #7a5c00; }
.tr-contact-last  { display: inline-block; margin-left: 6px; font-size: 11px;
                    color: var(--tr-muted); vertical-align: middle; }

.tr-row-extras { font-size: 12px; color: #2c3e50; margin-top: 3px; line-height: 1.45; }
.tr-extra-name { color: var(--tr-muted); font-weight: 600; }
.tr-extra-val  { color: #1e293b; }
.tr-extra-kw   { color: var(--tr-primary); font-weight: 700; }
.tr-empty { padding: 16px; color: var(--tr-muted); font-style: italic; text-align: center; }

.tr-banner-error { background: #fde7e9; border: 1px solid #f5c6cb;
                   color: var(--tr-danger); padding: 10px 14px; border-radius: 6px;
                   margin-bottom: 12px; font-size: 13px; }
.tr-as-of { font-size: 11px; color: #999; margin-top: 8px; text-align: right; }
.tr-tools { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 10px; }
.tr-btn { padding: 5px 10px; font-size: 13px; border: 1px solid var(--tr-primary);
          background: var(--tr-primary); color: #fff; border-radius: 4px; cursor: pointer; }
.tr-btn-secondary { background: #fff; color: var(--tr-primary); }
.tr-btn:disabled { opacity: 0.5; cursor: not-allowed; }
</style>

<div class="tr-root" id="trRoot">
  <div>
    <div class="tr-h1">Task Runner <span style="font-size:11px;color:#888;font-weight:400;">v''' + APP_VERSION + '''</span></div>
    <div class="tr-sub" id="trSub">Loading...</div>
  </div>
  <div class="tr-tools">
    <button class="tr-btn tr-btn-secondary" onclick="TRApp.reload()">Refresh</button>
    <button class="tr-btn tr-btn-secondary" onclick="TRApp.openSettings()" title="Configure contact methods">&#9881; Settings</button>
    <div class="tr-view-toggle" id="trViewToggle" style="display:none;">
      <button id="trViewMyBtn" class="tr-view-btn active" onclick="TRApp.setView('mine')">My View</button>
      <button id="trViewTeamBtn" class="tr-view-btn" onclick="TRApp.setView('team')">Team View</button>
    </div>
    <span id="trCount" class="tr-sub" style="margin:0;"></span>
  </div>
  <div id="trFocus"></div>
  <div id="trCapacity"></div>
  <div id="trBody">
    <div class="tr-loading"><span class="tr-spinner"></span>Loading tasks...</div>
  </div>
  <div id="trSnoozedFooter"></div>
  <div class="tr-as-of" id="trAsOf"></div>

  <!-- Log Contact modal -->
  <div class="tr-modal-overlay" id="trLogModal">
    <div class="tr-modal">
      <div class="tr-modal-header">
        <span class="tr-modal-title">Log a contact attempt</span>
        <button class="tr-modal-close" onclick="TRApp.closeModal('trLogModal')">&times;</button>
      </div>
      <div class="tr-modal-body">
        <div id="trLogTaskContext" style="font-size:12px;color:var(--tr-muted);margin-bottom:8px;"></div>
        <label>Method</label>
        <div class="tr-method-grid" id="trLogMethods"></div>
        <div id="trLogWarn"></div>
        <label>Note (optional)</label>
        <textarea id="trLogNote" placeholder="e.g. left voicemail, will follow up Thursday"></textarea>
        <div class="tr-modal-footer">
          <button class="tr-btn tr-btn-secondary" onclick="TRApp.closeModal('trLogModal')">Cancel</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Settings modal -->
  <div class="tr-modal-overlay" id="trSettingsModal">
    <div class="tr-modal" style="max-width: 760px;">
      <div class="tr-modal-header">
        <span class="tr-modal-title">&#9881; Contact Method Settings</span>
        <button class="tr-modal-close" onclick="TRApp.closeModal('trSettingsModal')">&times;</button>
      </div>
      <div class="tr-modal-body">
        <!-- Org Settings: staff involvement + subgroups (used by Team View) -->
        <div style="background:#f8f9fa; border:1px solid var(--tr-border); border-radius:6px;
                    padding:10px 12px; margin-bottom:14px;">
          <div style="font-weight:700; color:var(--tr-primary); font-size:13px; margin-bottom:6px;">
            Staff &amp; Departments
            <span style="font-weight:400; color:var(--tr-muted); font-size:11px;">(used by upcoming Team View)</span>
          </div>
          <label style="font-size:12px; font-weight:600; color:var(--tr-muted); margin-bottom:4px;">Staff Involvement</label>
          <div style="position:relative;">
            <input type="text" id="trSettingsStaffOrg" placeholder="Type to search involvements..." autocomplete="off" />
            <input type="hidden" id="trSettingsStaffOrgId" />
            <input type="hidden" id="trSettingsStaffOrgName" />
            <div class="tr-people-results" id="trSettingsOrgResults" style="display:none; position:absolute; top:100%; left:0; right:0; z-index:9;"></div>
          </div>
          <div id="trSettingsStaffOrgSelected" style="font-size:12px; color:var(--tr-accent); margin-top:4px;"></div>
          <label style="font-size:12px; font-weight:600; color:var(--tr-muted); margin-top:10px; display:block;">Team scope &mdash; whose tasks can I drill into?</label>
          <div style="font-size:13px;">
            <label style="display:block; padding:2px 0; cursor:pointer;">
              <input type="radio" name="trSettingsScope" value="off" style="margin-right:6px; vertical-align:middle;">
              Off &mdash; My View only, no Team View
            </label>
            <label style="display:block; padding:2px 0; cursor:pointer;">
              <input type="radio" name="trSettingsScope" value="subgroup" style="margin-right:6px; vertical-align:middle;">
              People who share at least one subgroup with me
            </label>
            <label style="display:block; padding:2px 0; cursor:pointer;">
              <input type="radio" name="trSettingsScope" value="all_staff" style="margin-right:6px; vertical-align:middle;">
              Everyone in the staff involvement
            </label>
          </div>

          <label style="font-size:12px; font-weight:600; color:var(--tr-muted); margin-top:10px; display:block;">Actions allowed on teammates' tasks</label>
          <div style="font-size:13px;">
            <label style="display:block; padding:2px 0; cursor:pointer;">
              <input type="radio" name="trSettingsActions" value="none" style="margin-right:6px; vertical-align:middle;">
              View only
            </label>
            <label style="display:block; padding:2px 0; cursor:pointer;">
              <input type="radio" name="trSettingsActions" value="reassign" style="margin-right:6px; vertical-align:middle;">
              Reassign only &mdash; lets a leader rebalance the load
            </label>
            <label style="display:block; padding:2px 0; cursor:pointer;">
              <input type="radio" name="trSettingsActions" value="reassign_complete" style="margin-right:6px; vertical-align:middle;">
              Reassign + Complete on behalf of teammates
            </label>
          </div>

          <label style="font-size:12px; font-weight:600; color:var(--tr-muted); margin-top:10px; display:block;">
            <input type="checkbox" id="trSettingsGroupDept" style="margin-right:6px; vertical-align:middle;">
            Group team list by subgroup (department) headers
          </label>
          <p style="font-size:11px; color:var(--tr-muted); margin:4px 0 0; font-style:italic;">
            When on, the team list separates members by MemberTag (e.g., "Pastoral", "Worship"). Otherwise it's a flat list.
          </p>
        </div>

        <div style="font-weight:700; color:var(--tr-primary); font-size:13px; margin-bottom:4px;">
          Contact Methods
        </div>
        <p style="font-size: 12px; color: var(--tr-muted); margin: 0 0 10px;">
          Define keyword codes for contact-effort tracking. Each method maps a short code
          (P, E, T, V, M, etc.) to a TouchPoint TaskNote keyword. These settings are shared
          with ProspectBuilder / Program Pulse, so configure once and both tools see the same list.
        </p>
        <div style="display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px solid var(--tr-border); margin-bottom:10px;">
          <label style="font-size:13px; font-weight:600; color:var(--tr-primary); margin:0;">Lookback window (days):</label>
          <input type="number" id="trSettingsLookback" min="7" max="1825" step="1"
                 style="width:90px; text-align:center;" />
          <span style="font-size:11px; color:var(--tr-muted);">how far back the contact badges count. 30-365 typical, 182 = ~6 months (default).</span>
        </div>
        <table style="width:100%; border-collapse: collapse; font-size: 13px;">
          <thead>
            <tr style="border-bottom: 2px solid var(--tr-border);">
              <th style="padding: 6px; text-align: left; width: 60px;">Code</th>
              <th style="padding: 6px; text-align: left; width: 140px;">Label</th>
              <th style="padding: 6px; text-align: left;">Keyword</th>
              <th style="padding: 6px; width: 60px;"></th>
            </tr>
          </thead>
          <tbody id="trSettingsRows"></tbody>
        </table>
        <button class="tr-btn tr-btn-secondary" style="margin-top: 10px;" onclick="TRApp.addMethod()">+ Add Method</button>
        <div class="tr-modal-footer">
          <button class="tr-btn tr-btn-secondary" onclick="TRApp.closeModal('trSettingsModal')">Cancel</button>
          <button class="tr-btn" onclick="TRApp.saveSettings()">Save Settings</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Reassign modal -->
  <div class="tr-modal-overlay" id="trReassignModal">
    <div class="tr-modal">
      <div class="tr-modal-header">
        <span class="tr-modal-title">Reassign task</span>
        <button class="tr-modal-close" onclick="TRApp.closeModal('trReassignModal')">&times;</button>
      </div>
      <div class="tr-modal-body">
        <div id="trReassignTaskContext" style="font-size:12px;color:var(--tr-muted);margin-bottom:8px;"></div>
        <label>Reassign to</label>
        <input type="text" id="trReassignSearch" placeholder="Name, partial OK (e.g. be swa or swa, b)" autocomplete="off" />
        <div class="tr-people-results" id="trReassignResults" style="display:none;"></div>
        <label>Note (optional)</label>
        <textarea id="trReassignNote" placeholder="Context for the new assignee"></textarea>
        <div class="tr-modal-footer">
          <button class="tr-btn tr-btn-secondary" onclick="TRApp.closeModal('trReassignModal')">Cancel</button>
          <button class="tr-btn" id="trReassignSubmit" onclick="TRApp.submitReassign()" disabled>Reassign</button>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
(function() {
"use strict";

var state = {
    user: null,
    tasks: [],
    hiddenTasks: [],
    peekingHidden: false,     // transient: just show the hidden bucket, don't un-hide
    asOf: '',
    diagnosis: null,
    snoozedCount: 0,
    stats: {},
    focusIds: [],
    contactMethods: [],
    suggestedContactMethods: [],
    contactLookbackDays: 182,
    orgSettings: {},
    activeTaskId: 0,          // task in the currently-open modal (log / reassign)

    // Team View state
    view: 'mine',             // 'mine' | 'team' | 'team-detail'
    viewAsId: 0,              // peopleId being drilled into when view === 'team-detail'
    viewAsName: '',
    team: null,               // {members, permissions, displayGrouped}
    hideEmptyTeam: false      // user pref: hide teammates with 0 open tasks
};

var scriptPath = (function() {
    var p = window.location.pathname;
    if (p.indexOf('/PyScriptForm/') > -1) return p;
    return p.replace('/PyScript/', '/PyScriptForm/');
})();

function ajax(action, params, cb) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', scriptPath, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    var data = 'action=' + encodeURIComponent(action);
    if (params) for (var k in params) {
        if (params.hasOwnProperty(k))
            data += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status >= 200 && xhr.status < 300) {
            try { cb(null, JSON.parse(xhr.responseText)); }
            catch (e) { cb('parse-error: ' + e.message, null); }
        } else { cb('HTTP ' + xhr.status, null); }
    };
    xhr.send(data);
}

function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function todayIso() {
    var d = new Date();
    var m = String(d.getMonth() + 1); if (m.length === 1) m = '0' + m;
    var dd = String(d.getDate());     if (dd.length === 1) dd = '0' + dd;
    return d.getFullYear() + '-' + m + '-' + dd;
}

function endOfWeekIso() {
    // ISO week ending Saturday (local)
    var d = new Date();
    var add = 6 - d.getDay();   // Sun=0 ... Sat=6
    d.setDate(d.getDate() + add);
    var m = String(d.getMonth() + 1); if (m.length === 1) m = '0' + m;
    var dd = String(d.getDate());     if (dd.length === 1) dd = '0' + dd;
    return d.getFullYear() + '-' + m + '-' + dd;
}

function bucketize(tasks) {
    var today = todayIso();
    var weekEnd = endOfWeekIso();
    var buckets = { overdue: [], today: [], week: [], later: [], undated: [] };
    for (var i = 0; i < tasks.length; i++) {
        var t = tasks[i];
        if (!t.dueDate) { buckets.undated.push(t); continue; }
        if (t.dueDate < today)    buckets.overdue.push(t);
        else if (t.dueDate === today) buckets.today.push(t);
        else if (t.dueDate <= weekEnd) buckets.week.push(t);
        else buckets.later.push(t);
    }
    return buckets;
}

function daysBetween(isoA, isoB) {
    // Returns isoB - isoA in days (positive if B after A). Simple, dates only.
    if (!isoA || !isoB) return 0;
    var a = new Date(isoA + 'T00:00:00');
    var b = new Date(isoB + 'T00:00:00');
    return Math.round((b - a) / 86400000);
}

function renderContactBadges(efforts) {
    // efforts: { methods: {P: 3, E: 1, ...}, total, lastDate, other }
    // Returns a small inline block of P(3) E(1) T(0) ... badges, plus
    // "Last X" date. Empty string when nothing to show.
    if (!efforts) return '';
    var methods = state.contactMethods || [];
    if (!methods.length) return '';
    var parts = '';
    var any = false;
    for (var i = 0; i < methods.length; i++) {
        var code = methods[i].code || '?';
        var cnt = (efforts.methods && efforts.methods[code]) || 0;
        if (cnt > 0) any = true;
        var cls = cnt > 0 ? 'tr-contact-code has-count' : 'tr-contact-code';
        parts += '<span class="' + cls + '" title="' + esc(methods[i].label || code)
              + ': ' + cnt + '">' + esc(code) + '(' + cnt + ')</span>';
    }
    if (efforts.other && efforts.other > 0) {
        any = true;
        parts += '<span class="tr-contact-code other" title="Other (TaskNotes without a configured keyword)">O(' + efforts.other + ')</span>';
    }
    // Skip the whole block when there's nothing to communicate -- avoids
    // visual noise for people who've never been contacted at all.
    if (!any) return '';
    var html = '<span class="tr-row-contact">' + parts + '</span>';
    if (efforts.lastDate) {
        html += '<span class="tr-contact-last">Last ' + esc(efforts.lastDate) + '</span>';
    }
    return html;
}

function renderRow(t) {
    var today = todayIso();
    var dueChip = '';
    if (t.dueDate) {
        var diff = daysBetween(today, t.dueDate);
        var cls = 'tr-chip';
        var label = t.dueDate;
        if (diff < 0) {
            cls = 'tr-chip tr-chip-overdue';
            label = Math.abs(diff) + 'd overdue';
        } else if (diff === 0) {
            cls = 'tr-chip tr-chip-today';
            label = 'Today';
        } else if (diff === 1) {
            label = 'Tomorrow';
        } else if (diff <= 6) {
            label = 'in ' + diff + 'd';
        }
        dueChip = '<span class="' + cls + '">' + esc(label) + '</span>';
    } else {
        dueChip = '<span class="tr-chip">No due date</span>';
    }

    var titleText = t.instructions || '(no instructions)';

    // Meta line: about-person link, org name, keyword chips, contact-effort badges
    var meta = [];
    if (t.aboutPersonId && t.aboutName) {
        meta.push('<a href="/Person2/' + t.aboutPersonId + '" target="_blank">' + esc(t.aboutName) + '</a>');
    } else if (t.aboutName) {
        meta.push(esc(t.aboutName));
    }
    if (t.orgName) meta.push(esc(t.orgName));
    if (t.ownerName) meta.push('<span style="opacity:0.75">from ' + esc(t.ownerName) + '</span>');
    for (var k = 0; k < (t.keywords || []).length; k++) {
        meta.push('<span class="tr-chip">' + esc(t.keywords[k]) + '</span>');
    }

    // Contact-effort badges: per-method counts over the last 26 weeks.
    // Only renders when the task has an about-person and contact methods
    // are configured. Codes with 0 count still show (muted) so the user
    // can see what hasn't been tried.
    var contactHtml = renderContactBadges(t.contactEfforts);
    if (contactHtml) meta.push(contactHtml);

    // Extras line: grouped by keyword, each pair "Name: Value".
    var extrasHtml = '';
    if (t.extras && t.extras.length) {
        var byKw = {}; var order = [];
        for (var ex = 0; ex < t.extras.length; ex++) {
            var e = t.extras[ex];
            if (!byKw[e.keyword]) { byKw[e.keyword] = []; order.push(e.keyword); }
            byKw[e.keyword].push(e);
        }
        var groups = [];
        for (var gi = 0; gi < order.length; gi++) {
            var kwName = order[gi];
            var parts = [];
            for (var ei = 0; ei < byKw[kwName].length; ei++) {
                var item = byKw[kwName][ei];
                parts.push('<span class="tr-extra-name">' + esc(item.name) + ':</span> '
                         + '<span class="tr-extra-val">' + esc(item.value) + '</span>');
            }
            // Show keyword label only if multiple keyword groups, so a single
            // group stays compact (the keyword already shows as a chip above).
            if (order.length > 1) {
                groups.push('<span class="tr-extra-kw">' + esc(kwName) + ':</span> ' + parts.join(' &middot; '));
            } else {
                groups.push(parts.join(' &middot; '));
            }
        }
        extrasHtml = '<div class="tr-row-extras">' + groups.join(' &nbsp;|&nbsp; ') + '</div>';
    }

    // Action strip varies by mode:
    //   mine          -> full actions (Complete / Log / Reassign / Hide)
    //   team-detail   -> only what permissions allow (Reassign +/- Complete);
    //                    no Hide (personal triage), no Log (uses TouchPoint
    //                    auth and the user isn't the about-person's caregiver
    //                    in this context).
    var inDrillin = state.view === 'team-detail';
    var perms = (state.team && state.team.permissions) || {};
    var actionsHtml = '<div class="tr-row-actions">';

    if (!inDrillin) {
        // Personal view: full action strip.
        actionsHtml += '<button class="tr-act tr-act-complete" onclick="TRApp.complete(' + t.id + ')" title="Mark complete">&#10003; Complete</button>';
        if (t.aboutPersonId && (state.contactMethods || []).length > 0) {
            actionsHtml += '<button class="tr-act" onclick="TRApp.openLog(' + t.id + ')" title="Log a contact attempt (phone, email, visit) on this person">Log</button>';
        }
        actionsHtml += '<button class="tr-act" onclick="TRApp.openReassign(' + t.id + ')" title="Reassign this task to someone else">Reassign</button>';
        var hideTooltip = 'Hide this task from Task Runner for a while. The actual due date in TouchPoint is unchanged -- this is just your personal triage view.';
        actionsHtml += '<div class="tr-snooze-wrap">'
            +   '<button class="tr-act tr-act-snooze" onclick="TRApp.toggleSnoozeMenu(' + t.id + ', event)" title="' + hideTooltip + '">Hide &#9662;</button>'
            +   '<div class="tr-snooze-menu" id="trSnoozeMenu' + t.id + '">'
            +     '<button onclick="TRApp.snooze(' + t.id + ', 1)">Hide until tomorrow</button>'
            +     '<button onclick="TRApp.snooze(' + t.id + ', 3)">Hide for 3 days</button>'
            +     '<button onclick="TRApp.snooze(' + t.id + ', 7)">Hide for a week</button>'
            +     '<button onclick="TRApp.snooze(' + t.id + ', 14)">Hide for 2 weeks</button>'
            +     '<button onclick="TRApp.snooze(' + t.id + ', 30)">Hide for 30 days</button>'
            +   '</div>'
            + '</div>';
    } else {
        // Drill-in: only what the configured policy allows.
        if (perms.canComplete) {
            actionsHtml += '<button class="tr-act tr-act-complete" onclick="TRApp.complete(' + t.id + ')" title="Complete on behalf of ' + esc(state.viewAsName || 'this person') + '">&#10003; Complete</button>';
        }
        if (perms.canReassign) {
            actionsHtml += '<button class="tr-act" onclick="TRApp.openReassign(' + t.id + ')" title="Reassign this task">Reassign</button>';
        }
        if (!perms.canComplete && !perms.canReassign) {
            actionsHtml += '<span class="tr-sub" style="font-style:italic;">View only</span>';
        }
    }
    actionsHtml += '</div>';

    return '<div class="tr-row" id="trRow' + t.id + '" data-task-id="' + t.id + '">'
         +   '<div class="tr-row-main">'
         +     '<div class="tr-row-title">' + esc(titleText) + '</div>'
         +     '<div class="tr-row-meta">' + dueChip + (meta.length ? ' &middot; ' + meta.join(' &middot; ') : '') + '</div>'
         +     extrasHtml
         +   '</div>'
         +   actionsHtml
         + '</div>';
}

function renderBucket(name, label, tasks, modifier) {
    if (!tasks.length && (modifier === 'undated' || modifier === 'later')) {
        return ''; // hide empty trailing buckets
    }
    var html = '<div class="tr-bucket tr-bucket-' + modifier + '">';
    html += '<div class="tr-bucket-header">';
    html += '<span>' + esc(label) + '</span>';
    html += '<span class="tr-bucket-count">' + tasks.length + '</span>';
    html += '</div>';
    if (!tasks.length) {
        html += '<div class="tr-empty">Nothing here.</div>';
    } else {
        for (var i = 0; i < tasks.length; i++) html += renderRow(tasks[i]);
    }
    html += '</div>';
    return html;
}

function renderFocus() {
    var el = document.getElementById('trFocus');
    if (!el) return;
    // Focus card is a personal triage surface -- doesn't apply when looking
    // at someone else's plate.
    if (state.view === 'team-detail') { el.innerHTML = ''; return; }
    if (!state.tasks.length || !state.focusIds || !state.focusIds.length) {
        el.innerHTML = '';
        return;
    }
    var byId = {};
    for (var i = 0; i < state.tasks.length; i++) byId[state.tasks[i].id] = state.tasks[i];
    var picks = [];
    for (var j = 0; j < state.focusIds.length; j++) {
        var t = byId[state.focusIds[j]];
        if (t) picks.push(t);
    }
    if (!picks.length) { el.innerHTML = ''; return; }

    var today = todayIso();
    var html = '<div class="tr-focus" title="The top 3 of your open tasks, ranked by how overdue they are and whether any urgent keywords are on them (hospital, death, emergency, urgent, baptism, salvation, bereavement). Same tasks also appear in the buckets below.">';
    html += '<div class="tr-focus-title">&#127919; Where to start</div>';
    html += '<div class="tr-focus-help">Your top 3 picks by urgency. Tap Complete to mark done.</div>';
    html += '<div class="tr-focus-list">';
    for (var p = 0; p < picks.length; p++) {
        var t = picks[p];
        var hint = '';
        if (t.dueDate) {
            var diff = daysBetween(today, t.dueDate);
            if (diff < 0) hint = Math.abs(diff) + 'd overdue';
            else if (diff === 0) hint = 'Due today';
            else if (diff === 1) hint = 'Due tomorrow';
            else hint = 'Due in ' + diff + 'd';
        } else {
            hint = 'No due date';
        }
        if (t.aboutName) hint += ' \\u00B7 ' + esc(t.aboutName);
        var title = t.instructions || '(no instructions)';
        html += '<div class="tr-focus-item">';
        html += '<div class="tr-focus-item-main">';
        html += '<div class="tr-focus-item-title">' + esc(title.substring(0, 120)) + (title.length > 120 ? '...' : '') + '</div>';
        html += '<div class="tr-focus-item-meta">' + hint + '</div>';
        html += '</div>';
        html += '<button onclick="TRApp.complete(' + t.id + ')" title="Mark this task complete">Complete</button>';
        html += '</div>';
    }
    html += '</div></div>';
    el.innerHTML = html;
}

function renderCapacity() {
    var el = document.getElementById('trCapacity');
    if (!el) return;
    var s = state.stats || {};
    var openCount = state.tasks.length;
    var today  = s.completedToday || 0;
    var d7     = s.completed7d || 0;
    var d30    = s.completed30d || 0;
    var avg7   = d7 / 7.0;
    var avg30  = d30 / 30.0;
    var avg    = avg7 > 0 ? avg7 : avg30;   // 7d is more responsive when present

    var parts = [];
    parts.push('<span><strong>' + openCount + '</strong> open</span>');
    parts.push('<span><strong>' + today + '</strong> completed today</span>');
    if (d7 > 0)  parts.push('<span><strong>' + d7 + '</strong> completed in 7 days</span>');
    if (avg > 0 && openCount > 0) {
        var daysToZero = Math.ceil(openCount / avg);
        var label;
        if (daysToZero <= 7) label = '~' + daysToZero + ' days to zero at current pace';
        else if (daysToZero <= 60) label = '~' + Math.round(daysToZero / 7) + ' weeks to zero at current pace';
        else label = 'pace puts zero ' + Math.round(daysToZero / 30) + '+ months out';
        parts.push('<span>' + label + '</span>');
    }
    el.innerHTML = '<div class="tr-capacity">' + parts.join(' \\u00B7 ') + '</div>';
}

function renderSnoozedFooter() {
    var el = document.getElementById('trSnoozedFooter');
    if (!el) return;
    // Hidden tasks are personal -- not relevant when drilling into someone else.
    if (state.view === 'team-detail') { el.innerHTML = ''; return; }
    if (!state.snoozedCount) { el.innerHTML = ''; return; }
    var peekLabel = state.peekingHidden ? 'Hide them again' : 'Peek';
    el.innerHTML = '<div class="tr-snoozed-footer" title="Hidden tasks are still active in TouchPoint; the hide only affects Task Runner.">'
        + '<span>' + state.snoozedCount + ' task' + (state.snoozedCount === 1 ? '' : 's')
        + ' hidden from this view.</span>'
        + '<span>'
        +   '<a onclick="TRApp.togglePeek()" style="margin-right:14px;">' + peekLabel + '</a>'
        +   '<a onclick="TRApp.revealAll()">Reveal all</a>'
        + '</span>'
        + '</div>';
}

function renderRowHidden(t) {
    // A muted variant of renderRow for the Hidden bucket. Shows the
    // hidden-until date plus an Unhide button instead of Complete/Hide.
    var title = t.instructions || '(no instructions)';
    var meta = [];
    if (t.aboutPersonId && t.aboutName) {
        meta.push('<a href="/Person2/' + t.aboutPersonId + '" target="_blank">' + esc(t.aboutName) + '</a>');
    } else if (t.aboutName) {
        meta.push(esc(t.aboutName));
    }
    if (t.dueDate) meta.push('Due ' + esc(t.dueDate));
    return '<div class="tr-row" id="trRow' + t.id + '">'
         +   '<div class="tr-row-main">'
         +     '<div class="tr-row-title">' + esc(title) + '</div>'
         +     '<div class="tr-row-meta">'
         +       '<span class="tr-hidden-until">Hidden until ' + esc(t.hiddenUntil || '?') + '</span>'
         +       (meta.length ? meta.join(' &middot; ') : '')
         +     '</div>'
         +   '</div>'
         +   '<div class="tr-row-actions">'
         +     '<button class="tr-act tr-act-complete" onclick="TRApp.complete(' + t.id + ')" title="Mark complete">&#10003; Complete</button>'
         +     '<button class="tr-act tr-act-unhide" onclick="TRApp.unhide(' + t.id + ')" title="Show this task again in the main list">Unhide</button>'
         +   '</div>'
         + '</div>';
}

function renderHiddenBucket() {
    if (!state.peekingHidden || !state.hiddenTasks.length) return '';
    // Sub-group hidden tasks by their actual due-date urgency so the user
    // can see at a glance whether their hidden pile contains overdue work
    // (and bulk-unhide a specific bucket).
    var sub = bucketize(state.hiddenTasks);
    var groupDefs = [
        {key: 'overdue',  label: 'Overdue (hidden)',     tasks: sub.overdue},
        {key: 'today',    label: 'Today (hidden)',       tasks: sub.today},
        {key: 'week',     label: 'This Week (hidden)',   tasks: sub.week},
        {key: 'later',    label: 'Later (hidden)',       tasks: sub.later},
        {key: 'undated',  label: 'No Due Date (hidden)', tasks: sub.undated}
    ];

    var html = '<div class="tr-bucket tr-bucket-hidden">';
    html += '<div class="tr-bucket-header">'
         +    '<span>&#128064; Hidden (peeking)</span>'
         +    '<span class="tr-bucket-count">' + state.hiddenTasks.length + '</span>'
         +  '</div>';

    // Skip sub-headers entirely when only one group has anything in it --
    // keeps the visual quiet when there's nothing interesting to compare.
    var populated = [];
    for (var g = 0; g < groupDefs.length; g++) {
        if (groupDefs[g].tasks.length) populated.push(groupDefs[g]);
    }
    var showSubHeaders = populated.length > 1;

    for (var i = 0; i < populated.length; i++) {
        var grp = populated[i];
        if (showSubHeaders) {
            var pids = grp.tasks.map(function(t) { return t.id; }).join(',');
            html += '<div class="tr-hidden-subhead">'
                 +    '<span class="tr-hidden-subhead-label">' + esc(grp.label) + '</span>'
                 +    '<span>'
                 +      '<span class="tr-hidden-subhead-count">' + grp.tasks.length + '</span>'
                 +      '<a class="tr-hidden-subhead-action" '
                 +        'onclick="TRApp.unhideGroup(\\'' + pids + '\\')">Unhide all ' + grp.tasks.length + '</a>'
                 +    '</span>'
                 +  '</div>';
        }
        for (var ti = 0; ti < grp.tasks.length; ti++) {
            html += renderRowHidden(grp.tasks[ti]);
        }
    }

    html += '</div>';
    return html;
}

function renderAll() {
    var body = document.getElementById('trBody');
    if (!state.tasks.length) {
        var msg = '<div class="tr-bucket"><div class="tr-empty">'
                + 'You have no open tasks. Nice work.</div></div>';
        // If the server tells us why, show it so we can debug install-specific
        // misfilters. Most useful field: asAssignee_open / asSoloOwner_open.
        if (state.diagnosis) {
            var d = state.diagnosis;
            var rows = '';
            for (var k in d) {
                if (!d.hasOwnProperty(k)) continue;
                rows += '<div><strong>' + esc(k) + ':</strong> ' + esc(d[k]) + '</div>';
            }
            msg += '<div class="tr-banner-error" style="background:#fff4ce;color:#7a5c00;border-color:#f4d35e;">'
                 + '<div style="font-weight:700;margin-bottom:4px;">Diagnostic</div>'
                 + rows
                 + '<div style="margin-top:6px;font-size:11px;">If <code>asAssignee_open</code> or '
                 + '<code>asSoloOwner_open</code> is &gt; 0, the filter is missing something. '
                 + 'Send this to the script author.</div></div>';
        }
        body.innerHTML = msg + renderHiddenBucket();
        document.getElementById('trCount').textContent = '';
        renderFocus();
        renderCapacity();
        renderSnoozedFooter();
        return;
    }
    var b = bucketize(state.tasks);
    var html = '';
    html += renderBucket('overdue', 'Overdue',    b.overdue, 'overdue');
    html += renderBucket('today',   'Today',      b.today,   'today');
    html += renderBucket('week',    'This Week',  b.week,    'week');
    html += renderBucket('later',   'Later',      b.later,   'later');
    html += renderBucket('undated', 'No Due Date',b.undated, 'undated');
    html += renderHiddenBucket();
    body.innerHTML = renderDrillInBanner() + html;

    var total = state.tasks.length;
    var open  = total;
    var parts = [open + ' open'];
    if (b.overdue.length) parts.push(b.overdue.length + ' overdue');
    if (state.snoozedCount) parts.push(state.snoozedCount + ' hidden');
    document.getElementById('trCount').textContent = parts.join(' \\u00B7 ');

    renderFocus();
    renderCapacity();
    renderSnoozedFooter();
}

function loadAndRender() {
    var body = document.getElementById('trBody');
    body.innerHTML = '<div class="tr-loading"><span class="tr-spinner"></span>Loading tasks...</div>';
    var params = {};
    if (state.view === 'team-detail' && state.viewAsId) {
        params.view_as_id = state.viewAsId;
    }
    ajax('load_my_tasks', params, function(err, d) {
        if (err || !d || !d.success) {
            body.innerHTML = '<div class="tr-banner-error">'
                          + 'Failed to load tasks: ' + esc((d && d.message) || err) + '</div>';
            return;
        }
        state.tasks = d.tasks || [];
        state.hiddenTasks = d.hiddenTasks || [];
        state.asOf  = d.asOf || '';
        state.diagnosis = d.diagnosis || null;
        state.snoozedCount = d.snoozedCount || 0;
        state.stats = d.stats || {};
        state.focusIds = d.focusIds || [];
        state.contactMethods = d.contactMethods || [];
        state.suggestedContactMethods = d.suggestedContactMethods || [];
        state.contactLookbackDays = d.contactLookbackDays || 182;
        state.orgSettings = d.orgSettings || {};
        applyViewVisibility();
        renderAll();
        var asOf = document.getElementById('trAsOf');
        if (asOf) asOf.textContent = 'As of ' + state.asOf.replace('T', ' ');
    });
}

// ---- View management ----
function applyViewVisibility() {
    // Show / hide the Team toggle based on org settings. The toggle stays
    // hidden until the admin configures a staff org + scope != 'off'.
    var os = state.orgSettings || {};
    var teamEnabled = (os.drillin_scope && os.drillin_scope !== 'off' && os.staff_org_id);
    var toggle = document.getElementById('trViewToggle');
    if (toggle) toggle.style.display = teamEnabled ? 'inline-flex' : 'none';
    // Sync the active button highlight
    var myBtn = document.getElementById('trViewMyBtn');
    var teamBtn = document.getElementById('trViewTeamBtn');
    var activeView = (state.view === 'mine') ? 'mine' : 'team';
    if (myBtn) myBtn.classList.toggle('active', activeView === 'mine');
    if (teamBtn) teamBtn.classList.toggle('active', activeView === 'team');
}

function setView(view) {
    if (view === state.view) return;
    if (view === 'mine') {
        state.view = 'mine';
        state.viewAsId = 0;
        state.viewAsName = '';
        loadAndRender();
    } else if (view === 'team') {
        state.view = 'team';
        state.viewAsId = 0;
        state.viewAsName = '';
        loadTeam();
    }
    applyViewVisibility();
}

function drillInto(peopleId, name) {
    state.view = 'team-detail';
    state.viewAsId = peopleId;
    state.viewAsName = name || '';
    loadAndRender();
}

function backToTeam() {
    state.view = 'team';
    state.viewAsId = 0;
    state.viewAsName = '';
    loadTeam();
}

function loadTeam() {
    // Render the team list. Reuses the existing #trBody area; clears the
    // focus/capacity/snoozed footers since they only apply to "mine".
    var body = document.getElementById('trBody');
    body.innerHTML = '<div class="tr-loading"><span class="tr-spinner"></span>Loading team...</div>';
    document.getElementById('trFocus').innerHTML = '';
    document.getElementById('trCapacity').innerHTML = '';
    document.getElementById('trSnoozedFooter').innerHTML = '';
    applyViewVisibility();

    ajax('load_my_team', {}, function(err, d) {
        if (err || !d || !d.success) {
            body.innerHTML = '<div class="tr-banner-error">'
                          + 'Failed to load team: ' + esc((d && d.message) || err) + '</div>';
            return;
        }
        state.team = d.team || {members: [], permissions: {}};
        renderTeam();
    });
}

function renderTeam() {
    var body = document.getElementById('trBody');
    var team = state.team || {};
    document.getElementById('trCount').textContent = '';

    if (!team.enabled) {
        body.innerHTML = '<div class="tr-team-disabled">'
            + 'Team View is disabled. An admin can enable it in '
            + '<a onclick="TRApp.openSettings();return false;" style="color:#7a5c00;cursor:pointer;text-decoration:underline;">Settings</a> '
            + 'by setting a Staff Involvement and a drill-in scope.</div>';
        return;
    }
    var allMembers = team.members || [];
    if (!allMembers.length) {
        body.innerHTML = '<div class="tr-team-section"><div class="tr-team-empty">'
            + 'No teammates found. Check your staff involvement membership and the drill-in scope in Settings.'
            + '</div></div>';
        return;
    }

    // Toolbar: filter toggle for hiding teammates with zero open tasks.
    var totalCount = allMembers.length;
    var emptyCount = 0;
    for (var ec = 0; ec < allMembers.length; ec++) {
        if (!(allMembers[ec].open || 0)) emptyCount++;
    }
    var members = allMembers;
    if (state.hideEmptyTeam) {
        var filtered = [];
        for (var fi = 0; fi < allMembers.length; fi++) {
            if ((allMembers[fi].open || 0) > 0) filtered.push(allMembers[fi]);
        }
        members = filtered;
    }
    var hideStateLabel = emptyCount === 0
        ? 'Hide teammates with no open tasks'
        : 'Hide teammates with no open tasks (' + emptyCount + ')';
    var toolbarHtml = '<div class="tr-team-toolbar">'
        + '<label class="tr-team-toolbar-toggle">'
        +   '<input type="checkbox" id="trTeamHideEmpty"' + (state.hideEmptyTeam ? ' checked' : '') + ' onchange="TRApp.toggleHideEmptyTeam(this.checked)">'
        +   '<span>' + esc(hideStateLabel) + '</span>'
        + '</label>'
        + '<span class="tr-team-toolbar-meta">' + members.length + ' of ' + totalCount + ' shown</span>'
        + '</div>';

    if (state.hideEmptyTeam && !members.length) {
        body.innerHTML = toolbarHtml
            + '<div class="tr-team-section"><div class="tr-team-empty">'
            + 'All teammates currently have zero open tasks. Uncheck the filter above to see everyone.'
            + '</div></div>';
        return;
    }

    // Group by subgroup when configured AND every member has at least one.
    var grouped = !!team.displayGrouped;
    var groups = {};
    var order = [];
    if (grouped) {
        for (var i = 0; i < members.length; i++) {
            var m = members[i];
            var subs = m.subgroups && m.subgroups.length ? m.subgroups : ['(no department)'];
            for (var s = 0; s < subs.length; s++) {
                var key = subs[s];
                if (!groups[key]) { groups[key] = []; order.push(key); }
                groups[key].push(m);
            }
        }
    } else {
        groups['_'] = members;
        order = ['_'];
    }
    order.sort(function(a, b) {
        if (a === '(no department)') return 1;
        if (b === '(no department)') return -1;
        return a.localeCompare(b);
    });

    var html = '';
    for (var oi = 0; oi < order.length; oi++) {
        var label = order[oi];
        var list = groups[label];
        html += '<div class="tr-team-section">';
        if (grouped) {
            html += '<div class="tr-team-section-header">' + esc(label) + ' &middot; ' + list.length + '</div>';
        }
        for (var li = 0; li < list.length; li++) {
            html += renderTeamCard(list[li]);
        }
        html += '</div>';
    }
    body.innerHTML = toolbarHtml + html;
}

function toggleHideEmptyTeam(checked) {
    state.hideEmptyTeam = !!checked;
    // Optimistic: re-render immediately, persist in background.
    renderTeam();
    ajax('save_user_pref', {
        key: 'hideEmptyTeam',
        value: state.hideEmptyTeam ? 'true' : 'false'
    }, function(err, d) {
        if (err || !d || !d.success) {
            // Don't revert -- the local view is what the user just clicked --
            // but warn that the pref didn't stick.
            showToast('Could not save preference', 'error');
        }
    });
}

function renderTeamCard(m) {
    var open = m.open || 0, overdue = m.overdue || 0, today = m.completedToday || 0;
    var lastLabel = '';
    if (m.lastCompleted) {
        lastLabel = 'Last completed ' + m.lastCompleted;
    } else {
        lastLabel = 'No recent completions';
    }
    var safeName = String(m.name || '').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    var subInline = '';
    if (m.subgroups && m.subgroups.length) {
        var sgs = [];
        for (var i = 0; i < m.subgroups.length; i++) {
            sgs.push('<span class="tr-chip">' + esc(m.subgroups[i]) + '</span>');
        }
        subInline = ' ' + sgs.join(' ');
    }
    return '<div class="tr-team-card" onclick="TRApp.drillInto(' + m.peopleId + ',&quot;' + safeName + '&quot;)">'
         + '<div class="tr-team-card-main">'
         +   '<div class="tr-team-name">' + esc(m.name) + subInline + '</div>'
         +   '<div class="tr-team-meta">' + esc(lastLabel) + '</div>'
         + '</div>'
         + '<div class="tr-team-stats">'
         +   '<div class="tr-team-stat"><div class="tr-team-stat-num' + (open === 0 ? ' zero' : '') + '">' + open + '</div><div class="tr-team-stat-label">Open</div></div>'
         +   '<div class="tr-team-stat"><div class="tr-team-stat-num' + (overdue > 0 ? ' overdue' : ' zero') + '">' + overdue + '</div><div class="tr-team-stat-label">Overdue</div></div>'
         +   '<div class="tr-team-stat"><div class="tr-team-stat-num' + (today === 0 ? ' zero' : '') + '">' + today + '</div><div class="tr-team-stat-label">Today</div></div>'
         + '</div>'
         + '</div>';
}

function renderDrillInBanner() {
    // Inserts a banner at the top of trBody when we're in drill-in mode.
    if (state.view !== 'team-detail' || !state.viewAsId) return '';
    return '<div class="tr-drillin-banner">'
         +   '<span>Viewing tasks for <strong>' + esc(state.viewAsName || ('#' + state.viewAsId)) + '</strong></span>'
         +   '<button onclick="TRApp.backToTeam()">&larr; Back to team</button>'
         + '</div>';
}

function bootstrap() {
    attachReassignSearch();
    ajax('who_am_i', {}, function(err, d) {
        var sub = document.getElementById('trSub');
        if (err || !d || !d.success) {
            sub.textContent = 'Could not identify the current user';
            return;
        }
        state.user = d.user;
        // Seed user prefs from server.
        var prefs = (state.user && state.user.prefs) || {};
        state.hideEmptyTeam = !!prefs.hideEmptyTeam;
        sub.textContent = 'Signed in as ' + (state.user.name || state.user.username || 'unknown');
        loadAndRender();
    });
}

// ---- Toast ----
var toastTimer = null;
function showToast(msg, kind) {
    kind = kind || 'info';
    var el = document.getElementById('trToast');
    if (!el) {
        el = document.createElement('div');
        el.id = 'trToast';
        document.body.appendChild(el);
    }
    el.className = 'tr-toast tr-toast-' + kind;
    el.textContent = msg;
    // Force reflow so the transition replays even on rapid clicks
    void el.offsetWidth;
    el.classList.add('show');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function() {
        el.classList.remove('show');
    }, 2400);
}

// ---- Snooze submenu toggle ----
function closeAllSnoozeMenus() {
    var menus = document.querySelectorAll('.tr-snooze-menu.open');
    for (var i = 0; i < menus.length; i++) menus[i].classList.remove('open');
}

function toggleSnoozeMenu(taskId, ev) {
    if (ev) {
        ev.stopPropagation();
        ev.preventDefault();
    }
    var menu = document.getElementById('trSnoozeMenu' + taskId);
    if (!menu) return;
    var wasOpen = menu.classList.contains('open');
    closeAllSnoozeMenus();
    if (!wasOpen) menu.classList.add('open');
}

document.addEventListener('click', function(e) {
    // Click anywhere outside an open snooze menu closes them.
    if (!e.target.closest || !e.target.closest('.tr-snooze-wrap')) {
        closeAllSnoozeMenus();
    }
});

// ---- Row removal + bucket count update ----
function removeRowAndReflow(taskId) {
    // Drop the task from state and re-render. Simpler than DOM splicing and
    // ensures bucket counts + empty messages stay accurate.
    var keep = [];
    for (var i = 0; i < state.tasks.length; i++) {
        if (state.tasks[i].id !== taskId) keep.push(state.tasks[i]);
    }
    state.tasks = keep;
    renderAll();
}

// ---- Action handlers ----
function completeTask(taskId) {
    if (!confirm('Mark this task complete?')) return;
    closeAllSnoozeMenus();
    var row = document.getElementById('trRow' + taskId);
    var btns = row ? row.querySelectorAll('button') : [];
    for (var i = 0; i < btns.length; i++) btns[i].disabled = true;
    ajax('complete_task', {task_id: taskId}, function(err, d) {
        if (err || !d || !d.success) {
            showToast('Complete failed: ' + ((d && d.message) || err), 'error');
            for (var j = 0; j < btns.length; j++) btns[j].disabled = false;
            return;
        }
        showToast('Task completed', 'success');
        removeRowAndReflow(taskId);
    });
}

function snoozeTask(taskId, days) {
    closeAllSnoozeMenus();
    var row = document.getElementById('trRow' + taskId);
    var btns = row ? row.querySelectorAll('button') : [];
    for (var i = 0; i < btns.length; i++) btns[i].disabled = true;
    ajax('snooze_task', {task_id: taskId, days: days}, function(err, d) {
        if (err || !d || !d.success) {
            showToast('Snooze failed: ' + ((d && d.message) || err), 'error');
            for (var j = 0; j < btns.length; j++) btns[j].disabled = false;
            return;
        }
        // Snooze is local triage state -- the task disappears from view
        // until d.snoozedUntil. Its real DueDate is unchanged in TouchPoint.
        var label = days === 1 ? 'tomorrow'
                  : days === 7 ? 'next week'
                  : days + ' days';
        showToast('Hidden for ' + label, 'success');
        state.snoozedCount = (state.snoozedCount || 0) + 1;
        removeRowAndReflow(taskId);
    });
}

function togglePeekHidden() {
    // Pure UI toggle -- shows the hidden bucket without modifying any flag.
    // No server call needed since the server already returned hiddenTasks.
    state.peekingHidden = !state.peekingHidden;
    renderAll();
}

function unhideTask(taskId) {
    // Per-row "Unhide" -- moves this one task back into the main list.
    var row = document.getElementById('trRow' + taskId);
    var btns = row ? row.querySelectorAll('button') : [];
    for (var i = 0; i < btns.length; i++) btns[i].disabled = true;
    ajax('unsnooze_task', {task_id: taskId}, function(err, d) {
        if (err || !d || !d.success) {
            showToast('Unhide failed: ' + ((d && d.message) || err), 'error');
            for (var j = 0; j < btns.length; j++) btns[j].disabled = false;
            return;
        }
        // Move task from hiddenTasks back into tasks, then re-render so it
        // lands in the right bucket. No re-fetch needed.
        for (var k = 0; k < state.hiddenTasks.length; k++) {
            if (state.hiddenTasks[k].id === taskId) {
                var t = state.hiddenTasks[k];
                delete t.hiddenUntil;
                state.tasks.push(t);
                state.hiddenTasks.splice(k, 1);
                break;
            }
        }
        state.snoozedCount = state.hiddenTasks.length;
        showToast('Task unhidden', 'success');
        renderAll();
    });
}

function revealAllHidden() {
    if (!confirm('Reveal all hidden tasks?\\n\\nThis clears every hidden flag at once. Hiding only affects Task Runner -- the tasks are still active in TouchPoint either way.')) return;
    ajax('clear_all_snoozed', {}, function(err, d) {
        if (err || !d || !d.success) {
            showToast('Failed to reveal', 'error');
            return;
        }
        showToast('All hidden tasks revealed', 'info');
        state.peekingHidden = false;
        loadAndRender();
    });
}

function unhideGroup(taskIdsCsv) {
    // Per-sub-section bulk unhide. Fires one server call per task -- could be
    // a single batch action later, but the per-call cost is small enough
    // (~50ms each over the same connection) that this is fine for v1.
    var ids = taskIdsCsv.split(',').map(function(s) { return parseInt(s, 10); })
                                   .filter(function(n) { return !isNaN(n); });
    if (!ids.length) return;
    var pending = ids.length;
    var failures = 0;
    ids.forEach(function(id) {
        ajax('unsnooze_task', {task_id: id}, function(err, d) {
            if (err || !d || !d.success) failures++;
            else {
                // Mirror locally so the UI updates without a full reload.
                for (var k = 0; k < state.hiddenTasks.length; k++) {
                    if (state.hiddenTasks[k].id === id) {
                        var t = state.hiddenTasks[k];
                        delete t.hiddenUntil;
                        state.tasks.push(t);
                        state.hiddenTasks.splice(k, 1);
                        break;
                    }
                }
            }
            pending--;
            if (pending === 0) {
                state.snoozedCount = state.hiddenTasks.length;
                if (failures) {
                    showToast(ids.length - failures + ' unhidden, ' + failures + ' failed', 'error');
                } else {
                    showToast(ids.length + ' task' + (ids.length === 1 ? '' : 's') + ' unhidden', 'success');
                }
                renderAll();
            }
        });
    });
}

// ---- Modal helpers ----
function openModal(id) {
    var el = document.getElementById(id);
    if (el) el.classList.add('open');
}
function closeModal(id) {
    var el = document.getElementById(id);
    if (el) el.classList.remove('open');
}

function findTask(taskId) {
    for (var i = 0; i < state.tasks.length; i++) {
        if (state.tasks[i].id === taskId) return state.tasks[i];
    }
    for (var j = 0; j < state.hiddenTasks.length; j++) {
        if (state.hiddenTasks[j].id === taskId) return state.hiddenTasks[j];
    }
    return null;
}

function describeTaskShort(t) {
    if (!t) return '';
    var snippet = (t.instructions || '(no instructions)').substring(0, 80);
    if ((t.instructions || '').length > 80) snippet += '...';
    if (t.aboutName) snippet += ' \\u2014 ' + t.aboutName;
    return snippet;
}

// ---- Log Contact ----
function openLog(taskId) {
    var t = findTask(taskId);
    if (!t) return;
    if (!t.aboutPersonId) {
        showToast('No about-person on this task to log against', 'error');
        return;
    }
    state.activeTaskId = taskId;
    document.getElementById('trLogTaskContext').textContent = describeTaskShort(t);
    document.getElementById('trLogNote').value = '';

    var grid = document.getElementById('trLogMethods');
    var methods = state.contactMethods || [];
    if (!methods.length) {
        grid.innerHTML = '<div class="tr-method-warn">No contact methods configured. Ask your admin to set them up in ProspectBuilder or Task Runner settings.</div>';
    } else {
        var html = '';
        var unconfigured = 0;
        for (var i = 0; i < methods.length; i++) {
            var m = methods[i];
            var code = m.code || '?';
            var label = m.label || code;
            if (!m.keywordId) unconfigured++;
            html += '<button class="tr-method-btn" onclick="TRApp.submitLog(\\'' + esc(code) + '\\')">'
                  + '<span class="tr-method-code">' + esc(code) + '</span>'
                  + '<span>' + esc(label) + '</span>'
                  + '</button>';
        }
        grid.innerHTML = html;
        var warn = document.getElementById('trLogWarn');
        if (unconfigured) {
            warn.innerHTML = '<div class="tr-method-warn">' + unconfigured
                           + ' method(s) have no TouchPoint keyword bound. Note will still log, but it won\\'t show on contact-history dashboards. Configure via Settings.</div>';
        } else {
            warn.innerHTML = '';
        }
    }
    openModal('trLogModal');
}

function submitLog(code) {
    var taskId = state.activeTaskId;
    if (!taskId) return;
    var note = document.getElementById('trLogNote').value.trim();
    ajax('log_contact_attempt', {task_id: taskId, code: code, note: note}, function(err, d) {
        if (err || !d || !d.success) {
            showToast('Log failed: ' + ((d && d.message) || err), 'error');
            return;
        }
        closeModal('trLogModal');
        showToast('Contact logged (' + code + ')', 'success');
    });
}

// ---- Reassign ----
var reassignTimer = null;
var reassignPick = null;
function openReassign(taskId) {
    var t = findTask(taskId);
    if (!t) return;
    state.activeTaskId = taskId;
    reassignPick = null;
    document.getElementById('trReassignTaskContext').textContent = describeTaskShort(t);
    document.getElementById('trReassignSearch').value = '';
    document.getElementById('trReassignResults').style.display = 'none';
    document.getElementById('trReassignResults').innerHTML = '';
    document.getElementById('trReassignNote').value = '';
    document.getElementById('trReassignSubmit').disabled = true;
    openModal('trReassignModal');
    setTimeout(function() {
        var i = document.getElementById('trReassignSearch');
        if (i) i.focus();
    }, 50);
}

function attachReassignSearch() {
    var inp = document.getElementById('trReassignSearch');
    if (!inp || inp.__attached) return;
    inp.__attached = true;
    inp.addEventListener('input', function() {
        if (reassignTimer) clearTimeout(reassignTimer);
        var term = inp.value.trim();
        if (term.length < 2) {
            document.getElementById('trReassignResults').style.display = 'none';
            reassignPick = null;
            document.getElementById('trReassignSubmit').disabled = true;
            return;
        }
        reassignTimer = setTimeout(function() {
            ajax('search_people', {term: term}, function(err, d) {
                var box = document.getElementById('trReassignResults');
                if (err || !d || !d.success) {
                    box.style.display = 'none';
                    return;
                }
                var people = d.people || [];
                if (!people.length) {
                    box.innerHTML = '<div class="tr-people-result" style="cursor:default;color:var(--tr-muted);">No matches</div>';
                    box.style.display = '';
                    return;
                }
                var html = '';
                for (var i = 0; i < people.length; i++) {
                    var p = people[i];
                    html += '<div class="tr-people-result" data-pid="' + p.peopleId + '" data-name="'
                          + esc(p.name).replace(/"/g, '&quot;') + '">' + esc(p.name) + '</div>';
                }
                box.innerHTML = html;
                box.style.display = '';
                // Click handlers
                var rows = box.querySelectorAll('.tr-people-result');
                for (var j = 0; j < rows.length; j++) {
                    rows[j].addEventListener('click', function(e) {
                        var pid = parseInt(this.getAttribute('data-pid'), 10);
                        var name = this.getAttribute('data-name');
                        if (!pid) return;
                        reassignPick = {peopleId: pid, name: name};
                        inp.value = name;
                        box.style.display = 'none';
                        document.getElementById('trReassignSubmit').disabled = false;
                    });
                }
            });
        }, 250);
    });
}

function submitReassign() {
    var taskId = state.activeTaskId;
    if (!taskId || !reassignPick) return;
    var note = document.getElementById('trReassignNote').value.trim();
    var btn = document.getElementById('trReassignSubmit');
    btn.disabled = true; btn.textContent = 'Reassigning...';
    ajax('reassign_task', {
        task_id: taskId,
        new_assignee_id: reassignPick.peopleId,
        note: note
    }, function(err, d) {
        btn.textContent = 'Reassign';
        if (err || !d || !d.success) {
            btn.disabled = false;
            showToast('Reassign failed: ' + ((d && d.message) || err), 'error');
            return;
        }
        closeModal('trReassignModal');
        showToast('Reassigned to ' + reassignPick.name, 'success');
        // The task is no longer on this user's plate, so remove from view.
        removeRowAndReflow(taskId);
    });
}

// ---- Settings ----
var settingsKeywords = null;     // cached keyword list for the dropdown
var settingsDraft = null;        // editable copy of the methods list

function openSettings() {
    // Start from the current contact methods (deep copy so cancel doesn't mutate state).
    settingsDraft = JSON.parse(JSON.stringify(state.contactMethods || []));
    // Pre-fill the lookback input with the current value.
    var lb = document.getElementById('trSettingsLookback');
    if (lb) lb.value = state.contactLookbackDays || 182;
    // Pre-fill org settings.
    var os = state.orgSettings || {};
    var staffInput = document.getElementById('trSettingsStaffOrg');
    var staffIdInput = document.getElementById('trSettingsStaffOrgId');
    var staffNameInput = document.getElementById('trSettingsStaffOrgName');
    var staffSel = document.getElementById('trSettingsStaffOrgSelected');
    if (staffInput) staffInput.value = os.staff_org_name || '';
    if (staffIdInput) staffIdInput.value = os.staff_org_id || '';
    if (staffNameInput) staffNameInput.value = os.staff_org_name || '';
    if (staffSel) staffSel.textContent = os.staff_org_id
        ? 'Saved: ' + (os.staff_org_name || '(#' + os.staff_org_id + ')')
        : '';

    var scope = os.drillin_scope || 'off';
    var scopeRadios = document.getElementsByName('trSettingsScope');
    for (var i = 0; i < scopeRadios.length; i++) {
        scopeRadios[i].checked = (scopeRadios[i].value === scope);
    }
    var actions = os.drillin_actions || 'reassign';
    var actionRadios = document.getElementsByName('trSettingsActions');
    for (var j = 0; j < actionRadios.length; j++) {
        actionRadios[j].checked = (actionRadios[j].value === actions);
    }
    var grpEl = document.getElementById('trSettingsGroupDept');
    if (grpEl) grpEl.checked = !!os.team_display_grouped;

    attachStaffOrgSearch();
    // Lazy-load keyword list once per page session.
    if (settingsKeywords === null) {
        ajax('list_keywords', {}, function(err, d) {
            if (d && d.success) settingsKeywords = d.keywords || [];
            else settingsKeywords = [];
            renderSettingsRows();
        });
    } else {
        renderSettingsRows();
    }
    openModal('trSettingsModal');
}

// Type-ahead search for the Staff Involvement picker.
var staffOrgSearchTimer = null;
function attachStaffOrgSearch() {
    var inp = document.getElementById('trSettingsStaffOrg');
    if (!inp || inp.__attached) return;
    inp.__attached = true;
    inp.addEventListener('input', function() {
        if (staffOrgSearchTimer) clearTimeout(staffOrgSearchTimer);
        var term = inp.value.trim();
        var box = document.getElementById('trSettingsOrgResults');
        if (!box) return;
        if (term.length < 2) {
            box.style.display = 'none';
            return;
        }
        staffOrgSearchTimer = setTimeout(function() {
            ajax('search_involvements', {term: term}, function(err, d) {
                if (err || !d || !d.success) { box.style.display = 'none'; return; }
                var orgs = d.involvements || [];
                if (!orgs.length) {
                    box.innerHTML = '<div class="tr-people-result" style="cursor:default;color:var(--tr-muted);">No matches</div>';
                    box.style.display = '';
                    return;
                }
                var html = '';
                for (var i = 0; i < orgs.length; i++) {
                    var o = orgs[i];
                    var sub = o.program + (o.division ? ' / ' + o.division : '');
                    html += '<div class="tr-people-result" data-oid="' + o.orgId
                          + '" data-oname="' + esc(o.name).replace(/"/g, '&quot;') + '">'
                          + esc(o.name)
                          + (sub.trim() ? '<div style="font-size:11px;color:var(--tr-muted);">' + esc(sub) + '</div>' : '')
                          + '</div>';
                }
                box.innerHTML = html;
                box.style.display = '';
                var rows = box.querySelectorAll('.tr-people-result');
                for (var j = 0; j < rows.length; j++) {
                    rows[j].addEventListener('click', function() {
                        var oid = this.getAttribute('data-oid');
                        var oname = this.getAttribute('data-oname');
                        document.getElementById('trSettingsStaffOrgId').value = oid;
                        document.getElementById('trSettingsStaffOrgName').value = oname;
                        document.getElementById('trSettingsStaffOrg').value = oname;
                        document.getElementById('trSettingsStaffOrgSelected').textContent = '\\u2713 Selected: ' + oname;
                        box.style.display = 'none';
                    });
                }
            });
        }, 220);
    });
    // Clicking outside closes the dropdown
    document.addEventListener('click', function(e) {
        var box = document.getElementById('trSettingsOrgResults');
        if (!box) return;
        if (e.target === inp || (box.contains && box.contains(e.target))) return;
        box.style.display = 'none';
    });
}

function renderSettingsRows() {
    var tbody = document.getElementById('trSettingsRows');
    if (!tbody) return;
    var rows = '';
    for (var i = 0; i < settingsDraft.length; i++) {
        var m = settingsDraft[i];
        var keywordOpts = '<option value="">(none -- log without keyword)</option>';
        if (settingsKeywords) {
            for (var k = 0; k < settingsKeywords.length; k++) {
                var kw = settingsKeywords[k];
                var sel = (m.keywordId && String(m.keywordId) === String(kw.keywordId)) ? ' selected' : '';
                keywordOpts += '<option value="' + kw.keywordId + '"' + sel + '>'
                             + esc(kw.description) + '</option>';
            }
        }
        rows += '<tr style="border-bottom: 1px solid #f0f0f0;">'
              + '<td style="padding: 6px;"><input type="text" maxlength="2" value="' + esc(m.code || '') + '"'
              +    ' style="width: 50px; text-align: center; font-weight: 700;" data-idx="' + i + '" data-field="code"></td>'
              + '<td style="padding: 6px;"><input type="text" value="' + esc(m.label || '') + '"'
              +    ' style="width: 120px;" data-idx="' + i + '" data-field="label"></td>'
              + '<td style="padding: 6px;"><select style="width: 100%;" data-idx="' + i + '" data-field="keywordId">'
              +    keywordOpts
              + '</select></td>'
              + '<td style="padding: 6px; text-align: center;">'
              +   '<button class="tr-btn tr-btn-secondary" onclick="TRApp.removeMethod(' + i + ')" style="padding: 4px 10px;">Remove</button>'
              + '</td>'
              + '</tr>';
    }
    if (!settingsDraft.length) {
        // Honest empty state: don't pre-suppose any naming. Offer a one-click
        // path to the conventional starters and a "from scratch" path.
        rows = '<tr><td colspan="4" style="padding: 18px; text-align: center;">'
             + '<div style="color: var(--tr-muted); margin-bottom: 10px;">'
             +   'No contact methods configured yet. Each church chooses its own short codes '
             +   '(e.g. <strong>P</strong>hone, <strong>E</strong>mail, <strong>V</strong>isit) and binds them to TouchPoint keywords.'
             + '</div>'
             + '<button class="tr-btn tr-btn-secondary" style="margin-right:6px;" onclick="TRApp.addMethod()">+ Add your first method</button>'
             + '<button class="tr-btn tr-btn-secondary" onclick="TRApp.applySuggestedMethods()">Use suggested starters</button>'
             + '</td></tr>';
    }
    tbody.innerHTML = rows;

    // Wire input listeners
    var inputs = tbody.querySelectorAll('input, select');
    for (var i2 = 0; i2 < inputs.length; i2++) {
        inputs[i2].addEventListener('change', onSettingsFieldChange);
        if (inputs[i2].tagName === 'INPUT') inputs[i2].addEventListener('input', onSettingsFieldChange);
    }
}

function onSettingsFieldChange(e) {
    var idx = parseInt(e.target.getAttribute('data-idx'), 10);
    var field = e.target.getAttribute('data-field');
    if (isNaN(idx) || !settingsDraft[idx]) return;
    var val = e.target.value;
    if (field === 'keywordId') {
        settingsDraft[idx].keywordId = val ? parseInt(val, 10) : null;
        // Mirror the keyword description so the saved methods carry it
        if (settingsKeywords && val) {
            for (var k = 0; k < settingsKeywords.length; k++) {
                if (String(settingsKeywords[k].keywordId) === String(val)) {
                    settingsDraft[idx].keyword = settingsKeywords[k].description;
                    break;
                }
            }
        } else {
            settingsDraft[idx].keyword = '';
        }
    } else {
        settingsDraft[idx][field] = val;
    }
}

function addMethod() {
    settingsDraft.push({code: '', label: '', keyword: '', keywordId: null});
    renderSettingsRows();
}

function applySuggestedMethods() {
    // Populate the draft from server-provided suggested starters. Admin still
    // has to bind keywords + Save -- nothing is auto-applied.
    var seed = state.suggestedContactMethods || [];
    settingsDraft = JSON.parse(JSON.stringify(seed));
    renderSettingsRows();
}

function removeMethod(idx) {
    settingsDraft.splice(idx, 1);
    renderSettingsRows();
}

function saveSettings() {
    var lb = document.getElementById('trSettingsLookback');
    var lookback = lb && lb.value ? parseInt(lb.value, 10) : '';
    if (lookback !== '' && (isNaN(lookback) || lookback < 7)) {
        showToast('Lookback must be at least 7 days', 'error');
        return;
    }

    // Collect org settings. These save through a separate action so the
    // two save paths can be retried independently if one fails.
    var staffIdEl = document.getElementById('trSettingsStaffOrgId');
    var staffNameEl = document.getElementById('trSettingsStaffOrgName');
    var staffId = staffIdEl ? (staffIdEl.value || '') : '';
    var staffName = staffNameEl ? (staffNameEl.value || '') : '';

    var scopeChecked = document.querySelector('input[name="trSettingsScope"]:checked');
    var scope = scopeChecked ? scopeChecked.value : 'off';
    var actionChecked = document.querySelector('input[name="trSettingsActions"]:checked');
    var actions = actionChecked ? actionChecked.value : 'reassign';
    var grpEl = document.getElementById('trSettingsGroupDept');
    var groupDept = grpEl && grpEl.checked ? 'true' : 'false';

    // Two sequential saves so we can return a single toast to the user.
    ajax('save_contact_methods', {
        methods_json: JSON.stringify(settingsDraft),
        lookback_days: lookback === '' ? '' : String(lookback)
    }, function(err, d) {
        if (err || !d || !d.success) {
            showToast('Save failed: ' + ((d && d.message) || err), 'error');
            return;
        }
        state.contactMethods = d.methods || [];
        if (d.lookbackDays) state.contactLookbackDays = d.lookbackDays;

        ajax('save_org_settings', {
            staff_org_id: staffId,
            staff_org_name: staffName,
            drillin_scope: scope,
            drillin_actions: actions,
            team_display_grouped: groupDept
        }, function(err2, d2) {
            if (err2 || !d2 || !d2.success) {
                showToast('Saved contact settings, but staff settings failed: '
                          + ((d2 && d2.message) || err2), 'error');
                return;
            }
            state.orgSettings = d2.orgSettings || {};
            applyViewVisibility();
            closeModal('trSettingsModal');
            showToast('Settings saved -- refresh to apply new lookback', 'success');
        });
    });
}

window.TRApp = {
    reload: loadAndRender,
    complete: completeTask,
    snooze: snoozeTask,
    toggleSnoozeMenu: toggleSnoozeMenu,
    togglePeek: togglePeekHidden,
    unhide: unhideTask,
    unhideGroup: unhideGroup,
    revealAll: revealAllHidden,
    openLog: openLog,
    submitLog: submitLog,
    openReassign: openReassign,
    submitReassign: submitReassign,
    openSettings: openSettings,
    addMethod: addMethod,
    applySuggestedMethods: applySuggestedMethods,
    removeMethod: removeMethod,
    saveSettings: saveSettings,
    setView: setView,
    drillInto: drillInto,
    backToTeam: backToTeam,
    toggleHideEmptyTeam: toggleHideEmptyTeam,
    closeModal: closeModal
};

bootstrap();

})();
</script>
'''
