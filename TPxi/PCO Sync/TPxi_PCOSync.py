#roles=Edit
#----------------------------------------------------------------------
# TPxi_PCOSync.py
#
# Planning Center Online (PCO) -> TouchPoint attendance sync.
#
# Worship teams schedule and take attendance in PCO Services. TouchPoint
# is the authoritative people database, but attendance data lives in two
# places and worship admins end up double-entering. This script gives them
# one place inside TouchPoint to pull recent PCO plans, match attendees
# to TouchPoint people (one-time mapping persisted as an Extra Value),
# and write attendance back to the corresponding TouchPoint involvement
# in one click.
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
#   Single .py file SPA. Python AJAX handlers for POST, HTML SPA for GET.
#   Settings + mappings + audit log stored in Special Content (Text).
#   PCO Person ID is stored on each TouchPoint person as a text Extra Value
#   so the mapping is per-person and survives across sessions.
#
# Storage Keys:
#   PCOSync_Settings        - PCO app_id, secret (PAT), last-sync stamps
#   PCOSync_OrgMappings     - {pcoServiceTypeId: tpOrgId}
#   PCOSync_Log_<YYYYMM>    - per-month audit log of writes
#
# Extra Value (per TouchPoint Person):
#   PCO_PersonId  (text)    - links a TP person to their PCO person record
#
# CSS Prefix: pco-
# Root Class: .pco-root
#
# Reference:
#   TPxi_AttendanceMarkings.py  - post-verify write pattern, SPA dispatch
#   TPxi_PaymentManager.py      - verified-apply-update pattern
#   TPxi_RollSheet.py           - safe_json transliteration pattern
#
# Changelog:
#   1.0.3  (2026-06-05)  First production release.
#
#                        Settings & connection
#                          - PCO Personal Access Token auth (App ID +
#                            Secret), saved to Special Content, with a
#                            Test Connection button.
#                          - Verify-after-write on every settings save
#                            so silent permission failures surface
#                            immediately.
#
#                        Sync Mappings tab -- three mapping types
#                          - All People: a singleton mapping that
#                            walks the entire PCO People directory
#                            into one TP involvement.
#                          - Service Type Mappings: one PCO Service
#                            Type -> one umbrella TP involvement.
#                            Optional layers: teams-as-subgroups,
#                            per-plan attendance.
#                          - Team Mappings: one PCO Team -> one TP
#                            involvement. Optional layers: positions-
#                            as-subgroups, per-plan attendance.
#                          - Per-mapping toggles for autoAddMember,
#                            teams/positions-as-subgroups, and per-
#                            plan attendance, all inline + auto-saved.
#                          - Inline "Check PCO positions" diagnostic
#                            on Team Mapping rows.
#
#                        Sync Dashboard tab
#                          - Cards for All People, Service Type,
#                            Team. Each shows last-synced pill and
#                            (when scheduled) a next-run pill.
#                          - Per-plan cards collapse multiple plans
#                            for the same date under SUNDAY-style day
#                            headers; auto-scrolls to today.
#                          - Preview & Sync modal shows match counts,
#                            sync-mode banner, mirror-removal banner
#                            (with red "will be removed" pill), and
#                            per-attendee rows with manual search-and-
#                            link for anything unmatched.
#                          - Confirm dialog spells out every action
#                            (add, drop, subgroup write, subgroup
#                            drop) before sync.
#
#                        PCO -> TP sync writes
#                          - Roster add: JoinOrg for every matched
#                            PCO person not already on the TP
#                            involvement.
#                          - Subgroup add: AddSubGroup for each PCO
#                            position (team) or team (service type)
#                            the person holds.
#                          - Per-plan attendance: creates / finds the
#                            TP meeting on the plan date and marks
#                            Confirmed attendees Present.
#                          - Mirror removal: PCO is source of truth.
#                            TP members whose PCO_PersonId is no
#                            longer in scope get RemoveFromOrg'd;
#                            subgroup memberships whose name matches
#                            a current PCO position/team but the
#                            person no longer holds get
#                            RemoveSubGroup'd. Manually-added
#                            members (no PCO link) and unrelated
#                            subgroups are left alone.
#                          - PCO team-level assignment endpoint with
#                            per-position fallback so subgroups sync
#                            even when one endpoint flakes.
#
#                        People Matching tab
#                          - Proposed Matches review: scores name +
#                            email + birthdate signals against TP for
#                            every unmatched PCO record. Tiered
#                            Strong / Medium / Weak. Per-row Apply or
#                            Skip Forever; bulk Apply for high-
#                            confidence tier. Scoped (subset from a
#                            preview) or full-directory walk.
#                            Client-side cache so tier/search/
#                            pagination changes don't re-walk PCO.
#                          - Verify Person Link: search TP by name/
#                            email, see the current PCO_PersonId
#                            side-by-side with PCO's record (with red
#                            cells where they disagree and a one-line
#                            verdict), Unlink or Replace with another
#                            PCO person.
#                          - Pending Data Reviews queue for field
#                            diffs flagged by Person Data Sync.
#
#                        Person Data Sync (Settings tab)
#                          - Per-field directional rules (PCO -> TP).
#                            Each field independently: off, auto-
#                            apply, or queue-for-review. Defaults
#                            off so TP stays authoritative unless
#                            opted in.
#                          - Apply queue surfaces on People Matching
#                            -> Pending Data Reviews.
#
#                        Scheduled Sync
#                          - One-click install adds a managed block
#                            to TouchPoint's ScheduledTasks special
#                            content (matches ProspectBuilder
#                            pattern). Per-mapping schedule editor
#                            is gated on global install both client-
#                            and server-side.
#                          - Per-mapping: Daily or Weekly, day-of-
#                            week + hour, notify TouchPoint user
#                            (typeahead picker with name / username /
#                            email search), include-issues toggle.
#                          - Scheduler runner walks all mappings each
#                            invocation, fires anything whose (day,
#                            hour) match now and hasn't run this hour
#                            slot.
#                          - Each fire syncs fully server-side and
#                            emails the configured user: summary
#                            counts (joined, already, subgroup
#                            writes, members removed, stale
#                            subgroups removed), optional issues
#                            list (unmatched, ambiguous, warnings),
#                            and a link back to PCO Sync.
#
#                        Storage migration
#                          - First v3 POST silently migrates any
#                            legacy v2 PCOSync_OrgMappings rows into
#                            PCOSync_PeopleMappings, preserving toggles.
#
#                        Diagnostics & safety
#                          - Verify-after-write on every settings /
#                            mapping / install write.
#                          - Audit log in PCOSync_Log_YYYYMM with per-
#                            sync counters (joined, dropped, subgroup
#                            adds/drops, failures, scheduler runs,
#                            mapping edits, link write/unlink).
#                          - Dashboard health panel surfaces broken
#                            mappings (deleted PCO resource, archived
#                            TP org) before staff hit Sync.
#                          - SQL helpers wrap large-table walks with
#                            NOLOCK + WHERE filters per CLAUDE.md
#                            performance rules.
#----------------------------------------------------------------------

import json
import datetime
import re
import base64

model.Header = 'PCO Sync'

# --- Version / Auto-update -------------------------------------------
APP_VERSION = '1.0.3'
DC_SCRIPT_ID = 'TPxi_PCOSync'

# v3.3.1: scheduler install (matches ProspectBuilder pattern). We
# inject a marker-delimited block into TouchPoint's ScheduledTasks
# special content. When ScheduledTasks fires (TP runs it hourly /
# nightly / however the admin configured the global cron), our block
# sets Data.scheduler='true' and CallScripts back into TPxi_PCOSync,
# which our top-of-file scheduler check picks up and routes to
# handle_run_scheduled_syncs.
_SCHED_MARKER_START = "# >>> TPxi_PCOSync schedule start (managed by app, do not edit) >>>"
_SCHED_MARKER_END   = "# <<< TPxi_PCOSync schedule end <<<"
_SCHED_CONTENT_SLOT = "ScheduledTasks"
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'

# --- Storage keys ----------------------------------------------------
SETTINGS_KEY = 'PCOSync_Settings'
ORG_MAPPINGS_KEY = 'PCOSync_OrgMappings'
TEAM_MAPPINGS_KEY = 'PCOSync_TeamMappings'  # v2.0+: per-Team mappings (Team Sync mode)
PEOPLE_MAPPINGS_KEY = 'PCOSync_PeopleMappings'  # v2.1+: per-Service-Type umbrella (People Sync mode)
ALL_PEOPLE_MAPPING_KEY = 'PCOSync_AllPeopleMapping'  # v2.2+: singleton PCO People directory -> one TP involvement
ALL_PEOPLE_SKIP_KEY = 'PCOSync_AllPeopleSkip'  # v2.5+: PCO Person IDs known to have no TP equivalent
PERSON_SYNC_RULES_KEY = 'PCOSync_PersonSyncRules'
PERSON_PENDING_KEY = 'PCOSync_PendingPersonChanges'
LOG_KEY_PREFIX = 'PCOSync_Log_'   # suffixed with YYYYMM

# v1 scope: PCO -> TP for these three. TP -> PCO and phone are v1.1 (need
# direct PCO People API writes / sub-resource fetches). Mapping: PCO field
# name on the team_members include=person object -> TouchPoint UpdatePerson
# field name.
PERSON_SYNC_FIELDS = [
    {'key': 'first_name', 'tpField': 'FirstName', 'label': 'First Name'},
    {'key': 'last_name',  'tpField': 'LastName',  'label': 'Last Name'},
    {'key': 'email',      'tpField': 'EmailAddress', 'label': 'Email'},
]

def default_person_sync_rules():
    """Default = no syncing. TP is authoritative until the admin opts in."""
    rules = {}
    for f in PERSON_SYNC_FIELDS:
        rules[f['key']] = {'direction': 'none', 'mode': 'review'}
    return rules

def normalize_person_sync_rules(raw):
    """Coerce a loaded rules dict to the canonical shape so missing fields
    or older saves still work."""
    rules = default_person_sync_rules()
    if isinstance(raw, dict):
        for f in PERSON_SYNC_FIELDS:
            r = raw.get(f['key'])
            if isinstance(r, dict):
                d = str(r.get('direction', 'none')).lower()
                if d not in ('none', 'pco_to_tp', 'tp_to_pco'):
                    d = 'none'
                m = str(r.get('mode', 'review')).lower()
                if m not in ('auto', 'review'):
                    m = 'review'
                rules[f['key']] = {'direction': d, 'mode': m}
    return rules

def person_sync_rules_active(rules):
    """True if at least one field has direction != none."""
    if not isinstance(rules, dict):
        return False
    for f in PERSON_SYNC_FIELDS:
        r = rules.get(f['key']) or {}
        if str(r.get('direction', 'none')).lower() != 'none':
            return True
    return False
PCO_PERSON_ID_FIELD = 'PCO_PersonId'  # Extra Value name on People

# --- PCO API ---------------------------------------------------------
PCO_BASE_URL = 'https://api.planningcenteronline.com'

# =====================================================================
# Latin-1 -> ASCII transliteration (CLAUDE.md pattern)
# IronPython's str()/unicode()/json.dumps mishandle some bytes coming back
# from SQL and .NET interop -- accented chars in person names blow up at
# the boundary. Walk values through safe_str() at read time so json.dumps
# only ever sees pure ASCII.
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
    """Convert any value to pure-ASCII string. Handles unicode, .NET
    System.String, byte strings in multiple encodings without raising."""
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
        try: return _to_ascii(s.decode('utf-8'))
        except: pass
        try: return _to_ascii(s.decode('latin-1'))
        except: pass
        try: return _to_ascii(s.decode('cp1252'))
        except: pass
        return ''.join(c if ord(c) < 128 else '?' for c in s)
    except:
        pass
    try: return repr(val)
    except: return ''

def safe_int(val, default=0):
    try: return int(val)
    except: return default

def _truthy(v, default):
    """Lenient bool parser for query/form params -- accepts 0/1, true/false,
    yes/no, on/off (case-insensitive). Returns default for unrecognised /
    None values so callers can preserve existing state when a checkbox is
    omitted from the POST."""
    if v is None:
        return default
    s = str(v).strip().lower()
    if s in ('1', 'true', 'yes', 'on'): return True
    if s in ('0', 'false', 'no', 'off', ''): return False
    return default

def html_escape(text):
    if text is None or text == '':
        return ''
    s = safe_str(text)
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    s = s.replace('"', '&quot;').replace("'", '&#39;')
    return s

# =====================================================================
# Storage helpers
# =====================================================================

def load_json(key, default=None):
    """Safely read a Text Content key as JSON. Returns default on any error."""
    try:
        raw = model.TextContent(key) or ''
        if not raw.strip():
            return default if default is not None else {}
        return json.loads(raw)
    except:
        return default if default is not None else {}

def save_json(key, obj):
    """Write a JSON-serializable object to a Text Content key."""
    try:
        model.WriteContentText(key, json.dumps(obj), '')
        return True
    except:
        return False

def now_iso():
    return datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')

def _parse_mapping(v):
    """Pull (orgId, syncAttendance, autoAddMember) out of a mapping value.
    Supports both the legacy int format (bare org id meant 'do everything')
    and the new dict format with per-mapping toggles. Returns ints/bools so
    callers don't have to defensively cast."""
    if isinstance(v, dict):
        try:
            org_id = int(v.get('orgId') or 0)
        except:
            org_id = 0
        return (
            org_id,
            bool(v.get('syncAttendance', True)),
            bool(v.get('autoAddMember', True)),
        )
    try:
        return int(v or 0), True, True
    except:
        return 0, True, True

def _mapping_org_id(v):
    """Shortcut when only the org id is needed (most read paths)."""
    return _parse_mapping(v)[0]

def _v3_migrate_org_to_people_mappings():
    """v3.0 migration: take any Service Plan Mappings (PCOSync_OrgMappings)
    and fold their per-plan attendance + auto-add-member settings into
    the matching Service Type Mapping (PCOSync_PeopleMappings). Runs at
    most once -- after the first successful merge, OrgMappings is left
    empty and this function is a no-op.

    Conflict resolution (rare): if the same service type is mapped to
    DIFFERENT TP involvements in the two storages, we keep the People
    Mapping's involvement and surface attendance there. The old org
    mapping's involvement reference is logged + dropped."""
    try:
        org_map = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(org_map, dict) or not org_map:
            return
        people_map = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(people_map, dict):
            people_map = {}
        migrated = 0
        for st_id, v in list(org_map.items()):
            old_org_id, sync_att, add_mem = _parse_mapping(v)
            if old_org_id <= 0:
                continue
            existing_p = people_map.get(st_id)
            if existing_p:
                p = _parse_people_mapping(existing_p)
                # Same involvement -- straight merge: lift attendance.
                if p['orgId'] == old_org_id:
                    p['perPlanAttendance'] = p['perPlanAttendance'] or bool(sync_att)
                    p['autoAddMember'] = p['autoAddMember'] and bool(add_mem)
                else:
                    # Different involvement -- keep People Mapping side.
                    # The old Service Plan Mapping target gets dropped.
                    pass
                people_map[st_id] = {
                    'pcoServiceTypeId': p['pcoServiceTypeId'] or st_id,
                    'pcoServiceTypeName': p['pcoServiceTypeName'],
                    'orgId': p['orgId'],
                    'autoAddMember': p['autoAddMember'],
                    'teamsAsSubgroups': p['teamsAsSubgroups'],
                    'perPlanAttendance': p['perPlanAttendance'],
                }
            else:
                # No People Mapping for this service type -- promote the
                # Service Plan Mapping into a People Mapping with
                # attendance on, no team subgroups (since the old
                # mapping had no concept of those).
                people_map[st_id] = {
                    'pcoServiceTypeId': st_id,
                    'pcoServiceTypeName': '',
                    'orgId': old_org_id,
                    'autoAddMember': bool(add_mem),
                    'teamsAsSubgroups': False,
                    'perPlanAttendance': bool(sync_att),
                }
            migrated += 1
        if migrated > 0:
            save_json(PEOPLE_MAPPINGS_KEY, people_map)
            save_json(ORG_MAPPINGS_KEY, {})  # clear old store
            append_audit({
                'action': 'v3_migrate_org_to_people_mappings',
                'migrated': migrated,
                'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
            })
    except Exception as e:
        # Swallow -- migration is best-effort; UI still works without it.
        try: model.DebugPrint('v3 migration failed: ' + str(e))
        except: pass

def _parse_schedule(v):
    """v3.3: shared schedule shape on every mapping.
    Picked off the mapping dict, never stored separately. Keeps storage
    schema flat so the existing save / load paths stay intact.

    Frequency: 'daily' (every day at hour) or 'weekly' (one specific
    day-of-week at hour). dayOfWeek follows Python convention 0=Mon..6=Sun
    so datetime.weekday() comparisons work directly."""
    if not isinstance(v, dict):
        v = {}
    sch = v.get('schedule') or {}
    if not isinstance(sch, dict):
        sch = {}
    try:
        dow = int(sch.get('dayOfWeek', 6))
    except:
        dow = 6
    try:
        hr = int(sch.get('hour', 6))
    except:
        hr = 6
    freq = safe_str(sch.get('frequency', 'weekly')).lower()
    if freq not in ('daily', 'weekly'):
        freq = 'weekly'
    if dow < 0 or dow > 6:
        dow = 6
    if hr < 0 or hr > 23:
        hr = 6
    return {
        'enabled': bool(sch.get('enabled', False)),
        'frequency': freq,
        'dayOfWeek': dow,
        'hour': hr,
        'notifyUsername': safe_str(sch.get('notifyUsername', '')),
        'includeIssues': bool(sch.get('includeIssues', True)),
        'lastScheduledRunAt': safe_str(sch.get('lastScheduledRunAt', '')),
    }

def _parse_all_people_mapping(v):
    """All People Sync: PCO People directory -> ONE TP involvement.
    Singleton -- only one mapping per install. TouchPoint is authoritative:
    no new TP people are ever created; PCO records without a TP match are
    surfaced as unmatched instead."""
    if not isinstance(v, dict):
        return {'orgId': 0, 'autoAddMember': True, 'includeInactive': False, 'schedule': _parse_schedule({})}
    try:
        org_id = int(v.get('orgId') or 0)
    except:
        org_id = 0
    return {
        'orgId': org_id,
        'autoAddMember': bool(v.get('autoAddMember', True)),
        # PCO People can include archived/inactive directory records.
        # Default OFF -- most churches want active records only in TP.
        'includeInactive': bool(v.get('includeInactive', False)),
        'schedule': _parse_schedule(v),
    }

def _parse_people_mapping(v):
    """Service Type Mapping: one PCO Service Type -> one umbrella TouchPoint
    involvement. Roster sync is implicit (the whole point of the mapping).
    Optional layers: teams as subgroups, per-plan attendance writes.
    v3.0+: perPlanAttendance toggle absorbs the old Service Plan Mappings."""
    if not isinstance(v, dict):
        return {
            'orgId': 0,
            'autoAddMember': True,
            'teamsAsSubgroups': True,
            'perPlanAttendance': False,
            'pcoServiceTypeId': '',
            'pcoServiceTypeName': '',
            'schedule': _parse_schedule({}),
        }
    try:
        org_id = int(v.get('orgId') or 0)
    except:
        org_id = 0
    return {
        'orgId': org_id,
        'autoAddMember': bool(v.get('autoAddMember', True)),
        'teamsAsSubgroups': bool(v.get('teamsAsSubgroups', True)),
        'perPlanAttendance': bool(v.get('perPlanAttendance', False)),
        'pcoServiceTypeId': safe_str(v.get('pcoServiceTypeId', '')),
        'pcoServiceTypeName': safe_str(v.get('pcoServiceTypeName', '')),
        'schedule': _parse_schedule(v),
    }

def _parse_team_mapping(v):
    """Team Mapping: one PCO Team -> one TouchPoint involvement. Roster
    sync is implicit. Optional layers: positions as subgroups, per-plan
    attendance writes filtered to this team's members."""
    if not isinstance(v, dict):
        return {
            'orgId': 0,
            'autoAddMember': True,
            'positionsAsSubgroups': True,
            'perPlanAttendance': False,
            'pcoTeamId': '',
            'pcoTeamName': '',
            'pcoServiceTypeId': '',
            'pcoServiceTypeName': '',
            'schedule': _parse_schedule({}),
        }
    try:
        org_id = int(v.get('orgId') or 0)
    except:
        org_id = 0
    return {
        'orgId': org_id,
        'autoAddMember': bool(v.get('autoAddMember', True)),
        'positionsAsSubgroups': bool(v.get('positionsAsSubgroups', True)),
        'perPlanAttendance': bool(v.get('perPlanAttendance', False)),
        'pcoTeamId': safe_str(v.get('pcoTeamId', '')),
        'pcoTeamName': safe_str(v.get('pcoTeamName', '')),
        'pcoServiceTypeId': safe_str(v.get('pcoServiceTypeId', '')),
        'pcoServiceTypeName': safe_str(v.get('pcoServiceTypeName', '')),
        'schedule': _parse_schedule(v),
    }

def _next_scheduled_run(sch, ref=None):
    """v3.3: compute the next datetime this schedule will fire after ref
    (defaults to now). Returns None if scheduling is disabled.

    Daily: next occurrence of (hour, 0) >= ref.
    Weekly: next occurrence of (dayOfWeek, hour, 0) >= ref."""
    if not isinstance(sch, dict) or not sch.get('enabled'):
        return None
    now = ref or datetime.datetime.now()
    hr = sch.get('hour', 6)
    freq = sch.get('frequency', 'weekly')
    candidate = now.replace(hour=hr, minute=0, second=0, microsecond=0)
    if freq == 'daily':
        if candidate <= now:
            candidate = candidate + datetime.timedelta(days=1)
        return candidate
    # Weekly.
    target_dow = sch.get('dayOfWeek', 6)
    days_ahead = (target_dow - now.weekday()) % 7
    candidate = candidate + datetime.timedelta(days=days_ahead)
    if candidate <= now:
        candidate = candidate + datetime.timedelta(days=7)
    return candidate

def _is_due_now(sch, last_run_iso, ref=None):
    """v3.3: True when this schedule should fire on the current invocation.

    Day matches AND hour matches AND we haven't already fired in the
    past hour. Last-run check prevents the scheduler firing twice if
    the runner runs more than once per hour."""
    if not isinstance(sch, dict) or not sch.get('enabled'):
        return False
    now = ref or datetime.datetime.now()
    if now.hour != sch.get('hour', 6):
        return False
    if sch.get('frequency', 'weekly') == 'weekly':
        if now.weekday() != sch.get('dayOfWeek', 6):
            return False
    if last_run_iso:
        try:
            last = datetime.datetime.strptime(last_run_iso[:19], '%Y-%m-%dT%H:%M:%S')
            if (now - last).total_seconds() < 60 * 55:
                return False  # already fired this hour-slot
        except:
            pass
    return True

def _resolve_username_email(username):
    """v3.3: look up a TouchPoint user's email by username. Returns
    {peopleId, name, email, found} so callers can validate before save
    and pull at run-time. Uses Users + People (Users.PeopleId joins
    People for the email)."""
    out = {'peopleId': 0, 'name': '', 'email': '', 'found': False}
    if not username:
        return out
    safe_u = safe_str(username).replace("'", "''")
    sql = """
        SELECT TOP 1 u.UserId, u.PeopleId, p.Name2, ISNULL(p.EmailAddress, '') AS Email
        FROM Users u WITH (NOLOCK)
        JOIN People p WITH (NOLOCK) ON p.PeopleId = u.PeopleId
        WHERE u.Username = '%s'
    """ % safe_u
    try:
        for r in q.QuerySql(sql):
            out['peopleId'] = int(r.PeopleId)
            out['name'] = safe_str(r.Name2)
            out['email'] = safe_str(r.Email)
            out['found'] = bool(out['email'])
            break
    except:
        pass
    return out

def _scheduler_installed():
    """v3.3.1: True if our marker block is in ScheduledTasks. Used by
    list/load handlers so the row UI can gate the per-mapping schedule
    editor on global install."""
    try:
        existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ''
        return (_SCHED_MARKER_START in existing)
    except:
        return False

def _format_schedule_label(sch):
    """v3.3: human-readable label like 'Sun 6:00 AM' or 'Daily 6:00 AM'."""
    if not isinstance(sch, dict) or not sch.get('enabled'):
        return ''
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    h = sch.get('hour', 6)
    ampm = 'AM' if h < 12 else 'PM'
    h12 = h if h <= 12 else h - 12
    if h12 == 0:
        h12 = 12
    when = '%d:00 %s' % (h12, ampm)
    if sch.get('frequency') == 'daily':
        return 'Daily ' + when
    return days[sch.get('dayOfWeek', 6)] + ' ' + when

def _send_sync_email(notify_username, subject, html_body):
    """Send the post-sync summary email. Looks up the user's email at
    send time (not at save) so role changes propagate. Returns
    (sent, errMsg).

    TouchPoint's email API is model.Email (NOT SendEmail -- which
    doesn't exist on PythonModel despite the older CLAUDE.md
    suggesting otherwise). Signature observed in TPxi_OpsCheckList:

        model.Email(pid, queuedBy, fromEmail, fromName, subject, body)

    'pid' can be a single PeopleId int (matches Search Builder p0=N
    semantics) and the email goes to that person. queuedBy needs to
    be a valid PeopleId so it shows up in the email queue with an
    attributable sender."""
    if not notify_username:
        return False, 'no notify username configured'
    info = _resolve_username_email(notify_username)
    if not info['found'] or not info['email']:
        return False, 'username "%s" has no email on file' % notify_username
    try:
        recipient_pid = int(info['peopleId'])
        # queued_by: prefer the logged-in user (foreground / UI path),
        # fall back to the recipient (scheduled / background path
        # where UserPeopleId may not be set). Always a valid PeopleId.
        queued_by = recipient_pid
        try:
            upid = model.UserPeopleId
            if upid:
                queued_by = int(upid)
        except:
            pass
        # From == recipient: PCO Sync has no configured sender, so we
        # self-notify. Email shows up in the queue from the recipient's
        # own address (domain-verified, won't bounce on SendGrid).
        from_email = info['email']
        from_name = 'PCO Sync'
        model.Email(recipient_pid, queued_by, from_email, from_name, subject, html_body)
        return True, ''
    except Exception as e:
        return False, 'send failed: ' + str(e)

def log_key_for_now():
    return LOG_KEY_PREFIX + datetime.datetime.now().strftime('%Y%m')

def _get_last_sync_times():
    """Walk the audit log (current + previous month) and return the most
    recent successful-sync timestamp per mapping key. Used by the
    Dashboard to show "Last synced X ago" pills.

    Audit shape (per action):
      sync_plan_attendance -> keyed by planId
      sync_roster          -> keyed by pcoServiceTypeId
      sync_team            -> keyed by pcoTeamId
      sync_people          -> keyed by pcoServiceTypeId
      sync_all_people      -> singleton timestamp"""
    times = {
        'all_people': None,
        'sync_team': {},
        'sync_people': {},
        'sync_roster': {},
        'sync_plan': {},
    }
    # Build [current_key, previous_key]. Only walk two months -- past that
    # the data is too stale to be useful as "last synced".
    keys = [log_key_for_now()]
    try:
        now = datetime.datetime.now()
        if now.month == 1:
            keys.append(LOG_KEY_PREFIX + str(now.year - 1) + '12')
        else:
            keys.append(LOG_KEY_PREFIX + ('%04d%02d' % (now.year, now.month - 1)))
    except:
        pass

    def _take(bucket, k, at):
        if not k:
            return
        if k not in bucket or at > bucket[k]:
            bucket[k] = at

    for key in keys:
        data = load_json(key, {})
        if not isinstance(data, dict):
            continue
        for e in (data.get('entries') or []):
            action = e.get('action', '')
            at = e.get('at', '')
            if not at or not action:
                continue
            if action == 'sync_all_people':
                if not times['all_people'] or at > times['all_people']:
                    times['all_people'] = at
            elif action == 'sync_team':
                _take(times['sync_team'], safe_str(e.get('pcoTeamId', '')), at)
            elif action == 'sync_people':
                _take(times['sync_people'], safe_str(e.get('pcoServiceTypeId', '')), at)
            elif action == 'sync_roster':
                _take(times['sync_roster'], safe_str(e.get('pcoServiceTypeId', '')), at)
            elif action == 'sync_plan_attendance':
                _take(times['sync_plan'], safe_str(e.get('planId', '')), at)
    return times

def append_audit(entry):
    """Append an audit entry to the current month's log file."""
    key = log_key_for_now()
    data = load_json(key, {'entries': []})
    if not isinstance(data, dict):
        data = {'entries': []}
    data.setdefault('entries', [])
    entry['at'] = now_iso()
    data['entries'].append(entry)
    # Trim oldest if huge (5000 entries / month = generous headroom)
    if len(data['entries']) > 5000:
        data['entries'] = data['entries'][-5000:]
    save_json(key, data)

# =====================================================================
# Form data helpers
# =====================================================================

def has_data(name):
    return hasattr(Data, name) and getattr(Data, name) not in (None, '')

def get_data(name, default=''):
    if has_data(name):
        return getattr(Data, name)
    return default

# =====================================================================
# PCO API client
# =====================================================================
# Auth: HTTP Basic with "app_id:secret" base64 encoded. The PAT pair is
# generated under My Apps -> Personal Access Tokens in PCO. Token lives in
# PCOSync_Settings; we don't take it from the request, so a stolen request
# can't leak it.

def _get_pco_credentials():
    """Return (app_id, secret) or (None, None) if not configured."""
    s = load_json(SETTINGS_KEY, {})
    if not isinstance(s, dict):
        return None, None
    app_id = (s.get('pco_app_id') or '').strip()
    secret = (s.get('pco_secret') or '').strip()
    if not app_id or not secret:
        return None, None
    return app_id, secret

def _pco_auth_header():
    """Build the Authorization header value for the configured PAT.
    Returns None if credentials aren't configured."""
    app_id, secret = _get_pco_credentials()
    if not app_id or not secret:
        return None
    raw = app_id + ':' + secret
    try:
        encoded = base64.b64encode(raw.encode('utf-8'))
    except:
        encoded = base64.b64encode(raw)
    return 'Basic ' + str(encoded)

def pco_get(path):
    """GET a path from the PCO API. Returns (parsed_json, error_message).
    error_message is None on success."""
    auth = _pco_auth_header()
    if not auth:
        return None, 'PCO credentials not configured. Open Settings tab and enter your Personal Access Token.'
    if not path.startswith('/'):
        path = '/' + path
    url = PCO_BASE_URL + path
    try:
        headers = {'Authorization': auth, 'Accept': 'application/json'}
        body = model.RestGet(url, headers)
        if body is None:
            return None, 'PCO API returned no body for ' + path
        try:
            return json.loads(str(body)), None
        except Exception as je:
            return None, 'PCO API returned non-JSON for ' + path + ': ' + str(je)
    except Exception as e:
        return None, 'PCO API call failed (' + path + '): ' + str(e)

# =====================================================================
# AJAX HANDLERS (POST)
# =====================================================================
# Each action is a small isolated handler. Per-tab actions are grouped
# together for grep-ability. Most return JSON; the response shape is
# {success: bool, message?: str, ...}.

def handle_test_connection():
    """Settings tab: verify the configured PAT can reach the PCO API."""
    try:
        data, err = pco_get('/services/v2/service_types?per_page=1')
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        # Successful auth returns a JSONAPI envelope with 'data' key
        if isinstance(data, dict) and ('data' in data or 'meta' in data):
            print json.dumps({'success': True, 'message': 'Connected to PCO successfully.'})
            return
        print json.dumps({'success': False, 'message': 'Unexpected response shape from PCO.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Connection test failed: ' + str(e)})

def handle_load_settings():
    """Return the non-secret parts of settings for the UI."""
    s = load_json(SETTINGS_KEY, {})
    if not isinstance(s, dict):
        s = {}
    out = {
        'success': True,
        'hasCredentials': bool(s.get('pco_app_id') and s.get('pco_secret')),
        'appIdMasked': '',
        'lastSyncAt': s.get('lastSyncAt', ''),
    }
    if s.get('pco_app_id'):
        aid = str(s['pco_app_id'])
        if len(aid) > 8:
            out['appIdMasked'] = aid[:4] + '...' + aid[-4:]
        else:
            out['appIdMasked'] = aid[:2] + '...'
    print json.dumps(out)

def handle_save_settings():
    """Save PCO PAT credentials. Empty values clear them."""
    try:
        app_id = str(get_data('pco_app_id', '')).strip()
        secret = str(get_data('pco_secret', '')).strip()
        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict):
            s = {}
        if app_id:
            s['pco_app_id'] = app_id
        elif get_data('clear_credentials', '') == 'true':
            s.pop('pco_app_id', None)
        if secret:
            s['pco_secret'] = secret
        elif get_data('clear_credentials', '') == 'true':
            s.pop('pco_secret', None)
        save_json(SETTINGS_KEY, s)
        print json.dumps({'success': True, 'message': 'Settings saved.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save failed: ' + str(e)})

# =====================================================================
# Service Type Mappings tab
# =====================================================================
# A PCO "Service Type" is the top-level container in PCO Services -- e.g.
# "Sunday Morning Worship", "Wednesday Night", "Christmas Eve". Each one
# maps to ONE TouchPoint involvement (the worship team org for that
# service). Stored in PCOSync_OrgMappings as {pcoServiceTypeId: tpOrgId}.

def handle_list_service_types():
    """Pull PCO service types and merge with our saved mappings so the
    UI can show mapped vs unmapped status in one round trip."""
    try:
        data, err = pco_get('/services/v2/service_types?per_page=100&order=name')
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        mappings = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}

        # Hydrate mapped org names in one SQL pass to avoid N+1.
        mapped_org_ids = set()
        for v in mappings.values():
            oid = _mapping_org_id(v)
            if oid > 0:
                mapped_org_ids.add(oid)
        org_name_by_id = {}
        if mapped_org_ids:
            ids_csv = ','.join(str(i) for i in mapped_org_ids)
            sql = """
                SELECT o.OrganizationId, o.OrganizationName,
                       ISNULL(d.Name, '') AS DivisionName,
                       ISNULL(p.Name, '') AS ProgramName
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE o.OrganizationId IN (%s)
            """ % ids_csv
            for r in q.QuerySql(sql):
                org_name_by_id[int(r.OrganizationId)] = {
                    'name': safe_str(r.OrganizationName),
                    'division': safe_str(r.DivisionName),
                    'program': safe_str(r.ProgramName),
                }

        out = []
        for item in (data.get('data') or []):
            pco_id = str(item.get('id') or '')
            attrs = item.get('attributes') or {}
            name = safe_str(attrs.get('name', ''))
            raw_mapping = mappings.get(pco_id)
            moid, sync_att, add_mem = _parse_mapping(raw_mapping)
            row = {
                'pcoId': pco_id,
                'pcoName': name,
                'sequence': safe_int(attrs.get('sequence', 0), 0),
                'mappedOrgId': None,
                'mappedOrgName': '',
                'mappedDivision': '',
                'mappedProgram': '',
                'syncAttendance': sync_att,
                'autoAddMember': add_mem,
            }
            if raw_mapping is not None and moid > 0:
                row['mappedOrgId'] = moid
                if moid in org_name_by_id:
                    row['mappedOrgName'] = org_name_by_id[moid]['name']
                    row['mappedDivision'] = org_name_by_id[moid]['division']
                    row['mappedProgram'] = org_name_by_id[moid]['program']
                else:
                    # Mapping exists but org isn't reachable (deleted? archived?)
                    row['mappedOrgName'] = '[Org #' + str(moid) + ' not found]'
            out.append(row)
        # Mapped first (the things already configured), then unmapped. Within
        # each group, alphabetical by PCO service type name. Lets the worship
        # admin see at a glance what's already wired up and what still needs
        # attention.
        out.sort(key=lambda x: (
            0 if x.get('mappedOrgId') else 1,
            (x.get('pcoName') or '').lower(),
        ))
        print json.dumps({'success': True, 'serviceTypes': out, 'mappedCount': len([r for r in out if r['mappedOrgId']])})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List service types failed: ' + str(e)})

def handle_search_involvements():
    """For the mapping picker. Returns up to 25 active TP orgs matching the
    search term, with division + program context."""
    try:
        term = safe_str(get_data('search_term', '')).strip()
        if not term:
            print json.dumps({'success': True, 'orgs': []})
            return
        safe_term = term.replace("'", "''")
        sql = """
            SELECT TOP 25
                o.OrganizationId, o.OrganizationName,
                ISNULL(d.Name, '') AS DivisionName,
                ISNULL(p.Name, '') AS ProgramName,
                (SELECT COUNT(*) FROM OrganizationMembers om
                 WHERE om.OrganizationId = o.OrganizationId) AS MemberCount
            FROM Organizations o WITH (NOLOCK)
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program p ON d.ProgId = p.Id
            WHERE o.OrganizationStatusId = 30
              AND (o.OrganizationName LIKE '%%%s%%'
                   OR CAST(o.OrganizationId AS VARCHAR) = '%s')
            ORDER BY o.OrganizationName
        """ % (safe_term, safe_term)
        orgs = []
        for r in q.QuerySql(sql):
            orgs.append({
                'orgId': int(r.OrganizationId),
                'orgName': safe_str(r.OrganizationName),
                'divisionName': safe_str(r.DivisionName),
                'programName': safe_str(r.ProgramName),
                'memberCount': safe_int(r.MemberCount, 0),
            })
        print json.dumps({'success': True, 'orgs': orgs})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Search failed: ' + str(e)})

def handle_save_org_mapping():
    """Save or remove one PCO Service Type -> TP Org mapping.
    org_id == 0 means delete the mapping."""
    try:
        pco_id = str(get_data('pco_service_type_id', '')).strip()
        org_id = safe_int(get_data('tp_org_id', 0), 0)
        if not pco_id:
            print json.dumps({'success': False, 'message': 'Missing PCO service type id'})
            return
        mappings = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        if org_id <= 0:
            # Delete
            mappings.pop(pco_id, None)
            save_json(ORG_MAPPINGS_KEY, mappings)
            append_audit({
                'action': 'unmap_service_type',
                'pcoServiceTypeId': pco_id,
                'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
            })
            print json.dumps({'success': True, 'message': 'Mapping removed.'})
            return
        # Verify the org exists + is active before saving the mapping.
        verify_sql = """
            SELECT TOP 1 o.OrganizationId, o.OrganizationName
            FROM Organizations o WITH (NOLOCK)
            WHERE o.OrganizationId = %s
              AND o.OrganizationStatusId = 30
        """ % str(org_id)
        org_name = ''
        for r in q.QuerySql(verify_sql):
            org_name = safe_str(r.OrganizationName)
            break
        if not org_name:
            print json.dumps({'success': False, 'message': 'Org #' + str(org_id) + ' is not active or does not exist.'})
            return
        # Preserve existing per-mapping toggles if we're re-pointing an
        # already-mapped service type. New mappings default both flags on
        # (match prior behavior so existing churches don't suddenly stop
        # joining members or writing attendance).
        _existing = mappings.get(pco_id)
        _, sync_att, add_mem = _parse_mapping(_existing) if _existing is not None else (0, True, True)
        mappings[pco_id] = {
            'orgId': org_id,
            'syncAttendance': sync_att,
            'autoAddMember': add_mem,
        }
        save_json(ORG_MAPPINGS_KEY, mappings)
        append_audit({
            'action': 'map_service_type',
            'pcoServiceTypeId': pco_id,
            'tpOrgId': org_id,
            'tpOrgName': org_name,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'message': 'Mapped to ' + org_name + '.', 'orgName': org_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save mapping failed: ' + str(e)})

def handle_load_org_mappings():
    """Return raw mappings dict (for diagnostics)."""
    m = load_json(ORG_MAPPINGS_KEY, {})
    if not isinstance(m, dict):
        m = {}
    print json.dumps({'success': True, 'mappings': m})

# ----- Team Mappings (v2.0+) -----------------------------------------

def handle_list_pco_teams_for_service_type():
    """For the Team Mapping picker. After the admin picks a Service Type
    they see the list of Teams to choose from."""
    try:
        service_type_id = safe_str(get_data('service_type_id', '')).strip()
        if not service_type_id:
            print json.dumps({'success': False, 'message': 'Missing service_type_id'})
            return
        teams, err = _pco_teams_for_service_type(service_type_id)
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        print json.dumps({'success': True, 'teams': teams})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List teams failed: ' + str(e)})

def handle_list_team_mappings():
    """List all saved Team Sync mappings with TP org name hydrated.
    Used by the Team Mappings section under the Mappings tab."""
    try:
        raw = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(raw, dict):
            raw = {}
        # Hydrate TP org details.
        org_ids = set()
        for v in raw.values():
            info = _parse_team_mapping(v)
            if info['orgId'] > 0:
                org_ids.add(info['orgId'])
        org_info = {}
        if org_ids:
            ids_csv = ','.join(str(i) for i in org_ids)
            sql = """
                SELECT o.OrganizationId, o.OrganizationName,
                       ISNULL(d.Name, '') AS DivisionName,
                       ISNULL(p.Name, '') AS ProgramName
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE o.OrganizationId IN (%s)
            """ % ids_csv
            for r in q.QuerySql(sql):
                org_info[int(r.OrganizationId)] = {
                    'name': safe_str(r.OrganizationName),
                    'division': safe_str(r.DivisionName),
                    'program': safe_str(r.ProgramName),
                }
        out = []
        for team_id, v in raw.items():
            info = _parse_team_mapping(v)
            oi = org_info.get(info['orgId'], {}) if info['orgId'] > 0 else {}
            sch = info.get('schedule', _parse_schedule({}))
            nxt = _next_scheduled_run(sch)
            out.append({
                'pcoTeamId': info['pcoTeamId'] or team_id,
                'pcoTeamName': info['pcoTeamName'],
                'pcoServiceTypeId': info['pcoServiceTypeId'],
                'pcoServiceTypeName': info['pcoServiceTypeName'],
                'tpOrgId': info['orgId'],
                'tpOrgName': oi.get('name', '[Org #' + str(info['orgId']) + ' not found]'),
                'tpDivision': oi.get('division', ''),
                'tpProgram': oi.get('program', ''),
                'autoAddMember': info['autoAddMember'],
                'positionsAsSubgroups': info['positionsAsSubgroups'],
                'perPlanAttendance': info['perPlanAttendance'],
                'schedule': sch,
                'scheduleLabel': _format_schedule_label(sch),
                'nextRunIso': nxt.strftime('%Y-%m-%dT%H:%M:%S') if nxt else '',
            })
        out.sort(key=lambda r: ((r.get('pcoServiceTypeName') or '').lower(),
                                (r.get('pcoTeamName') or '').lower()))
        print json.dumps({'success': True, 'mappings': out, 'schedulerInstalled': _scheduler_installed()})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List team mappings failed: ' + str(e)})

def handle_save_team_mapping():
    """Save or update a Team Sync mapping. The PCO team is identified by
    its PCO team id; the org by TP id; metadata is stored so the UI can
    render team/service-type names without re-fetching from PCO."""
    try:
        pco_team_id = safe_str(get_data('pco_team_id', '')).strip()
        pco_team_name = safe_str(get_data('pco_team_name', '')).strip()
        pco_st_id = safe_str(get_data('pco_service_type_id', '')).strip()
        pco_st_name = safe_str(get_data('pco_service_type_name', '')).strip()
        org_id = safe_int(get_data('tp_org_id', 0), 0)
        if not pco_team_id or org_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing pco_team_id or tp_org_id'})
            return
        verify_sql = """
            SELECT TOP 1 o.OrganizationId, o.OrganizationName
            FROM Organizations o WITH (NOLOCK)
            WHERE o.OrganizationId = %s
              AND o.OrganizationStatusId = 30
        """ % str(org_id)
        org_name = ''
        for r in q.QuerySql(verify_sql):
            org_name = safe_str(r.OrganizationName)
            break
        if not org_name:
            print json.dumps({'success': False, 'message': 'Org #' + str(org_id) + ' is not active or does not exist.'})
            return
        mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        existing = _parse_team_mapping(mappings.get(pco_team_id))
        mappings[pco_team_id] = {
            'pcoTeamId': pco_team_id,
            'pcoTeamName': pco_team_name or existing['pcoTeamName'],
            'pcoServiceTypeId': pco_st_id or existing['pcoServiceTypeId'],
            'pcoServiceTypeName': pco_st_name or existing['pcoServiceTypeName'],
            'orgId': org_id,
            'autoAddMember': existing['autoAddMember'] if existing['orgId'] > 0 else True,
            'positionsAsSubgroups': existing['positionsAsSubgroups'] if existing['orgId'] > 0 else True,
            'perPlanAttendance': existing['perPlanAttendance'] if existing['orgId'] > 0 else False,
        }
        save_ok = save_json(TEAM_MAPPINGS_KEY, mappings)
        verify = load_json(TEAM_MAPPINGS_KEY, {})
        verified = (save_ok and isinstance(verify, dict)
                    and _parse_team_mapping(verify.get(pco_team_id))['orgId'] == org_id)
        append_audit({
            'action': 'save_team_mapping',
            'pcoTeamId': pco_team_id,
            'pcoTeamName': pco_team_name,
            'tpOrgId': org_id,
            'tpOrgName': org_name,
            'saveOk': save_ok,
            'verified': verified,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        if not verified:
            print json.dumps({'success': False, 'message': 'Save did not persist -- check role on PCOSync_TeamMappings content.'})
            return
        print json.dumps({'success': True, 'message': 'Team mapped to ' + org_name + '.', 'orgName': org_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save team mapping failed: ' + str(e)})

def handle_delete_team_mapping():
    try:
        pco_team_id = safe_str(get_data('pco_team_id', '')).strip()
        if not pco_team_id:
            print json.dumps({'success': False, 'message': 'Missing pco_team_id'})
            return
        mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        mappings.pop(pco_team_id, None)
        save_json(TEAM_MAPPINGS_KEY, mappings)
        append_audit({
            'action': 'delete_team_mapping',
            'pcoTeamId': pco_team_id,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'message': 'Team mapping removed.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Delete team mapping failed: ' + str(e)})

def handle_check_scheduler_install():
    """v3.3.1: report whether our scheduler block is present in
    TouchPoint's ScheduledTasks special content. The UI gates the
    per-mapping schedule editor on this -- no point letting staff
    configure schedules that will never fire."""
    try:
        try:
            existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ''
        except Exception as _re:
            print json.dumps({'success': False, 'installed': False, 'message': 'Could not read ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_re)})
            return
        installed = (_SCHED_MARKER_START in existing)
        # Surface whether the script name is referenced anywhere outside
        # our managed block -- helps catch hand-edits the admin made.
        ref_outside = False
        if not installed and DC_SCRIPT_ID in existing:
            ref_outside = True
        print json.dumps({
            'success': True,
            'installed': installed,
            'referencedOutsideBlock': ref_outside,
            'contentSlot': _SCHED_CONTENT_SLOT,
            'scriptName': DC_SCRIPT_ID,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Check install failed: ' + str(e)})

def handle_install_scheduler():
    """v3.3.1: inject our scheduler block into the ScheduledTasks
    special content. Idempotent -- if our marker is already present
    we report 'already installed'. Otherwise we append after the
    existing content (preserves any other tools' blocks)."""
    try:
        try:
            existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ''
        except Exception as _re:
            print json.dumps({'success': False, 'message': 'Could not read ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_re)})
            return
        if _SCHED_MARKER_START in existing:
            print json.dumps({'success': True, 'installed': True, 'alreadyInstalled': True,
                              'message': 'Already installed in ' + _SCHED_CONTENT_SLOT + '.'})
            return
        # The block sets the scheduler flag then calls back into our
        # script. CallScript reuses model.Data, so our entry-point
        # check picks up the flag. Errors get printed (TouchPoint's
        # ScheduledTasks runner captures them).
        block = (
            _SCHED_MARKER_START + "\n"
            "try:\n"
            "    Data.scheduler = 'true'\n"
            "    model.CallScript('" + DC_SCRIPT_ID + "')\n"
            "except Exception as _pco_e:\n"
            "    print 'PCO Sync scheduler error: ' + str(_pco_e)\n"
            + _SCHED_MARKER_END + "\n"
        )
        new_content = (existing.rstrip() + ('\n\n' if existing.strip() else '') + block)
        try:
            model.WriteContentPython(_SCHED_CONTENT_SLOT, new_content)
        except Exception as _we:
            print json.dumps({'success': False, 'message': 'Could not write ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_we)})
            return
        # Verify-after-write. Re-read and confirm the marker is there.
        try:
            verify = model.PythonContent(_SCHED_CONTENT_SLOT) or ''
            if _SCHED_MARKER_START not in verify:
                print json.dumps({'success': False, 'message': 'Write did not persist. Check role on ' + _SCHED_CONTENT_SLOT + '.'})
                return
        except:
            pass
        append_audit({
            'action': 'install_scheduler',
            'contentSlot': _SCHED_CONTENT_SLOT,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'installed': True,
                          'message': 'Installed into ' + _SCHED_CONTENT_SLOT + '. PCO Sync will now run on TouchPoint\'s scheduled task cycle.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Install failed: ' + str(e)})

def handle_uninstall_scheduler():
    """v3.3.1: remove our scheduler block. Leaves anything else in
    ScheduledTasks untouched."""
    try:
        try:
            existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ''
        except Exception as _re:
            print json.dumps({'success': False, 'message': 'Could not read ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_re)})
            return
        if _SCHED_MARKER_START not in existing:
            print json.dumps({'success': True, 'installed': False, 'notInstalled': True,
                              'message': 'Not installed in ' + _SCHED_CONTENT_SLOT + '.'})
            return
        import re as _re_mod
        pat = _re_mod.escape(_SCHED_MARKER_START) + r".*?" + _re_mod.escape(_SCHED_MARKER_END) + r"\n?"
        new_content = _re_mod.sub(pat, '', existing, flags=_re_mod.DOTALL)
        new_content = _re_mod.sub(r"\n{3,}", "\n\n", new_content).rstrip() + "\n"
        try:
            model.WriteContentPython(_SCHED_CONTENT_SLOT, new_content)
        except Exception as _we:
            print json.dumps({'success': False, 'message': 'Could not write ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_we)})
            return
        append_audit({
            'action': 'uninstall_scheduler',
            'contentSlot': _SCHED_CONTENT_SLOT,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'installed': False,
                          'message': 'Removed from ' + _SCHED_CONTENT_SLOT + '. Scheduled syncs will no longer fire.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Uninstall failed: ' + str(e)})

def handle_search_tp_users():
    """v3.3.2: typeahead for the schedule Notify field. Searches TP
    Users by name (Name / Name2 first-last and last-first), username,
    or email. Only returns people who actually HAVE a Users record
    AND an email -- both are required for the scheduler email to fire.
    Caps at 12 hits to keep the dropdown short."""
    try:
        term = safe_str(get_data('search_term', '')).strip()
        if not term or len(term) < 2:
            print json.dumps({'success': True, 'users': []})
            return
        safe_term = term.replace("'", "''")
        # Per-word AND clause against Name so "Ben Bax" or "Bax Ben" both
        # match. Mirrors the TP people search pattern.
        words = [w for w in re.split(r'\s+', term) if w]
        word_clauses = []
        for w in words:
            sw = w.replace("'", "''")
            word_clauses.append("p.Name LIKE '%%%s%%'" % sw)
        per_word_clause = ' AND '.join(word_clauses) if word_clauses else "1 = 1"
        sql = ("""
            SELECT TOP 12 u.Username, p.PeopleId, p.Name, p.Name2,
                   ISNULL(p.EmailAddress, '') AS Email
            FROM Users u WITH (NOLOCK)
            JOIN People p WITH (NOLOCK) ON p.PeopleId = u.PeopleId
            WHERE (
                    u.Username LIKE '%%%s%%'
                 OR p.Name LIKE '%%%s%%'
                 OR p.Name2 LIKE '%%%s%%'
                 OR (%s)
                 OR p.EmailAddress LIKE '%%%s%%'
                  )
              AND p.IsDeceased = 0
              AND p.ArchivedFlag = 0
              AND ISNULL(p.EmailAddress, '') <> ''
              AND ISNULL(u.Username, '') <> ''
            ORDER BY p.Name2
        """ % (safe_term, safe_term, safe_term, per_word_clause, safe_term))
        out = []
        try:
            for r in q.QuerySql(sql):
                out.append({
                    'username': safe_str(r.Username),
                    'peopleId': int(r.PeopleId),
                    'name': safe_str(r.Name) or safe_str(r.Name2),
                    'email': safe_str(r.Email),
                })
        except:
            pass
        print json.dumps({'success': True, 'users': out})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Search users failed: ' + str(e)})

def handle_resolve_username():
    """v3.3: validate a TouchPoint username and return their email.
    Used by the inline scheduler config to give immediate feedback when
    staff type a username."""
    try:
        username = safe_str(get_data('username', '')).strip()
        info = _resolve_username_email(username)
        print json.dumps({
            'success': True,
            'found': info['found'],
            'name': info['name'],
            'email': info['email'],
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Resolve username failed: ' + str(e)})

def handle_set_schedule_options():
    """v3.3: save schedule settings on a mapping. Accepts a 'kind' so
    one handler covers all three mapping types (people/team/all_people).

    Inputs:
      kind: 'people' | 'team' | 'all_people'
      key:  PCO id (service type id for people, team id for team, blank for all_people)
      enabled, frequency, day_of_week, hour, notify_username, include_issues

    Reads + writes via the same load_json/save_json path as existing
    setters so storage stays consistent."""
    try:
        kind = safe_str(get_data('kind', '')).strip()
        key = safe_str(get_data('key', '')).strip()
        if kind == 'people':
            store_key = PEOPLE_MAPPINGS_KEY
        elif kind == 'team':
            store_key = TEAM_MAPPINGS_KEY
        elif kind == 'all_people':
            store_key = ALL_PEOPLE_MAPPING_KEY
        else:
            print json.dumps({'success': False, 'message': 'Unknown kind: ' + kind})
            return
        # Build the new schedule dict.
        sch = {
            'enabled': _truthy(get_data('enabled', '0'), False),
            'frequency': safe_str(get_data('frequency', 'weekly')).lower(),
            'dayOfWeek': safe_int(get_data('day_of_week', 6), 6),
            'hour': safe_int(get_data('hour', 6), 6),
            'notifyUsername': safe_str(get_data('notify_username', '')).strip(),
            'includeIssues': _truthy(get_data('include_issues', '1'), True),
        }
        if sch['frequency'] not in ('daily', 'weekly'):
            sch['frequency'] = 'weekly'
        if sch['dayOfWeek'] < 0 or sch['dayOfWeek'] > 6:
            sch['dayOfWeek'] = 6
        if sch['hour'] < 0 or sch['hour'] > 23:
            sch['hour'] = 6
        # v3.3.1: must have global scheduler installed before any
        # per-mapping schedule can be enabled. Otherwise the schedule
        # will never fire and staff would wonder why.
        if sch['enabled'] and not _scheduler_installed():
            print json.dumps({'success': False, 'message': 'Cannot enable schedule: the global scheduler is not installed. Open Settings -> Scheduled Sync and click Install.'})
            return
        # Email lookup (informational on save).
        email_info = _resolve_username_email(sch['notifyUsername']) if sch['notifyUsername'] else {'found': False, 'email': '', 'name': ''}
        if sch['enabled'] and not email_info['found']:
            print json.dumps({'success': False, 'message': 'Cannot enable schedule: username "%s" has no email on file.' % sch['notifyUsername']})
            return
        # Persist into the mapping.
        if kind == 'all_people':
            cur = load_json(store_key, {})
            if not isinstance(cur, dict):
                cur = {}
            existing_sch = (cur.get('schedule') or {})
            # Preserve last-run timestamp -- never overwrite via UI save.
            sch['lastScheduledRunAt'] = safe_str(existing_sch.get('lastScheduledRunAt', ''))
            cur['schedule'] = sch
            save_json(store_key, cur)
        else:
            mappings = load_json(store_key, {})
            if not isinstance(mappings, dict):
                mappings = {}
            v = mappings.get(key)
            if not isinstance(v, dict):
                print json.dumps({'success': False, 'message': 'Mapping not found for key: ' + key})
                return
            existing_sch = (v.get('schedule') or {})
            sch['lastScheduledRunAt'] = safe_str(existing_sch.get('lastScheduledRunAt', ''))
            v['schedule'] = sch
            mappings[key] = v
            save_json(store_key, mappings)
        # Compute next-run label for the response so the row can show it
        # immediately without re-fetching.
        label = _format_schedule_label(sch)
        next_run = _next_scheduled_run(sch)
        next_iso = next_run.strftime('%Y-%m-%dT%H:%M:%S') if next_run else ''
        print json.dumps({
            'success': True,
            'schedule': sch,
            'scheduleLabel': label,
            'nextRunIso': next_iso,
            'recipientEmail': email_info['email'],
            'recipientName': email_info['name'],
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save schedule failed: ' + str(e)})

def handle_set_team_mapping_options():
    """Toggle the per-team-mapping options (autoAddMember,
    positionsAsSubgroups). autoAddMember is included for symmetry, but
    in v2.0 Team Sync only does roster + subgroups -- there's no
    attendance mode, so autoAddMember off effectively disables the
    mapping. Kept for future hooks (e.g. notify-on-roster-change)."""
    try:
        pco_team_id = safe_str(get_data('pco_team_id', '')).strip()
        if not pco_team_id:
            print json.dumps({'success': False, 'message': 'Missing pco_team_id'})
            return
        mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        if pco_team_id not in mappings:
            print json.dumps({'success': False, 'message': 'No team mapping found for this PCO team.'})
            return
        current = _parse_team_mapping(mappings[pco_team_id])
        new_add = _truthy(get_data('auto_add_member', None), current['autoAddMember'])
        new_pos = _truthy(get_data('positions_as_subgroups', None), current['positionsAsSubgroups'])
        new_att = _truthy(get_data('per_plan_attendance', None), current['perPlanAttendance'])
        mappings[pco_team_id] = {
            'pcoTeamId': current['pcoTeamId'] or pco_team_id,
            'pcoTeamName': current['pcoTeamName'],
            'pcoServiceTypeId': current['pcoServiceTypeId'],
            'pcoServiceTypeName': current['pcoServiceTypeName'],
            'orgId': current['orgId'],
            'autoAddMember': new_add,
            'positionsAsSubgroups': new_pos,
            'perPlanAttendance': new_att,
            # v3.3: preserve scheduler config across option saves.
            'schedule': current.get('schedule', _parse_schedule({})),
        }
        save_json(TEAM_MAPPINGS_KEY, mappings)
        append_audit({
            'action': 'update_team_mapping_options',
            'pcoTeamId': pco_team_id,
            'autoAddMember': new_add,
            'positionsAsSubgroups': new_pos,
            'perPlanAttendance': new_att,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'autoAddMember': new_add, 'positionsAsSubgroups': new_pos, 'perPlanAttendance': new_att})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Update team mapping options failed: ' + str(e)})

def handle_set_mapping_options():
    """Toggle the per-mapping syncAttendance / autoAddMember flags without
    changing which org the service type points to. Called when the user
    flips a checkbox on the Mappings tab."""
    try:
        pco_id = str(get_data('pco_service_type_id', '')).strip()
        if not pco_id:
            print json.dumps({'success': False, 'message': 'Missing PCO service type id'})
            return
        sync_att_raw = get_data('sync_attendance', None)
        add_mem_raw = get_data('auto_add_member', None)
        mappings = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        if pco_id not in mappings:
            print json.dumps({'success': False, 'message': 'No mapping exists for this service type yet -- pick an involvement first.'})
            return
        current_org, current_sync, current_add = _parse_mapping(mappings[pco_id])
        new_sync = _truthy(sync_att_raw, current_sync)
        new_add = _truthy(add_mem_raw, current_add)
        new_entry = {
            'orgId': current_org,
            'syncAttendance': new_sync,
            'autoAddMember': new_add,
        }
        mappings[pco_id] = new_entry
        save_ok = save_json(ORG_MAPPINGS_KEY, mappings)
        # Verify-after-write: read it back and confirm the toggle stuck.
        # Without this, a silent WriteContentText failure looks like
        # success to the UI and the user sees the checkbox spring back on
        # the next page load.
        verify = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(verify, dict):
            verify = {}
        v_org, v_sync, v_add = _parse_mapping(verify.get(pco_id))
        verified = (
            save_ok and
            v_org == current_org and
            v_sync == new_sync and
            v_add == new_add
        )
        append_audit({
            'action': 'update_mapping_options',
            'pcoServiceTypeId': pco_id,
            'tpOrgId': current_org,
            'syncAttendance': new_sync,
            'autoAddMember': new_add,
            'saveOk': save_ok,
            'verified': verified,
            'rawSyncAttendance': str(sync_att_raw),
            'rawAutoAddMember': str(add_mem_raw),
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        if not verified:
            print json.dumps({
                'success': False,
                'message': ('Save did not persist. WriteContentText returned ' +
                            ('OK' if save_ok else 'ERROR') +
                            ' but readback shows syncAttendance=' + str(v_sync) +
                            ', autoAddMember=' + str(v_add) +
                            ' (wanted ' + str(new_sync) + '/' + str(new_add) + '). ' +
                            'Check the role on PCOSync_OrgMappings content.'),
                'syncAttendance': v_sync,
                'autoAddMember': v_add,
            })
            return
        print json.dumps({
            'success': True,
            'syncAttendance': new_sync,
            'autoAddMember': new_add,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Update options failed: ' + str(e)})

# =====================================================================
# Person Data Sync (directional, decision-based)
# =====================================================================
# Rules: per-field {direction: none|pco_to_tp|tp_to_pco, mode: auto|review}.
# Default = none (TouchPoint authoritative). When a rule is opted in:
#   auto    -> write happens silently during plan/roster sync
#   review  -> change queued to PCOSync_PendingPersonChanges for triage
# Skip Forever -> stored as person extra value PCOSync_Skip_<field> so a
# stale PCO record doesn't re-queue the same change every Sunday.

def handle_load_person_sync_rules():
    try:
        raw = load_json(PERSON_SYNC_RULES_KEY, {})
        rules = normalize_person_sync_rules(raw)
        print json.dumps({
            'success': True,
            'rules': rules,
            'fields': PERSON_SYNC_FIELDS,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load person sync rules failed: ' + str(e)})

def handle_save_person_sync_rules():
    try:
        raw_payload = get_data('rules_json', '{}')
        try:
            parsed = json.loads(raw_payload)
        except Exception as je:
            print json.dumps({'success': False, 'message': 'Invalid rules JSON: ' + str(je)})
            return
        rules = normalize_person_sync_rules(parsed)
        save_ok = save_json(PERSON_SYNC_RULES_KEY, rules)
        # Verify-after-write (same pattern as set_mapping_options).
        verify = load_json(PERSON_SYNC_RULES_KEY, {})
        verified = (save_ok and normalize_person_sync_rules(verify) == rules)
        append_audit({
            'action': 'save_person_sync_rules',
            'rules': rules,
            'saveOk': save_ok,
            'verified': verified,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        if not verified:
            print json.dumps({
                'success': False,
                'message': 'Save did not persist -- check role on PCOSync_PersonSyncRules content.'
            })
            return
        print json.dumps({'success': True, 'rules': rules})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save person sync rules failed: ' + str(e)})

def _person_skip_forever_field_name(field):
    """Stable extra-value field name used to record 'never queue this
    field for this person again' decisions."""
    return 'PCOSync_Skip_' + field

def _person_should_skip_forever(people_id, field):
    """Read the skip flag for one (person, field). Returns True if set."""
    try:
        v = model.ExtraValueText(int(people_id), _person_skip_forever_field_name(field))
        if v and str(v).strip().lower() in ('1', 'true', 'yes', 'on'):
            return True
    except:
        pass
    try:
        # Some setups put bool extras in the bit slot instead.
        b = model.ExtraValueBit(int(people_id), _person_skip_forever_field_name(field))
        if b is True or b == 1:
            return True
    except:
        pass
    return False

def _person_field_norm(field, value):
    """Normalize a value for comparison. Email is case-insensitive +
    trimmed; names are trimmed."""
    s = safe_str(value).strip()
    if field == 'email':
        return s.lower()
    return s

def _tp_person_field(person_obj, tp_field):
    """Read the TP value for a field via the Person object."""
    if not person_obj:
        return ''
    try:
        return safe_str(getattr(person_obj, tp_field, '') or '')
    except:
        return ''

def _queue_person_change(entry):
    """Append a pending change to the queue. We dedupe by
    (tpPeopleId, field) so a person who's been syncing for weeks doesn't
    accumulate 4 'change email' rows."""
    queue = load_json(PERSON_PENDING_KEY, {'entries': []})
    if not isinstance(queue, dict):
        queue = {'entries': []}
    entries = queue.setdefault('entries', [])
    # Drop any prior pending entry for the same (person, field) -- the
    # newer one is what's current.
    keep = []
    for e in entries:
        if (e.get('tpPeopleId') == entry['tpPeopleId'] and
            e.get('field') == entry['field'] and
            e.get('status', 'pending') == 'pending'):
            continue
        keep.append(e)
    entry['id'] = entry.get('id') or (str(entry['tpPeopleId']) + '_' + entry['field'] + '_' + now_iso())
    entry['queuedAt'] = entry.get('queuedAt') or now_iso()
    entry['status'] = 'pending'
    keep.append(entry)
    # Cap queue size to keep the content row sane.
    if len(keep) > 2000:
        keep = keep[-2000:]
    queue['entries'] = keep
    save_json(PERSON_PENDING_KEY, queue)

def _parse_people_json_payload(raw):
    """Decode the JS-supplied people array. Each entry is
    {tpPeopleId, isConfirmed, pcoFirstName, pcoLastName, pcoEmail}.
    Returns a dict keyed by tpPeopleId for quick lookup; missing/invalid
    entries silently drop (sync proceeds without person-data comparison).
    isConfirmed defaults to True for backward-compat with older callers
    that didn't send the flag (those always sent only Confirmed pids)."""
    out = {}
    if not raw:
        return out
    try:
        parsed = json.loads(raw)
    except:
        return out
    if not isinstance(parsed, list):
        return out
    for item in parsed:
        if not isinstance(item, dict):
            continue
        try:
            tp_id = int(item.get('tpPeopleId') or 0)
        except:
            tp_id = 0
        if tp_id <= 0:
            continue
        out[tp_id] = {
            'isConfirmed': bool(item.get('isConfirmed', True)),
            'first_name': safe_str(item.get('pcoFirstName', '')),
            'last_name':  safe_str(item.get('pcoLastName', '')),
            'email':      safe_str(item.get('pcoEmail', '')),
        }
    return out

def apply_person_sync_for_one(tp_people_id, pco_data, rules):
    """Compare PCO field values to TP for one matched person and
    either auto-write or queue a review. Returns counts dict.
    pco_data = {first_name, last_name, email}."""
    counts = {'auto': 0, 'queued': 0, 'skipped': 0, 'unchanged': 0}
    try:
        person = model.GetPerson(int(tp_people_id))
    except:
        person = None
    if not person:
        return counts
    for f in PERSON_SYNC_FIELDS:
        key = f['key']
        rule = (rules or {}).get(key) or {}
        direction = str(rule.get('direction', 'none')).lower()
        if direction != 'pco_to_tp':
            # v1 only writes PCO -> TP. TP -> PCO is reserved for v1.1.
            continue
        pco_val = _person_field_norm(key, pco_data.get(key, ''))
        if not pco_val:
            # Don't queue "clear out the email" changes -- empty PCO
            # values are almost always missing data, not intentional.
            counts['unchanged'] += 1
            continue
        tp_val = _person_field_norm(key, _tp_person_field(person, f['tpField']))
        if pco_val == tp_val:
            counts['unchanged'] += 1
            continue
        if _person_should_skip_forever(tp_people_id, key):
            counts['skipped'] += 1
            continue
        mode = str(rule.get('mode', 'review')).lower()
        if mode == 'auto':
            try:
                model.UpdatePerson(int(tp_people_id), f['tpField'], pco_data.get(key, ''))
                counts['auto'] += 1
                append_audit({
                    'action': 'person_data_auto_update',
                    'tpPeopleId': int(tp_people_id),
                    'field': key,
                    'oldValue': _tp_person_field(person, f['tpField']),
                    'newValue': pco_data.get(key, ''),
                    'direction': direction,
                    'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
                })
            except Exception as ue:
                # Failed write -> queue for review so a human can decide.
                _queue_person_change({
                    'tpPeopleId': int(tp_people_id),
                    'tpName': safe_str(getattr(person, 'Name2', '') or getattr(person, 'Name', '')),
                    'field': key,
                    'fieldLabel': f['label'],
                    'tpValue': _tp_person_field(person, f['tpField']),
                    'pcoValue': pco_data.get(key, ''),
                    'direction': direction,
                    'note': 'Auto-write failed: ' + str(ue),
                })
                counts['queued'] += 1
        else:
            _queue_person_change({
                'tpPeopleId': int(tp_people_id),
                'tpName': safe_str(getattr(person, 'Name2', '') or getattr(person, 'Name', '')),
                'field': key,
                'fieldLabel': f['label'],
                'tpValue': _tp_person_field(person, f['tpField']),
                'pcoValue': pco_data.get(key, ''),
                'direction': direction,
            })
            counts['queued'] += 1
    return counts

def handle_list_pending_person_changes():
    try:
        queue = load_json(PERSON_PENDING_KEY, {'entries': []})
        if not isinstance(queue, dict):
            queue = {'entries': []}
        entries = [e for e in (queue.get('entries') or []) if e.get('status', 'pending') == 'pending']
        # Newest first.
        entries.sort(key=lambda e: e.get('queuedAt') or '', reverse=True)
        print json.dumps({
            'success': True,
            'entries': entries,
            'pendingCount': len(entries),
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List pending failed: ' + str(e)})

def _find_pending(entries, entry_id):
    for e in entries:
        if e.get('id') == entry_id:
            return e
    return None

def handle_apply_person_change():
    try:
        entry_id = safe_str(get_data('entry_id', '')).strip()
        if not entry_id:
            print json.dumps({'success': False, 'message': 'Missing entry_id'})
            return
        queue = load_json(PERSON_PENDING_KEY, {'entries': []})
        if not isinstance(queue, dict): queue = {'entries': []}
        entries = queue.get('entries') or []
        entry = _find_pending(entries, entry_id)
        if not entry:
            print json.dumps({'success': False, 'message': 'Entry not found.'})
            return
        # Look up the TP field name for this key.
        tp_field = ''
        for f in PERSON_SYNC_FIELDS:
            if f['key'] == entry.get('field'):
                tp_field = f['tpField']; break
        if not tp_field:
            print json.dumps({'success': False, 'message': 'Unknown field on entry: ' + str(entry.get('field'))})
            return
        try:
            model.UpdatePerson(int(entry['tpPeopleId']), tp_field, entry.get('pcoValue', ''))
        except Exception as ue:
            print json.dumps({'success': False, 'message': 'UpdatePerson failed: ' + str(ue)})
            return
        entry['status'] = 'applied'
        entry['appliedAt'] = now_iso()
        save_json(PERSON_PENDING_KEY, queue)
        append_audit({
            'action': 'person_data_apply',
            'tpPeopleId': int(entry['tpPeopleId']),
            'field': entry.get('field'),
            'newValue': entry.get('pcoValue', ''),
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Apply failed: ' + str(e)})

def handle_skip_person_change():
    try:
        entry_id = safe_str(get_data('entry_id', '')).strip()
        forever = safe_str(get_data('forever', '0')).strip().lower() in ('1', 'true', 'yes', 'on')
        if not entry_id:
            print json.dumps({'success': False, 'message': 'Missing entry_id'})
            return
        queue = load_json(PERSON_PENDING_KEY, {'entries': []})
        if not isinstance(queue, dict): queue = {'entries': []}
        entries = queue.get('entries') or []
        entry = _find_pending(entries, entry_id)
        if not entry:
            print json.dumps({'success': False, 'message': 'Entry not found.'})
            return
        entry['status'] = 'skipped_forever' if forever else 'skipped'
        entry['skippedAt'] = now_iso()
        save_json(PERSON_PENDING_KEY, queue)
        if forever:
            try:
                model.AddExtraValueText(int(entry['tpPeopleId']),
                                        _person_skip_forever_field_name(entry.get('field') or ''),
                                        '1')
            except:
                pass
        append_audit({
            'action': 'person_data_skip' + ('_forever' if forever else ''),
            'tpPeopleId': int(entry['tpPeopleId']),
            'field': entry.get('field'),
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Skip failed: ' + str(e)})

# =====================================================================
# Sync Dashboard tab
# =====================================================================
# Lists recent PCO plans across all mapped service types. PCO doesn't have
# a single "list all my plans" endpoint -- you walk per service type. We
# only walk MAPPED service types (no point showing plans we can't sync).

def _pco_all_people_page(offset=0, per_page=100, include_inactive=False):
    """Pull one page of the PCO People directory. Returns
    (list_of_people, next_offset_or_None, error_or_None).

    PCO stores emails as a sub-resource (Email records linked back to
    Person), so we request include=emails to get them inline in the
    response's included[] array. Without this, every Person.email is
    blank and the matcher can't use the email-exact signal."""
    path = '/people/v2/people?per_page=' + str(per_page) + '&offset=' + str(offset) + '&include=emails'
    if not include_inactive:
        # PCO's "status" attribute. "active" is the normal case; archived
        # / inactive users we usually don't want flooding the TP roster.
        path += '&where[status]=active'
    data, err = pco_get(path)
    if err:
        return [], None, err
    # Build a personId -> primary email lookup from the included[] array.
    # Each Email row has a relationship back to Person and an address +
    # optional primary flag. Prefer primary; fall back to first non-empty.
    email_by_pid = {}
    primary_seen = set()
    for inc in (data.get('included') or []):
        if inc.get('type') != 'Email':
            continue
        e_attrs = inc.get('attributes') or {}
        addr = safe_str(e_attrs.get('address', '')).strip().lower()
        if not addr:
            continue
        is_primary = bool(e_attrs.get('primary', False))
        rels = inc.get('relationships') or {}
        person_rel = ((rels.get('person') or {}).get('data') or {})
        pid = safe_str(person_rel.get('id', ''))
        if not pid:
            continue
        if is_primary:
            email_by_pid[pid] = addr
            primary_seen.add(pid)
        elif pid not in primary_seen and pid not in email_by_pid:
            email_by_pid[pid] = addr
    people = []
    for row in (data.get('data') or []):
        attrs = row.get('attributes') or {}
        bdate_raw = safe_str(attrs.get('birthdate', ''))
        bdate_iso = ''
        if bdate_raw and len(bdate_raw) >= 10:
            bdate_iso = bdate_raw[:10]
        pid = safe_str(row.get('id', ''))
        people.append({
            'pcoPersonId': pid,
            'first_name': safe_str(attrs.get('first_name', '')),
            'last_name':  safe_str(attrs.get('last_name', '')),
            'name': safe_str(attrs.get('name', '')),
            # Prefer email from the included[] Email rows; fall back to
            # any inline attribute (some PCO setups expose this directly).
            'email': email_by_pid.get(pid) or safe_str(attrs.get('email_address', '') or attrs.get('email', '')).strip().lower(),
            'birthdate': bdate_iso,
            'status': safe_str(attrs.get('status', 'active')).lower(),
        })
    # Pagination: PCO returns meta.next.offset when more pages exist.
    next_offset = None
    meta = (data.get('meta') or {})
    next_meta = meta.get('next') or {}
    try:
        next_offset = int(next_meta.get('offset')) if next_meta.get('offset') is not None else None
    except:
        next_offset = None
    return people, next_offset, None

def _pco_all_people_walk(include_inactive=False, page_cap=200):
    """Walk every page of the PCO People directory. page_cap is a safety
    net -- 200 pages * 100/page = 20K records, enough for nearly any
    church. Returns (all_people, errors)."""
    out = []
    errors = []
    offset = 0
    pages = 0
    while True:
        page, next_offset, err = _pco_all_people_page(offset=offset, per_page=100, include_inactive=include_inactive)
        if err:
            errors.append('Offset ' + str(offset) + ': ' + err)
            break
        out.extend(page)
        pages += 1
        if next_offset is None or pages >= page_cap:
            break
        offset = next_offset
    return out, errors

def _pco_teams_for_service_type(service_type_id):
    """List teams under a service type. Used by the Team Mapping picker
    (after the user picks a Service Type, show its teams). Returns
    [{teamId, teamName}, ...]."""
    data, err = pco_get('/services/v2/service_types/' + str(service_type_id) + '/teams?per_page=100')
    if err:
        return [], err
    out = []
    for t in (data.get('data') or []):
        attrs = t.get('attributes') or {}
        out.append({
            'teamId': safe_str(t.get('id', '')),
            'teamName': safe_str(attrs.get('name', '')),
        })
    out.sort(key=lambda x: (x.get('teamName') or '').lower())
    return out, None

def _pco_team_positions(team_id):
    """List Team Positions on a team. Returns [{positionId, positionName}, ...]."""
    data, err = pco_get('/services/v2/teams/' + str(team_id) + '/team_positions?per_page=100')
    if err:
        return [], err
    out = []
    for p in (data.get('data') or []):
        attrs = p.get('attributes') or {}
        out.append({
            'positionId': safe_str(p.get('id', '')),
            'positionName': safe_str(attrs.get('name', '')),
        })
    return out, None

def _pco_team_people(team_id):
    """List Person resources currently on a team. The relationship is
    'person added to team' -- distinct from 'scheduled in a plan'.
    Each result has the linked Person via include=person which gives
    name/email without a separate /people/{id} fetch."""
    data, err = pco_get('/services/v2/teams/' + str(team_id) + '/people?per_page=100&include=person')
    if err:
        return [], err
    # Map included Person attrs by id.
    person_attrs_by_id = {}
    for inc in (data.get('included') or []):
        if inc.get('type') != 'Person':
            continue
        pid = safe_str(inc.get('id', ''))
        if not pid:
            continue
        a = inc.get('attributes') or {}
        person_attrs_by_id[pid] = {
            'first_name': safe_str(a.get('first_name', '')),
            'last_name':  safe_str(a.get('last_name', '')),
            'nickname':   safe_str(a.get('nickname', '')),
        }
    # PCO returns one Person object per team row -- but the team membership
    # row itself is a "Person" type with relationships. Be lenient about
    # the shape: try 'attributes.first_name' on the row, fall back to the
    # included Person if needed.
    out = []
    seen_ids = set()
    for row in (data.get('data') or []):
        # On /teams/{id}/people, each row's id IS the person id and its
        # attributes ARE the person attrs.
        pid = safe_str(row.get('id', ''))
        if not pid or pid in seen_ids:
            continue
        seen_ids.add(pid)
        attrs = row.get('attributes') or {}
        person_extra = person_attrs_by_id.get(pid, {})
        first = safe_str(attrs.get('first_name', '') or person_extra.get('first_name', ''))
        last = safe_str(attrs.get('last_name', '') or person_extra.get('last_name', ''))
        # Email is on the Person resource but isn't always inline; left
        # empty here -- per-person sync still works for name fields.
        name = (first + ' ' + last).strip() or safe_str(attrs.get('name', '') or '(Unknown)')
        out.append({
            'pcoPersonId': pid,
            'name': name,
            'pcoFirstName': first,
            'pcoLastName': last,
            'email': safe_str(attrs.get('email_address', '')).strip().lower(),
        })
    return out, None

def _pco_team_position_assignments(position_id):
    """List person assignments for a Team Position. Used to figure out
    which positions each person on the team is eligible for. Returns
    list of person ids."""
    path = '/services/v2/team_positions/' + str(position_id) + '/person_team_position_assignments?per_page=100&include=person'
    data, err = pco_get(path)
    if err:
        return [], err
    out = []
    for row in (data.get('data') or []):
        rels = row.get('relationships') or {}
        person_rel = ((rels.get('person') or {}).get('data') or {})
        pid = safe_str(person_rel.get('id', ''))
        if pid:
            out.append(pid)
    return out, None

def _pco_team_all_position_assignments(team_id):
    """Walk every PersonTeamPositionAssignment under a team in ONE call.
    Returns (position_id -> [person_id, ...], err).

    The per-position endpoint sometimes silently returns empty even when
    the UI shows clear assignments (observed v3.1+). The team-level
    endpoint with includes returns everything in one shot and has been
    more reliable. We try it first; the caller falls back to per-position
    iteration if it errors or returns empty.

    The included resources include both Person and TeamPosition rows,
    so we can join the assignment back to its position via the
    relationships.team_position.data.id field."""
    path = '/services/v2/teams/' + str(team_id) + '/person_team_position_assignments?per_page=100&include=person,team_position'
    data, err = pco_get(path)
    if err:
        return {}, err
    by_position = {}
    page_count = 0
    while True:
        page_count += 1
        for row in (data.get('data') or []):
            rels = row.get('relationships') or {}
            person_rel = ((rels.get('person') or {}).get('data') or {})
            pos_rel = ((rels.get('team_position') or {}).get('data') or {})
            pid = safe_str(person_rel.get('id', ''))
            tpid = safe_str(pos_rel.get('id', ''))
            if pid and tpid:
                by_position.setdefault(tpid, []).append(pid)
        # Follow pagination if PCO paginated us. Most teams stay under
        # one page (100), but worship teams with 20+ positions x 10+
        # assignees can exceed that.
        links = (data.get('links') or {})
        nxt = links.get('next', '')
        if not nxt or page_count >= 10:
            break
        # links.next is a fully-qualified URL; pco_get expects a path
        # starting with /services/...
        try:
            from urlparse import urlparse  # py2
        except ImportError:
            from urllib.parse import urlparse  # py3
        try:
            parsed = urlparse(nxt)
            nxt_path = parsed.path + ('?' + parsed.query if parsed.query else '')
            data, err = pco_get(nxt_path)
            if err:
                break
        except:
            break
    return by_position, None

def _tp_org_members_with_pco_link(org_id):
    """Fetch active TP org members and their PCO_PersonId extra value.
    Returns dict {tpPeopleId: pcoPersonId or ''}.
    Used by mirror sync to find TP members whose PCO link no longer
    exists on the PCO side, so they can be dropped from the org."""
    sql = ("""
        SELECT om.PeopleId,
               COALESCE(NULLIF(pe.Data, ''), pe.StrValue, '') AS PCOID
        FROM OrganizationMembers om WITH (NOLOCK)
        LEFT JOIN PeopleExtra pe WITH (NOLOCK)
            ON pe.PeopleId = om.PeopleId AND pe.Field = '%s'
        WHERE om.OrganizationId = %s
          AND om.InactiveDate IS NULL
    """ % (PCO_PERSON_ID_FIELD, int(org_id)))
    out = {}
    try:
        for r in q.QuerySql(sql):
            out[int(r.PeopleId)] = safe_str(getattr(r, 'PCOID', ''))
    except:
        pass
    return out

def _tp_org_subgroups(org_id):
    """Fetch every (peopleId, subgroupName) row for active members of
    the org. Returns dict {tpPeopleId: set(subgroupName lower)} so we
    can do case-insensitive membership checks against PCO position
    names. Subgroup membership lives in OrgMemMemTags + MemberTags."""
    sql = ("""
        SELECT ommt.PeopleId, mt.Name AS SubName
        FROM OrgMemMemTags ommt WITH (NOLOCK)
        JOIN MemberTags mt WITH (NOLOCK) ON mt.Id = ommt.MemberTagId
        JOIN OrganizationMembers om WITH (NOLOCK)
            ON om.OrganizationId = mt.OrgId AND om.PeopleId = ommt.PeopleId
        WHERE mt.OrgId = %s
          AND om.InactiveDate IS NULL
    """ % int(org_id))
    out = {}
    try:
        for r in q.QuerySql(sql):
            pid = int(r.PeopleId)
            name = safe_str(getattr(r, 'SubName', '')).strip()
            if not name:
                continue
            out.setdefault(pid, set()).add(name.lower())
    except:
        pass
    return out

def _pco_plans_for_service_type(service_type_id, days_back, days_forward):
    """Pull recent + upcoming plans for one service type. PCO returns
    plans ordered by sort_date desc. We use the date filter to bound the
    window and per_page to cap returned rows."""
    # PCO date filter format: filter[sort_date]=YYYY-MM-DD...YYYY-MM-DD is
    # NOT supported; use filter=past + per_page or pass after/before in
    # the query (depending on API version). Simpler: just pull the most
    # recent N and let the Python side trim by date.
    path = ('/services/v2/service_types/' + str(service_type_id) +
            '/plans?order=-sort_date&per_page=20')
    data, err = pco_get(path)
    if err:
        return [], err
    plans = data.get('data') or []
    # Trim to the requested window. PCO sort_date is "YYYY-MM-DDTHH:MM:SSZ".
    today = datetime.datetime.now().date()
    earliest = today - datetime.timedelta(days=days_back)
    latest = today + datetime.timedelta(days=days_forward)
    out = []
    for p in plans:
        attrs = p.get('attributes') or {}
        sort_date_raw = attrs.get('sort_date', '')
        plan_date = None
        if sort_date_raw:
            try:
                plan_date = datetime.datetime.strptime(sort_date_raw[:10], '%Y-%m-%d').date()
            except:
                plan_date = None
        if plan_date is None:
            continue
        if plan_date < earliest or plan_date > latest:
            continue
        # Return raw fields so the client can compose the display the same
        # way PCO's own UI does: service type name = primary title, plan
        # title (if any) = secondary subtitle. PCO plans almost never have
        # a populated title field for routine services, so the service
        # type name is what the worship director thinks of as "the title."
        raw_title = safe_str(attrs.get('title', ''))
        series = safe_str(attrs.get('series_title', ''))
        short_dates = safe_str(attrs.get('short_dates', ''))
        out.append({
            'planId': str(p.get('id') or ''),
            'serviceTypeId': str(service_type_id),
            'planTitle': raw_title or series,  # often empty
            'shortDates': short_dates,
            'sortDate': sort_date_raw,
            'planDateIso': plan_date.strftime('%Y-%m-%d'),
        })
    return out, None

def handle_list_recent_plans():
    """v3.0+: Aggregates dashboard data across the three mapping types
    (All People, Service Type, Team). Plan cards come from Service Type
    Mappings that have perPlanAttendance=True. Roster cards come from
    every Service Type / Team / All People mapping. No more separate
    Service Plan Mappings storage."""
    try:
        days_back = safe_int(get_data('days_back', 30), 30)
        days_forward = safe_int(get_data('days_forward', 7), 7)

        # Load all three mapping stores up front.
        people_map_raw = load_json(PEOPLE_MAPPINGS_KEY, {})
        team_map_raw = load_json(TEAM_MAPPINGS_KEY, {})
        all_people_raw = load_json(ALL_PEOPLE_MAPPING_KEY, {})
        if not isinstance(people_map_raw, dict): people_map_raw = {}
        if not isinstance(team_map_raw, dict): team_map_raw = {}
        ap_info = _parse_all_people_mapping(all_people_raw)

        # Early exit -- nothing mapped.
        nothing_mapped = (not people_map_raw and not team_map_raw and ap_info['orgId'] <= 0)
        if nothing_mapped:
            print json.dumps({
                'success': True, 'plans': [], 'peopleMappings': [], 'teamMappings': [],
                'message': 'No mappings configured yet. Open the Sync Mappings tab to set one up before syncing.'
            })
            return

        # Service type names -- needed both for plan cards and roster
        # card labels.
        st_name_by_id = {}
        st_data, st_err = pco_get('/services/v2/service_types?per_page=100')
        if not st_err and isinstance(st_data, dict):
            for st in (st_data.get('data') or []):
                sid = str(st.get('id') or '')
                attrs = st.get('attributes') or {}
                st_name_by_id[sid] = safe_str(attrs.get('name', ''))

        # Hydrate every referenced TP org in one SQL pass.
        all_org_ids = set()
        for v in people_map_raw.values():
            i = _parse_people_mapping(v)
            if i['orgId'] > 0: all_org_ids.add(i['orgId'])
        for v in team_map_raw.values():
            i = _parse_team_mapping(v)
            if i['orgId'] > 0: all_org_ids.add(i['orgId'])
        if ap_info['orgId'] > 0:
            all_org_ids.add(ap_info['orgId'])
        org_info_by_id = {}
        if all_org_ids:
            ids_csv = ','.join(str(i) for i in all_org_ids)
            try:
                for r in q.QuerySql("SELECT OrganizationId, OrganizationName FROM Organizations WITH (NOLOCK) WHERE OrganizationId IN (%s)" % ids_csv):
                    org_info_by_id[int(r.OrganizationId)] = safe_str(r.OrganizationName)
            except:
                pass

        # Walk plans for any Service Type Mapping with perPlanAttendance.
        per_plan_plans = []
        api_errors = []
        for pco_st_id, raw_mapping in people_map_raw.items():
            info = _parse_people_mapping(raw_mapping)
            if not info['perPlanAttendance']:
                continue
            if info['orgId'] <= 0:
                continue
            plans, err = _pco_plans_for_service_type(pco_st_id, days_back, days_forward)
            if err:
                api_errors.append('Service type ' + str(pco_st_id) + ': ' + err)
                continue
            st_name = info['pcoServiceTypeName'] or st_name_by_id.get(pco_st_id, '(unknown)')
            tp_org_name = org_info_by_id.get(info['orgId'], '[Org #' + str(info['orgId']) + ' not found]')
            for plan in plans:
                plan['serviceTypeName'] = st_name
                plan['tpOrgId'] = info['orgId']
                plan['tpOrgName'] = tp_org_name
                # Surface the mapping's behavior so the Preview modal
                # can compose the right action set.
                plan['syncAttendance'] = True
                plan['autoAddMember'] = info['autoAddMember']
                plan['teamsAsSubgroups'] = info['teamsAsSubgroups']
                per_plan_plans.append(plan)
        per_plan_plans.sort(key=lambda p: p.get('planDateIso', ''), reverse=True)
        per_plan_plans = per_plan_plans[:100]

        # Last-sync timestamps per mapping.
        last_sync = _get_last_sync_times()
        for p in per_plan_plans:
            p['lastSyncedAt'] = last_sync['sync_plan'].get(safe_str(p.get('planId', '')))

        # roster_rollups concept dropped in v3.0 -- the rollup section
        # was a workaround for Service Plan Mappings with attendance off.
        # With v3.0's unified model, every Service Type Mapping always
        # syncs the durable roster, and attendance is a layered toggle.
        # The "Service Type Sync" section on the Dashboard now covers
        # everything.
        rollup_list = []
        extra_org_names = org_info_by_id  # alias to avoid renaming below

        extra_org_ids = set()
        for v in people_map_raw.values():
            i = _parse_people_mapping(v)
            if i['orgId'] > 0: extra_org_ids.add(i['orgId'])
        for v in team_map_raw.values():
            i = _parse_team_mapping(v)
            if i['orgId'] > 0: extra_org_ids.add(i['orgId'])
        ap_info = _parse_all_people_mapping(all_people_raw)
        if ap_info['orgId'] > 0:
            extra_org_ids.add(ap_info['orgId'])

        extra_org_names = {}
        if extra_org_ids:
            ids_csv = ','.join(str(i) for i in extra_org_ids)
            esql = "SELECT OrganizationId, OrganizationName FROM Organizations WITH (NOLOCK) WHERE OrganizationId IN (%s)" % ids_csv
            try:
                for r in q.QuerySql(esql):
                    extra_org_names[int(r.OrganizationId)] = safe_str(r.OrganizationName)
            except:
                pass

        people_mapping_cards = []
        for st_id, v in people_map_raw.items():
            info = _parse_people_mapping(v)
            key = info['pcoServiceTypeId'] or st_id
            sch = info.get('schedule', _parse_schedule({}))
            nxt = _next_scheduled_run(sch)
            people_mapping_cards.append({
                'pcoServiceTypeId': key,
                'pcoServiceTypeName': info['pcoServiceTypeName'],
                'tpOrgId': info['orgId'],
                'tpOrgName': extra_org_names.get(info['orgId'], '[Org #' + str(info['orgId']) + ' not found]') if info['orgId'] > 0 else '',
                'teamsAsSubgroups': info['teamsAsSubgroups'],
                'lastSyncedAt': last_sync['sync_people'].get(safe_str(key)),
                'scheduleLabel': _format_schedule_label(sch),
                'nextRunIso': nxt.strftime('%Y-%m-%dT%H:%M:%S') if nxt else '',
            })
        people_mapping_cards.sort(key=lambda r: (r.get('pcoServiceTypeName') or '').lower())

        team_mapping_cards = []
        for team_id, v in team_map_raw.items():
            info = _parse_team_mapping(v)
            key = info['pcoTeamId'] or team_id
            sch = info.get('schedule', _parse_schedule({}))
            nxt = _next_scheduled_run(sch)
            team_mapping_cards.append({
                'pcoTeamId': key,
                'pcoTeamName': info['pcoTeamName'],
                'pcoServiceTypeName': info['pcoServiceTypeName'],
                'tpOrgId': info['orgId'],
                'tpOrgName': extra_org_names.get(info['orgId'], '[Org #' + str(info['orgId']) + ' not found]') if info['orgId'] > 0 else '',
                'positionsAsSubgroups': info['positionsAsSubgroups'],
                'lastSyncedAt': last_sync['sync_team'].get(safe_str(key)),
                'scheduleLabel': _format_schedule_label(sch),
                'nextRunIso': nxt.strftime('%Y-%m-%dT%H:%M:%S') if nxt else '',
            })
        team_mapping_cards.sort(key=lambda r: ((r.get('pcoServiceTypeName') or '').lower(),
                                                (r.get('pcoTeamName') or '').lower()))

        all_people_card = None
        if ap_info['orgId'] > 0:
            sch = ap_info.get('schedule', _parse_schedule({}))
            nxt = _next_scheduled_run(sch)
            all_people_card = {
                'tpOrgId': ap_info['orgId'],
                'tpOrgName': extra_org_names.get(ap_info['orgId'], '[Org #' + str(ap_info['orgId']) + ' not found]'),
                'includeInactive': ap_info['includeInactive'],
                'lastSyncedAt': last_sync['all_people'],
                'scheduleLabel': _format_schedule_label(sch),
                'nextRunIso': nxt.strftime('%Y-%m-%dT%H:%M:%S') if nxt else '',
            }

        resp = {
            'success': True,
            'plans': per_plan_plans,
            'planCount': len(per_plan_plans),
            'rosterRollups': rollup_list,
            'rosterRollupCount': len(rollup_list),
            'peopleMappings': people_mapping_cards,
            'teamMappings': team_mapping_cards,
            'allPeopleMapping': all_people_card,
        }
        if api_errors:
            resp['warnings'] = api_errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List plans failed: ' + str(e)})

# =====================================================================
# Per-plan review + sync write
# =====================================================================
# When the user clicks "Preview & Sync" on a plan card, we fetch the plan's
# people from PCO, match each to a TouchPoint PeopleId using PCO_PersonId
# extra value (canonical) or email (fallback). Unmatched rows get a manual
# search picker. Confirming a manual match writes the PCO_PersonId extra
# value so the same person is auto-matched on every future plan.
# On Sync: GetMeetingIdByDateTime resolves the TP meeting for the plan
# date, then we EditPersonAttendance(True) per matched attendee with
# post-verify, and write an audit log row per write.

# Confirmed = scheduled and the person said they'd be there. Other letters
# (Unconfirmed, Declined) exist but we don't sync them as attended.
PCO_STATUS_CONFIRMED = 'C'

def _tp_match_by_pco_id(pco_person_ids):
    """One-shot lookup of TP PeopleIds with a matching PCO_PersonId extra
    value. Returns {pcoPersonId(str): {peopleId, name}}.

    Storage column: model.AddExtraValueText -> Person.AddEditExtraData,
    which writes to PeopleExtra.Data (NOT StrValue). StrValue is used by
    Code-type extras only. We check both so external admin edits made
    through the TouchPoint UI (which may pick a slightly different column)
    are still picked up."""
    out = {}
    if not pco_person_ids:
        return out
    safe_ids = []
    for pid in pco_person_ids:
        sp = safe_str(pid).replace("'", "''")
        if sp:
            safe_ids.append("'" + sp + "'")
    if not safe_ids:
        return out
    ids_csv = ','.join(safe_ids)
    sql = ("""
        SELECT COALESCE(NULLIF(pe.Data, ''), pe.StrValue) AS PcoPersonId,
               p.PeopleId, p.Name2
        FROM PeopleExtra pe WITH (NOLOCK)
        JOIN People p WITH (NOLOCK) ON p.PeopleId = pe.PeopleId
        WHERE pe.Field = '%s'
          AND (pe.Data IN (%s) OR pe.StrValue IN (%s))
          AND p.IsDeceased = 0 AND p.ArchivedFlag = 0
    """ % (PCO_PERSON_ID_FIELD, ids_csv, ids_csv))
    try:
        for r in q.QuerySql(sql):
            out[safe_str(r.PcoPersonId)] = {
                'peopleId': int(r.PeopleId),
                'name': safe_str(r.Name2),
            }
    except:
        pass
    return out

def _tp_match_by_name(name_pairs):
    """Batched candidate lookup by (first, last). Returns dict keyed by
    '<first_lower>|<last_lower>' -> list of TP records with email +
    birthdate so the caller can score the match.

    Looks at both FirstName and NickName (so "Bob" matches "Robert"
    with NickName='Bob') and against LastName. Uses the existing
    IsDeceased + ArchivedFlag filters to keep results clean."""
    out = {}
    if not name_pairs:
        return out
    seen = set()
    where_clauses = []
    for first, last in name_pairs:
        first = safe_str(first).strip().lower()
        last = safe_str(last).strip().lower()
        if not first or not last:
            continue
        key = first + '|' + last
        if key in seen:
            continue
        seen.add(key)
        f_esc = first.replace("'", "''")
        l_esc = last.replace("'", "''")
        where_clauses.append("(LOWER(p.LastName) = '" + l_esc + "' AND "
                             + "(LOWER(p.FirstName) = '" + f_esc + "' OR LOWER(p.NickName) = '" + f_esc + "'))")
    if not where_clauses:
        return out
    # Chunk to keep the SQL statement size sane.
    CHUNK = 100
    for i in range(0, len(where_clauses), CHUNK):
        chunk = where_clauses[i:i+CHUNK]
        sql = ("""
            SELECT p.PeopleId, p.Name2, p.FirstName, p.NickName, p.LastName,
                   ISNULL(p.EmailAddress, '') AS Email,
                   p.BirthYear, p.BirthMonth, p.BirthDay
            FROM People p WITH (NOLOCK)
            WHERE (%s)
              AND p.IsDeceased = 0 AND p.ArchivedFlag = 0
        """ % ' OR '.join(chunk))
        try:
            for r in q.QuerySql(sql):
                # Build matching keys for both FirstName and NickName.
                last_lower = safe_str(r.LastName).lower()
                tp_first = safe_str(r.FirstName).lower()
                tp_nick = safe_str(r.NickName).lower()
                # Build a YYYY-MM-DD birthdate if possible (year may be
                # 0 for kids whose year wasn't recorded).
                by = safe_int(r.BirthYear, 0)
                bm = safe_int(r.BirthMonth, 0)
                bd = safe_int(r.BirthDay, 0)
                tp_bdate = ''
                if bm > 0 and bd > 0:
                    tp_bdate = '%04d-%02d-%02d' % (by if by > 0 else 0, bm, bd)
                rec = {
                    'peopleId': int(r.PeopleId),
                    'name': safe_str(r.Name2),
                    'firstName': safe_str(r.FirstName),
                    'nickName': safe_str(r.NickName),
                    'lastName': safe_str(r.LastName),
                    'email': safe_str(r.Email).strip().lower(),
                    'birthdate': tp_bdate,
                }
                for f in [tp_first, tp_nick]:
                    if f and last_lower:
                        out.setdefault(f + '|' + last_lower, []).append(rec)
        except:
            pass
    return out

def _bdate_match_score(pco_bd, tp_bd):
    """Compare two YYYY-MM-DD strings. Allows year-missing matches
    (0000-MM-DD) -- PCO sometimes has full birthdate when TP doesn't,
    so we accept month+day as a partial signal."""
    if not pco_bd or not tp_bd or len(pco_bd) < 10 or len(tp_bd) < 10:
        return 0
    p_year, p_md = pco_bd[:4], pco_bd[5:10]
    t_year, t_md = tp_bd[:4], tp_bd[5:10]
    if p_md != t_md:
        return 0
    # Month+day match. Strong if years also match (or one side is 0000).
    if p_year == t_year:
        return 2  # full match
    if p_year == '0000' or t_year == '0000':
        return 1  # month/day match, year missing on one side
    return 0  # mm/dd same, but year differs -> probably different person

def _score_pco_to_tp_candidates(pco_person, by_name_lookup, by_email_lookup):
    """Build the proposed candidate list for one unmatched PCO person.
    Each candidate has score (0-100) and tier label so the UI can sort/
    bulk-act. Tiers used by the review UI:
      strong  -- score >= 90 (almost certainly the same person)
      medium  -- 70-89 (single name hit, plausible)
      weak    -- below 70 or ambiguous"""
    candidates = {}  # tp_people_id -> candidate dict
    first = (pco_person.get('first_name') or '').strip().lower()
    last = (pco_person.get('last_name') or '').strip().lower()
    email = (pco_person.get('email') or '').strip().lower()
    pco_bd = pco_person.get('birthdate') or ''

    def _add(tp_rec, base_score, signals):
        pid = tp_rec['peopleId']
        if pid in candidates:
            if base_score > candidates[pid]['score']:
                candidates[pid]['score'] = base_score
                candidates[pid]['signals'] = signals
            return
        candidates[pid] = {
            'tpPeopleId': pid,
            'tpName': tp_rec['name'],
            'tpFirstName': tp_rec.get('firstName', ''),
            'tpLastName': tp_rec.get('lastName', ''),
            'tpEmail': tp_rec['email'],
            'tpBirthdate': tp_rec.get('birthdate', ''),
            'score': base_score,
            'signals': signals,
        }

    # Email-exact path -- highest non-canonical confidence.
    if email and email in by_email_lookup:
        hits = by_email_lookup[email]
        for h in hits:
            score = 95 if len(hits) == 1 else 75
            sig = ['email-exact']
            if len(hits) > 1:
                sig.append('email-ambiguous')
            _add(h, score, sig)

    # Name path -- score boosted by birthdate / email signals.
    if first and last:
        key = first + '|' + last
        name_hits = by_name_lookup.get(key, [])
        ambiguous = len(name_hits) > 1
        for h in name_hits:
            base = 70 if not ambiguous else 55
            sig = ['name-exact']
            if ambiguous:
                sig.append('name-ambiguous')
            bd_score = _bdate_match_score(pco_bd, h.get('birthdate', ''))
            if bd_score == 2:
                base = 92
                sig.append('birthdate-full')
            elif bd_score == 1:
                base = 84
                sig.append('birthdate-md')
            if email and h.get('email') == email:
                base = max(base, 95)
                sig.append('email-also')
            _add(h, base, sig)

    out = list(candidates.values())
    out.sort(key=lambda c: -c['score'])
    return out

def _tier_for_score(score):
    if score >= 90: return 'strong'
    if score >= 70: return 'medium'
    return 'weak'

def load_all_people_skip():
    raw = load_json(ALL_PEOPLE_SKIP_KEY, {})
    if not isinstance(raw, dict):
        return set()
    return set(safe_str(p) for p in (raw.get('pcoIds') or []))

def save_all_people_skip(skip_set):
    save_json(ALL_PEOPLE_SKIP_KEY, {'pcoIds': sorted(list(skip_set))})

def _tp_match_by_email(emails):
    """One-shot lookup by email. Returns {email_lower: [{peopleId, name,
    firstName, lastName, email, birthdate}, ...]}. v2.5.7+ includes the
    name + birthdate fields so the Proposed Matches scorer has a
    consistent shape across email and name lookups. Older callers that
    just read peopleId/name/email keep working unchanged."""
    out = {}
    if not emails:
        return out
    safe_emails = []
    seen = set()
    for e in emails:
        e = safe_str(e).strip().lower()
        if e and e not in seen:
            seen.add(e)
            safe_emails.append("'" + e.replace("'", "''") + "'")
    if not safe_emails:
        return out
    sql = ("""
        SELECT p.PeopleId, p.Name2, p.FirstName, p.NickName, p.LastName,
               p.EmailAddress, p.EmailAddress2,
               p.BirthYear, p.BirthMonth, p.BirthDay
        FROM People p WITH (NOLOCK)
        WHERE (LOWER(p.EmailAddress) IN (%s) OR LOWER(p.EmailAddress2) IN (%s))
          AND p.IsDeceased = 0 AND p.ArchivedFlag = 0
    """ % (','.join(safe_emails), ','.join(safe_emails)))
    try:
        for r in q.QuerySql(sql):
            by = safe_int(r.BirthYear, 0)
            bm = safe_int(r.BirthMonth, 0)
            bd = safe_int(r.BirthDay, 0)
            tp_bdate = ''
            if bm > 0 and bd > 0:
                tp_bdate = '%04d-%02d-%02d' % (by if by > 0 else 0, bm, bd)
            base = {
                'peopleId': int(r.PeopleId),
                'name': safe_str(r.Name2),
                'firstName': safe_str(r.FirstName),
                'nickName': safe_str(r.NickName),
                'lastName': safe_str(r.LastName),
                'birthdate': tp_bdate,
            }
            for em_attr in ('EmailAddress', 'EmailAddress2'):
                em = safe_str(getattr(r, em_attr, '')).lower()
                if em:
                    rec = dict(base)
                    rec['email'] = em
                    out.setdefault(em, []).append(rec)
    except:
        pass
    return out

def handle_load_plan_preview():
    """Fetch a plan's people from PCO and resolve each to a TP PeopleId
    using PCO_PersonId extra value first, then email fallback.

    Plan title fallback: PCO plans frequently have empty title and
    series_title. PCO's own UI shows "<Service Type Name> -- <Short Dates>".
    We mirror that so the modal header is always meaningful.

    De-dup by PCO Person ID: a person can hold multiple positions on the
    same plan (Welcome + Invitation Counselor, etc.). Show one row per
    unique person with all positions/statuses combined. Sync once."""
    try:
        plan_id = safe_str(get_data('plan_id', '')).strip()
        service_type_id = safe_str(get_data('service_type_id', '')).strip()
        if not plan_id or not service_type_id:
            print json.dumps({'success': False, 'message': 'Missing plan_id or service_type_id'})
            return

        # 1. Plan info.
        plan_path = '/services/v2/service_types/' + service_type_id + '/plans/' + plan_id
        plan_data, err = pco_get(plan_path)
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        plan_attrs = ((plan_data or {}).get('data') or {}).get('attributes') or {}

        # 2. Service type name (one extra call, but worth it for a readable title).
        st_name = ''
        try:
            st_data, st_err = pco_get('/services/v2/service_types/' + service_type_id)
            if not st_err:
                st_attrs = ((st_data or {}).get('data') or {}).get('attributes') or {}
                st_name = safe_str(st_attrs.get('name', ''))
        except:
            pass

        raw_title = safe_str(plan_attrs.get('title', ''))
        series_title = safe_str(plan_attrs.get('series_title', ''))
        short_dates = safe_str(plan_attrs.get('short_dates', ''))
        # Title composition mirrors PCO's own UI: service type name is the
        # primary identifier (always shown big in PCO's plan list); the
        # plan's raw title only shows up if explicitly set, as a subtitle.
        plan_info = {
            'planId': plan_id,
            'title': st_name or '(Untitled Plan)',  # primary = service type
            'planTitle': raw_title or series_title,  # secondary, often empty
            'serviceTypeName': st_name,
            'sortDate': safe_str(plan_attrs.get('sort_date', '')),
            'shortDates': short_dates,
            'planDateIso': '',
        }
        if plan_info['sortDate']:
            try:
                plan_info['planDateIso'] = plan_info['sortDate'][:10]
            except:
                pass

        # v3.0+: pull the toggles from the unified Service Type Mapping.
        # Plan cards on the Dashboard only appear when the mapping has
        # perPlanAttendance ON, so syncAttendance is True by definition
        # for any plan we're previewing.
        mappings_for_mode = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings_for_mode, dict):
            mappings_for_mode = {}
        _info = _parse_people_mapping(mappings_for_mode.get(service_type_id))
        plan_info['syncAttendance'] = bool(_info['perPlanAttendance'])
        plan_info['autoAddMember'] = bool(_info['autoAddMember'])

        # 3. Plan team members. include=person pulls the linked Person
        # resource into the response's "included" array -- gives us
        # first/last name without a separate /people/{id} call per row.
        pp_path = '/services/v2/service_types/' + service_type_id + '/plans/' + plan_id + '/team_members?per_page=100&include=person'
        pp_data, err = pco_get(pp_path)
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        rows = pp_data.get('data') or []
        # Build {pcoPersonId: {first_name, last_name}} from the included array.
        person_attrs_by_id = {}
        for inc in (pp_data.get('included') or []):
            if inc.get('type') != 'Person':
                continue
            pid = safe_str(inc.get('id', ''))
            if not pid:
                continue
            a = inc.get('attributes') or {}
            person_attrs_by_id[pid] = {
                'first_name': safe_str(a.get('first_name', '')),
                'last_name': safe_str(a.get('last_name', '')),
                'nickname': safe_str(a.get('nickname', '') or a.get('given_name', '')),
            }

        # Group by PCO Person ID. Each unique person becomes one attendee
        # row with positions[] and statuses[]. Confirmed if ANY assignment
        # is status='C'. (TouchPoint attendance is per-person-per-meeting,
        # not per-position.)
        by_pid = {}            # pcoPersonId -> attendee row
        no_pid_rows = []        # team_member entries without a linked person
        for tm in rows:
            tm_attrs = tm.get('attributes') or {}
            rels = tm.get('relationships') or {}
            person_rel = ((rels.get('person') or {}).get('data') or {})
            pco_person_id = safe_str(person_rel.get('id', ''))
            name = safe_str(tm_attrs.get('name', '') or tm_attrs.get('person_name', '') or '(Unknown)')
            status = safe_str(tm_attrs.get('status', ''))
            team_pos = safe_str(tm_attrs.get('team_position_name', ''))
            email = safe_str(tm_attrs.get('email', '') or tm_attrs.get('email_address', '')).strip().lower()

            assignment = {'teamPosition': team_pos, 'status': status,
                          'pcoTeamMemberId': safe_str(tm.get('id', ''))}

            person_extra = person_attrs_by_id.get(pco_person_id, {}) if pco_person_id else {}
            pco_first = person_extra.get('first_name', '')
            pco_last = person_extra.get('last_name', '')

            if not pco_person_id:
                # No linked Person record -- can't dedupe. Keep as its own row.
                no_pid_rows.append({
                    'pcoPersonId': '',
                    'name': name,
                    'email': email,
                    'pcoFirstName': pco_first,
                    'pcoLastName': pco_last,
                    'positions': [assignment],
                    'isConfirmed': (status == PCO_STATUS_CONFIRMED),
                    'tpPeopleId': None, 'tpName': '', 'matchSource': '',
                    'emailAmbiguous': False,
                })
                continue

            if pco_person_id not in by_pid:
                by_pid[pco_person_id] = {
                    'pcoPersonId': pco_person_id,
                    'name': name,
                    'email': email,
                    'pcoFirstName': pco_first,
                    'pcoLastName': pco_last,
                    'positions': [],
                    'isConfirmed': False,
                    'tpPeopleId': None, 'tpName': '', 'matchSource': '',
                    'emailAmbiguous': False,
                }
            row = by_pid[pco_person_id]
            row['positions'].append(assignment)
            if status == PCO_STATUS_CONFIRMED:
                row['isConfirmed'] = True
            # Prefer non-empty email if a later row has one.
            if email and not row['email']:
                row['email'] = email

        # Bulk resolve in one round trip each.
        pco_person_ids = list(by_pid.keys())
        emails_to_check = [r['email'] for r in by_pid.values() if r['email']]
        emails_to_check += [r['email'] for r in no_pid_rows if r['email']]
        by_pco = _tp_match_by_pco_id(pco_person_ids)
        by_email = _tp_match_by_email(emails_to_check)

        def _resolve(row):
            if row['pcoPersonId'] and row['pcoPersonId'] in by_pco:
                tp = by_pco[row['pcoPersonId']]
                row['tpPeopleId'] = tp['peopleId']
                row['tpName'] = tp['name']
                row['matchSource'] = 'pco_id'
                return 'matched'
            if row['email'] and row['email'] in by_email:
                hits = by_email[row['email']]
                if len(hits) == 1:
                    row['tpPeopleId'] = hits[0]['peopleId']
                    row['tpName'] = hits[0]['name']
                    row['matchSource'] = 'email'
                    return 'matched'
                row['emailAmbiguous'] = True
                return 'ambiguous'
            return 'unmatched'

        attendees = list(by_pid.values()) + no_pid_rows

        # Resolve TP match BEFORE sorting so we can group by status. Rows
        # needing manual action (unmatched + ambiguous) go on top so the
        # worship admin can resolve them without scrolling past the already-
        # matched names. Inside each group, alphabetical by name.
        matched = ambiguous = unmatched = 0
        for a in attendees:
            outcome = _resolve(a)
            if outcome == 'matched': matched += 1
            elif outcome == 'ambiguous': ambiguous += 1
            else: unmatched += 1
            a['matchStatus'] = outcome

        # Group order: unmatched first, then ambiguous, then matched.
        # Alphabetical by name inside each group.
        _group_rank = {'unmatched': 0, 'ambiguous': 1, 'matched': 2}
        attendees.sort(key=lambda a: (
            _group_rank.get(a.get('matchStatus'), 99),
            (a.get('name') or '').lower(),
        ))

        print json.dumps({
            'success': True,
            'planInfo': plan_info,
            'attendees': attendees,
            'summary': {
                'total': len(attendees),
                'matched': matched,
                'ambiguousEmail': ambiguous,
                'unmatched': unmatched,
                'confirmed': sum(1 for a in attendees if a['isConfirmed']),
                'rawAssignments': len(rows),  # before grouping
            },
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load plan preview failed: ' + str(e)})

def handle_load_roster_preview():
    """Aggregate team members across ALL plans in the window for one
    service type. Used by the add-only roster sync flow -- the worship
    admin wants 'who is Confirmed somewhere in upcoming plans', not a
    per-plan dance."""
    try:
        service_type_id = safe_str(get_data('service_type_id', '')).strip()
        if not service_type_id:
            print json.dumps({'success': False, 'message': 'Missing service_type_id'})
            return
        days_back = safe_int(get_data('days_back', 30), 30)
        days_forward = safe_int(get_data('days_forward', 7), 7)

        # Verify mapping + capture flags. Even though this is the add-only
        # flow, we surface the flags so the modal can confirm.
        mappings = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        if service_type_id not in mappings:
            print json.dumps({'success': False, 'message': 'Service type is not mapped.'})
            return
        tp_org_id, sync_att, add_mem = _parse_mapping(mappings[service_type_id])
        if tp_org_id <= 0:
            print json.dumps({'success': False, 'message': 'Mapping has no TouchPoint org.'})
            return

        # Pull service type name + plans in window.
        st_name = ''
        try:
            st_data, st_err = pco_get('/services/v2/service_types/' + service_type_id)
            if not st_err:
                st_attrs = ((st_data or {}).get('data') or {}).get('attributes') or {}
                st_name = safe_str(st_attrs.get('name', ''))
        except:
            pass
        plans, err = _pco_plans_for_service_type(service_type_id, days_back, days_forward)
        if err:
            print json.dumps({'success': False, 'message': err})
            return

        # Walk plans -> team_members. Aggregate by PCO Person ID. Track:
        #   isConfirmed (True if confirmed in ANY plan)
        #   positions (all unique position+status across plans)
        #   plansSeen (count of distinct plans they appear in)
        # Same matching/grouping logic as load_plan_preview, just rolled up.
        agg = {}
        no_pid_rows = []
        api_errors = []
        for plan in plans:
            tm_path = ('/services/v2/service_types/' + service_type_id +
                       '/plans/' + plan['planId'] + '/team_members?per_page=100&include=person')
            tm_data, tm_err = pco_get(tm_path)
            if tm_err:
                api_errors.append('Plan ' + plan['planId'] + ': ' + tm_err)
                continue
            person_attrs_by_id = {}
            for inc in (tm_data.get('included') or []):
                if inc.get('type') != 'Person':
                    continue
                pid = safe_str(inc.get('id', ''))
                if not pid:
                    continue
                a = inc.get('attributes') or {}
                person_attrs_by_id[pid] = {
                    'first_name': safe_str(a.get('first_name', '')),
                    'last_name': safe_str(a.get('last_name', '')),
                }
            for tm in (tm_data.get('data') or []):
                tm_attrs = tm.get('attributes') or {}
                rels = tm.get('relationships') or {}
                person_rel = ((rels.get('person') or {}).get('data') or {})
                pco_person_id = safe_str(person_rel.get('id', ''))
                name = safe_str(tm_attrs.get('name', '') or tm_attrs.get('person_name', '') or '(Unknown)')
                status = safe_str(tm_attrs.get('status', ''))
                team_pos = safe_str(tm_attrs.get('team_position_name', ''))
                email = safe_str(tm_attrs.get('email', '') or tm_attrs.get('email_address', '')).strip().lower()
                person_extra = person_attrs_by_id.get(pco_person_id, {}) if pco_person_id else {}
                is_conf = (status == 'C')
                key = pco_person_id or ('NOPID_' + name + '_' + email)
                row = agg.get(key)
                if row is None:
                    row = {
                        'pcoPersonId': pco_person_id,
                        'name': name,
                        'email': email,
                        'pcoFirstName': person_extra.get('first_name', ''),
                        'pcoLastName': person_extra.get('last_name', ''),
                        'positions': [],
                        'isConfirmed': False,
                        'plansSeen': 0,
                        '_plan_ids': set(),
                        '_pos_keys': set(),
                    }
                    agg[key] = row
                row['isConfirmed'] = row['isConfirmed'] or is_conf
                if email and not row['email']:
                    row['email'] = email
                if person_extra.get('first_name') and not row.get('pcoFirstName'):
                    row['pcoFirstName'] = person_extra['first_name']
                if person_extra.get('last_name') and not row.get('pcoLastName'):
                    row['pcoLastName'] = person_extra['last_name']
                pos_key = (team_pos, status)
                if pos_key not in row['_pos_keys']:
                    row['_pos_keys'].add(pos_key)
                    row['positions'].append({'position': team_pos, 'status': status})
                if plan['planId'] not in row['_plan_ids']:
                    row['_plan_ids'].add(plan['planId'])
                    row['plansSeen'] += 1

        # Drop the sets before serializing; they're internal-only.
        for row in agg.values():
            row.pop('_plan_ids', None)
            row.pop('_pos_keys', None)

        # Match TP people (same as load_plan_preview).
        pco_person_ids = [r['pcoPersonId'] for r in agg.values() if r['pcoPersonId']]
        emails_to_check = [r['email'] for r in agg.values() if r['email']]
        by_pco = _tp_match_by_pco_id(pco_person_ids)
        by_email = _tp_match_by_email(emails_to_check)

        matched = ambiguous = unmatched = 0
        attendees = list(agg.values())
        for row in attendees:
            row['tpPeopleId'] = None
            row['tpName'] = ''
            row['matchSource'] = ''
            row['emailAmbiguous'] = False
            if row['pcoPersonId'] and row['pcoPersonId'] in by_pco:
                tp = by_pco[row['pcoPersonId']]
                row['tpPeopleId'] = tp['peopleId']
                row['tpName'] = tp['name']
                row['matchSource'] = 'pco_id'
                row['matchStatus'] = 'matched'
                matched += 1
                continue
            if row['email'] and row['email'] in by_email:
                hits = by_email[row['email']]
                if len(hits) == 1:
                    row['tpPeopleId'] = hits[0]['peopleId']
                    row['tpName'] = hits[0]['name']
                    row['matchSource'] = 'email'
                    row['matchStatus'] = 'matched'
                    matched += 1
                    continue
                row['emailAmbiguous'] = True
                row['matchStatus'] = 'ambiguous'
                ambiguous += 1
                continue
            row['matchStatus'] = 'unmatched'
            unmatched += 1

        # Sort: unmatched first (need action), then ambiguous, then matched. Name alpha within.
        _group_rank = {'unmatched': 0, 'ambiguous': 1, 'matched': 2}
        attendees.sort(key=lambda r: (
            _group_rank.get(r.get('matchStatus'), 99),
            (r.get('name') or '').lower(),
        ))

        # Compose plan_info so the JS modal can reuse the same render path
        # as load_plan_preview. isRosterSync=True tells the JS to swap the
        # action on Sync (no plan_id, no meeting).
        latest = ''
        for p in plans:
            iso = p.get('planDateIso', '')
            if iso and (not latest or iso > latest):
                latest = iso
        plan_info = {
            'isRosterSync': True,
            'serviceTypeId': service_type_id,
            'serviceTypeName': st_name,
            'title': st_name + ' (Roster Sync)',
            'planTitle': '',
            'planCount': len(plans),
            'latestPlanIso': latest,
            'shortDates': str(len(plans)) + ' plan(s)' + (' through ' + latest if latest else ''),
            'syncAttendance': sync_att,
            'autoAddMember': add_mem,
            'tpOrgId': tp_org_id,
        }

        resp = {
            'success': True,
            'planInfo': plan_info,
            'attendees': attendees,
            'summary': {
                'total': len(attendees),
                'matched': matched,
                'ambiguousEmail': ambiguous,
                'unmatched': unmatched,
                'confirmed': sum(1 for r in attendees if r['isConfirmed']),
                'rawAssignments': 0,
            },
        }
        if api_errors:
            resp['warnings'] = api_errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load roster preview failed: ' + str(e)})

def handle_sync_roster():
    """Add Confirmed + matched people to the involvement roster. No
    meeting, no attendance. Used by add-only roster sync."""
    try:
        service_type_id = safe_str(get_data('service_type_id', '')).strip()
        tp_org_id = safe_int(get_data('tp_org_id', 0), 0)
        people_ids_csv = safe_str(get_data('people_ids', '')).strip()
        if not service_type_id or tp_org_id <= 0 or not people_ids_csv:
            print json.dumps({'success': False, 'message': 'Missing required args'})
            return
        # Verify the mapping says auto-add. Defensive -- the dashboard
        # only shows roster cards for add-only mappings, but check anyway.
        mappings = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        mapped_org, _sync_att, add_mem = _parse_mapping(mappings.get(service_type_id))
        if mapped_org != tp_org_id:
            print json.dumps({'success': False, 'message': 'Mapping mismatch -- service type is mapped to a different org now.'})
            return
        if not add_mem:
            print json.dumps({'success': False, 'message': 'Auto-add is disabled for this mapping.'})
            return

        pids = []
        for tok in people_ids_csv.split(','):
            pid = safe_int(tok.strip(), 0)
            if pid > 0:
                pids.append(pid)
        if not pids:
            print json.dumps({'success': False, 'message': 'No people ids to sync'})
            return

        # Active-member pre-check (same as sync_plan_attendance fix).
        existing_member_ids = set()
        try:
            mem_sql = ("""
                SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                WHERE OrganizationId = %s
                  AND PeopleId IN (%s)
                  AND InactiveDate IS NULL
            """ % (int(tp_org_id), ','.join(str(p) for p in pids)))
            for r in q.QuerySql(mem_sql):
                existing_member_ids.add(int(r.PeopleId))
        except:
            pass

        # Person data sync hook (same pattern as sync_plan_attendance).
        person_rules = normalize_person_sync_rules(load_json(PERSON_SYNC_RULES_KEY, {}))
        person_rules_active = person_sync_rules_active(person_rules)
        pco_data_by_pid = _parse_people_json_payload(safe_str(get_data('people_json', '')))
        person_auto = 0
        person_queued = 0

        joined = 0
        already = 0
        skipped = 0
        per_person = []
        for pid in pids:
            try:
                if pid in existing_member_ids:
                    already += 1
                    per_person.append({'peopleId': pid, 'status': 'already_member'})
                else:
                    person = model.GetPerson(pid)
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(tp_org_id), person)
                            joined += 1
                            per_person.append({'peopleId': pid, 'status': 'joined'})
                        except Exception as je:
                            jm = str(je).lower()
                            if 'already' in jm:
                                already += 1
                                per_person.append({'peopleId': pid, 'status': 'already_member'})
                            else:
                                skipped += 1
                                per_person.append({'peopleId': pid, 'status': 'join_failed', 'message': str(je)})
                                continue
                    else:
                        skipped += 1
                        per_person.append({'peopleId': pid, 'status': 'no_person'})
                        continue
                # Person data sync after the roster decision is recorded.
                if person_rules_active and pid in pco_data_by_pid:
                    try:
                        c = apply_person_sync_for_one(pid, pco_data_by_pid[pid], person_rules)
                        person_auto += c.get('auto', 0)
                        person_queued += c.get('queued', 0)
                    except:
                        pass
            except Exception as we:
                skipped += 1
                per_person.append({'peopleId': pid, 'status': 'error', 'message': str(we)})

        append_audit({
            'action': 'sync_roster',
            'pcoServiceTypeId': service_type_id,
            'tpOrgId': tp_org_id,
            'joined': joined,
            'alreadyMember': already,
            'skipped': skipped,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })

        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict):
            s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)

        if joined == 0 and already > 0:
            msg = 'Everyone was already an active member -- nothing new to add.'
        else:
            msg = 'Added ' + str(joined) + ' member(s) to the involvement.'
            if already > 0:
                msg += ' (' + str(already) + ' already on the roster.)'
        if person_auto:
            msg += ' Updated ' + str(person_auto) + ' person field(s).'
        if person_queued:
            msg += ' Queued ' + str(person_queued) + ' person change(s) for review.'
        print json.dumps({
            'success': True,
            'message': msg,
            'joinedOrg': joined,
            'alreadyMember': already,
            'skipped': skipped,
            'personAuto': person_auto,
            'personQueued': person_queued,
            'perPerson': per_person,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Sync roster failed: ' + str(e)})

# =====================================================================
# Team Sync (v2.0+)
# =====================================================================
# Sources the durable PCO Team membership (people added to the team, not
# people scheduled in a specific plan). Optional: positions become TP
# subgroups so the same data is filter-able in the involvement view.

def handle_load_team_sync_preview():
    """Preview a Team Sync: list team members + their position
    eligibility, resolved to TP people. Same modal shape as the per-plan
    preview so the JS can reuse renderPlanPreviewBody."""
    try:
        pco_team_id = safe_str(get_data('pco_team_id', '')).strip()
        if not pco_team_id:
            print json.dumps({'success': False, 'message': 'Missing pco_team_id'})
            return
        mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        if pco_team_id not in mappings:
            print json.dumps({'success': False, 'message': 'No mapping found for this team.'})
            return
        info = _parse_team_mapping(mappings[pco_team_id])
        if info['orgId'] <= 0:
            print json.dumps({'success': False, 'message': 'Mapping has no TouchPoint org.'})
            return

        people, err = _pco_team_people(pco_team_id)
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        api_errors = []
        # Pull positions and per-position assignments so we know which
        # positions each person is eligible for. Cheap: one call per
        # position (a team typically has 3-10 positions).
        positions_by_id = {}
        positions = []
        total_assignments = 0
        if info['positionsAsSubgroups']:
            pos_list, perr = _pco_team_positions(pco_team_id)
            if perr:
                api_errors.append('Team positions: ' + perr)
            else:
                for p in pos_list:
                    positions.append(p)
                    positions_by_id[p['positionId']] = p['positionName']
                # Walk all assignments under the team in one shot. Per-
                # position iteration occasionally returned empty even
                # when assignments exist; the team-level endpoint with
                # includes is the source of truth.
                by_position, terr = _pco_team_all_position_assignments(pco_team_id)
                if terr:
                    api_errors.append('Team assignments: ' + terr)
                    by_position = {}
                # If the team-level returned nothing but positions exist,
                # try the per-position fallback for safety.
                if not by_position and pos_list:
                    for p in pos_list:
                        aids, aerr = _pco_team_position_assignments(p['positionId'])
                        if aerr:
                            api_errors.append('Position ' + p['positionName'] + ': ' + aerr)
                            continue
                        if aids:
                            by_position[p['positionId']] = aids
                # Aggregate person -> [position names].
                for tpid, pids in by_position.items():
                    pname = positions_by_id.get(tpid, '')
                    if not pname:
                        continue
                    total_assignments += len(pids)
                    for pid in pids:
                        positions_by_id.setdefault('__personPositions__', {}).setdefault(pid, []).append(pname)

        person_positions = positions_by_id.pop('__personPositions__', {}) if info['positionsAsSubgroups'] else {}

        # Attach positions and match to TP.
        pco_ids = [p['pcoPersonId'] for p in people if p['pcoPersonId']]
        emails = [p['email'] for p in people if p['email']]
        by_pco = _tp_match_by_pco_id(pco_ids)
        by_email = _tp_match_by_email(emails)

        attendees = []
        matched = ambiguous = unmatched = 0
        for p in people:
            row = {
                'pcoPersonId': p['pcoPersonId'],
                'name': p['name'],
                'email': p['email'],
                'pcoFirstName': p['pcoFirstName'],
                'pcoLastName': p['pcoLastName'],
                # Team Sync has no plan-day RSVP. Treat everyone on the team
                # as in-scope for the sync; the JS-side "Will Sync" pill
                # works off mode + tpPeopleId, not isConfirmed.
                'isConfirmed': True,
                # Match the per-plan preview's field naming ('teamPosition')
                # so renderAttendeeRow + the sync-write reader both find the
                # name. 'M' = "team member" (no plan-day status applies).
                'positions': [{'teamPosition': name, 'status': 'M'} for name in person_positions.get(p['pcoPersonId'], [])],
                'tpPeopleId': None, 'tpName': '', 'matchSource': '',
                'emailAmbiguous': False,
            }
            if row['pcoPersonId'] and row['pcoPersonId'] in by_pco:
                tp = by_pco[row['pcoPersonId']]
                row['tpPeopleId'] = tp['peopleId']
                row['tpName'] = tp['name']
                row['matchSource'] = 'pco_id'
                row['matchStatus'] = 'matched'
                matched += 1
            elif row['email'] and row['email'] in by_email:
                hits = by_email[row['email']]
                if len(hits) == 1:
                    row['tpPeopleId'] = hits[0]['peopleId']
                    row['tpName'] = hits[0]['name']
                    row['matchSource'] = 'email'
                    row['matchStatus'] = 'matched'
                    matched += 1
                else:
                    row['emailAmbiguous'] = True
                    row['matchStatus'] = 'ambiguous'
                    ambiguous += 1
            else:
                row['matchStatus'] = 'unmatched'
                unmatched += 1
            attendees.append(row)

        _rank = {'unmatched': 0, 'ambiguous': 1, 'matched': 2}
        attendees.sort(key=lambda r: (_rank.get(r.get('matchStatus'), 99),
                                       (r.get('name') or '').lower()))

        # v3.2: mirror-removal counts -- TP rows that have a PCO_PersonId
        # but no longer appear in PCO's team roster / position list.
        # PCO is the source of truth until we wire TP-side writes back.
        pco_team_id_set = set([p['pcoPersonId'] for p in people if p['pcoPersonId']])
        # Map PCO position name -> lowercase form so subgroup comparisons
        # ignore case (TP's AddSubGroup is case-insensitive on writes).
        pco_position_names_lower = set([(positions_by_id.get(p['positionId']) or '').lower() for p in positions])
        pco_position_names_lower.discard('')
        # Per-person current position set (lowercase) for diffing against
        # TP subgroup memberships.
        person_positions_lower = {}
        for pco_pid, names in person_positions.items():
            person_positions_lower[pco_pid] = set([n.lower() for n in names if n])
        # Pull current TP state.
        tp_members = _tp_org_members_with_pco_link(info['orgId'])
        tp_subgroups = _tp_org_subgroups(info['orgId'])
        # Roster drops: TP members with a PCO_PersonId NOT in current PCO team.
        roster_drop_ids = []
        for tp_pid, pco_pid in tp_members.items():
            if pco_pid and pco_pid not in pco_team_id_set:
                roster_drop_ids.append(tp_pid)
        # Subgroup drops (per-position only): walk every TP member's
        # subgroups; drop ones whose name matches an existing PCO
        # position name but the person no longer holds that PCO position.
        # Manually-added subgroups whose name doesn't match a PCO position
        # are left alone (out-of-band).
        subgroup_drops = []  # list of (tp_pid, subgroup_name_lower)
        if info['positionsAsSubgroups']:
            # Reverse map TP person id -> PCO person id for lookup.
            pco_id_by_tp = {}
            for tp_pid, pco_pid in tp_members.items():
                if pco_pid:
                    pco_id_by_tp[tp_pid] = pco_pid
            for tp_pid, sg_names in tp_subgroups.items():
                pco_pid = pco_id_by_tp.get(tp_pid, '')
                if not pco_pid:
                    continue  # untracked / manual member -- leave alone
                desired = person_positions_lower.get(pco_pid, set())
                for sg_name_lower in sg_names:
                    # Only consider dropping subgroups whose name matches
                    # a CURRENT PCO position (otherwise it's an unrelated
                    # subgroup -- leave it).
                    if sg_name_lower in pco_position_names_lower and sg_name_lower not in desired:
                        subgroup_drops.append((tp_pid, sg_name_lower))

        # Show the team name as the primary title (not the service type --
        # one service type can host several teams). The service type name
        # rides along in parens.
        team_label = info['pcoTeamName'] or 'Team'
        if info['pcoServiceTypeName']:
            primary_title = team_label + ' (' + info['pcoServiceTypeName'] + ')'
        else:
            primary_title = team_label
        plan_info = {
            'isTeamSync': True,
            'pcoTeamId': pco_team_id,
            'serviceTypeName': primary_title,
            'pcoTeamName': info['pcoTeamName'],
            'pcoServiceTypeName': info['pcoServiceTypeName'],
            'title': team_label + ' (Team Sync)',
            'planTitle': '',
            'shortDates': str(len(people)) + ' team member(s)',
            'syncAttendance': False,
            'autoAddMember': info['autoAddMember'],
            'positionsAsSubgroups': info['positionsAsSubgroups'],
            'tpOrgId': info['orgId'],
            'positions': positions,
            # v3.1.1: surface position discovery counts so the user can
            # tell "no chips" apart from "PCO has no positions assigned".
            'positionCount': len(positions),
            'positionAssignmentCount': total_assignments,
            'peopleWithPositionsCount': len(person_positions),
            # v3.2: mirror drop counts so the preview can warn before sync.
            'rosterDropCount': len(roster_drop_ids),
            'subgroupDropCount': len(subgroup_drops),
        }
        resp = {
            'success': True,
            'planInfo': plan_info,
            'attendees': attendees,
            'summary': {
                'total': len(attendees),
                'matched': matched,
                'ambiguousEmail': ambiguous,
                'unmatched': unmatched,
                'confirmed': len(attendees),
                'rawAssignments': 0,
            },
        }
        if api_errors:
            resp['warnings'] = api_errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load team sync preview failed: ' + str(e)})

def _run_team_sync_server_side(pco_team_id):
    """v3.3: fully server-side Team Sync. Used by the scheduler. Walks
    PCO, matches to TP, applies adds + mirror drops, returns a result
    dict the email builder can consume. Mirrors the same write logic
    used by handle_sync_team but skips JS-state assumptions."""
    out = {
        'success': False,
        'kind': 'team',
        'key': pco_team_id,
        'tpOrgId': 0,
        'orgName': '',
        'joined': 0,
        'alreadyMember': 0,
        'subgroupAdds': 0,
        'subgroupFailures': 0,
        'rosterDrops': 0,
        'rosterDropFailures': 0,
        'subgroupDrops': 0,
        'subgroupDropFailures': 0,
        'unmatched': [],
        'ambiguous': [],
        'warnings': [],
        'message': '',
    }
    try:
        mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict) or pco_team_id not in mappings:
            out['message'] = 'No team mapping for ' + pco_team_id
            return out
        info = _parse_team_mapping(mappings[pco_team_id])
        if info['orgId'] <= 0 or not info['autoAddMember']:
            out['message'] = 'Mapping disabled or has no TP org.'
            return out
        out['tpOrgId'] = info['orgId']
        out['mappingLabel'] = info['pcoTeamName'] + (' (' + info['pcoServiceTypeName'] + ')' if info['pcoServiceTypeName'] else '')
        # Resolve org name for the email body.
        try:
            for r in q.QuerySql("SELECT TOP 1 OrganizationName FROM Organizations WHERE OrganizationId = %s" % int(info['orgId'])):
                out['orgName'] = safe_str(r.OrganizationName)
                break
        except:
            pass

        # PCO state.
        people, perr = _pco_team_people(pco_team_id)
        if perr:
            out['warnings'].append('Team people fetch: ' + perr)
        pco_team_id_set = set([p['pcoPersonId'] for p in (people or []) if p.get('pcoPersonId')])
        positions = []
        positions_by_id_local = {}
        by_position = {}
        if info['positionsAsSubgroups']:
            pos_list, perr2 = _pco_team_positions(pco_team_id)
            if perr2:
                out['warnings'].append('Team positions: ' + perr2)
            else:
                positions = pos_list
                positions_by_id_local = dict([(p['positionId'], p['positionName']) for p in pos_list])
                by_position, terr = _pco_team_all_position_assignments(pco_team_id)
                if terr:
                    out['warnings'].append('Team assignments: ' + terr)
                if not by_position and pos_list:
                    for p in pos_list:
                        aids, _aerr = _pco_team_position_assignments(p['positionId'])
                        if aids:
                            by_position[p['positionId']] = aids
        # PCO person -> list of position names (case-preserved + lower).
        person_positions_display = {}
        person_positions_lower = {}
        for tpid_pos, pids2 in by_position.items():
            pname = positions_by_id_local.get(tpid_pos, '')
            if not pname:
                continue
            for pp in pids2:
                person_positions_display.setdefault(pp, []).append(pname)
                person_positions_lower.setdefault(pp, set()).add(pname.lower())

        # Match PCO people -> TP.
        pco_ids = [p['pcoPersonId'] for p in (people or []) if p.get('pcoPersonId')]
        emails = [p['email'] for p in (people or []) if p.get('email')]
        by_pco = _tp_match_by_pco_id(pco_ids)
        by_email = _tp_match_by_email(emails)
        matched_pids = []
        matched_pid_to_positions = {}
        for p in (people or []):
            tp_id = 0
            if p.get('pcoPersonId') and p['pcoPersonId'] in by_pco:
                tp_id = int(by_pco[p['pcoPersonId']]['peopleId'])
            elif p.get('email') and p['email'] in by_email:
                hits = by_email[p['email']]
                if len(hits) == 1:
                    tp_id = int(hits[0]['peopleId'])
                else:
                    out['ambiguous'].append({
                        'pcoPersonId': p.get('pcoPersonId', ''),
                        'name': p.get('name', ''),
                        'email': p.get('email', ''),
                    })
                    continue
            if tp_id:
                matched_pids.append(tp_id)
                pos_names = person_positions_display.get(p['pcoPersonId'], [])
                if pos_names:
                    matched_pid_to_positions[tp_id] = pos_names
            else:
                out['unmatched'].append({
                    'pcoPersonId': p.get('pcoPersonId', ''),
                    'name': p.get('name', ''),
                    'email': p.get('email', ''),
                })

        # Current TP state.
        existing_member_ids = set()
        if matched_pids:
            try:
                for i in range(0, len(matched_pids), 500):
                    chunk = matched_pids[i:i+500]
                    mem_sql = ("""
                        SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                        WHERE OrganizationId = %s
                          AND PeopleId IN (%s)
                          AND InactiveDate IS NULL
                    """ % (int(info['orgId']), ','.join(str(p) for p in chunk)))
                    for r in q.QuerySql(mem_sql):
                        existing_member_ids.add(int(r.PeopleId))
            except:
                pass

        # Apply adds.
        for pid in matched_pids:
            try:
                if pid not in existing_member_ids:
                    person = model.GetPerson(int(pid))
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(info['orgId']), person)
                            out['joined'] += 1
                        except Exception as je:
                            if 'already' in str(je).lower():
                                out['alreadyMember'] += 1
                            else:
                                continue
                else:
                    out['alreadyMember'] += 1
                if info['positionsAsSubgroups']:
                    for pos_name in matched_pid_to_positions.get(pid, []):
                        try:
                            model.AddSubGroup(int(pid), int(info['orgId']), pos_name)
                            out['subgroupAdds'] += 1
                        except:
                            out['subgroupFailures'] += 1
            except:
                pass

        # Mirror drops (same logic as handle_sync_team).
        try:
            tp_members = _tp_org_members_with_pco_link(int(info['orgId']))
            for tp_pid, pco_pid in tp_members.items():
                if pco_pid and pco_pid not in pco_team_id_set:
                    try:
                        person = model.GetPerson(int(tp_pid))
                        if person and person.PeopleId:
                            model.RemoveFromOrg(person, int(info['orgId']))
                            out['rosterDrops'] += 1
                    except:
                        out['rosterDropFailures'] += 1
            if info['positionsAsSubgroups']:
                pco_position_names_lower = set([n.lower() for n in positions_by_id_local.values() if n])
                pco_lower_to_display = {}
                for nm in positions_by_id_local.values():
                    if nm:
                        pco_lower_to_display[nm.lower()] = nm
                tp_subgroups_now = _tp_org_subgroups(int(info['orgId']))
                for tp_pid, sg_lower_set in tp_subgroups_now.items():
                    pco_pid = tp_members.get(tp_pid, '')
                    if not pco_pid:
                        continue
                    desired = person_positions_lower.get(pco_pid, set())
                    for sg_lower in sg_lower_set:
                        if sg_lower in pco_position_names_lower and sg_lower not in desired:
                            display_name = pco_lower_to_display.get(sg_lower, sg_lower)
                            try:
                                model.RemoveSubGroup(int(tp_pid), int(info['orgId']), display_name)
                                out['subgroupDrops'] += 1
                            except:
                                out['subgroupDropFailures'] += 1
        except:
            pass

        out['success'] = True
        out['message'] = 'Team sync complete.'
        # Update audit + last-sync setting.
        append_audit({
            'action': 'sync_team',
            'pcoTeamId': pco_team_id,
            'tpOrgId': info['orgId'],
            'joined': out['joined'],
            'alreadyMember': out['alreadyMember'],
            'subgroupAdds': out['subgroupAdds'],
            'rosterDrops': out['rosterDrops'],
            'subgroupDrops': out['subgroupDrops'],
            'by': 'scheduler',
        })
        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict): s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)
        return out
    except Exception as e:
        out['warnings'].append('Run failed: ' + str(e))
        out['message'] = 'Run failed: ' + str(e)
        return out

def _run_people_sync_server_side(pco_st_id):
    """v3.3: fully server-side Service Type / People Sync."""
    out = {
        'success': False, 'kind': 'people', 'key': pco_st_id, 'tpOrgId': 0, 'orgName': '', 'mappingLabel': '',
        'joined': 0, 'alreadyMember': 0, 'subgroupAdds': 0, 'subgroupFailures': 0,
        'rosterDrops': 0, 'rosterDropFailures': 0, 'subgroupDrops': 0, 'subgroupDropFailures': 0,
        'unmatched': [], 'ambiguous': [], 'warnings': [], 'message': '',
    }
    try:
        mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict) or pco_st_id not in mappings:
            out['message'] = 'No people mapping for ' + pco_st_id
            return out
        info = _parse_people_mapping(mappings[pco_st_id])
        if info['orgId'] <= 0 or not info['autoAddMember']:
            out['message'] = 'Mapping disabled.'
            return out
        out['tpOrgId'] = info['orgId']
        out['mappingLabel'] = info['pcoServiceTypeName'] or pco_st_id
        try:
            for r in q.QuerySql("SELECT TOP 1 OrganizationName FROM Organizations WHERE OrganizationId = %s" % int(info['orgId'])):
                out['orgName'] = safe_str(r.OrganizationName)
                break
        except:
            pass
        teams, terr = _pco_teams_for_service_type(pco_st_id)
        if terr:
            out['warnings'].append('Teams fetch: ' + terr)
            return out
        # Aggregate everyone across teams.
        person_data = {}  # pco_pid -> {name, email, teams:[name]}
        for t in teams:
            ppl, perr = _pco_team_people(t['teamId'])
            if perr:
                out['warnings'].append('Team ' + t['teamName'] + ': ' + perr)
                continue
            for p in ppl:
                pp = p.get('pcoPersonId')
                if not pp:
                    continue
                row = person_data.get(pp)
                if row is None:
                    row = {'name': p.get('name', ''), 'email': p.get('email', ''), 'teams': []}
                    person_data[pp] = row
                if t['teamName'] not in row['teams']:
                    row['teams'].append(t['teamName'])
                if not row['email'] and p.get('email'):
                    row['email'] = p.get('email', '')
        pco_in_scope = set(person_data.keys())
        # Match.
        by_pco = _tp_match_by_pco_id(list(pco_in_scope))
        emails_list = [d.get('email') for d in person_data.values() if d.get('email')]
        by_email = _tp_match_by_email(emails_list)
        matched_pids = []
        matched_pid_to_teams = {}
        for pp, d in person_data.items():
            tp_id = 0
            if pp in by_pco:
                tp_id = int(by_pco[pp]['peopleId'])
            elif d.get('email') and d['email'] in by_email:
                hits = by_email[d['email']]
                if len(hits) == 1:
                    tp_id = int(hits[0]['peopleId'])
                else:
                    out['ambiguous'].append({'pcoPersonId': pp, 'name': d.get('name', ''), 'email': d.get('email', '')})
                    continue
            if tp_id:
                matched_pids.append(tp_id)
                if d.get('teams'):
                    matched_pid_to_teams[tp_id] = d['teams']
            else:
                out['unmatched'].append({'pcoPersonId': pp, 'name': d.get('name', ''), 'email': d.get('email', '')})
        # Existing members.
        existing_member_ids = set()
        if matched_pids:
            try:
                for i in range(0, len(matched_pids), 500):
                    chunk = matched_pids[i:i+500]
                    mem_sql = ("""
                        SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                        WHERE OrganizationId = %s AND PeopleId IN (%s) AND InactiveDate IS NULL
                    """ % (int(info['orgId']), ','.join(str(p) for p in chunk)))
                    for r in q.QuerySql(mem_sql):
                        existing_member_ids.add(int(r.PeopleId))
            except:
                pass
        # Apply adds.
        for pid in matched_pids:
            try:
                if pid not in existing_member_ids:
                    person = model.GetPerson(int(pid))
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(info['orgId']), person)
                            out['joined'] += 1
                        except Exception as je:
                            if 'already' in str(je).lower():
                                out['alreadyMember'] += 1
                            else:
                                continue
                else:
                    out['alreadyMember'] += 1
                if info['teamsAsSubgroups']:
                    for team_name in matched_pid_to_teams.get(pid, []):
                        try:
                            model.AddSubGroup(int(pid), int(info['orgId']), team_name)
                            out['subgroupAdds'] += 1
                        except:
                            out['subgroupFailures'] += 1
            except:
                pass
        # Mirror drops.
        try:
            tp_members = _tp_org_members_with_pco_link(int(info['orgId']))
            for tp_pid, pco_pid in tp_members.items():
                if pco_pid and pco_pid not in pco_in_scope:
                    try:
                        person = model.GetPerson(int(tp_pid))
                        if person and person.PeopleId:
                            model.RemoveFromOrg(person, int(info['orgId']))
                            out['rosterDrops'] += 1
                    except:
                        out['rosterDropFailures'] += 1
            if info['teamsAsSubgroups']:
                pco_team_names_lower = set([t['teamName'].lower() for t in teams if t.get('teamName')])
                pco_lower_to_display = {}
                for t in teams:
                    if t.get('teamName'):
                        pco_lower_to_display[t['teamName'].lower()] = t['teamName']
                # Desired per TP (lowercase).
                person_teams_lower_by_pco = {}
                for pp, d in person_data.items():
                    person_teams_lower_by_pco[pp] = set([nm.lower() for nm in d.get('teams', [])])
                tp_subgroups_now = _tp_org_subgroups(int(info['orgId']))
                for tp_pid, sg_lower_set in tp_subgroups_now.items():
                    pco_pid = tp_members.get(tp_pid, '')
                    if not pco_pid:
                        continue
                    desired = person_teams_lower_by_pco.get(pco_pid, set())
                    for sg_lower in sg_lower_set:
                        if sg_lower in pco_team_names_lower and sg_lower not in desired:
                            display_name = pco_lower_to_display.get(sg_lower, sg_lower)
                            try:
                                model.RemoveSubGroup(int(tp_pid), int(info['orgId']), display_name)
                                out['subgroupDrops'] += 1
                            except:
                                out['subgroupDropFailures'] += 1
        except:
            pass
        out['success'] = True
        out['message'] = 'People sync complete.'
        append_audit({'action': 'sync_people', 'pcoServiceTypeId': pco_st_id, 'tpOrgId': info['orgId'],
                      'joined': out['joined'], 'rosterDrops': out['rosterDrops'], 'by': 'scheduler'})
        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict): s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)
        return out
    except Exception as e:
        out['warnings'].append('Run failed: ' + str(e))
        out['message'] = 'Run failed: ' + str(e)
        return out

def _run_all_people_sync_server_side():
    """v3.3: fully server-side All People Sync."""
    out = {
        'success': False, 'kind': 'all_people', 'key': '', 'tpOrgId': 0, 'orgName': '', 'mappingLabel': 'All People',
        'joined': 0, 'alreadyMember': 0, 'subgroupAdds': 0, 'subgroupFailures': 0,
        'rosterDrops': 0, 'rosterDropFailures': 0, 'subgroupDrops': 0, 'subgroupDropFailures': 0,
        'unmatched': [], 'ambiguous': [], 'warnings': [], 'message': '',
    }
    try:
        info = _parse_all_people_mapping(load_json(ALL_PEOPLE_MAPPING_KEY, {}))
        if info['orgId'] <= 0 or not info['autoAddMember']:
            out['message'] = 'All People mapping disabled.'
            return out
        out['tpOrgId'] = info['orgId']
        try:
            for r in q.QuerySql("SELECT TOP 1 OrganizationName FROM Organizations WHERE OrganizationId = %s" % int(info['orgId'])):
                out['orgName'] = safe_str(r.OrganizationName)
                break
        except:
            pass
        people, errors = _pco_all_people_walk(include_inactive=info['includeInactive'])
        if errors:
            for e in errors: out['warnings'].append(e)
        pco_ids = [p['pcoPersonId'] for p in people if p.get('pcoPersonId')]
        emails = [p['email'] for p in people if p.get('email')]
        by_pco = _tp_match_by_pco_id(pco_ids)
        by_email = _tp_match_by_email(emails)
        matched_pids = []
        for p in people:
            tp_id = 0
            if p['pcoPersonId'] and p['pcoPersonId'] in by_pco:
                tp_id = int(by_pco[p['pcoPersonId']]['peopleId'])
            elif p['email'] and p['email'] in by_email:
                hits = by_email[p['email']]
                if len(hits) == 1:
                    tp_id = int(hits[0]['peopleId'])
                else:
                    out['ambiguous'].append({'pcoPersonId': p.get('pcoPersonId', ''), 'name': (p.get('first_name', '') + ' ' + p.get('last_name', '')).strip(), 'email': p.get('email', '')})
                    continue
            if tp_id:
                matched_pids.append(tp_id)
            else:
                if len(out['unmatched']) < 200:
                    out['unmatched'].append({'pcoPersonId': p.get('pcoPersonId', ''), 'name': (p.get('first_name', '') + ' ' + p.get('last_name', '')).strip(), 'email': p.get('email', '')})
        tp_org_id = info['orgId']
        existing_member_ids = set()
        if matched_pids:
            for i in range(0, len(matched_pids), 500):
                chunk = matched_pids[i:i+500]
                try:
                    for r in q.QuerySql("SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK) WHERE OrganizationId = %s AND PeopleId IN (%s) AND InactiveDate IS NULL" % (int(tp_org_id), ','.join(str(p) for p in chunk))):
                        existing_member_ids.add(int(r.PeopleId))
                except:
                    pass
        for pid in matched_pids:
            try:
                if pid in existing_member_ids:
                    out['alreadyMember'] += 1
                else:
                    person = model.GetPerson(int(pid))
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(tp_org_id), person)
                            out['joined'] += 1
                        except Exception as je:
                            if 'already' in str(je).lower():
                                out['alreadyMember'] += 1
            except:
                pass
        # Mirror drops.
        try:
            pco_in_scope = set([p['pcoPersonId'] for p in people if p.get('pcoPersonId')])
            tp_members = _tp_org_members_with_pco_link(int(tp_org_id))
            for tp_pid, pco_pid in tp_members.items():
                if pco_pid and pco_pid not in pco_in_scope:
                    try:
                        person = model.GetPerson(int(tp_pid))
                        if person and person.PeopleId:
                            model.RemoveFromOrg(person, int(tp_org_id))
                            out['rosterDrops'] += 1
                    except:
                        out['rosterDropFailures'] += 1
        except:
            pass
        out['success'] = True
        out['message'] = 'All People sync complete.'
        append_audit({'action': 'sync_all_people', 'tpOrgId': tp_org_id,
                      'joined': out['joined'], 'rosterDrops': out['rosterDrops'], 'by': 'scheduler'})
        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict): s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)
        return out
    except Exception as e:
        out['warnings'].append('Run failed: ' + str(e))
        return out

def _format_sync_email_html(result, sch, include_issues, link_back):
    """Render the post-sync email body. Always shows counts; optionally
    shows issues list (unmatched, ambiguous, warnings)."""
    org_name = result.get('orgName', '') or ('Org #' + str(result.get('tpOrgId', '')))
    title = (result.get('mappingLabel') or result.get('kind', 'PCO Sync')) + ' &rarr; ' + org_name
    rows = [
        ('Joined', result.get('joined', 0)),
        ('Already member', result.get('alreadyMember', 0)),
        ('Subgroup writes', result.get('subgroupAdds', 0)),
        ('Members removed (mirror)', result.get('rosterDrops', 0)),
        ('Stale subgroups removed', result.get('subgroupDrops', 0)),
    ]
    counts_html = '<table style="border-collapse:collapse;font-size:13px;">'
    for label, n in rows:
        color = '#1f6b3a' if n else '#777'
        counts_html += '<tr><td style="padding:3px 10px 3px 0;color:#555;">' + label + '</td>'
        counts_html += '<td style="padding:3px 0;color:' + color + ';font-weight:600;">' + str(n) + '</td></tr>'
    counts_html += '</table>'

    issues_html = ''
    if include_issues:
        unm = result.get('unmatched', []) or []
        amb = result.get('ambiguous', []) or []
        warns = result.get('warnings', []) or []
        if unm or amb or warns:
            issues_html += '<h3 style="color:#8a2020;margin-top:18px;">Issues</h3>'
        if warns:
            issues_html += '<p><strong>API warnings:</strong></p><ul>'
            for w in warns[:25]:
                issues_html += '<li>' + html_escape(safe_str(w)) + '</li>'
            issues_html += '</ul>'
        if unm:
            issues_html += '<p><strong>Unmatched PCO records (' + str(len(unm)) + '):</strong> these PCO people had no TP match. Use the Proposed Matches review in PCO Sync.</p>'
            issues_html += '<ul>'
            for u in unm[:50]:
                issues_html += '<li>' + html_escape(u.get('name', '') or '(unknown)') + (' &mdash; ' + html_escape(u.get('email', '')) if u.get('email') else '') + ' <span style="color:#999;">PCO #' + html_escape(u.get('pcoPersonId', '')) + '</span></li>'
            issues_html += '</ul>'
            if len(unm) > 50:
                issues_html += '<p style="color:#999;font-size:12px;">...and ' + str(len(unm) - 50) + ' more.</p>'
        if amb:
            issues_html += '<p><strong>Ambiguous email matches (' + str(len(amb)) + '):</strong> email matches multiple TP people. Use the Verify Person Link tool.</p>'
            issues_html += '<ul>'
            for a in amb[:25]:
                issues_html += '<li>' + html_escape(a.get('name', '')) + ' &mdash; ' + html_escape(a.get('email', '')) + '</li>'
            issues_html += '</ul>'

    return ('<div style="font-family:Segoe UI,Arial,sans-serif;color:#333;">'
            + '<h2 style="color:#1f4e79;margin-bottom:4px;">PCO Sync &mdash; ' + html_escape(title) + '</h2>'
            + '<p style="color:#666;font-size:13px;margin-top:0;">Scheduled run &middot; ' + _format_schedule_label(sch or {}) + '</p>'
            + counts_html
            + issues_html
            + '<p style="margin-top:18px;"><a href="' + link_back + '" style="color:#1f4e79;">Open PCO Sync &rarr;</a></p>'
            + '<p style="color:#999;font-size:11px;margin-top:20px;">PCO Sync v' + APP_VERSION + ' &middot; Scheduled task</p>'
            + '</div>')

def _audit_email_outcome(action, kind, key, recipient_user, sent, err):
    """Persist whether the post-sync email actually went out so staff
    can confirm via the audit log (or future Audit tab) when the
    sync ran but the email vanished."""
    append_audit({
        'action': 'email_' + ('sent' if sent else 'failed'),
        'kind': kind,
        'key': key,
        'notifyUsername': recipient_user,
        'syncAction': action,
        'error': err or '',
    })

def handle_run_scheduled_syncs():
    """Scheduler entry point. Walks every mapping, fires due ones,
    captures result + emails the configured user. Called by the
    ScheduledTasks block (Data.scheduler='true') OR by the manual
    Run Scheduler test button in the UI.

    Accepts an optional 'force' param: when truthy, ignores the
    per-hour-slot dedup so the test button can re-run a mapping that
    already fired earlier this hour."""
    try:
        force_run = _truthy(get_data('force', '0'), False)
        results = []
        skipped_already_fired = []
        # Build origin URL for email backlinks. CmsHost is sometimes
        # 'https://myfbch.com', sometimes 'myfbch.com'. Normalize to a
        # single absolute URL so we don't end up with 'https://https://'.
        try:
            host = safe_str(getattr(model, 'CmsHost', '') or '').strip()
        except:
            host = ''
        if host:
            if host.startswith('http://') or host.startswith('https://'):
                link_back = host.rstrip('/') + '/PyScriptForm/TPxi_PCOSync'
            else:
                link_back = 'https://' + host.rstrip('/') + '/PyScriptForm/TPxi_PCOSync'
        else:
            link_back = '/PyScriptForm/TPxi_PCOSync'
        now = datetime.datetime.now()

        def _due_or_skip(sch, label):
            if force_run:
                # Force-run still requires a valid schedule (enabled).
                return sch.get('enabled', False), 'disabled'
            if not _is_due_now(sch, sch.get('lastScheduledRunAt', ''), now):
                # Two reasons we'd skip: not enabled, or day/hour mismatch,
                # or already-fired-this-hour. Distinguish the last so
                # the UI can be honest.
                if not sch.get('enabled', False):
                    return False, 'disabled'
                last = sch.get('lastScheduledRunAt', '')
                if last:
                    try:
                        last_dt = datetime.datetime.strptime(last[:19], '%Y-%m-%dT%H:%M:%S')
                        if (now - last_dt).total_seconds() < 60 * 55 and now.hour == sch.get('hour', 6):
                            return False, 'already_fired_this_hour'
                    except:
                        pass
                return False, 'not_due'
            return True, ''

        # Team mappings.
        team_mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if isinstance(team_mappings, dict):
            for tk, tv in team_mappings.items():
                info = _parse_team_mapping(tv)
                sch = info.get('schedule', {})
                due, reason = _due_or_skip(sch, info.get('pcoTeamName', ''))
                if not due:
                    if reason == 'already_fired_this_hour':
                        skipped_already_fired.append({
                            'kind': 'team', 'key': tk,
                            'label': info['pcoTeamName'] or 'Team',
                            'lastRunAt': sch.get('lastScheduledRunAt', ''),
                        })
                    continue
                result = _run_team_sync_server_side(tk)
                tv['schedule'] = dict(sch)
                tv['schedule']['lastScheduledRunAt'] = now.strftime('%Y-%m-%dT%H:%M:%S')
                team_mappings[tk] = tv
                if sch.get('notifyUsername'):
                    html_body = _format_sync_email_html(result, sch, sch.get('includeIssues', True), link_back)
                    subject = '[PCO Sync] ' + (info['pcoTeamName'] or 'Team') + ' -- joined ' + str(result['joined']) + ', removed ' + str(result['rosterDrops'])
                    sent, err = _send_sync_email(sch['notifyUsername'], subject, html_body)
                    result['emailSent'] = sent
                    if err:
                        result['emailError'] = err
                    _audit_email_outcome('sync_team', 'team', tk, sch['notifyUsername'], sent, err)
                else:
                    result['emailSent'] = False
                    result['emailError'] = 'No notify username configured.'
                results.append(result)
            save_json(TEAM_MAPPINGS_KEY, team_mappings)

        # People mappings.
        people_mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if isinstance(people_mappings, dict):
            for pk, pv in people_mappings.items():
                info = _parse_people_mapping(pv)
                sch = info.get('schedule', {})
                due, reason = _due_or_skip(sch, info.get('pcoServiceTypeName', ''))
                if not due:
                    if reason == 'already_fired_this_hour':
                        skipped_already_fired.append({
                            'kind': 'people', 'key': pk,
                            'label': info['pcoServiceTypeName'] or 'Service Type',
                            'lastRunAt': sch.get('lastScheduledRunAt', ''),
                        })
                    continue
                result = _run_people_sync_server_side(pk)
                pv['schedule'] = dict(sch)
                pv['schedule']['lastScheduledRunAt'] = now.strftime('%Y-%m-%dT%H:%M:%S')
                people_mappings[pk] = pv
                if sch.get('notifyUsername'):
                    html_body = _format_sync_email_html(result, sch, sch.get('includeIssues', True), link_back)
                    subject = '[PCO Sync] ' + (info['pcoServiceTypeName'] or 'Service Type') + ' -- joined ' + str(result['joined']) + ', removed ' + str(result['rosterDrops'])
                    sent, err = _send_sync_email(sch['notifyUsername'], subject, html_body)
                    result['emailSent'] = sent
                    if err: result['emailError'] = err
                    _audit_email_outcome('sync_people', 'people', pk, sch['notifyUsername'], sent, err)
                else:
                    result['emailSent'] = False
                    result['emailError'] = 'No notify username configured.'
                results.append(result)
            save_json(PEOPLE_MAPPINGS_KEY, people_mappings)

        # All People mapping.
        ap_cur = load_json(ALL_PEOPLE_MAPPING_KEY, {})
        if isinstance(ap_cur, dict):
            info = _parse_all_people_mapping(ap_cur)
            sch = info.get('schedule', {})
            due, reason = _due_or_skip(sch, 'All People')
            if due:
                result = _run_all_people_sync_server_side()
                ap_cur['schedule'] = dict(sch)
                ap_cur['schedule']['lastScheduledRunAt'] = now.strftime('%Y-%m-%dT%H:%M:%S')
                save_json(ALL_PEOPLE_MAPPING_KEY, ap_cur)
                if sch.get('notifyUsername'):
                    html_body = _format_sync_email_html(result, sch, sch.get('includeIssues', True), link_back)
                    subject = '[PCO Sync] All People -- joined ' + str(result['joined']) + ', removed ' + str(result['rosterDrops'])
                    sent, err = _send_sync_email(sch['notifyUsername'], subject, html_body)
                    result['emailSent'] = sent
                    if err: result['emailError'] = err
                    _audit_email_outcome('sync_all_people', 'all_people', '', sch['notifyUsername'], sent, err)
                else:
                    result['emailSent'] = False
                    result['emailError'] = 'No notify username configured.'
                results.append(result)
            elif reason == 'already_fired_this_hour':
                skipped_already_fired.append({
                    'kind': 'all_people', 'key': '',
                    'label': 'All People',
                    'lastRunAt': sch.get('lastScheduledRunAt', ''),
                })

        # Summarize email status for the test-button caller.
        emails_attempted = sum(1 for r in results if r.get('emailSent') is not None and r.get('emailError') != 'No notify username configured.')
        emails_sent = sum(1 for r in results if r.get('emailSent'))
        emails_failed = [{'mapping': r.get('mappingLabel', r.get('key', '')), 'error': r.get('emailError', '')} for r in results if r.get('emailSent') is False and r.get('emailError') != 'No notify username configured.']

        print json.dumps({
            'success': True,
            'firedCount': len(results),
            'results': results,
            'skippedAlreadyFired': skipped_already_fired,
            'emailsAttempted': emails_attempted,
            'emailsSent': emails_sent,
            'emailsFailed': emails_failed,
            'force': force_run,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Scheduler run failed: ' + str(e)})

def handle_check_pco_team_positions():
    """v3.1.1: Inline diagnostic for a team mapping. Walks PCO once and
    reports position discovery counts so the user can tell apart 'no
    positions exist' vs 'positions exist but unassigned' vs 'all good'
    without opening the full preview modal.

    Uses team-level assignment endpoint with per-position fallback so
    we see assignments even when the per-position endpoint is flaky.
    Reports which path returned data for diagnostic clarity."""
    try:
        pco_team_id = safe_str(get_data('pco_team_id', '')).strip()
        if not pco_team_id:
            print json.dumps({'success': False, 'message': 'Missing pco_team_id'})
            return
        # Position list.
        pos_list, perr = _pco_team_positions(pco_team_id)
        if perr:
            print json.dumps({'success': False, 'message': 'Positions fetch: ' + perr})
            return
        positions = [{'positionId': p['positionId'], 'positionName': p['positionName']} for p in pos_list]
        pos_name_by_id = dict([(p['positionId'], p['positionName']) for p in pos_list])

        # Primary path: team-level assignment endpoint.
        source = 'team-level'
        by_position, terr = _pco_team_all_position_assignments(pco_team_id)
        if terr:
            source = 'team-level errored, fell back to per-position'
            by_position = {}
        # Fallback path: per-position iteration.
        if not by_position and pos_list:
            source = 'per-position (team-level returned empty)'
            for p in pos_list:
                aids, aerr = _pco_team_position_assignments(p['positionId'])
                if aids:
                    by_position[p['positionId']] = aids

        total_assignments = 0
        person_position_set = set()
        per_pos = []
        for p in pos_list:
            aids = by_position.get(p['positionId'], [])
            cnt = len(aids)
            total_assignments += cnt
            for pid in aids:
                person_position_set.add(pid)
            per_pos.append({
                'positionId': p['positionId'],
                'positionName': p['positionName'],
                'assignmentCount': cnt,
                'error': '',
            })

        # Cross-check against the team's people roster to surface drift.
        roster, rerr = _pco_team_people(pco_team_id)
        roster_ids = set([r['pcoPersonId'] for r in (roster or []) if r.get('pcoPersonId')])
        matched_in_roster = len(person_position_set & roster_ids)
        print json.dumps({
            'success': True,
            'positionCount': len(positions),
            'positionAssignmentCount': total_assignments,
            'peopleWithPositionsCount': matched_in_roster,
            'positions': positions,
            'perPosition': per_pos,
            'rosterSize': len(roster_ids),
            'source': source,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Check positions failed: ' + str(e)})

def handle_sync_team():
    """Apply a Team Sync write: JoinOrg + (optional) subgroup-per-position.
    Same active-member filter and Person Sync hook as sync_roster."""
    try:
        pco_team_id = safe_str(get_data('pco_team_id', '')).strip()
        tp_org_id = safe_int(get_data('tp_org_id', 0), 0)
        people_ids_csv = safe_str(get_data('people_ids', '')).strip()
        if not pco_team_id or tp_org_id <= 0 or not people_ids_csv:
            print json.dumps({'success': False, 'message': 'Missing required args'})
            return
        mappings = load_json(TEAM_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        info = _parse_team_mapping(mappings.get(pco_team_id))
        if info['orgId'] != tp_org_id:
            print json.dumps({'success': False, 'message': 'Mapping mismatch -- team is mapped to a different org now.'})
            return
        if not info['autoAddMember']:
            print json.dumps({'success': False, 'message': 'Auto-add is disabled for this team mapping.'})
            return

        pids = []
        for tok in people_ids_csv.split(','):
            pid = safe_int(tok.strip(), 0)
            if pid > 0:
                pids.append(pid)
        if not pids:
            print json.dumps({'success': False, 'message': 'No people ids to sync'})
            return

        # Optional positions payload: JS bundles personId -> [position
        # names] so we know which subgroups to add per person.
        positions_by_pid = {}
        positions_raw = safe_str(get_data('positions_json', ''))
        if positions_raw:
            try:
                pp = json.loads(positions_raw)
                if isinstance(pp, dict):
                    for k, v in pp.items():
                        try:
                            tp_id = int(k)
                        except:
                            continue
                        if isinstance(v, list):
                            positions_by_pid[tp_id] = [str(x) for x in v if x]
            except:
                pass

        # Active-member pre-check.
        existing_member_ids = set()
        try:
            mem_sql = ("""
                SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                WHERE OrganizationId = %s
                  AND PeopleId IN (%s)
                  AND InactiveDate IS NULL
            """ % (int(tp_org_id), ','.join(str(p) for p in pids)))
            for r in q.QuerySql(mem_sql):
                existing_member_ids.add(int(r.PeopleId))
        except:
            pass

        # Person Sync hook.
        person_rules = normalize_person_sync_rules(load_json(PERSON_SYNC_RULES_KEY, {}))
        person_rules_active = person_sync_rules_active(person_rules)
        pco_data_by_pid = _parse_people_json_payload(safe_str(get_data('people_json', '')))
        person_auto = 0
        person_queued = 0

        joined = 0
        already = 0
        subgroup_adds = 0
        subgroup_failures = 0
        skipped = 0
        per_person = []
        for pid in pids:
            try:
                if pid in existing_member_ids:
                    already += 1
                    per_person.append({'peopleId': pid, 'status': 'already_member'})
                else:
                    person = model.GetPerson(pid)
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(tp_org_id), person)
                            joined += 1
                            per_person.append({'peopleId': pid, 'status': 'joined'})
                        except Exception as je:
                            jm = str(je).lower()
                            if 'already' in jm:
                                already += 1
                                per_person.append({'peopleId': pid, 'status': 'already_member'})
                            else:
                                skipped += 1
                                per_person.append({'peopleId': pid, 'status': 'join_failed', 'message': str(je)})
                                continue
                    else:
                        skipped += 1
                        per_person.append({'peopleId': pid, 'status': 'no_person'})
                        continue
                # Subgroup writes per position the person is eligible for.
                # AddSubGroup creates the group if missing and is idempotent
                # for already-member rows in most TP builds, so we don't
                # need an InSubGroup check here.
                if info['positionsAsSubgroups']:
                    for pos_name in positions_by_pid.get(pid, []):
                        try:
                            model.AddSubGroup(int(pid), int(tp_org_id), pos_name)
                            subgroup_adds += 1
                        except Exception as se:
                            subgroup_failures += 1
                # Person data sync hook.
                if person_rules_active and pid in pco_data_by_pid:
                    try:
                        c = apply_person_sync_for_one(pid, pco_data_by_pid[pid], person_rules)
                        person_auto += c.get('auto', 0)
                        person_queued += c.get('queued', 0)
                    except:
                        pass
            except Exception as we:
                skipped += 1
                per_person.append({'peopleId': pid, 'status': 'error', 'message': str(we)})

        # v3.2: mirror removals. PCO is the source of truth -- TP members
        # whose PCO_PersonId no longer appears in the PCO team get
        # dropped, and subgroup memberships matching a PCO position name
        # that the person doesn't currently hold get dropped too.
        # Re-walk PCO state here (don't trust JS) so the drop set is
        # always computed from the latest server-side data.
        roster_drops = 0
        roster_drop_failures = 0
        subgroup_drops = 0
        subgroup_drop_failures = 0
        try:
            pco_people, _perr = _pco_team_people(pco_team_id)
            pco_team_id_set = set([p['pcoPersonId'] for p in (pco_people or []) if p.get('pcoPersonId')])
            tp_members = _tp_org_members_with_pco_link(int(tp_org_id))
            for tp_pid, pco_pid in tp_members.items():
                if pco_pid and pco_pid not in pco_team_id_set:
                    try:
                        person = model.GetPerson(int(tp_pid))
                        if person and person.PeopleId:
                            model.RemoveFromOrg(person, int(tp_org_id))
                            roster_drops += 1
                    except Exception:
                        roster_drop_failures += 1
        except Exception:
            pass
        # Subgroup mirror -- only when positionsAsSubgroups is on.
        if info['positionsAsSubgroups']:
            try:
                # Pull positions + assignments to know what PCO claims now.
                pos_list, _perr2 = _pco_team_positions(pco_team_id)
                positions_by_id_local = dict([(p['positionId'], p['positionName']) for p in (pos_list or [])])
                pco_position_names_lower = set([n.lower() for n in positions_by_id_local.values() if n])
                by_position, _terr = _pco_team_all_position_assignments(pco_team_id)
                if not by_position and pos_list:
                    by_position = {}
                    for p in pos_list:
                        aids, _aerr = _pco_team_position_assignments(p['positionId'])
                        if aids:
                            by_position[p['positionId']] = aids
                # PCO person id -> set of current PCO position names (lower).
                current_positions_lower = {}
                for tpid, pids2 in by_position.items():
                    pname = positions_by_id_local.get(tpid, '')
                    if not pname:
                        continue
                    pname_l = pname.lower()
                    for pp in pids2:
                        current_positions_lower.setdefault(pp, set()).add(pname_l)
                # Walk TP subgroups; drop ones that match a PCO position
                # name but aren't held by the person now.
                tp_subgroups_now = _tp_org_subgroups(int(tp_org_id))
                tp_members_now = _tp_org_members_with_pco_link(int(tp_org_id))
                # Need the actual subgroup name (case preserved) to call
                # RemoveSubGroup. Pull a name map from MemberTags for
                # any subgroup we plan to touch.
                pco_pos_lower_to_display = {}
                for nm in positions_by_id_local.values():
                    if nm:
                        pco_pos_lower_to_display[nm.lower()] = nm
                for tp_pid, sg_lower_set in tp_subgroups_now.items():
                    pco_pid = tp_members_now.get(tp_pid, '')
                    if not pco_pid:
                        continue  # untracked / manual member
                    desired = current_positions_lower.get(pco_pid, set())
                    for sg_lower in sg_lower_set:
                        if sg_lower in pco_position_names_lower and sg_lower not in desired:
                            display_name = pco_pos_lower_to_display.get(sg_lower, sg_lower)
                            try:
                                model.RemoveSubGroup(int(tp_pid), int(tp_org_id), display_name)
                                subgroup_drops += 1
                            except Exception:
                                subgroup_drop_failures += 1
            except Exception:
                pass

        append_audit({
            'action': 'sync_team',
            'pcoTeamId': pco_team_id,
            'tpOrgId': tp_org_id,
            'joined': joined,
            'alreadyMember': already,
            'subgroupAdds': subgroup_adds,
            'subgroupFailures': subgroup_failures,
            'rosterDrops': roster_drops,
            'rosterDropFailures': roster_drop_failures,
            'subgroupDrops': subgroup_drops,
            'subgroupDropFailures': subgroup_drop_failures,
            'skipped': skipped,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })

        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict): s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)

        if joined == 0 and already > 0:
            msg = 'Everyone was already an active member.'
        else:
            msg = 'Added ' + str(joined) + ' member(s) to the involvement.'
            if already > 0:
                msg += ' (' + str(already) + ' already on the roster.)'
        if subgroup_adds:
            msg += ' Wrote ' + str(subgroup_adds) + ' position-subgroup association(s).'
        elif info['positionsAsSubgroups'] and not positions_by_pid:
            # Toggle is on but the preview shipped zero position data --
            # means PCO returned no Team Position assignments. The
            # involvement still got the people, but subgroups are blank.
            msg += ' (Positions toggle is on, but PCO returned no position assignments for these people -- check that the team in PCO has people assigned to specific positions, not just added to the team.)'
        if roster_drops:
            msg += ' Removed ' + str(roster_drops) + ' member(s) no longer in PCO.'
        if subgroup_drops:
            msg += ' Removed ' + str(subgroup_drops) + ' stale position-subgroup association(s).'
        if person_auto:
            msg += ' Updated ' + str(person_auto) + ' person field(s).'
        if person_queued:
            msg += ' Queued ' + str(person_queued) + ' person change(s) for review.'
        print json.dumps({
            'success': True,
            'message': msg,
            'joinedOrg': joined,
            'alreadyMember': already,
            'subgroupAdds': subgroup_adds,
            'subgroupFailures': subgroup_failures,
            'rosterDrops': roster_drops,
            'rosterDropFailures': roster_drop_failures,
            'subgroupDrops': subgroup_drops,
            'subgroupDropFailures': subgroup_drop_failures,
            'skipped': skipped,
            'personAuto': person_auto,
            'personQueued': person_queued,
            'perPerson': per_person,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Sync team failed: ' + str(e)})

# =====================================================================
# People Sync (v2.1+)
# =====================================================================
# Maps one PCO Service Type to a single TP umbrella involvement and
# turns each Team under that service type into a subgroup. Use case:
# you want one "Worship Team" involvement in TP whose members are
# anyone on any of the worship teams, with subgroups showing which
# teams they're on.

def handle_list_people_mappings():
    try:
        raw = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(raw, dict):
            raw = {}
        org_ids = set()
        for v in raw.values():
            info = _parse_people_mapping(v)
            if info['orgId'] > 0:
                org_ids.add(info['orgId'])
        org_info = {}
        if org_ids:
            ids_csv = ','.join(str(i) for i in org_ids)
            sql = """
                SELECT o.OrganizationId, o.OrganizationName,
                       ISNULL(d.Name, '') AS DivisionName,
                       ISNULL(p.Name, '') AS ProgramName
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE o.OrganizationId IN (%s)
            """ % ids_csv
            for r in q.QuerySql(sql):
                org_info[int(r.OrganizationId)] = {
                    'name': safe_str(r.OrganizationName),
                    'division': safe_str(r.DivisionName),
                    'program': safe_str(r.ProgramName),
                }
        out = []
        for st_id, v in raw.items():
            info = _parse_people_mapping(v)
            oi = org_info.get(info['orgId'], {}) if info['orgId'] > 0 else {}
            sch = info.get('schedule', _parse_schedule({}))
            nxt = _next_scheduled_run(sch)
            out.append({
                'pcoServiceTypeId': info['pcoServiceTypeId'] or st_id,
                'pcoServiceTypeName': info['pcoServiceTypeName'],
                'tpOrgId': info['orgId'],
                'tpOrgName': oi.get('name', '[Org #' + str(info['orgId']) + ' not found]'),
                'tpDivision': oi.get('division', ''),
                'tpProgram': oi.get('program', ''),
                'autoAddMember': info['autoAddMember'],
                'teamsAsSubgroups': info['teamsAsSubgroups'],
                'perPlanAttendance': info['perPlanAttendance'],
                'schedule': sch,
                'scheduleLabel': _format_schedule_label(sch),
                'nextRunIso': nxt.strftime('%Y-%m-%dT%H:%M:%S') if nxt else '',
            })
        out.sort(key=lambda r: (r.get('pcoServiceTypeName') or '').lower())
        print json.dumps({'success': True, 'mappings': out, 'schedulerInstalled': _scheduler_installed()})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List people mappings failed: ' + str(e)})

def handle_save_people_mapping():
    try:
        pco_st_id = safe_str(get_data('pco_service_type_id', '')).strip()
        pco_st_name = safe_str(get_data('pco_service_type_name', '')).strip()
        org_id = safe_int(get_data('tp_org_id', 0), 0)
        if not pco_st_id or org_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing pco_service_type_id or tp_org_id'})
            return
        verify_sql = """
            SELECT TOP 1 o.OrganizationId, o.OrganizationName
            FROM Organizations o WITH (NOLOCK)
            WHERE o.OrganizationId = %s AND o.OrganizationStatusId = 30
        """ % str(org_id)
        org_name = ''
        for r in q.QuerySql(verify_sql):
            org_name = safe_str(r.OrganizationName)
            break
        if not org_name:
            print json.dumps({'success': False, 'message': 'Org #' + str(org_id) + ' is not active or does not exist.'})
            return
        mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        existing = _parse_people_mapping(mappings.get(pco_st_id))
        mappings[pco_st_id] = {
            'pcoServiceTypeId': pco_st_id,
            'pcoServiceTypeName': pco_st_name or existing['pcoServiceTypeName'],
            'orgId': org_id,
            'autoAddMember': existing['autoAddMember'] if existing['orgId'] > 0 else True,
            'teamsAsSubgroups': existing['teamsAsSubgroups'] if existing['orgId'] > 0 else True,
            'perPlanAttendance': existing['perPlanAttendance'] if existing['orgId'] > 0 else False,
        }
        save_ok = save_json(PEOPLE_MAPPINGS_KEY, mappings)
        verify = load_json(PEOPLE_MAPPINGS_KEY, {})
        verified = (save_ok and isinstance(verify, dict)
                    and _parse_people_mapping(verify.get(pco_st_id))['orgId'] == org_id)
        append_audit({
            'action': 'save_people_mapping',
            'pcoServiceTypeId': pco_st_id,
            'pcoServiceTypeName': pco_st_name,
            'tpOrgId': org_id,
            'tpOrgName': org_name,
            'saveOk': save_ok,
            'verified': verified,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        if not verified:
            print json.dumps({'success': False, 'message': 'Save did not persist -- check role on PCOSync_PeopleMappings content.'})
            return
        print json.dumps({'success': True, 'message': 'Mapped service type "' + (pco_st_name or pco_st_id) + '" to ' + org_name + '.', 'orgName': org_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save people mapping failed: ' + str(e)})

def handle_delete_people_mapping():
    try:
        pco_st_id = safe_str(get_data('pco_service_type_id', '')).strip()
        if not pco_st_id:
            print json.dumps({'success': False, 'message': 'Missing pco_service_type_id'})
            return
        mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        mappings.pop(pco_st_id, None)
        save_json(PEOPLE_MAPPINGS_KEY, mappings)
        append_audit({
            'action': 'delete_people_mapping',
            'pcoServiceTypeId': pco_st_id,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'message': 'People mapping removed.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Delete people mapping failed: ' + str(e)})

def handle_set_people_mapping_options():
    """v3.0+: per-mapping toggle save. Updates teamsAsSubgroups and/or
    perPlanAttendance without changing which org the service type points
    to. Called when the user flips a checkbox on the mapping row."""
    try:
        pco_st_id = safe_str(get_data('pco_service_type_id', '')).strip()
        if not pco_st_id:
            print json.dumps({'success': False, 'message': 'Missing pco_service_type_id'})
            return
        mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        if pco_st_id not in mappings:
            print json.dumps({'success': False, 'message': 'No mapping exists for this service type.'})
            return
        existing = _parse_people_mapping(mappings[pco_st_id])
        new_subg = _truthy(get_data('teams_as_subgroups', None), existing['teamsAsSubgroups'])
        new_att  = _truthy(get_data('per_plan_attendance', None), existing['perPlanAttendance'])
        mappings[pco_st_id] = {
            'pcoServiceTypeId': existing['pcoServiceTypeId'] or pco_st_id,
            'pcoServiceTypeName': existing['pcoServiceTypeName'],
            'orgId': existing['orgId'],
            'autoAddMember': existing['autoAddMember'],
            'teamsAsSubgroups': new_subg,
            'perPlanAttendance': new_att,
            # v3.3: preserve scheduler config across option saves.
            'schedule': existing.get('schedule', _parse_schedule({})),
        }
        save_ok = save_json(PEOPLE_MAPPINGS_KEY, mappings)
        verify = _parse_people_mapping(load_json(PEOPLE_MAPPINGS_KEY, {}).get(pco_st_id))
        verified = (save_ok and verify['teamsAsSubgroups'] == new_subg
                              and verify['perPlanAttendance'] == new_att)
        append_audit({
            'action': 'set_people_mapping_options',
            'pcoServiceTypeId': pco_st_id,
            'teamsAsSubgroups': new_subg,
            'perPlanAttendance': new_att,
            'saveOk': save_ok,
            'verified': verified,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        if not verified:
            print json.dumps({'success': False, 'message': 'Save did not persist -- check role on PCOSync_PeopleMappings.'})
            return
        print json.dumps({'success': True, 'teamsAsSubgroups': new_subg, 'perPlanAttendance': new_att})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Set options failed: ' + str(e)})

def handle_load_people_sync_preview():
    """Aggregate everyone on every team in the service type into a single
    cross-team roster. Each person row carries the teams they're on so
    the JS can show team pills + the sync can write subgroups."""
    try:
        pco_st_id = safe_str(get_data('pco_service_type_id', '')).strip()
        if not pco_st_id:
            print json.dumps({'success': False, 'message': 'Missing pco_service_type_id'})
            return
        mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        info = _parse_people_mapping(mappings.get(pco_st_id))
        if info['orgId'] <= 0:
            print json.dumps({'success': False, 'message': 'No people mapping for this service type.'})
            return

        teams, terr = _pco_teams_for_service_type(pco_st_id)
        if terr:
            print json.dumps({'success': False, 'message': terr})
            return
        api_errors = []
        # Walk each team -> people, aggregating by PCO Person ID. Each
        # person carries which teams they're on (these become subgroup
        # memberships in TP).
        by_pid = {}
        no_pid_rows = []
        for team in teams:
            people, perr = _pco_team_people(team['teamId'])
            if perr:
                api_errors.append('Team ' + team['teamName'] + ': ' + perr)
                continue
            for p in people:
                pid = p['pcoPersonId']
                if not pid:
                    no_pid_rows.append({
                        'pcoPersonId': '',
                        'name': p['name'],
                        'email': p['email'],
                        'pcoFirstName': p['pcoFirstName'],
                        'pcoLastName': p['pcoLastName'],
                        'positions': [{'position': team['teamName'], 'status': 'M'}],
                        'isConfirmed': True,
                        'tpPeopleId': None, 'tpName': '', 'matchSource': '', 'emailAmbiguous': False,
                    })
                    continue
                row = by_pid.get(pid)
                if row is None:
                    row = {
                        'pcoPersonId': pid,
                        'name': p['name'],
                        'email': p['email'],
                        'pcoFirstName': p['pcoFirstName'],
                        'pcoLastName': p['pcoLastName'],
                        'positions': [],
                        'isConfirmed': True,
                        'tpPeopleId': None, 'tpName': '', 'matchSource': '', 'emailAmbiguous': False,
                        '_team_set': set(),
                    }
                    by_pid[pid] = row
                if team['teamName'] not in row['_team_set']:
                    row['_team_set'].add(team['teamName'])
                    row['positions'].append({'position': team['teamName'], 'status': 'M'})
                # Backfill missing fields if a later team has them.
                if p['email'] and not row['email']:
                    row['email'] = p['email']
                if p['pcoFirstName'] and not row.get('pcoFirstName'):
                    row['pcoFirstName'] = p['pcoFirstName']
                if p['pcoLastName'] and not row.get('pcoLastName'):
                    row['pcoLastName'] = p['pcoLastName']

        # Drop internal set fields before serialization.
        for r in by_pid.values():
            r.pop('_team_set', None)

        # Match to TP.
        pco_ids = [r['pcoPersonId'] for r in by_pid.values() if r['pcoPersonId']]
        emails = [r['email'] for r in by_pid.values() if r['email']]
        emails += [r['email'] for r in no_pid_rows if r['email']]
        by_pco = _tp_match_by_pco_id(pco_ids)
        by_email = _tp_match_by_email(emails)

        attendees = list(by_pid.values()) + no_pid_rows
        matched = ambiguous = unmatched = 0
        for row in attendees:
            if row['pcoPersonId'] and row['pcoPersonId'] in by_pco:
                tp = by_pco[row['pcoPersonId']]
                row['tpPeopleId'] = tp['peopleId']
                row['tpName'] = tp['name']
                row['matchSource'] = 'pco_id'
                row['matchStatus'] = 'matched'
                matched += 1
                continue
            if row['email'] and row['email'] in by_email:
                hits = by_email[row['email']]
                if len(hits) == 1:
                    row['tpPeopleId'] = hits[0]['peopleId']
                    row['tpName'] = hits[0]['name']
                    row['matchSource'] = 'email'
                    row['matchStatus'] = 'matched'
                    matched += 1
                    continue
                row['emailAmbiguous'] = True
                row['matchStatus'] = 'ambiguous'
                ambiguous += 1
                continue
            row['matchStatus'] = 'unmatched'
            unmatched += 1

        _rank = {'unmatched': 0, 'ambiguous': 1, 'matched': 2}
        attendees.sort(key=lambda r: (_rank.get(r.get('matchStatus'), 99),
                                       (r.get('name') or '').lower()))

        # v3.2: mirror-removal counts for Service Type Sync.
        pco_in_scope_ids = set([r['pcoPersonId'] for r in by_pid.values() if r.get('pcoPersonId')])
        pco_team_names_lower = set([t['teamName'].lower() for t in teams if t.get('teamName')])
        # Per-PCO-person current team set (lowercase).
        person_teams_lower = {}
        for r in by_pid.values():
            pid = r.get('pcoPersonId')
            if not pid:
                continue
            person_teams_lower[pid] = set([(p.get('position') or '').lower() for p in r.get('positions', []) if p.get('position')])
        tp_members_p = _tp_org_members_with_pco_link(info['orgId'])
        tp_subgroups_p = _tp_org_subgroups(info['orgId'])
        roster_drop_ids_p = []
        for tp_pid, pco_pid in tp_members_p.items():
            if pco_pid and pco_pid not in pco_in_scope_ids:
                roster_drop_ids_p.append(tp_pid)
        subgroup_drops_p = []
        if info['teamsAsSubgroups']:
            pco_id_by_tp_p = dict([(tpid, pcoid) for tpid, pcoid in tp_members_p.items() if pcoid])
            for tp_pid, sg_names in tp_subgroups_p.items():
                pco_pid = pco_id_by_tp_p.get(tp_pid, '')
                if not pco_pid:
                    continue
                desired = person_teams_lower.get(pco_pid, set())
                for sg_lower in sg_names:
                    if sg_lower in pco_team_names_lower and sg_lower not in desired:
                        subgroup_drops_p.append((tp_pid, sg_lower))

        plan_info = {
            'isPeopleSync': True,
            'pcoServiceTypeId': pco_st_id,
            'pcoServiceTypeName': info['pcoServiceTypeName'],
            'serviceTypeName': info['pcoServiceTypeName'] or 'People Sync',
            'title': (info['pcoServiceTypeName'] or 'Service Type') + ' (People Sync)',
            'planTitle': '',
            'shortDates': str(len(attendees)) + ' people across ' + str(len(teams)) + ' team(s)',
            'syncAttendance': False,
            'autoAddMember': info['autoAddMember'],
            'teamsAsSubgroups': info['teamsAsSubgroups'],
            'tpOrgId': info['orgId'],
            'teams': teams,
            'rosterDropCount': len(roster_drop_ids_p),
            'subgroupDropCount': len(subgroup_drops_p),
        }
        resp = {
            'success': True,
            'planInfo': plan_info,
            'attendees': attendees,
            'summary': {
                'total': len(attendees),
                'matched': matched,
                'ambiguousEmail': ambiguous,
                'unmatched': unmatched,
                'confirmed': len(attendees),
                'rawAssignments': 0,
            },
        }
        if api_errors:
            resp['warnings'] = api_errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load people sync preview failed: ' + str(e)})

def handle_sync_people():
    """Write People Sync: JoinOrg + (optional) team-as-subgroup writes."""
    try:
        pco_st_id = safe_str(get_data('pco_service_type_id', '')).strip()
        tp_org_id = safe_int(get_data('tp_org_id', 0), 0)
        people_ids_csv = safe_str(get_data('people_ids', '')).strip()
        if not pco_st_id or tp_org_id <= 0 or not people_ids_csv:
            print json.dumps({'success': False, 'message': 'Missing required args'})
            return
        mappings = load_json(PEOPLE_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict):
            mappings = {}
        info = _parse_people_mapping(mappings.get(pco_st_id))
        if info['orgId'] != tp_org_id:
            print json.dumps({'success': False, 'message': 'Mapping mismatch -- service type is mapped to a different org now.'})
            return
        if not info['autoAddMember']:
            print json.dumps({'success': False, 'message': 'Auto-add is disabled for this people mapping.'})
            return

        pids = []
        for tok in people_ids_csv.split(','):
            pid = safe_int(tok.strip(), 0)
            if pid > 0:
                pids.append(pid)
        if not pids:
            print json.dumps({'success': False, 'message': 'No people ids to sync'})
            return

        teams_by_pid = {}
        teams_raw = safe_str(get_data('teams_json', ''))
        if teams_raw:
            try:
                tp = json.loads(teams_raw)
                if isinstance(tp, dict):
                    for k, v in tp.items():
                        try:
                            tp_id = int(k)
                        except:
                            continue
                        if isinstance(v, list):
                            teams_by_pid[tp_id] = [str(x) for x in v if x]
            except:
                pass

        existing_member_ids = set()
        try:
            mem_sql = ("""
                SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                WHERE OrganizationId = %s
                  AND PeopleId IN (%s)
                  AND InactiveDate IS NULL
            """ % (int(tp_org_id), ','.join(str(p) for p in pids)))
            for r in q.QuerySql(mem_sql):
                existing_member_ids.add(int(r.PeopleId))
        except:
            pass

        person_rules = normalize_person_sync_rules(load_json(PERSON_SYNC_RULES_KEY, {}))
        person_rules_active = person_sync_rules_active(person_rules)
        pco_data_by_pid = _parse_people_json_payload(safe_str(get_data('people_json', '')))
        person_auto = 0
        person_queued = 0

        joined = 0
        already = 0
        subgroup_adds = 0
        subgroup_failures = 0
        skipped = 0
        per_person = []
        for pid in pids:
            try:
                if pid in existing_member_ids:
                    already += 1
                    per_person.append({'peopleId': pid, 'status': 'already_member'})
                else:
                    person = model.GetPerson(pid)
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(tp_org_id), person)
                            joined += 1
                            per_person.append({'peopleId': pid, 'status': 'joined'})
                        except Exception as je:
                            jm = str(je).lower()
                            if 'already' in jm:
                                already += 1
                                per_person.append({'peopleId': pid, 'status': 'already_member'})
                            else:
                                skipped += 1
                                per_person.append({'peopleId': pid, 'status': 'join_failed', 'message': str(je)})
                                continue
                    else:
                        skipped += 1
                        per_person.append({'peopleId': pid, 'status': 'no_person'})
                        continue
                # Subgroup writes per team the person belongs to.
                if info['teamsAsSubgroups']:
                    for team_name in teams_by_pid.get(pid, []):
                        try:
                            model.AddSubGroup(int(pid), int(tp_org_id), team_name)
                            subgroup_adds += 1
                        except Exception as se:
                            subgroup_failures += 1
                if person_rules_active and pid in pco_data_by_pid:
                    try:
                        c = apply_person_sync_for_one(pid, pco_data_by_pid[pid], person_rules)
                        person_auto += c.get('auto', 0)
                        person_queued += c.get('queued', 0)
                    except:
                        pass
            except Exception as we:
                skipped += 1
                per_person.append({'peopleId': pid, 'status': 'error', 'message': str(we)})

        # v3.2: mirror removals for Service Type Sync. Re-walk PCO to
        # compute the in-scope set and current team membership.
        roster_drops = 0
        roster_drop_failures = 0
        subgroup_drops = 0
        subgroup_drop_failures = 0
        try:
            teams_now, _terr = _pco_teams_for_service_type(pco_st_id)
            pco_in_scope = set()
            person_teams_now = {}  # pco_pid -> set(team_name_lower)
            pco_team_names_lower = set([t['teamName'].lower() for t in (teams_now or []) if t.get('teamName')])
            for t in (teams_now or []):
                tname_l = (t['teamName'] or '').lower()
                ppl, _perr = _pco_team_people(t['teamId'])
                for p in (ppl or []):
                    pp = p.get('pcoPersonId')
                    if not pp:
                        continue
                    pco_in_scope.add(pp)
                    if tname_l:
                        person_teams_now.setdefault(pp, set()).add(tname_l)
            tp_members_now = _tp_org_members_with_pco_link(int(tp_org_id))
            for tp_pid, pco_pid in tp_members_now.items():
                if pco_pid and pco_pid not in pco_in_scope:
                    try:
                        person = model.GetPerson(int(tp_pid))
                        if person and person.PeopleId:
                            model.RemoveFromOrg(person, int(tp_org_id))
                            roster_drops += 1
                    except Exception:
                        roster_drop_failures += 1
            if info['teamsAsSubgroups']:
                tp_subgroups_now = _tp_org_subgroups(int(tp_org_id))
                pco_lower_to_display = {}
                for t in (teams_now or []):
                    if t.get('teamName'):
                        pco_lower_to_display[t['teamName'].lower()] = t['teamName']
                for tp_pid, sg_lower_set in tp_subgroups_now.items():
                    pco_pid = tp_members_now.get(tp_pid, '')
                    if not pco_pid:
                        continue
                    desired = person_teams_now.get(pco_pid, set())
                    for sg_lower in sg_lower_set:
                        if sg_lower in pco_team_names_lower and sg_lower not in desired:
                            display_name = pco_lower_to_display.get(sg_lower, sg_lower)
                            try:
                                model.RemoveSubGroup(int(tp_pid), int(tp_org_id), display_name)
                                subgroup_drops += 1
                            except Exception:
                                subgroup_drop_failures += 1
        except Exception:
            pass

        append_audit({
            'action': 'sync_people',
            'pcoServiceTypeId': pco_st_id,
            'tpOrgId': tp_org_id,
            'joined': joined,
            'alreadyMember': already,
            'subgroupAdds': subgroup_adds,
            'subgroupFailures': subgroup_failures,
            'rosterDrops': roster_drops,
            'rosterDropFailures': roster_drop_failures,
            'subgroupDrops': subgroup_drops,
            'subgroupDropFailures': subgroup_drop_failures,
            'skipped': skipped,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })

        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict): s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)

        if joined == 0 and already > 0:
            msg = 'Everyone was already an active member.'
        else:
            msg = 'Added ' + str(joined) + ' member(s) to the involvement.'
            if already > 0:
                msg += ' (' + str(already) + ' already on the roster.)'
        if subgroup_adds:
            msg += ' Wrote ' + str(subgroup_adds) + ' team-subgroup association(s).'
        if roster_drops:
            msg += ' Removed ' + str(roster_drops) + ' member(s) no longer in PCO.'
        if subgroup_drops:
            msg += ' Removed ' + str(subgroup_drops) + ' stale team-subgroup association(s).'
        if person_auto:
            msg += ' Updated ' + str(person_auto) + ' person field(s).'
        if person_queued:
            msg += ' Queued ' + str(person_queued) + ' person change(s) for review.'
        print json.dumps({
            'success': True,
            'message': msg,
            'joinedOrg': joined,
            'alreadyMember': already,
            'subgroupAdds': subgroup_adds,
            'subgroupFailures': subgroup_failures,
            'rosterDrops': roster_drops,
            'rosterDropFailures': roster_drop_failures,
            'subgroupDrops': subgroup_drops,
            'subgroupDropFailures': subgroup_drop_failures,
            'skipped': skipped,
            'personAuto': person_auto,
            'personQueued': person_queued,
            'perPerson': per_person,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Sync people failed: ' + str(e)})

# =====================================================================
# All People Sync (v2.2+)
# =====================================================================
# Singleton mapping: PCO People directory -> ONE TP involvement. Used by
# churches who want every PCO record reflected as a TP involvement member.
# TouchPoint stays authoritative -- we never create TP records; PCO
# records without a TP match are surfaced as unmatched (counted in the
# response so staff know how much manual matching is left).

def handle_load_all_people_mapping():
    try:
        v = load_json(ALL_PEOPLE_MAPPING_KEY, {})
        info = _parse_all_people_mapping(v)
        org_name = ''
        org_division = ''
        org_program = ''
        if info['orgId'] > 0:
            sql = """
                SELECT TOP 1 o.OrganizationName,
                       ISNULL(d.Name, '') AS DivisionName,
                       ISNULL(p.Name, '') AS ProgramName
                FROM Organizations o WITH (NOLOCK)
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE o.OrganizationId = %s
            """ % int(info['orgId'])
            for r in q.QuerySql(sql):
                org_name = safe_str(r.OrganizationName)
                org_division = safe_str(r.DivisionName)
                org_program = safe_str(r.ProgramName)
                break
        sch = info.get('schedule', _parse_schedule({}))
        nxt = _next_scheduled_run(sch)
        print json.dumps({
            'success': True,
            'mapping': info,
            'tpOrgName': org_name,
            'tpDivision': org_division,
            'tpProgram': org_program,
            'schedule': sch,
            'scheduleLabel': _format_schedule_label(sch),
            'nextRunIso': nxt.strftime('%Y-%m-%dT%H:%M:%S') if nxt else '',
            'schedulerInstalled': _scheduler_installed(),
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load all-people mapping failed: ' + str(e)})

def handle_save_all_people_mapping():
    try:
        org_id = safe_int(get_data('tp_org_id', 0), 0)
        include_inactive = safe_str(get_data('include_inactive', '0')).strip().lower() in ('1', 'true', 'yes', 'on')
        if org_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing tp_org_id'})
            return
        verify_sql = """
            SELECT TOP 1 o.OrganizationId, o.OrganizationName
            FROM Organizations o WITH (NOLOCK)
            WHERE o.OrganizationId = %s AND o.OrganizationStatusId = 30
        """ % str(org_id)
        org_name = ''
        for r in q.QuerySql(verify_sql):
            org_name = safe_str(r.OrganizationName)
            break
        if not org_name:
            print json.dumps({'success': False, 'message': 'Org #' + str(org_id) + ' is not active or does not exist.'})
            return
        existing = _parse_all_people_mapping(load_json(ALL_PEOPLE_MAPPING_KEY, {}))
        new_val = {
            'orgId': org_id,
            'autoAddMember': existing['autoAddMember'] if existing['orgId'] > 0 else True,
            'includeInactive': include_inactive,
            # v3.3: preserve scheduler config across mapping saves.
            'schedule': existing.get('schedule', _parse_schedule({})),
        }
        save_ok = save_json(ALL_PEOPLE_MAPPING_KEY, new_val)
        verify = _parse_all_people_mapping(load_json(ALL_PEOPLE_MAPPING_KEY, {}))
        verified = (save_ok and verify['orgId'] == org_id)
        append_audit({
            'action': 'save_all_people_mapping',
            'tpOrgId': org_id,
            'tpOrgName': org_name,
            'includeInactive': include_inactive,
            'saveOk': save_ok,
            'verified': verified,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        if not verified:
            print json.dumps({'success': False, 'message': 'Save did not persist -- check role on PCOSync_AllPeopleMapping content.'})
            return
        print json.dumps({'success': True, 'message': 'All People mapped to ' + org_name + '.', 'orgName': org_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Save all-people mapping failed: ' + str(e)})

def handle_delete_all_people_mapping():
    try:
        save_json(ALL_PEOPLE_MAPPING_KEY, {})
        append_audit({
            'action': 'delete_all_people_mapping',
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'message': 'All People mapping removed.'})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Delete all-people mapping failed: ' + str(e)})

def handle_load_all_people_preview():
    """Walk the PCO People directory, match each person against TP, and
    return counts + sample of unmatched. Used by the modal before the
    user commits a sync. Heavy operation -- can take 10-30 seconds for a
    large church (5K+ records)."""
    try:
        info = _parse_all_people_mapping(load_json(ALL_PEOPLE_MAPPING_KEY, {}))
        if info['orgId'] <= 0:
            print json.dumps({'success': False, 'message': 'No All People mapping configured.'})
            return
        people, errors = _pco_all_people_walk(include_inactive=info['includeInactive'])
        pco_ids = [p['pcoPersonId'] for p in people if p['pcoPersonId']]
        emails = [p['email'] for p in people if p['email']]
        by_pco = _tp_match_by_pco_id(pco_ids)
        by_email = _tp_match_by_email(emails)
        matched = 0
        ambiguous = 0
        unmatched_sample = []
        for p in people:
            pid = p['pcoPersonId']
            if pid and pid in by_pco:
                matched += 1
                continue
            if p['email'] and p['email'] in by_email:
                hits = by_email[p['email']]
                if len(hits) == 1:
                    matched += 1
                    continue
                ambiguous += 1
                continue
            if len(unmatched_sample) < 50:
                unmatched_sample.append({
                    'pcoPersonId': pid,
                    'name': (p['first_name'] + ' ' + p['last_name']).strip() or p['name'] or '(Unknown)',
                    'email': p['email'],
                })
        unmatched = len(people) - matched - ambiguous
        # v3.2: mirror-removal count -- TP members with PCO_PersonId that
        # no longer appear in the walked PCO directory.
        pco_in_scope = set([p['pcoPersonId'] for p in people if p.get('pcoPersonId')])
        tp_members_a = _tp_org_members_with_pco_link(info['orgId'])
        roster_drop_count_a = 0
        for tp_pid, pco_pid in tp_members_a.items():
            if pco_pid and pco_pid not in pco_in_scope:
                roster_drop_count_a += 1
        resp = {
            'success': True,
            'totalPco': len(people),
            'matched': matched,
            'ambiguous': ambiguous,
            'unmatched': unmatched,
            'unmatchedSample': unmatched_sample,
            'mapping': info,
            'rosterDropCount': roster_drop_count_a,
        }
        if errors:
            resp['warnings'] = errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load all-people preview failed: ' + str(e)})

def handle_sync_all_people():
    """Walk the PCO People directory, match each person against TP, and
    JoinOrg the matched ones into the designated involvement. Person Sync
    rules apply per the existing infrastructure."""
    try:
        info = _parse_all_people_mapping(load_json(ALL_PEOPLE_MAPPING_KEY, {}))
        if info['orgId'] <= 0:
            print json.dumps({'success': False, 'message': 'No All People mapping configured.'})
            return
        if not info['autoAddMember']:
            print json.dumps({'success': False, 'message': 'Auto-add is disabled for the All People mapping.'})
            return
        people, errors = _pco_all_people_walk(include_inactive=info['includeInactive'])
        pco_ids = [p['pcoPersonId'] for p in people if p['pcoPersonId']]
        emails = [p['email'] for p in people if p['email']]
        by_pco = _tp_match_by_pco_id(pco_ids)
        by_email = _tp_match_by_email(emails)

        # Build the list of TP person ids to act on + the PCO data payload
        # for the Person Sync hook. Skip ambiguous-email rows -- they need
        # manual disambiguation, not blind auto-join.
        matched_pids = []
        pco_data_by_pid = {}
        matched_count = 0
        ambiguous_count = 0
        unmatched_count = 0
        for p in people:
            pid = p['pcoPersonId']
            tp_id = 0
            if pid and pid in by_pco:
                tp_id = int(by_pco[pid]['peopleId'])
                matched_count += 1
            elif p['email'] and p['email'] in by_email:
                hits = by_email[p['email']]
                if len(hits) == 1:
                    tp_id = int(hits[0]['peopleId'])
                    matched_count += 1
                else:
                    ambiguous_count += 1
                    continue
            else:
                unmatched_count += 1
                continue
            matched_pids.append(tp_id)
            pco_data_by_pid[tp_id] = {
                'isConfirmed': True,
                'first_name': p['first_name'],
                'last_name': p['last_name'],
                'email': p['email'],
            }

        if not matched_pids:
            print json.dumps({
                'success': True,
                'message': 'No matched PCO people to sync. ' + str(unmatched_count) + ' unmatched, ' + str(ambiguous_count) + ' ambiguous.',
                'totalPco': len(people),
                'matched': 0,
                'joined': 0,
                'alreadyMember': 0,
                'unmatched': unmatched_count,
                'ambiguous': ambiguous_count,
            })
            return

        tp_org_id = info['orgId']
        existing_member_ids = set()
        # Pre-check is large here; chunk the IN clause.
        for i in range(0, len(matched_pids), 500):
            chunk = matched_pids[i:i+500]
            try:
                mem_sql = ("""
                    SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                    WHERE OrganizationId = %s
                      AND PeopleId IN (%s)
                      AND InactiveDate IS NULL
                """ % (int(tp_org_id), ','.join(str(p) for p in chunk)))
                for r in q.QuerySql(mem_sql):
                    existing_member_ids.add(int(r.PeopleId))
            except:
                pass

        person_rules = normalize_person_sync_rules(load_json(PERSON_SYNC_RULES_KEY, {}))
        person_rules_active = person_sync_rules_active(person_rules)
        person_auto = 0
        person_queued = 0

        joined = 0
        already = 0
        skipped = 0
        for pid in matched_pids:
            try:
                if pid in existing_member_ids:
                    already += 1
                else:
                    person = model.GetPerson(pid)
                    if person and person.PeopleId:
                        try:
                            model.JoinOrg(int(tp_org_id), person)
                            joined += 1
                        except Exception as je:
                            jm = str(je).lower()
                            if 'already' in jm:
                                already += 1
                            else:
                                skipped += 1
                                continue
                    else:
                        skipped += 1
                        continue
                if person_rules_active and pid in pco_data_by_pid:
                    try:
                        c = apply_person_sync_for_one(pid, pco_data_by_pid[pid], person_rules)
                        person_auto += c.get('auto', 0)
                        person_queued += c.get('queued', 0)
                    except:
                        pass
            except:
                skipped += 1

        # v3.2: mirror removals for All People Sync. PCO directory is
        # source of truth -- TP members whose PCO_PersonId no longer
        # appears in the walked directory get dropped. No subgroup
        # handling here (All People has no subgroup concept).
        roster_drops = 0
        roster_drop_failures = 0
        try:
            pco_in_scope_set = set([p['pcoPersonId'] for p in people if p.get('pcoPersonId')])
            tp_members_now = _tp_org_members_with_pco_link(int(tp_org_id))
            for tp_pid, pco_pid in tp_members_now.items():
                if pco_pid and pco_pid not in pco_in_scope_set:
                    try:
                        person = model.GetPerson(int(tp_pid))
                        if person and person.PeopleId:
                            model.RemoveFromOrg(person, int(tp_org_id))
                            roster_drops += 1
                    except Exception:
                        roster_drop_failures += 1
        except Exception:
            pass

        append_audit({
            'action': 'sync_all_people',
            'tpOrgId': tp_org_id,
            'totalPco': len(people),
            'matched': matched_count,
            'joined': joined,
            'alreadyMember': already,
            'unmatched': unmatched_count,
            'ambiguous': ambiguous_count,
            'rosterDrops': roster_drops,
            'rosterDropFailures': roster_drop_failures,
            'skipped': skipped,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })

        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict): s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)

        msg = 'Walked ' + str(len(people)) + ' PCO record(s). '
        msg += 'Matched ' + str(matched_count) + ', joined ' + str(joined) + ', already member ' + str(already) + '. '
        if roster_drops:
            msg += 'Removed ' + str(roster_drops) + ' TP member(s) no longer in PCO. '
        if unmatched_count:
            msg += str(unmatched_count) + ' unmatched (PCO record has no TP match -- check Unmatched People). '
        if ambiguous_count:
            msg += str(ambiguous_count) + ' ambiguous (email matches multiple TP people). '
        if person_auto:
            msg += 'Updated ' + str(person_auto) + ' person field(s). '
        if person_queued:
            msg += 'Queued ' + str(person_queued) + ' person change(s) for review. '
        resp = {
            'success': True,
            'message': msg.strip(),
            'totalPco': len(people),
            'matched': matched_count,
            'joined': joined,
            'alreadyMember': already,
            'unmatched': unmatched_count,
            'ambiguous': ambiguous_count,
            'rosterDrops': roster_drops,
            'rosterDropFailures': roster_drop_failures,
            'skipped': skipped,
            'personAuto': person_auto,
            'personQueued': person_queued,
        }
        if errors:
            resp['warnings'] = errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Sync All People failed: ' + str(e)})

# =====================================================================
# Sync Mappings Health (v2.3+)
# =====================================================================
# Walks every saved mapping and verifies both endpoints still exist:
#   TP side -- OrganizationStatusId = 30 (active)
#   PCO side -- referenced service type / team still appears in fresh
#               API responses
# Anything broken gets returned with severity + recommended action so the
# dashboard can surface it before staff hit Sync and wonder why nothing
# happens.

# =====================================================================
# All People Proposed-Match Review (v2.5+)
# =====================================================================
# Walks the full PCO People directory, tries to match each unmatched PCO
# record against TP People using name + birthdate + email signals, and
# returns proposed candidates with confidence tier so staff can review
# in bulk. NEVER applies a match silently -- everything goes through
# Apply, even strong-tier proposals.

def handle_load_all_people_proposed_matches():
    try:
        # v2.5.6+: server sends ALL scored rows on a single load. The
        # JS caches the dataset and handles tier/search/page changes
        # locally so flipping filters doesn't re-walk PCO every time.
        # Old tier_filter / search_term / page params are read but
        # IGNORED here -- kept silent for back-compat with stale tabs.
        # Scoped mode: JS passes a subset_json bundle of PCO people from
        # a specific preview (e.g. "the 199 unmatched in this People
        # Sync preview"). We still walk the full PCO People directory
        # to enrich each row with email + birthdate (those aren't on
        # team_members), then FILTER to the scoped IDs before scoring.
        # Same walk cost as full mode but the result table only shows
        # records relevant to where the user came from.
        subset_raw = safe_str(get_data('subset_json', ''))
        scope_ids = None
        scope_fallback = {}  # pcoId -> dict (data from subset, used if PCO walk skips them)
        scoped = False
        if subset_raw:
            try:
                subset = json.loads(subset_raw)
            except:
                subset = None
            if isinstance(subset, list) and subset:
                scoped = True
                scope_ids = set()
                for it in subset:
                    if not isinstance(it, dict):
                        continue
                    pid = safe_str(it.get('pcoPersonId', ''))
                    if not pid:
                        continue
                    scope_ids.add(pid)
                    scope_fallback[pid] = {
                        'pcoPersonId': pid,
                        'first_name': safe_str(it.get('first_name', '')),
                        'last_name':  safe_str(it.get('last_name', '')),
                        'name': safe_str(it.get('name', '')),
                        'email': safe_str(it.get('email', '')).strip().lower(),
                        'birthdate': safe_str(it.get('birthdate', '')),
                        'status': 'active',
                    }

        info = _parse_all_people_mapping(load_json(ALL_PEOPLE_MAPPING_KEY, {}))
        include_inactive_default = False
        if info['orgId'] > 0:
            include_inactive_default = info['includeInactive']
        # Both scoped and full modes walk PCO People for the enriched
        # data. Full mode REQUIRES the All People mapping (so the user
        # can eventually Sync). Scoped mode does not -- the user is
        # just resolving matches from a preview, which has its own
        # downstream sync action.
        if scope_ids is None and info['orgId'] <= 0:
            print json.dumps({'success': False, 'message': 'No All People mapping configured. Either configure one on the Sync Mappings tab, or click "Open Proposed Matches" from a specific preview to scope this to those records.'})
            return
        people, errors = _pco_all_people_walk(include_inactive=include_inactive_default)
        if scope_ids is not None:
            # Filter the walk to scoped IDs only. Anyone in the scope
            # who wasn't in the PCO walk (deleted? archived? excluded
            # by include_inactive=False?) gets the fallback row so the
            # user can still see and act on them.
            found_ids = set()
            filtered = []
            for p in people:
                if p['pcoPersonId'] in scope_ids:
                    filtered.append(p)
                    found_ids.add(p['pcoPersonId'])
            for pid, fb in scope_fallback.items():
                if pid not in found_ids:
                    filtered.append(fb)
            people = filtered

        # Drop people already locked (PCO_PersonId match) -- those don't
        # belong in the review queue.
        skip_set = load_all_people_skip()
        pco_ids = [p['pcoPersonId'] for p in people if p['pcoPersonId']]
        by_pco_locked = _tp_match_by_pco_id(pco_ids)

        unmatched_pco = []
        for p in people:
            pid = p['pcoPersonId']
            if pid and pid in by_pco_locked:
                continue
            unmatched_pco.append(p)

        # Bulk lookups for name + email matching.
        name_pairs = [(p.get('first_name'), p.get('last_name')) for p in unmatched_pco]
        emails = [p['email'] for p in unmatched_pco if p['email']]
        by_name = _tp_match_by_name(name_pairs)
        by_email = _tp_match_by_email(emails)

        # Score each row + compute the tier. Client paginates / filters
        # / searches the whole dataset from cache after this single load.
        rows = []
        counts = {'strong': 0, 'medium': 0, 'weak': 0, 'none': 0, 'skipped': 0}
        for p in unmatched_pco:
            cands = _score_pco_to_tp_candidates(p, by_name, by_email)
            top_score = cands[0]['score'] if cands else 0
            tier = _tier_for_score(top_score) if cands else 'none'
            is_skipped = p['pcoPersonId'] in skip_set
            row = {
                'pcoPersonId': p['pcoPersonId'],
                'pcoName': (p['first_name'] + ' ' + p['last_name']).strip() or p['name'] or '(Unknown)',
                'pcoFirstName': p['first_name'],
                'pcoLastName': p['last_name'],
                'pcoEmail': p['email'],
                'pcoBirthdate': p['birthdate'],
                'candidates': cands[:3],  # top 3 only -- keeps payload sane
                'topScore': top_score,
                'tier': tier,
                'skipped': is_skipped,
            }
            rows.append(row)
            if is_skipped:
                counts['skipped'] += 1
            else:
                counts[tier] += 1
        # Sort: strongest first, then name alpha. Client preserves this order.
        rows.sort(key=lambda r: (-r['topScore'], r['pcoName'].lower()))

        resp = {
            'success': True,
            'totalPco': len(people),
            'lockedCount': len(by_pco_locked),
            'unmatchedCount': len(unmatched_pco),
            'counts': counts,
            'rows': rows,  # ALL rows -- client paginates/filters
            'scoped': scoped,
        }
        if errors:
            resp['warnings'] = errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load proposed matches failed: ' + str(e)})

def handle_apply_proposed_match():
    """Write PCO_PersonId extra value on the chosen TP person. Single
    pair -- the JS sends one (pcoPersonId, tpPeopleId) per Apply click."""
    try:
        pco_person_id = safe_str(get_data('pco_person_id', '')).strip()
        tp_people_id = safe_int(get_data('tp_people_id', 0), 0)
        if not pco_person_id or tp_people_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing pco_person_id or tp_people_id'})
            return
        try:
            person = model.GetPerson(int(tp_people_id))
            if not person or not person.PeopleId:
                print json.dumps({'success': False, 'message': 'TP person not found.'})
                return
            model.AddExtraValueText(int(tp_people_id), 'PCO_PersonId', pco_person_id)
        except Exception as ue:
            print json.dumps({'success': False, 'message': 'Write failed: ' + str(ue)})
            return
        append_audit({
            'action': 'apply_proposed_match',
            'pcoPersonId': pco_person_id,
            'tpPeopleId': int(tp_people_id),
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Apply proposed match failed: ' + str(e)})

def handle_bulk_apply_proposed_matches():
    """Apply many (pco_id, tp_id) pairs in one call. JS sends pairs_json
    as [[pcoId, tpId], ...]."""
    try:
        pairs_raw = safe_str(get_data('pairs_json', ''))
        if not pairs_raw:
            print json.dumps({'success': False, 'message': 'Missing pairs_json'})
            return
        try:
            pairs = json.loads(pairs_raw)
        except:
            pairs = []
        if not isinstance(pairs, list) or not pairs:
            print json.dumps({'success': False, 'message': 'No pairs supplied'})
            return
        applied = 0
        skipped = 0
        per_row = []
        for pair in pairs:
            if not isinstance(pair, list) or len(pair) < 2:
                skipped += 1
                continue
            pco_id = safe_str(pair[0]).strip()
            try:
                tp_id = int(pair[1])
            except:
                tp_id = 0
            if not pco_id or tp_id <= 0:
                skipped += 1
                continue
            try:
                person = model.GetPerson(tp_id)
                if not person or not person.PeopleId:
                    skipped += 1
                    per_row.append({'pcoPersonId': pco_id, 'tpPeopleId': tp_id, 'status': 'tp_missing'})
                    continue
                model.AddExtraValueText(tp_id, 'PCO_PersonId', pco_id)
                applied += 1
                per_row.append({'pcoPersonId': pco_id, 'tpPeopleId': tp_id, 'status': 'applied'})
            except Exception as ue:
                skipped += 1
                per_row.append({'pcoPersonId': pco_id, 'tpPeopleId': tp_id, 'status': 'error', 'message': str(ue)})
        append_audit({
            'action': 'bulk_apply_proposed_matches',
            'applied': applied,
            'skipped': skipped,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'applied': applied, 'skipped': skipped, 'perRow': per_row})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Bulk apply failed: ' + str(e)})

def handle_skip_pco_person_forever():
    try:
        pco_person_id = safe_str(get_data('pco_person_id', '')).strip()
        unskip = safe_str(get_data('unskip', '0')).strip() in ('1', 'true', 'yes', 'on')
        if not pco_person_id:
            print json.dumps({'success': False, 'message': 'Missing pco_person_id'})
            return
        skip_set = load_all_people_skip()
        if unskip:
            skip_set.discard(pco_person_id)
        else:
            skip_set.add(pco_person_id)
        save_all_people_skip(skip_set)
        append_audit({
            'action': 'skip_pco_person_forever' if not unskip else 'unskip_pco_person',
            'pcoPersonId': pco_person_id,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'skipped': not unskip})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Skip forever failed: ' + str(e)})

def handle_load_sync_mappings_health():
    try:
        issues = []
        checked = {'orgMappings': 0, 'teamMappings': 0, 'peopleMappings': 0, 'allPeopleMapping': 0}

        # Pull all four mapping stores.
        org_map_raw = load_json(ORG_MAPPINGS_KEY, {})
        team_map_raw = load_json(TEAM_MAPPINGS_KEY, {})
        people_map_raw = load_json(PEOPLE_MAPPINGS_KEY, {})
        all_people_raw = load_json(ALL_PEOPLE_MAPPING_KEY, {})
        if not isinstance(org_map_raw, dict): org_map_raw = {}
        if not isinstance(team_map_raw, dict): team_map_raw = {}
        if not isinstance(people_map_raw, dict): people_map_raw = {}

        # Collect all TP org ids referenced anywhere so we can verify in
        # one SQL pass.
        tp_org_ids_needed = set()
        for v in org_map_raw.values():
            oid = _mapping_org_id(v)
            if oid > 0: tp_org_ids_needed.add(oid)
        for v in team_map_raw.values():
            info = _parse_team_mapping(v)
            if info['orgId'] > 0: tp_org_ids_needed.add(info['orgId'])
        for v in people_map_raw.values():
            info = _parse_people_mapping(v)
            if info['orgId'] > 0: tp_org_ids_needed.add(info['orgId'])
        ap_info = _parse_all_people_mapping(all_people_raw)
        if ap_info['orgId'] > 0:
            tp_org_ids_needed.add(ap_info['orgId'])

        tp_org_status = {}  # orgId -> {name, active}
        if tp_org_ids_needed:
            ids_csv = ','.join(str(i) for i in tp_org_ids_needed)
            sql = """
                SELECT OrganizationId, OrganizationName, OrganizationStatusId
                FROM Organizations WITH (NOLOCK)
                WHERE OrganizationId IN (%s)
            """ % ids_csv
            for r in q.QuerySql(sql):
                tp_org_status[int(r.OrganizationId)] = {
                    'name': safe_str(r.OrganizationName),
                    'active': int(r.OrganizationStatusId or 0) == 30,
                }

        # PCO side: pull current service-type list once. Used for both
        # org mappings and people mappings; team mappings need per-service
        # team lookups (do those on demand for the affected service types).
        pco_st_by_id = {}
        try:
            st_data, st_err = pco_get('/services/v2/service_types?per_page=100')
            if not st_err and isinstance(st_data, dict):
                for st in (st_data.get('data') or []):
                    sid = safe_str(st.get('id', ''))
                    attrs = st.get('attributes') or {}
                    if sid:
                        pco_st_by_id[sid] = {'name': safe_str(attrs.get('name', ''))}
        except:
            pass

        # Helper: report a TP-side issue (org missing or inactive).
        def _check_tp_org(org_id, label, mapping_type, mapping_key):
            if org_id <= 0:
                return False  # already handled by save layer
            status = tp_org_status.get(int(org_id))
            if status is None:
                issues.append({
                    'mappingType': mapping_type,
                    'mappingKey': mapping_key,
                    'mappingLabel': label,
                    'severity': 'error',
                    'side': 'tp',
                    'message': 'TouchPoint involvement #' + str(org_id) + ' no longer exists (was it deleted?).',
                    'tpOrgId': org_id,
                })
                return True
            if not status['active']:
                issues.append({
                    'mappingType': mapping_type,
                    'mappingKey': mapping_key,
                    'mappingLabel': label,
                    'severity': 'error',
                    'side': 'tp',
                    'message': 'TouchPoint involvement "' + status['name'] + '" (#' + str(org_id) + ') is archived/inactive.',
                    'tpOrgId': org_id,
                })
                return True
            return False

        # 1) Service Type Mappings (Service Plan Sync).
        for pco_st_id, v in org_map_raw.items():
            checked['orgMappings'] += 1
            org_id, _sa, _am = _parse_mapping(v)
            st_name = (pco_st_by_id.get(pco_st_id) or {}).get('name', '')
            label = (st_name or 'Service Type ' + str(pco_st_id)) + ' (Service Plan Sync)'
            _check_tp_org(org_id, label, 'service_type', pco_st_id)
            if pco_st_by_id and pco_st_id not in pco_st_by_id:
                issues.append({
                    'mappingType': 'service_type',
                    'mappingKey': pco_st_id,
                    'mappingLabel': label,
                    'severity': 'error',
                    'side': 'pco',
                    'message': 'PCO Service Type ' + str(pco_st_id) + ' no longer exists in PCO (deleted or PAT lost access).',
                })

        # 2) People Mappings (Service Type umbrella).
        for pco_st_id, v in people_map_raw.items():
            checked['peopleMappings'] += 1
            info = _parse_people_mapping(v)
            label = (info['pcoServiceTypeName'] or 'Service Type ' + str(pco_st_id)) + ' (People Sync)'
            _check_tp_org(info['orgId'], label, 'people', pco_st_id)
            if pco_st_by_id and pco_st_id not in pco_st_by_id:
                issues.append({
                    'mappingType': 'people',
                    'mappingKey': pco_st_id,
                    'mappingLabel': label,
                    'severity': 'error',
                    'side': 'pco',
                    'message': 'PCO Service Type ' + str(pco_st_id) + ' no longer exists in PCO.',
                })

        # 3) Team Mappings. Per-service-type team fetch is on-demand so
        # we don't blow API budget. Cache fetched team lists.
        team_lists_cache = {}  # service_type_id -> {team_id: name}
        def _teams_for(st_id):
            if not st_id: return None
            if st_id in team_lists_cache:
                return team_lists_cache[st_id]
            tlist, terr = _pco_teams_for_service_type(st_id)
            if terr:
                team_lists_cache[st_id] = None
                return None
            m = {}
            for t in tlist:
                tid = t.get('teamId') or ''
                if tid:
                    m[tid] = t.get('teamName', '')
            team_lists_cache[st_id] = m
            return m
        for pco_team_id, v in team_map_raw.items():
            checked['teamMappings'] += 1
            info = _parse_team_mapping(v)
            label = (info['pcoTeamName'] or 'Team ' + str(pco_team_id))
            if info['pcoServiceTypeName']:
                label += ' (' + info['pcoServiceTypeName'] + ')'
            label += ' (Team Sync)'
            _check_tp_org(info['orgId'], label, 'team', pco_team_id)
            # Verify service type + team still exist in PCO.
            st_id = info['pcoServiceTypeId']
            if pco_st_by_id and st_id and st_id not in pco_st_by_id:
                issues.append({
                    'mappingType': 'team',
                    'mappingKey': pco_team_id,
                    'mappingLabel': label,
                    'severity': 'error',
                    'side': 'pco',
                    'message': 'Parent PCO Service Type ' + str(st_id) + ' no longer exists -- its teams are unreachable.',
                })
                continue
            if st_id:
                team_map = _teams_for(st_id)
                if team_map is not None and pco_team_id not in team_map:
                    issues.append({
                        'mappingType': 'team',
                        'mappingKey': pco_team_id,
                        'mappingLabel': label,
                        'severity': 'error',
                        'side': 'pco',
                        'message': 'PCO Team ' + str(pco_team_id) + ' is gone from PCO (deleted or renamed at a new id).',
                    })

        # 4) All People Mapping (singleton -- just verify TP org).
        if ap_info['orgId'] > 0:
            checked['allPeopleMapping'] = 1
            _check_tp_org(ap_info['orgId'], 'All PCO People', 'all_people', '')

        # Roll up severity for the dashboard banner color.
        has_errors = any(i.get('severity') == 'error' for i in issues)
        has_warnings = any(i.get('severity') == 'warning' for i in issues)

        print json.dumps({
            'success': True,
            'issues': issues,
            'issueCount': len(issues),
            'hasErrors': has_errors,
            'hasWarnings': has_warnings,
            'checked': checked,
        })
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Health check failed: ' + str(e)})

# =====================================================================
# Verify Person Link (v3.1+)
# =====================================================================
# Once a PCO_PersonId is written, the link is opaque. These handlers let
# staff search a TP person, see their current PCO link with side-by-side
# detail comparison, and either Unlink or Replace it. Catches the
# reactive case ("Alice's info looks weird, what's she linked to?").

def handle_load_verify_details():
    """Given a TP people id, return the TP person details + their
    current PCO_PersonId + (if linked) the PCO person record from PCO.
    The JS shows a side-by-side comparison so staff can eyeball whether
    the link is correct."""
    try:
        tp_id = safe_int(get_data('tp_people_id', 0), 0)
        if tp_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing tp_people_id'})
            return
        sql = """
            SELECT TOP 1 p.PeopleId, p.Name2, p.FirstName, p.NickName, p.LastName,
                   ISNULL(p.EmailAddress, '') AS Email,
                   ISNULL(p.CellPhone, '') AS CellPhone,
                   ISNULL(p.HomePhone, '') AS HomePhone,
                   p.BirthYear, p.BirthMonth, p.BirthDay,
                   p.MemberStatusId, p.ArchivedFlag, p.IsDeceased
            FROM People p WITH (NOLOCK)
            WHERE p.PeopleId = %s
        """ % str(tp_id)
        tp_info = None
        for r in q.QuerySql(sql):
            by = safe_int(r.BirthYear, 0)
            bm = safe_int(r.BirthMonth, 0)
            bd = safe_int(r.BirthDay, 0)
            tp_bdate = ''
            if bm > 0 and bd > 0:
                tp_bdate = '%04d-%02d-%02d' % (by if by > 0 else 0, bm, bd)
            tp_info = {
                'peopleId': int(r.PeopleId),
                'name': safe_str(r.Name2),
                'firstName': safe_str(r.FirstName),
                'nickName': safe_str(r.NickName),
                'lastName': safe_str(r.LastName),
                'email': safe_str(r.Email),
                'cellPhone': safe_str(r.CellPhone),
                'homePhone': safe_str(r.HomePhone),
                'birthdate': tp_bdate,
                'archived': bool(getattr(r, 'ArchivedFlag', 0)),
                'deceased': bool(getattr(r, 'IsDeceased', 0)),
            }
            break
        if not tp_info:
            print json.dumps({'success': False, 'message': 'TP person not found.'})
            return
        # Read the current PCO_PersonId extra value.
        pco_person_id = ''
        try:
            v = model.ExtraValueText(tp_id, 'PCO_PersonId')
            pco_person_id = safe_str(v).strip()
        except:
            pass
        # If they're linked, fetch the PCO record.
        pco_info = None
        pco_fetch_error = ''
        if pco_person_id:
            data, err = pco_get('/people/v2/people/' + pco_person_id + '?include=emails')
            if err:
                pco_fetch_error = err
            elif isinstance(data, dict):
                row = data.get('data') or {}
                attrs = row.get('attributes') or {}
                # Walk included Email rows to find primary.
                email_addr = ''
                for inc in (data.get('included') or []):
                    if inc.get('type') != 'Email':
                        continue
                    a = inc.get('attributes') or {}
                    addr = safe_str(a.get('address', '')).strip()
                    if not addr:
                        continue
                    if bool(a.get('primary', False)):
                        email_addr = addr
                        break
                    elif not email_addr:
                        email_addr = addr
                bdate_raw = safe_str(attrs.get('birthdate', ''))
                pco_info = {
                    'pcoPersonId': pco_person_id,
                    'name': (safe_str(attrs.get('first_name', '')) + ' ' + safe_str(attrs.get('last_name', ''))).strip() or safe_str(attrs.get('name', '')),
                    'firstName': safe_str(attrs.get('first_name', '')),
                    'lastName': safe_str(attrs.get('last_name', '')),
                    'nickname': safe_str(attrs.get('nickname', '')),
                    'email': email_addr,
                    'birthdate': bdate_raw[:10] if bdate_raw else '',
                    'status': safe_str(attrs.get('status', '')),
                }
        resp = {
            'success': True,
            'tp': tp_info,
            'pcoPersonId': pco_person_id,
            'pco': pco_info,
        }
        if pco_fetch_error:
            resp['pcoFetchError'] = pco_fetch_error
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Load verify details failed: ' + str(e)})

def handle_unlink_tp_person():
    """Clear the PCO_PersonId extra value on a TP person. Doesn't touch
    the TP person otherwise -- name/email/membership stays intact."""
    try:
        tp_id = safe_int(get_data('tp_people_id', 0), 0)
        if tp_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing tp_people_id'})
            return
        old_pco_id = ''
        try:
            old_pco_id = safe_str(model.ExtraValueText(tp_id, 'PCO_PersonId')).strip()
        except:
            pass
        try:
            # Writing an empty string clears the value for our purposes.
            model.AddExtraValueText(tp_id, 'PCO_PersonId', '')
        except Exception as ue:
            print json.dumps({'success': False, 'message': 'Unlink write failed: ' + str(ue)})
            return
        append_audit({
            'action': 'unlink_tp_person',
            'tpPeopleId': tp_id,
            'oldPcoPersonId': old_pco_id,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'oldPcoPersonId': old_pco_id})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Unlink failed: ' + str(e)})

def handle_search_pco_people_by_name():
    """Lightweight PCO People search for the Replace flow. Uses PCO's
    where[search_name_or_email] filter so we don't have to walk the
    whole directory just to relink one person."""
    try:
        term = safe_str(get_data('search_term', '')).strip()
        if not term or len(term) < 2:
            print json.dumps({'success': True, 'people': []})
            return
        # URL-encode the term defensively. PCO's filter accepts
        # spaces but other chars can break the query.
        enc = term.replace(' ', '+')
        path = '/people/v2/people?where[search_name_or_email]=' + enc + '&per_page=25&include=emails'
        data, err = pco_get(path)
        if err:
            print json.dumps({'success': False, 'message': err})
            return
        email_by_pid = {}
        for inc in (data.get('included') or []):
            if inc.get('type') != 'Email':
                continue
            a = inc.get('attributes') or {}
            addr = safe_str(a.get('address', '')).strip()
            if not addr:
                continue
            rels = inc.get('relationships') or {}
            person_rel = ((rels.get('person') or {}).get('data') or {})
            pid = safe_str(person_rel.get('id', ''))
            if not pid:
                continue
            if bool(a.get('primary', False)) or pid not in email_by_pid:
                email_by_pid[pid] = addr
        out = []
        for row in (data.get('data') or []):
            pid = safe_str(row.get('id', ''))
            attrs = row.get('attributes') or {}
            bdate_raw = safe_str(attrs.get('birthdate', ''))
            out.append({
                'pcoPersonId': pid,
                'name': (safe_str(attrs.get('first_name', '')) + ' ' + safe_str(attrs.get('last_name', ''))).strip() or safe_str(attrs.get('name', '')),
                'firstName': safe_str(attrs.get('first_name', '')),
                'lastName': safe_str(attrs.get('last_name', '')),
                'email': email_by_pid.get(pid, ''),
                'birthdate': bdate_raw[:10] if bdate_raw else '',
                'status': safe_str(attrs.get('status', '')),
            })
        print json.dumps({'success': True, 'people': out})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'PCO search failed: ' + str(e)})

def handle_search_tp_people():
    """Search TP People by name fragment or email for manual matching.

    TouchPoint stores names in two columns:
      Name    = "First Last"   (used for display)
      Name2   = "Last, First"  (used for sorting + queries)
    A search like "Ben Baxley" only matches Name (since Name2 starts with
    "Baxley,"). Need to LIKE both. Email/PeopleId stay exact-match.
    For multi-word terms ("Ben Bax"), also AND-match per word against
    Name so "Bax Ben" or "Ben Bax" both work."""
    try:
        term = safe_str(get_data('search_term', '')).strip()
        if not term or len(term) < 2:
            print json.dumps({'success': True, 'people': []})
            return
        safe_term = term.replace("'", "''")
        # Build a per-word AND clause against Name -- handles either name order.
        words = [w for w in re.split(r'\s+', term) if w]
        word_clauses = []
        for w in words:
            sw = w.replace("'", "''")
            word_clauses.append("p.Name LIKE '%%%s%%'" % sw)
        per_word_clause = ' AND '.join(word_clauses) if word_clauses else "1 = 1"
        sql = ("""
            SELECT TOP 25 p.PeopleId, p.Name, p.Name2,
                   ISNULL(p.EmailAddress, '') AS Email,
                   p.Age, ISNULL(g.Description, '') AS Gender,
                   ISNULL(pe.PCO_PersonId, '') AS PCO_PersonId
            FROM People p WITH (NOLOCK)
            LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
            LEFT JOIN (
                SELECT PeopleId,
                       COALESCE(NULLIF(Data, ''), StrValue) AS PCO_PersonId
                FROM PeopleExtra WITH (NOLOCK)
                WHERE Field = '%s'
            ) pe ON pe.PeopleId = p.PeopleId
            WHERE (
                    p.Name LIKE '%%%s%%'
                 OR p.Name2 LIKE '%%%s%%'
                 OR (%s)
                 OR p.EmailAddress = '%s'
                 OR p.EmailAddress2 = '%s'
                 OR CAST(p.PeopleId AS VARCHAR) = '%s'
                  )
              AND p.IsDeceased = 0 AND p.ArchivedFlag = 0
            ORDER BY p.Name2
        """ % (PCO_PERSON_ID_FIELD, safe_term, safe_term, per_word_clause, safe_term, safe_term, safe_term))
        out = []
        try:
            for r in q.QuerySql(sql):
                # Return the "First Last" Name for display -- matches what
                # the user typed and reads naturally in the picker.
                disp = safe_str(r.Name) or safe_str(r.Name2)
                out.append({
                    'peopleId': int(r.PeopleId),
                    'name': disp,
                    'email': safe_str(r.Email),
                    'age': safe_int(r.Age, 0) if r.Age is not None else None,
                    'gender': safe_str(r.Gender),
                    'pcoPersonId': safe_str(getattr(r, 'PCO_PersonId', '')),
                })
        except:
            pass
        print json.dumps({'success': True, 'people': out})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Search failed: ' + str(e)})

def handle_confirm_person_mapping():
    """Write the PCO_PersonId extra value on the TP person so they auto-match
    on every future plan. The mapping is per-person and survives forever."""
    try:
        pco_person_id = safe_str(get_data('pco_person_id', '')).strip()
        tp_people_id = safe_int(get_data('tp_people_id', 0), 0)
        pco_name = safe_str(get_data('pco_name', '')).strip()
        if not pco_person_id or tp_people_id <= 0:
            print json.dumps({'success': False, 'message': 'Missing pco_person_id or tp_people_id'})
            return
        # Verify the TP person exists and is reachable.
        tp_name = ''
        try:
            for r in q.QuerySql('SELECT TOP 1 PeopleId, Name2 FROM People WHERE PeopleId = ' + str(tp_people_id)):
                tp_name = safe_str(r.Name2)
                break
        except:
            pass
        if not tp_name:
            print json.dumps({'success': False, 'message': 'TP person #' + str(tp_people_id) + ' not found.'})
            return
        # Write the extra value. AddExtraValueText -> Person.AddEditExtraData,
        # which writes to PeopleExtra.Data and overwrites the existing row.
        try:
            model.AddExtraValueText(tp_people_id, PCO_PERSON_ID_FIELD, pco_person_id)
        except Exception as we:
            print json.dumps({'success': False, 'message': 'Extra value write failed: ' + str(we)})
            return
        # Verify the write actually persisted to the column we read from.
        # First version of this script wrote successfully but read from the
        # wrong column (StrValue vs Data) -- this catches that kind of
        # storage-column drift in the future.
        verify_ok = False
        try:
            verify_sql = ("""
                SELECT TOP 1 COALESCE(NULLIF(Data, ''), StrValue) AS V
                FROM PeopleExtra WITH (NOLOCK)
                WHERE PeopleId = %s AND Field = '%s'
            """ % (int(tp_people_id), PCO_PERSON_ID_FIELD))
            for r in q.QuerySql(verify_sql):
                if safe_str(r.V) == pco_person_id:
                    verify_ok = True
                break
        except:
            pass
        if not verify_ok:
            # Don't fail the user-visible operation -- the audit log will
            # capture it and the next page refresh will tell us if the
            # mapping really didn't stick. But include the warning so the
            # UI can surface it.
            append_audit({
                'action': 'confirm_person_mapping_verify_failed',
                'pcoPersonId': pco_person_id,
                'pcoName': pco_name,
                'tpPeopleId': tp_people_id,
                'tpName': tp_name,
                'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
            })
            print json.dumps({
                'success': False,
                'message': "Saved but couldn't verify the write came back -- check that the PCO_PersonId extra value field exists and that the role has write access. The audit log captured the attempt."
            })
            return
        append_audit({
            'action': 'confirm_person_mapping',
            'pcoPersonId': pco_person_id,
            'pcoName': pco_name,
            'tpPeopleId': tp_people_id,
            'tpName': tp_name,
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })
        print json.dumps({'success': True, 'message': 'Mapped to ' + tp_name + '.', 'tpName': tp_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Confirm mapping failed: ' + str(e)})

def _verify_attend_flag(meeting_id, people_id, expected_attended):
    """Read the Attend row back after a write to confirm AttendanceFlag
    matches what we set. Same pattern as TPxi_AttendanceMarkings.
    Returns (matched, actual_int_or_None)."""
    try:
        sql = "SELECT AttendanceFlag FROM Attend WHERE MeetingId = %s AND PeopleId = %s" % (
            int(meeting_id), int(people_id))
        for r in q.QuerySql(sql):
            f = r.AttendanceFlag
            if f is None:
                return False, None
            actual = bool(f)
            return actual == bool(expected_attended), 1 if actual else 0
    except:
        pass
    return False, None

def handle_sync_plan_attendance():
    """Write Attend rows for the matched attendees. Resolves the TP meeting
    via GetMeetingIdByDateTime using the plan's date + the org's first
    scheduled time of day (falls back to 9:00 AM if no schedule)."""
    try:
        plan_id = safe_str(get_data('plan_id', '')).strip()
        tp_org_id = safe_int(get_data('tp_org_id', 0), 0)
        plan_date_iso = safe_str(get_data('plan_date_iso', '')).strip()
        people_ids_csv = safe_str(get_data('people_ids', '')).strip()
        pco_service_type_id = safe_str(get_data('service_type_id', '')).strip()
        if not plan_id or tp_org_id <= 0 or not plan_date_iso or not people_ids_csv:
            print json.dumps({'success': False, 'message': 'Missing required args'})
            return
        try:
            plan_dt = datetime.datetime.strptime(plan_date_iso[:10], '%Y-%m-%d')
        except:
            print json.dumps({'success': False, 'message': 'Invalid plan_date_iso (expected YYYY-MM-DD)'})
            return

        # v3.0+: read from the unified Service Type Mapping. Plan cards
        # only render on the Dashboard when perPlanAttendance is on, so
        # syncAttendance is true by definition here; autoAddMember
        # follows the mapping's own toggle.
        sync_attendance = True
        auto_add_member = True
        if pco_service_type_id:
            mappings_now = load_json(PEOPLE_MAPPINGS_KEY, {})
            if isinstance(mappings_now, dict) and pco_service_type_id in mappings_now:
                _info = _parse_people_mapping(mappings_now[pco_service_type_id])
                # Honor mapping config; default to ON if missing.
                sync_attendance = bool(_info['perPlanAttendance'])
                auto_add_member = bool(_info['autoAddMember'])
        if not sync_attendance and not auto_add_member:
            print json.dumps({
                'success': False,
                'message': 'This service type mapping has neither attendance nor auto-add enabled. Nothing to do.'
            })
            return
        # Pick the meeting time. Try OrgSchedule for the plan's weekday, fall
        # back to 9:00 AM.
        sched_day = (plan_dt.weekday() + 1) % 7  # Python Mon=0..Sun=6 -> TP Sun=0..Sat=6
        sched_hour, sched_minute = 9, 0
        try:
            sched_sql = ("""
                SELECT TOP 1 SchedTime FROM OrgSchedule
                WHERE OrganizationId = %s AND SchedDay = %s
                ORDER BY SchedTime
            """ % (tp_org_id, sched_day))
            for r in q.QuerySql(sched_sql):
                st = r.SchedTime
                if st is not None:
                    try:
                        sched_hour = int(st.hour)
                        sched_minute = int(st.minute)
                    except:
                        try:
                            sched_hour = int(st.Hour)
                            sched_minute = int(st.Minute)
                        except:
                            pass
                break
        except:
            pass
        meeting_dt = datetime.datetime(plan_dt.year, plan_dt.month, plan_dt.day,
                                       sched_hour, sched_minute, 0)
        # Only create the TouchPoint meeting if we actually plan to write
        # attendance to it. With sync_attendance off, this mapping is
        # auto-add-member only and we never touch the meeting.
        meeting_id = 0
        if sync_attendance:
            try:
                meeting_id = model.GetMeetingIdByDateTime(int(tp_org_id), meeting_dt, True)
            except Exception as me:
                print json.dumps({'success': False, 'message': 'Could not resolve/create TP meeting: ' + str(me)})
                return
            if not meeting_id:
                print json.dumps({'success': False, 'message': 'GetMeetingIdByDateTime returned no id'})
                return

        # Parse the people ids csv.
        pids = []
        for tok in people_ids_csv.split(','):
            pid = safe_int(tok.strip(), 0)
            if pid > 0:
                pids.append(pid)
        if not pids:
            print json.dumps({'success': False, 'message': 'No people ids to sync'})
            return

        # Pre-resolve current OM membership for the org so we only call
        # JoinOrg when needed. CRITICAL: filter to active rows
        # (InactiveDate IS NULL). If someone was previously a member and
        # got dropped, their OM row sticks around with InactiveDate set --
        # if we skip JoinOrg for those, they get the attendance row but
        # remain Inactive and never re-appear on the involvement's roster.
        existing_member_ids = set()
        try:
            mem_sql = ("""
                SELECT PeopleId FROM OrganizationMembers WITH (NOLOCK)
                WHERE OrganizationId = %s
                  AND PeopleId IN (%s)
                  AND InactiveDate IS NULL
            """ % (int(tp_org_id), ','.join(str(p) for p in pids)))
            for r in q.QuerySql(mem_sql):
                existing_member_ids.add(int(r.PeopleId))
        except:
            pass

        # Person data sync hook: rules + PCO data per person come from JS.
        # If the admin hasn't opted any field in, person_rules_active is
        # False and we skip the comparison work entirely.
        person_rules = normalize_person_sync_rules(load_json(PERSON_SYNC_RULES_KEY, {}))
        person_rules_active = person_sync_rules_active(person_rules)
        pco_data_by_pid = _parse_people_json_payload(safe_str(get_data('people_json', '')))
        person_auto = 0
        person_queued = 0

        applied = 0
        skipped = 0
        joined = 0
        verify_failures = []
        per_person = []
        for pid in pids:
            try:
                # JoinOrg first if auto-add is on and they aren't already an
                # active member. Without this the EditPersonAttendance write
                # still creates an Attend row but the person is recorded as
                # a Visitor, which excludes them from member-attendance
                # reports. Auto-add OFF means the worship admin wants to
                # gate membership manually (audition-only choirs, etc.).
                if auto_add_member and pid not in existing_member_ids:
                    try:
                        person = model.GetPerson(pid)
                        if person and person.PeopleId:
                            model.JoinOrg(int(tp_org_id), person)
                            joined += 1
                    except Exception as je:
                        msg = str(je).lower()
                        # JoinOrg can throw "already a member" if we lost a
                        # race vs another writer -- treat as already-joined.
                        if 'already' not in msg:
                            per_person.append({'peopleId': pid, 'status': 'join_failed', 'message': str(je)})
                            skipped += 1
                            continue
                # Mark attendance: gated by syncAttendance AND by the
                # person's Confirmed status from the JS payload. The JS now
                # includes Unconfirmed/Declined team members (for the
                # auto_add_member side) but we only count Confirmed ones
                # as actually-attended.
                person_info = pco_data_by_pid.get(pid) or {}
                # If JS didn't include this person in people_json (legacy
                # caller), assume Confirmed -- the old behavior only sent
                # Confirmed pids in the first place.
                is_conf = bool(person_info.get('isConfirmed', True)) if pid in pco_data_by_pid else True
                if sync_attendance and is_conf:
                    model.EditPersonAttendance(int(meeting_id), pid, True)
                    ok, actual = _verify_attend_flag(meeting_id, pid, True)
                    if not ok:
                        verify_failures.append(pid)
                        per_person.append({'peopleId': pid, 'status': 'write_unverified', 'actualFlag': actual})
                    else:
                        per_person.append({'peopleId': pid, 'status': 'present'})
                    applied += 1
                elif sync_attendance and not is_conf:
                    # Unconfirmed/Declined got joined as a member but no
                    # Attend row -- they're on the team, just not present.
                    per_person.append({'peopleId': pid, 'status': 'member_only_not_confirmed'})
                else:
                    # Auto-add-only mode: record what we did per-person but
                    # don't bump 'applied' since nothing got marked Present.
                    per_person.append({'peopleId': pid, 'status': 'joined_only' if pid not in existing_member_ids else 'already_member'})
                # Person data sync runs AFTER attendance/membership writes
                # so a failed sync on, say, attendance doesn't also block
                # the person-data comparison. Failures here are swallowed
                # silently -- worst case the row gets re-queued next sync.
                if person_rules_active and pid in pco_data_by_pid:
                    try:
                        c = apply_person_sync_for_one(pid, pco_data_by_pid[pid], person_rules)
                        person_auto += c.get('auto', 0)
                        person_queued += c.get('queued', 0)
                    except:
                        pass
            except Exception as we:
                skipped += 1
                per_person.append({'peopleId': pid, 'status': 'error', 'message': str(we)})
        append_audit({
            'action': 'sync_plan_attendance',
            'planId': plan_id,
            'tpOrgId': tp_org_id,
            'meetingId': int(meeting_id) if meeting_id else 0,
            'meetingDateIso': plan_date_iso,
            'syncAttendance': sync_attendance,
            'autoAddMember': auto_add_member,
            'applied': applied,
            'skipped': skipped,
            'joinedOrg': joined,
            'verifyFailures': len(verify_failures),
            'by': safe_str(model.UserName) if hasattr(model, 'UserName') else '',
        })

        # Update last-sync timestamp.
        s = load_json(SETTINGS_KEY, {})
        if not isinstance(s, dict):
            s = {}
        s['lastSyncAt'] = now_iso()
        save_json(SETTINGS_KEY, s)

        if sync_attendance:
            msg = 'Synced ' + str(applied) + ' attendee(s) to TouchPoint meeting #' + str(meeting_id) + '.'
            if joined > 0:
                msg += ' Added ' + str(joined) + ' as member(s) of the involvement.'
        else:
            # auto-add-only mode
            msg = 'Added ' + str(joined) + ' member(s) to the involvement. ' + \
                  '(Attendance sync is disabled for this service type.)'
        if person_auto:
            msg += ' Updated ' + str(person_auto) + ' person field(s).'
        if person_queued:
            msg += ' Queued ' + str(person_queued) + ' person change(s) for review.'
        resp = {
            'success': True,
            'message': msg,
            'personAuto': person_auto,
            'personQueued': person_queued,
            'meetingId': int(meeting_id) if meeting_id else 0,
            'meetingDateTime': meeting_dt.strftime('%Y-%m-%dT%H:%M:%S'),
            'syncAttendance': sync_attendance,
            'autoAddMember': auto_add_member,
            'applied': applied,
            'skipped': skipped,
            'joinedOrg': joined,
            'perPerson': per_person,
        }
        if verify_failures:
            resp['verifyWarning'] = str(len(verify_failures)) + ' write(s) did not read back as Present.'
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Sync failed: ' + str(e)})

# =====================================================================
# Unmatched People tab
# =====================================================================
# Bulk triage view. Scans recent plans across mapped service types,
# collects every PCO person that doesn't have a PCO_PersonId match on
# any TouchPoint person yet, suggests email matches, and lets admins
# confirm in one place instead of going plan-by-plan.

def handle_list_unmatched_people():
    """Return a deduplicated list of PCO people seen in recent plans who
    don't have a PCO_PersonId match on any TP person."""
    try:
        mappings = load_json(ORG_MAPPINGS_KEY, {})
        if not isinstance(mappings, dict) or not mappings:
            print json.dumps({
                'success': True, 'people': [],
                'message': 'No service types are mapped. Map service types first so we know which plans to scan.'
            })
            return
        days_back = safe_int(get_data('days_back', 14), 14)
        days_forward = safe_int(get_data('days_forward', 7), 7)

        # 1. Walk plans across mapped service types -> list of plans.
        all_plans = []
        api_errors = []
        for pco_st_id in mappings.keys():
            plans, err = _pco_plans_for_service_type(pco_st_id, days_back, days_forward)
            if err:
                api_errors.append('Service type ' + str(pco_st_id) + ': ' + err)
                continue
            for p in plans:
                p['serviceTypeId'] = pco_st_id
                all_plans.append(p)

        # 2. For each plan, pull team_members and collect PCO Person IDs.
        # Aggregate: {pcoPersonId: {name, email, plans: [{title, dateIso}]}}
        agg = {}
        for plan in all_plans:
            tm_path = ('/services/v2/service_types/' + plan['serviceTypeId'] +
                       '/plans/' + plan['planId'] + '/team_members?per_page=100')
            tm_data, err = pco_get(tm_path)
            if err:
                api_errors.append('Plan ' + plan['planId'] + ': ' + err)
                continue
            for tm in (tm_data.get('data') or []):
                rels = tm.get('relationships') or {}
                person_rel = ((rels.get('person') or {}).get('data') or {})
                ppid = safe_str(person_rel.get('id', ''))
                if not ppid:
                    continue
                tm_attrs = tm.get('attributes') or {}
                name = safe_str(tm_attrs.get('name', '') or tm_attrs.get('person_name', '') or '(Unknown)')
                email = safe_str(tm_attrs.get('email', '') or tm_attrs.get('email_address', '')).strip().lower()
                if ppid not in agg:
                    agg[ppid] = {
                        'pcoPersonId': ppid,
                        'name': name,
                        'email': email,
                        'plans': [],
                    }
                # Prefer non-empty email if we get it from a later plan.
                if email and not agg[ppid]['email']:
                    agg[ppid]['email'] = email
                # _pco_plans_for_service_type emits planTitle/planDateIso
                # (not title/dateIso). Use .get so a missing key is silent
                # instead of breaking the whole tab.
                agg[ppid]['plans'].append({
                    'title': plan.get('planTitle') or '',
                    'dateIso': plan.get('planDateIso') or '',
                })

        # 3. Drop anyone already matched via PCO_PersonId.
        all_pco_ids = list(agg.keys())
        already_matched = _tp_match_by_pco_id(all_pco_ids)
        unmatched_ids = [pid for pid in all_pco_ids if pid not in already_matched]

        # 4. For the remainder, batch-resolve email matches as suggestions.
        emails_to_check = [agg[pid]['email'] for pid in unmatched_ids if agg[pid]['email']]
        by_email = _tp_match_by_email(emails_to_check)

        out = []
        for pid in unmatched_ids:
            row = agg[pid]
            row['plansSeen'] = len(row['plans'])
            row['suggestion'] = None
            row['suggestionAmbiguous'] = False
            if row['email'] and row['email'] in by_email:
                hits = by_email[row['email']]
                if len(hits) == 1:
                    row['suggestion'] = {
                        'tpPeopleId': hits[0]['peopleId'],
                        'tpName': hits[0]['name'],
                        'matchSource': 'email',
                    }
                else:
                    row['suggestionAmbiguous'] = True
            out.append(row)

        # Sort: rows with a confident suggestion first, then by name.
        out.sort(key=lambda r: (0 if r.get('suggestion') else 1, (r.get('name') or '').lower()))

        resp = {'success': True, 'people': out, 'unmatchedCount': len(out), 'plansScanned': len(all_plans)}
        if api_errors:
            resp['warnings'] = api_errors
        print json.dumps(resp)
    except Exception as e:
        print json.dumps({'success': False, 'message': 'List unmatched failed: ' + str(e)})

# Placeholder dispatch for tabs not yet built. Each prints a clear "not
# implemented yet" so the JS callers don't crash silently.

def handle_not_implemented(action_name):
    print json.dumps({
        'success': False,
        'message': 'Action "' + action_name + '" is not implemented yet (v' + APP_VERSION + ').'
    })

# =====================================================================
# DISPATCH
# =====================================================================

# v3.3: scheduler hook -- runs when invoked via ?scheduler=1 (TouchPoint
# Scheduled Task / cron URL). Prints JSON and skips the SPA render.
_is_scheduler_invocation = False
try:
    if hasattr(model.Data, 'scheduler') and safe_str(model.Data.scheduler) in ('1', 'true', 'yes'):
        _is_scheduler_invocation = True
    else:
        _u_check = safe_str(getattr(model, 'URL', '') or '')
        if 'scheduler=1' in _u_check:
            _is_scheduler_invocation = True
except:
    pass

if _is_scheduler_invocation:
    handle_run_scheduled_syncs()
elif model.HttpMethod == 'post':
    action = str(get_data('pco_action', ''))
    # v3.0 storage migration -- one-shot, idempotent. Runs once per
    # POST entry so we don't slow GET pageloads.
    _v3_migrate_org_to_people_mappings()
    if action == 'test_connection':
        handle_test_connection()
    elif action == 'load_settings':
        handle_load_settings()
    elif action == 'save_settings':
        handle_save_settings()
    # Service Type Mappings tab
    elif action == 'list_service_types':
        handle_list_service_types()
    # v3.0: Service Plan Mappings dropped -- save_org_mapping /
    # load_org_mappings / set_mapping_options no longer dispatched.
    # Handlers remain in source but are unreachable.
    elif action == 'search_involvements':
        handle_search_involvements()
    # Team Mappings (v2.0+)
    elif action == 'list_pco_teams_for_service_type':
        handle_list_pco_teams_for_service_type()
    elif action == 'list_team_mappings':
        handle_list_team_mappings()
    elif action == 'save_team_mapping':
        handle_save_team_mapping()
    elif action == 'delete_team_mapping':
        handle_delete_team_mapping()
    elif action == 'set_team_mapping_options':
        handle_set_team_mapping_options()
    elif action == 'load_team_sync_preview':
        handle_load_team_sync_preview()
    elif action == 'sync_team':
        handle_sync_team()
    # v3.1.1: inline "Check PCO positions" diagnostic
    elif action == 'check_pco_team_positions':
        handle_check_pco_team_positions()
    # v3.3: scheduler
    elif action == 'resolve_username':
        handle_resolve_username()
    elif action == 'search_tp_users':
        handle_search_tp_users()
    elif action == 'set_schedule_options':
        handle_set_schedule_options()
    elif action == 'run_scheduled_syncs':
        handle_run_scheduled_syncs()
    # v3.3.1: scheduler install (auto-add to ScheduledTasks)
    elif action == 'check_scheduler_install':
        handle_check_scheduler_install()
    elif action == 'install_scheduler':
        handle_install_scheduler()
    elif action == 'uninstall_scheduler':
        handle_uninstall_scheduler()
    # All People Mapping (v2.2+, singleton)
    elif action == 'load_all_people_mapping':
        handle_load_all_people_mapping()
    elif action == 'save_all_people_mapping':
        handle_save_all_people_mapping()
    elif action == 'delete_all_people_mapping':
        handle_delete_all_people_mapping()
    elif action == 'load_all_people_preview':
        handle_load_all_people_preview()
    elif action == 'sync_all_people':
        handle_sync_all_people()
    # All People proposed-match review queue (v2.5+)
    elif action == 'load_all_people_proposed_matches':
        handle_load_all_people_proposed_matches()
    elif action == 'apply_proposed_match':
        handle_apply_proposed_match()
    elif action == 'bulk_apply_proposed_matches':
        handle_bulk_apply_proposed_matches()
    elif action == 'skip_pco_person_forever':
        handle_skip_pco_person_forever()
    # Sync Mappings health roll-up (v2.3+)
    elif action == 'load_sync_mappings_health':
        handle_load_sync_mappings_health()
    # People Mappings (v2.1+)
    elif action == 'list_people_mappings':
        handle_list_people_mappings()
    elif action == 'save_people_mapping':
        handle_save_people_mapping()
    elif action == 'delete_people_mapping':
        handle_delete_people_mapping()
    elif action == 'set_people_mapping_options':
        handle_set_people_mapping_options()
    elif action == 'load_people_sync_preview':
        handle_load_people_sync_preview()
    elif action == 'sync_people':
        handle_sync_people()
    # Sync Dashboard tab
    elif action == 'list_recent_plans':
        handle_list_recent_plans()
    # Per-plan review + sync write
    elif action == 'load_plan_preview':
        handle_load_plan_preview()
    elif action == 'search_tp_people':
        handle_search_tp_people()
    elif action == 'confirm_person_mapping':
        handle_confirm_person_mapping()
    # v3.1: verify/repair an existing TP <-> PCO link
    elif action == 'load_verify_details':
        handle_load_verify_details()
    elif action == 'unlink_tp_person':
        handle_unlink_tp_person()
    elif action == 'search_pco_people_by_name':
        handle_search_pco_people_by_name()
    elif action == 'sync_plan_attendance':
        handle_sync_plan_attendance()
    # v3.0: load_roster_preview / sync_roster / list_unmatched_people
    # no longer dispatched. The add-only "rollup" workflow was folded
    # into Service Type Mappings, and the legacy Unmatched section is
    # superseded by People Matching's Proposed Matches. Handlers remain
    # in source but are unreachable.
    # Person data sync (v1: PCO -> TP only)
    elif action == 'load_person_sync_rules':
        handle_load_person_sync_rules()
    elif action == 'save_person_sync_rules':
        handle_save_person_sync_rules()
    elif action == 'list_pending_person_changes':
        handle_list_pending_person_changes()
    elif action == 'apply_person_change':
        handle_apply_person_change()
    elif action == 'skip_person_change':
        handle_skip_person_change()
    # Stubs for upcoming features.
    elif action in (
        'load_audit_log',
    ):
        handle_not_implemented(action)
    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + safe_str(action)})

else:
    # =================================================================
    # GET: render the SPA
    # =================================================================
    # /PyScript/ -> /PyScriptForm/ redirect, so the Run button in the
    # Special Content editor lands on the right route.
    _url = ''
    try:
        _url = str(getattr(model, 'URL', '') or '')
    except:
        _url = ''
    _needs_redirect = ('/PyScript/' in _url) and ('/PyScriptForm/' not in _url)

    if _needs_redirect:
        _name = DC_SCRIPT_ID
        _qs = ''
        try:
            _m = re.search(r'/PyScript/([^/?#&]+)', _url)
            if _m:
                _name = _m.group(1)
            _qi = _url.find('?')
            if _qi >= 0:
                _qs = _url[_qi:]
        except:
            pass
        _target = '/PyScriptForm/' + _name + _qs
        print '<!DOCTYPE html><html><head><title>Loading...</title>'
        print '<meta http-equiv="refresh" content="0;url=' + _target + '">'
        print '<script>window.location.replace(' + json.dumps(_target) + ');</script>'
        print '</head><body style="font-family:Segoe UI,sans-serif;padding:30px;">'
        print 'Loading <a href="' + _target + '">PCO Sync</a>...'
        print '</body></html>'

    css = """
    <style>
    .pco-root { font-family: 'Segoe UI', Arial, sans-serif; color: #222; max-width: 1300px; margin: 0 auto; padding: 12px; }
    .pco-root * { box-sizing: border-box; }
    .pco-h1 { font-size: 22px; font-weight: 700; margin: 0 0 4px 0; color: #1f4e79; }
    .pco-sub { font-size: 13px; color: #666; margin-bottom: 16px; }
    .pco-version { font-size: 12px; color: #888; font-weight: 400; margin-left: 8px; }

    .pco-tabs { display: flex; gap: 4px; border-bottom: 2px solid #e1e4e8; margin-bottom: 16px; flex-wrap: wrap; }
    .pco-tab { padding: 10px 16px; cursor: pointer; border: none; background: transparent; font-size: 14px; font-weight: 600; color: #555; border-bottom: 3px solid transparent; margin-bottom: -2px; transition: all 0.15s; }
    .pco-tab:hover { color: #1f4e79; }
    .pco-tab.active { color: #1f4e79; border-bottom-color: #1f4e79; }

    .pco-card { background: #fff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 16px; margin-bottom: 12px; }
    .pco-card h3 { margin: 0 0 10px 0; font-size: 16px; color: #1f4e79; }
    .pco-help { font-size: 12px; color: #777; margin: 6px 0; line-height: 1.5; }

    .pco-form-row { display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap; margin-bottom: 10px; }
    .pco-form-row > div { flex: 1 1 220px; }
    .pco-label { display: block; font-size: 12px; color: #555; font-weight: 600; margin-bottom: 4px; text-transform: uppercase; letter-spacing: 0.5px; }
    .pco-input { width: 100%; padding: 8px 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; background: #fff; }
    .pco-input:focus { outline: none; border-color: #1f4e79; box-shadow: 0 0 0 2px rgba(31,78,121,0.15); }

    .pco-btn { display: inline-block; padding: 8px 14px; font-size: 14px; font-weight: 600; border: 1px solid #1f4e79; background: #1f4e79; color: #fff; border-radius: 4px; cursor: pointer; text-decoration: none; }
    .pco-btn:hover { background: #2a5e8e; border-color: #2a5e8e; }
    .pco-btn.pco-secondary { background: #fff; color: #1f4e79; }
    .pco-btn.pco-secondary:hover { background: #f0f4f8; }
    .pco-btn:disabled { opacity: 0.5; cursor: not-allowed; }

    .pco-pill { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
    .pco-pill.pco-ok { background: #d4f0db; color: #1f6b3a; }
    .pco-pill.pco-warn { background: #fce7c2; color: #7a4a00; }
    .pco-pill.pco-err { background: #f9d6d6; color: #8a2020; }

    .pco-toast { position: fixed; bottom: 20px; right: 20px; padding: 12px 20px; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); font-size: 14px; z-index: 1000; opacity: 0; transform: translateY(10px); transition: all 0.2s; }
    .pco-toast.pco-show { opacity: 1; transform: translateY(0); }
    .pco-toast.pco-toast-ok { background: #d4f0db; color: #1f6b3a; border: 1px solid #b8e0c4; }
    .pco-toast.pco-toast-err { background: #f9d6d6; color: #8a2020; border: 1px solid #ecbcbc; }
    .pco-toast.pco-toast-info { background: #d6e6f5; color: #1f4e79; border: 1px solid #b8d4ed; }

    .pco-empty { text-align: center; color: #888; padding: 40px 20px; font-size: 14px; }
    .pco-muted { color: #888; font-size: 12px; }

    /* Loading spinner -- used by the Proposed Matches walk + similar
       multi-second waits where users otherwise think the page froze. */
    .pco-spinner {
      display: inline-block;
      width: 36px;
      height: 36px;
      border: 4px solid #e1e4e8;
      border-top-color: #0f7c84;
      border-radius: 50%;
      animation: pco-spin 0.9s linear infinite;
    }
    @keyframes pco-spin { to { transform: rotate(360deg); } }
    .pco-loading-block { text-align: center; padding: 36px 20px; color: #444; }
    .pco-loading-block .pco-loading-title { margin-top: 14px; font-weight: 600; font-size: 14px; }
    .pco-loading-block .pco-loading-elapsed { margin-top: 4px; font-size: 12px; color: #777; }
    </style>
    """

    body = """
    <div class="pco-root">
      <div class="pco-h1">PCO Sync <span class="pco-version">v__APP_VERSION__</span></div>
      <div class="pco-sub">One-way sync from Planning Center Online into TouchPoint: people, rosters, teams, and per-plan attendance.</div>

      <div class="pco-tabs" id="pcoTabs">
        <button class="pco-tab active" data-tab="sync">Sync Dashboard</button>
        <button class="pco-tab" data-tab="mappings">Sync Mappings</button>
        <button class="pco-tab" data-tab="people">People Matching</button>
        <button class="pco-tab" data-tab="settings">Settings</button>
      </div>

      <div id="pcoContent"></div>
      <div id="pcoToastHost"></div>
    </div>
    """

    js = r"""
    <script>
    (function(){
      'use strict';

      var APP_VERSION = '__APP_VERSION__';
      var state = {
        tab: 'sync',
        settings: null,
      };

      function $(id) { return document.getElementById(id); }

      function ajax(action, params, cb) {
        var data = 'pco_action=' + encodeURIComponent(action);
        if (params) {
          for (var k in params) {
            if (params.hasOwnProperty(k)) {
              data += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
            }
          }
        }
        var xhr = new XMLHttpRequest();
        xhr.open('POST', window.location.pathname, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function(){
          if (xhr.readyState !== 4) return;
          if (xhr.status !== 200) { cb('HTTP ' + xhr.status, null); return; }
          var txt = xhr.responseText || '';
          try {
            // Defensive: TouchPoint sometimes prepends whitespace/HTML
            var s = txt.indexOf('{');
            var e = txt.lastIndexOf('}');
            if (s >= 0 && e > s) txt = txt.substring(s, e + 1);
            cb(null, JSON.parse(txt));
          } catch(ex) {
            cb('Bad response: ' + ex.message + ' | raw: ' + xhr.responseText.substring(0, 200), null);
          }
        };
        xhr.send(data);
      }

      function toast(msg, kind) {
        kind = kind || 'info';
        var host = $('pcoToastHost');
        var el = document.createElement('div');
        el.className = 'pco-toast pco-toast-' + kind;
        el.textContent = msg;
        host.appendChild(el);
        setTimeout(function(){ el.classList.add('pco-show'); }, 10);
        setTimeout(function(){
          el.classList.remove('pco-show');
          setTimeout(function(){ host.removeChild(el); }, 250);
        }, 3500);
      }

      // ---- Tabs --------------------------------------------------

      function selectTab(name) {
        state.tab = name;
        var tabs = document.querySelectorAll('.pco-tab');
        for (var i = 0; i < tabs.length; i++) {
          if (tabs[i].getAttribute('data-tab') === name) tabs[i].classList.add('active');
          else tabs[i].classList.remove('active');
        }
        renderTab();
      }

      function renderTab() {
        if (state.tab === 'sync') return renderSyncTab();
        if (state.tab === 'mappings') return renderMappingsTab();
        if (state.tab === 'people') return renderPeopleSyncTab();
        if (state.tab === 'settings') return renderSettingsTab();
        // Back-compat for any stale state pointing at the removed tabs.
        if (state.tab === 'unmatched' || state.tab === 'reviews') {
          state.tab = 'people';
          return renderPeopleSyncTab();
        }
      }

      // ---- Settings tab (the only working tab in v1.0.0) ---------

      function renderSettingsTab() {
        var host = $('pcoContent');
        host.innerHTML = ''
          + '<div class="pco-card">'
          + '<h3>Planning Center Personal Access Token</h3>'
          + '<div class="pco-help">'
          + 'Generate a PAT in PCO under <strong>My Account &rarr; Applications &rarr; Personal Access Tokens</strong>. '
          + 'Paste the App ID + Secret here. They are stored in TouchPoint Special Content and are only used server-side '
          + 'to call the PCO API on your behalf.'
          + '</div>'
          + '<div id="pcoSettingsStatus" class="pco-help"></div>'
          + '<div class="pco-form-row">'
          + '  <div><label class="pco-label">App ID</label><input class="pco-input" id="pcoAppId" type="text" placeholder="Paste PCO App ID..."></div>'
          + '  <div><label class="pco-label">Secret</label><input class="pco-input" id="pcoSecret" type="password" placeholder="Paste PCO Secret..."></div>'
          + '</div>'
          + '<div style="display:flex;gap:8px;margin-top:10px;">'
          + '  <button class="pco-btn" id="pcoSaveBtn">Save</button>'
          + '  <button class="pco-btn pco-secondary" id="pcoTestBtn">Test Connection</button>'
          + '  <button class="pco-btn pco-secondary" id="pcoClearBtn" style="margin-left:auto;color:#a00;border-color:#a00;">Clear Credentials</button>'
          + '</div>'
          + '</div>'
          // Person Data Sync rules
          + '<div class="pco-card" style="margin-top:14px;">'
          + '<h3>Person Data Sync</h3>'
          + '<div class="pco-help">'
          + 'Default: TouchPoint is authoritative -- nothing flows from PCO into TP person records unless you opt in here. '
          + 'For each field, pick the direction and whether the write happens silently or queues for review on the '
          + '<strong>Data Reviews</strong> tab. '
          + '<em>v1 supports PCO &rarr; TP only.</em> Bidirectional (TP &rarr; PCO) is planned.'
          + '</div>'
          + '<div id="pcoPersonRulesGrid" style="margin-top:10px;">Loading rules...</div>'
          + '<div style="display:flex;gap:8px;margin-top:10px;">'
          + '  <button class="pco-btn" id="pcoRulesSaveBtn">Save Rules</button>'
          + '  <span id="pcoRulesStatus" class="pco-muted" style="align-self:center;font-size:13px;"></span>'
          + '</div>'
          + '</div>'
          // v3.3.1: Scheduler install/uninstall (auto-managed
          // ScheduledTasks block, matches ProspectBuilder pattern).
          + '<div class="pco-card" style="margin-top:14px;">'
          + '<h3>Scheduled Sync</h3>'
          + '<div class="pco-help">'
          + 'Each mapping can run on its own day/time schedule with an email summary to a TouchPoint user. '
          + 'Install the global scheduler here once -- it adds a managed block to the <code>ScheduledTasks</code> '
          + 'special content that TouchPoint runs on its standard cron. Per-mapping schedules then go on the '
          + '<strong>Sync Mappings</strong> tab.'
          + '</div>'
          + '<div id="pcoSchedInstallStatus" style="margin-top:10px;padding:10px 12px;background:#f4f7fa;border-radius:4px;font-size:13px;">Checking install status...</div>'
          + '<div style="margin-top:8px;">'
          + '<button class="pco-btn" id="pcoSchedTestBtn" style="font-size:12px;padding:4px 10px;" disabled>Run scheduler now (test)</button>'
          + '<span class="pco-muted" id="pcoSchedTestResult" style="margin-left:10px;font-size:12px;"></span>'
          + '</div>'
          + '</div>';

        $('pcoSaveBtn').onclick = function(){
          var aid = $('pcoAppId').value.trim();
          var sec = $('pcoSecret').value.trim();
          if (!aid && !sec) { toast('Enter App ID and Secret first.', 'err'); return; }
          ajax('save_settings', {pco_app_id: aid, pco_secret: sec}, function(err, d){
            if (err || !d || !d.success) { toast('Save failed: ' + (d && d.message || err), 'err'); return; }
            toast('Saved. You can now Test Connection.', 'ok');
            $('pcoAppId').value = '';
            $('pcoSecret').value = '';
            loadAndRenderSettingsStatus();
          });
        };

        $('pcoTestBtn').onclick = function(){
          var btn = this; btn.disabled = true; btn.textContent = 'Testing...';
          ajax('test_connection', {}, function(err, d){
            btn.disabled = false; btn.textContent = 'Test Connection';
            if (err || !d || !d.success) { toast('Connection failed: ' + ((d && d.message) || err), 'err'); return; }
            toast(d.message || 'Connection OK', 'ok');
          });
        };

        $('pcoClearBtn').onclick = function(){
          if (!confirm('Remove the saved PCO credentials? You will need to paste them again to sync.')) return;
          ajax('save_settings', {clear_credentials: 'true'}, function(err, d){
            if (err || !d || !d.success) { toast('Clear failed.', 'err'); return; }
            toast('Credentials cleared.', 'info');
            loadAndRenderSettingsStatus();
          });
        };

        loadAndRenderSettingsStatus();
        loadAndRenderPersonRules();

        // v3.3.1: scheduler install state + Install / Uninstall buttons.
        function renderSchedInstallStatus() {
          var host = $('pcoSchedInstallStatus');
          if (!host) return;
          host.innerHTML = 'Checking install status...';
          ajax('check_scheduler_install', {}, function(err, d){
            if (err || !d) {
              host.innerHTML = '<span style="color:#a00;">Could not check install status: ' + escHtml(err || 'unknown') + '</span>';
              return;
            }
            if (!d.success) {
              host.innerHTML = '<span style="color:#a00;">' + escHtml(d.message || 'Check failed') + '</span>';
              return;
            }
            var schedBtnLocal = $('pcoSchedTestBtn');
            if (d.installed) {
              host.innerHTML = ''
                + '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">'
                + '  <div>'
                + '    <span class="pco-pill pco-ok">Installed</span> '
                + '    in <code style="background:#fff;padding:1px 5px;border:1px solid #ccc;border-radius:3px;">' + escHtml(d.contentSlot) + '</code>. '
                + '    PCO Sync runs on TouchPoint\'s scheduled task cycle. '
                + '    Add per-mapping day/time on the <strong>Sync Mappings</strong> tab.'
                + '  </div>'
                + '  <button class="pco-btn pco-secondary" id="pcoSchedUninstallBtn" style="font-size:12px;padding:4px 10px;color:#a00;border-color:#a00;">Uninstall</button>'
                + '</div>';
              if (schedBtnLocal) schedBtnLocal.disabled = false;
              var uninst = $('pcoSchedUninstallBtn');
              if (uninst) uninst.onclick = function(){
                if (!confirm('Remove the PCO Sync block from ScheduledTasks?\n\nScheduled syncs will stop firing. Per-mapping schedule settings stay saved -- you can re-install later without losing them.')) return;
                uninst.disabled = true;
                ajax('uninstall_scheduler', {}, function(err, d){
                  uninst.disabled = false;
                  if (err || !d || !d.success) { toast('Uninstall failed: ' + ((d && d.message) || err), 'err'); return; }
                  toast(d.message || 'Uninstalled.', 'ok');
                  renderSchedInstallStatus();
                });
              };
            } else {
              var warnLine = d.referencedOutsideBlock
                ? '<div style="margin-top:6px;color:#8a6d3b;font-size:12px;"><strong>Note:</strong> the script name was found in ScheduledTasks outside our managed block. Hand-edits may conflict -- review before installing.</div>'
                : '';
              host.innerHTML = ''
                + '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">'
                + '  <div>'
                + '    <span class="pco-pill" style="background:#eef0f2;color:#666;">Not installed</span> '
                + '    Click Install to add a managed block to <code style="background:#fff;padding:1px 5px;border:1px solid #ccc;border-radius:3px;">' + escHtml(d.contentSlot) + '</code>. '
                + '    Until you install, per-mapping schedules cannot be enabled.'
                + '  </div>'
                + '  <button class="pco-btn" id="pcoSchedInstallBtn" style="font-size:12px;padding:4px 12px;">Install</button>'
                + '</div>'
                + warnLine;
              if (schedBtnLocal) schedBtnLocal.disabled = true;
              var inst = $('pcoSchedInstallBtn');
              if (inst) inst.onclick = function(){
                inst.disabled = true;
                inst.textContent = 'Installing...';
                ajax('install_scheduler', {}, function(err, d){
                  inst.disabled = false;
                  inst.textContent = 'Install';
                  if (err || !d || !d.success) { toast('Install failed: ' + ((d && d.message) || err), 'err'); return; }
                  toast(d.message || 'Installed.', 'ok');
                  renderSchedInstallStatus();
                });
              };
            }
          });
        }
        renderSchedInstallStatus();

        function runSchedulerTest(force) {
          var schedBtn2 = $('pcoSchedTestBtn');
          var res = $('pcoSchedTestResult');
          if (schedBtn2) schedBtn2.disabled = true;
          if (res) {
            res.innerHTML = (force ? 'Force re-running...' : 'Running...');
            res.style.color = '';
          }
          ajax('run_scheduled_syncs', force ? {force: '1'} : {}, function(err, d){
            if (schedBtn2) schedBtn2.disabled = false;
            if (err || !d || !d.success) {
              if (res) { res.textContent = (d && d.message) || err || 'Run failed.'; res.style.color = '#a00'; }
              return;
            }
            if (!res) return;
            var bits = [];
            if (d.firedCount > 0) {
              bits.push('<span style="color:#1f6b3a;"><strong>Fired ' + d.firedCount + ' mapping(s).</strong></span>');
              if (d.emailsAttempted > 0) {
                if (d.emailsSent === d.emailsAttempted) {
                  bits.push('<span style="color:#1f6b3a;">All ' + d.emailsSent + ' email(s) sent.</span>');
                } else {
                  bits.push('<span style="color:#a00;"><strong>' + (d.emailsAttempted - d.emailsSent) + '/' + d.emailsAttempted + ' email(s) failed:</strong></span>');
                  for (var i = 0; i < (d.emailsFailed || []).length; i++) {
                    var ef = d.emailsFailed[i];
                    bits.push('<span style="color:#a00;">&nbsp;&nbsp;&middot; ' + escHtml(ef.mapping) + ': ' + escHtml(ef.error) + '</span>');
                  }
                }
              } else {
                bits.push('<span class="pco-muted">(No notify username on fired mappings -- no emails sent.)</span>');
              }
            } else {
              bits.push('<span style="color:#8a6d3b;">No mappings fired.</span>');
            }
            if ((d.skippedAlreadyFired || []).length > 0) {
              var sk = d.skippedAlreadyFired;
              bits.push('<span style="color:#8a6d3b;"><strong>Skipped ' + sk.length + ' mapping(s)</strong> that already fired this hour' + (sk[0].lastRunAt ? ' (last: ' + escHtml(sk[0].lastRunAt.replace("T", " ")) + ')' : '') + '. <a href="#" id="pcoSchedForceLink">Force re-run anyway</a>.</span>');
            }
            res.innerHTML = bits.join('<br>');
            var fl = $('pcoSchedForceLink');
            if (fl) fl.onclick = function(ev){ ev.preventDefault(); runSchedulerTest(true); };
          });
        }

        var schedBtn = $('pcoSchedTestBtn');
        if (schedBtn) schedBtn.onclick = function(){ runSchedulerTest(false); };

        $('pcoRulesSaveBtn').onclick = function(){
          var rulesObj = {};
          var grid = $('pcoPersonRulesGrid');
          var rows = grid.querySelectorAll('.pco-rule-row');
          for (var i = 0; i < rows.length; i++) {
            var r = rows[i];
            var key = r.dataset.field;
            rulesObj[key] = {
              direction: r.querySelector('.pco-rule-direction').value,
              mode: r.querySelector('.pco-rule-mode').value,
            };
          }
          var status = $('pcoRulesStatus');
          status.textContent = 'Saving...';
          ajax('save_person_sync_rules', {rules_json: JSON.stringify(rulesObj)}, function(err, d){
            if (err || !d || !d.success) {
              status.textContent = 'Save failed: ' + ((d && d.message) || err);
              status.style.color = '#a00';
              return;
            }
            status.textContent = 'Saved.';
            status.style.color = '#1f6b3a';
            setTimeout(function(){ status.textContent = ''; }, 1500);
          });
        };
      }

      function loadAndRenderPersonRules() {
        ajax('load_person_sync_rules', {}, function(err, d){
          var grid = $('pcoPersonRulesGrid');
          if (err || !d || !d.success) {
            grid.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var fields = d.fields || [];
          var rules = d.rules || {};
          var html = ''
            + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
            + '<thead><tr style="background:#f0f3f6;text-align:left;">'
            + '<th style="padding:6px 8px;">Field</th>'
            + '<th style="padding:6px 8px;">Direction</th>'
            + '<th style="padding:6px 8px;">Mode</th>'
            + '</tr></thead><tbody>';
          for (var i = 0; i < fields.length; i++) {
            var f = fields[i];
            var r = rules[f.key] || {direction: 'none', mode: 'review'};
            var optDir = ''
              + '<option value="none"' + (r.direction === 'none' ? ' selected' : '') + '>No sync (TP is source of truth)</option>'
              + '<option value="pco_to_tp"' + (r.direction === 'pco_to_tp' ? ' selected' : '') + '>PCO &rarr; TP (pull from PCO)</option>'
              + '<option value="tp_to_pco" disabled' + (r.direction === 'tp_to_pco' ? ' selected' : '') + '>TP &rarr; PCO (coming in v1.1)</option>';
            var optMode = ''
              + '<option value="review"' + (r.mode === 'review' ? ' selected' : '') + '>Review first</option>'
              + '<option value="auto"' + (r.mode === 'auto' ? ' selected' : '') + '>Auto-apply</option>';
            html += '<tr class="pco-rule-row" data-field="' + escAttr(f.key) + '" style="border-top:1px solid #e1e4e8;">'
              + '<td style="padding:6px 8px;font-weight:600;">' + escHtml(f.label) + '</td>'
              + '<td style="padding:6px 8px;"><select class="pco-rule-direction" style="padding:4px 6px;">' + optDir + '</select></td>'
              + '<td style="padding:6px 8px;"><select class="pco-rule-mode" style="padding:4px 6px;">' + optMode + '</select></td>'
              + '</tr>';
          }
          html += '</tbody></table>';
          grid.innerHTML = html;
        });
      }

      function loadAndRenderSettingsStatus() {
        ajax('load_settings', {}, function(err, d){
          if (err || !d) { $('pcoSettingsStatus').innerHTML = '<span class="pco-pill pco-err">Error loading settings</span>'; return; }
          if (d.hasCredentials) {
            $('pcoSettingsStatus').innerHTML = '<span class="pco-pill pco-ok">Configured</span> '
              + '<span class="pco-muted">App ID: ' + (d.appIdMasked || '?') + '</span>';
          } else {
            $('pcoSettingsStatus').innerHTML = '<span class="pco-pill pco-warn">Not configured</span>'
              + ' <span class="pco-muted">Enter credentials below.</span>';
          }
        });
      }

      // ---- Placeholder tabs (built in upcoming versions) ---------

      function renderSyncTab() {
        $('pcoContent').innerHTML = ''
          + '<div id="pcoHealthPanel"></div>'
          + '<div class="pco-card">'
          + '<h3>Sync Dashboard <span style="background:#0f7c84;color:#fff;font-size:10px;font-weight:700;letter-spacing:0.5px;padding:2px 6px;border-radius:3px;margin-left:6px;">PCO &rarr; TP</span></h3>'
          + '<div class="pco-help">'
          + 'Every sync path in one place: All People, Service Type Sync (umbrella with team subgroups), Team Sync, '
          + 'Roster Sync (add-only service-type mappings), and per-plan attendance. Click any Preview &amp; Sync '
          + 'button to walk through what will change before it writes. '
          + '<strong>One-way:</strong> PCO is the source of truth; TouchPoint changes don\'t flow back.'
          + '</div>'
          + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">'
          + '  <button class="pco-btn pco-secondary" id="pcoSyncRefresh">Refresh</button>'
          + '  <span class="pco-muted">Window: last 30 days through next 7 days</span>'
          + '</div>'
          + '<div id="pcoSyncList" class="pco-empty">Loading plans from PCO...</div>'
          + '</div>';
        $('pcoSyncRefresh').onclick = function(){
          loadRecentPlans();
          loadSyncMappingsHealth();
        };
        loadSyncMappingsHealth();
        loadRecentPlans();
      }

      // Health panel sits above the Sync Dashboard card. Renders only
      // when there's at least one issue so a healthy install gets no
      // extra clutter.
      function loadSyncMappingsHealth() {
        var host = $('pcoHealthPanel');
        if (!host) return;
        host.innerHTML = '';
        ajax('load_sync_mappings_health', {}, function(err, d){
          if (err || !d || !d.success) {
            // Silent fail -- the dashboard itself still works; we just
            // don't surface health.
            return;
          }
          var issues = (d && d.issues) || [];
          if (!issues.length) return;
          var color = d.hasErrors ? '#a00' : '#a85a00';
          var headerLabel = d.hasErrors ? (issues.length + ' mapping issue(s) need attention') : (issues.length + ' mapping warning(s)');
          var html = ''
            + '<div class="pco-card" style="background:#fff4f0;border-color:' + color + ';border-left:4px solid ' + color + ';margin-bottom:14px;">'
            + '<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">'
            + '<h3 style="margin:0;color:' + color + ';font-size:15px;">' + escHtml(headerLabel) + '</h3>'
            + '<a href="#" id="pcoHealthRefresh" style="font-size:13px;">Re-check</a>'
            + '</div>'
            + '<div style="margin-top:8px;display:flex;flex-direction:column;gap:6px;">';
          for (var i = 0; i < issues.length; i++) {
            html += renderHealthIssueRow(issues[i]);
          }
          html += '</div></div>';
          host.innerHTML = html;
          $('pcoHealthRefresh').onclick = function(ev){ ev.preventDefault(); loadSyncMappingsHealth(); };
          host.addEventListener('click', function(ev){
            var rm = ev.target.closest && ev.target.closest('.pco-health-remove');
            var gomap = ev.target.closest && ev.target.closest('.pco-health-go-mappings');
            if (rm) {
              if (!confirm('Remove this mapping? Existing TP members/subgroups will NOT be removed.')) return;
              var type = rm.dataset.mtype;
              var key = rm.dataset.mkey;
              var ajaxAction = null, payload = null;
              if (type === 'service_type') { ajaxAction = 'save_org_mapping'; payload = {pco_service_type_id: key, tp_org_id: 0}; }
              else if (type === 'team')    { ajaxAction = 'delete_team_mapping'; payload = {pco_team_id: key}; }
              else if (type === 'people')  { ajaxAction = 'delete_people_mapping'; payload = {pco_service_type_id: key}; }
              else if (type === 'all_people') { ajaxAction = 'delete_all_people_mapping'; payload = {}; }
              if (!ajaxAction) return;
              ajax(ajaxAction, payload, function(err, d){
                if (err || !d || !d.success) { toast('Remove failed: ' + ((d && d.message) || err), 'err'); return; }
                toast('Removed.', 'ok');
                loadSyncMappingsHealth();
              });
            } else if (gomap) {
              selectTab('mappings');
            }
          });
        });
      }

      function renderHealthIssueRow(issue) {
        var sideTag = issue.side === 'tp' ? 'TP' : 'PCO';
        var sideColor = issue.side === 'tp' ? '#1f4e79' : '#1f6b3a';
        var sevColor = issue.severity === 'error' ? '#a00' : '#a85a00';
        return ''
          + '<div style="border:1px solid #eee;border-radius:4px;padding:8px 10px;background:#fff;">'
          + '<div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:600;color:' + sevColor + ';font-size:13px;">' + escHtml(issue.mappingLabel || 'Mapping')
          +   ' <span class="pco-pill" style="background:' + sideColor + ';color:#fff;font-size:10px;font-weight:600;margin-left:4px;">' + sideTag + ' side</span></div>'
          + '<div style="margin-top:2px;font-size:13px;color:#333;">' + escHtml(issue.message || '') + '</div>'
          + '</div>'
          + '<div style="display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<button class="pco-btn pco-secondary pco-health-go-mappings" style="font-size:12px;padding:4px 10px;">Open Sync Mappings</button>'
          + '<button class="pco-btn pco-secondary pco-health-remove" data-mtype="' + escAttr(issue.mappingType || '') + '" data-mkey="' + escAttr(issue.mappingKey || '') + '" style="font-size:12px;padding:4px 10px;color:#a00;border-color:#a00;">Remove</button>'
          + '</div>'
          + '</div>'
          + '</div>';
      }

      function loadRecentPlans() {
        var host = $('pcoSyncList');
        host.className = 'pco-empty';
        host.innerHTML = 'Loading plans from PCO...';
        ajax('list_recent_plans', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var plans = d.plans || [];
          var rosterRollups = d.rosterRollups || [];
          var peopleMappings = d.peopleMappings || [];
          var teamMappings = d.teamMappings || [];
          var allPeopleMapping = d.allPeopleMapping || null;
          var nothing = !plans.length && !rosterRollups.length && !peopleMappings.length && !teamMappings.length && !allPeopleMapping;
          if (nothing) {
            host.innerHTML = '<div>' + (d.message || 'No plans or mappings found.') + '</div>'
              + '<div class="pco-muted" style="margin-top:8px;">Open the Sync Mappings tab to wire up a mapping.</div>';
            return;
          }
          renderPlansList(host, plans, rosterRollups, peopleMappings, teamMappings, allPeopleMapping, d.warnings || []);
        });
      }

      function renderPlansList(host, plans, rosterRollups, peopleMappings, teamMappings, allPeopleMapping, warnings) {
        host.className = '';
        var html = '';
        if (warnings && warnings.length) {
          html += '<div class="pco-card" style="background:#fff8e0;border-color:#f0c870;margin-bottom:10px;padding:8px 12px;">'
            + '<strong>Warnings:</strong><ul style="margin:4px 0 0 20px;font-size:13px;">';
          for (var w = 0; w < warnings.length; w++) {
            html += '<li>' + escHtml(warnings[w]) + '</li>';
          }
          html += '</ul></div>';
        }
        // Group by planDateIso. Sort group keys ascending so the user can
        // scroll up to past days and down to future ones, with today
        // auto-anchored in the middle (PCO-style).
        var groups = {};
        for (var i = 0; i < plans.length; i++) {
          var key = plans[i].planDateIso || 'unknown';
          if (!groups[key]) groups[key] = [];
          groups[key].push(plans[i]);
        }
        var dayKeys = Object.keys(groups).sort();

        // Top-of-page summary line covers all sync types.
        var summaryBits = [];
        if (allPeopleMapping) summaryBits.push('1 all-people mapping');
        if (peopleMappings && peopleMappings.length) summaryBits.push(peopleMappings.length + ' people mapping(s)');
        if (teamMappings && teamMappings.length) summaryBits.push(teamMappings.length + ' team mapping(s)');
        if (rosterRollups && rosterRollups.length) summaryBits.push(rosterRollups.length + ' roster-sync mapping(s)');
        if (plans.length) summaryBits.push(plans.length + ' plan(s) across ' + dayKeys.length + ' day(s)');
        if (summaryBits.length) {
          html += '<div class="pco-muted" style="margin-bottom:8px;">' + summaryBits.join(' &middot; ') + '</div>';
        }

        // ALL PEOPLE (singleton, broadest scope).
        if (allPeopleMapping) {
          html += '<div style="margin-bottom:18px;">'
               + '<div class="pco-day-header" style="margin:0 0 6px 0;padding:6px 10px;background:#1f4e79;color:#fff;border-radius:4px;font-size:13px;font-weight:700;letter-spacing:0.5px;">ALL PEOPLE (entire PCO directory)</div>'
               + '<div style="display:flex;flex-direction:column;gap:6px;">'
               + renderDashAllPeopleCard(allPeopleMapping)
               + '</div></div>';
        }

        // SERVICE TYPE SYNC (umbrella -> involvement with team subgroups).
        if (peopleMappings && peopleMappings.length) {
          html += '<div style="margin-bottom:18px;">'
               + '<div class="pco-day-header" style="margin:0 0 6px 0;padding:6px 10px;background:#1f4e79;color:#fff;border-radius:4px;font-size:13px;font-weight:700;letter-spacing:0.5px;">SERVICE TYPE SYNC (umbrella &rarr; involvement)</div>'
               + '<div style="display:flex;flex-direction:column;gap:6px;">';
          for (var pm = 0; pm < peopleMappings.length; pm++) {
            html += renderDashPeopleCard(peopleMappings[pm]);
          }
          html += '</div></div>';
        }

        // TEAM MAPPINGS (one team -> dedicated involvement, positions as subgroups).
        if (teamMappings && teamMappings.length) {
          html += '<div style="margin-bottom:18px;">'
               + '<div class="pco-day-header" style="margin:0 0 6px 0;padding:6px 10px;background:#1f6b3a;color:#fff;border-radius:4px;font-size:13px;font-weight:700;letter-spacing:0.5px;">TEAM SYNC (one team &rarr; one involvement)</div>'
               + '<div style="display:flex;flex-direction:column;gap:6px;">';
          for (var tm = 0; tm < teamMappings.length; tm++) {
            html += renderDashTeamCard(teamMappings[tm]);
          }
          html += '</div></div>';
        }

        // ROSTER ROLLUP SECTION (add-only Service Type mappings).
        if (rosterRollups && rosterRollups.length) {
          html += '<div class="pco-roster-section" style="margin-bottom:18px;">'
               + '<div class="pco-day-header" style="margin:0 0 6px 0;padding:6px 10px;background:#a85a00;color:#fff;border-radius:4px;font-size:13px;font-weight:700;letter-spacing:0.5px;">'
               +   'ROSTER SYNC (add-member-only Service Type mappings)'
               + '</div>'
               + '<div style="display:flex;flex-direction:column;gap:6px;">';
          for (var rr = 0; rr < rosterRollups.length; rr++) {
            html += renderRosterRollupCard(rosterRollups[rr]);
          }
          html += '</div></div>';
        }

        // ATTENDANCE SYNC group (per-plan attendance mappings, grouped
        // by day). Bordered+indented container so the day sections
        // visually nest under the section banner instead of looking
        // peer-level. Color is teal -- distinct from ROSTER SYNC's
        // orange so the two never get visually confused.
        if (plans.length) {
          html += '<div style="margin-bottom:18px;border:1px solid #e1e4e8;border-left:3px solid #0f7c84;border-radius:6px;background:#fafbfc;">'
               + '<div class="pco-day-header" style="margin:0;padding:6px 10px;background:#0f7c84;color:#fff;border-radius:5px 5px 0 0;font-size:13px;font-weight:700;letter-spacing:0.5px;">ATTENDANCE SYNC (per-plan, by date)</div>'
               + '<div style="padding:8px 12px 10px 12px;display:flex;flex-direction:column;gap:0;">';
          var scrollTarget = pickScrollTargetDay(dayKeys);
          for (var k = 0; k < dayKeys.length; k++) {
            var dk = dayKeys[k];
            var dayPlans = groups[dk];
            dayPlans.sort(function(a, b){ return (a.sortDate || '').localeCompare(b.sortDate || ''); });
            var expanded = (dk === scrollTarget);
            var toggleChar = expanded ? '▼' : '▶';
            html += '<div class="pco-day-section" data-date="' + escAttr(dk) + '" style="margin-top:' + (k === 0 ? '0' : '6px') + ';">'
                 + '<div class="pco-day-header pco-day-toggle-header" style="cursor:pointer;margin:0 0 4px 0;padding:5px 10px;background:#eef2f5;color:#1f4e79;border:1px solid #d8dee3;border-radius:4px;font-size:13px;font-weight:600;display:flex;justify-content:space-between;align-items:center;">'
                 +   '<span><span class="pco-day-toggle" style="display:inline-block;width:14px;">' + toggleChar + '</span> ' + escHtml(formatDayHeader(dk)) + '</span>'
                 +   '<span style="font-weight:400;font-size:12px;color:#666;">' + dayPlans.length + ' plan(s)</span>'
                 + '</div>'
                 + '<div class="pco-day-body" style="display:' + (expanded ? 'flex' : 'none') + ';flex-direction:column;gap:6px;padding-left:14px;">';
            for (var i = 0; i < dayPlans.length; i++) {
              html += renderPlanCard(dayPlans[i]);
            }
            html += '</div></div>';
          }
          html += '</div></div>';
        }
        host.innerHTML = html;
        host.addEventListener('click', function(ev){
          var dayToggle = ev.target.closest && ev.target.closest('.pco-day-toggle-header');
          if (dayToggle) {
            var section = dayToggle.closest('.pco-day-section');
            if (section) {
              var body = section.querySelector('.pco-day-body');
              var tog = section.querySelector('.pco-day-toggle');
              var hidden = !body || body.style.display === 'none';
              if (body) body.style.display = hidden ? 'flex' : 'none';
              if (tog) tog.textContent = hidden ? '▼' : '▶';
            }
            return;
          }
          var planBtn = ev.target.closest && ev.target.closest('.pco-plan-preview-btn');
          var rosterBtn = ev.target.closest && ev.target.closest('.pco-roster-preview-btn');
          var dashPeopleBtn = ev.target.closest && ev.target.closest('.pco-dash-people-btn');
          var dashTeamBtn   = ev.target.closest && ev.target.closest('.pco-dash-team-btn');
          var dashAllPpl    = ev.target.closest && ev.target.closest('.pco-dash-allpeople-btn');
          if (dashPeopleBtn) {
            openPeopleSyncPreview(dashPeopleBtn.dataset.pcoStId, dashPeopleBtn.dataset.tpOrgId);
            return;
          }
          if (dashTeamBtn) {
            openTeamSyncPreview(dashTeamBtn.dataset.pcoTeamId, dashTeamBtn.dataset.tpOrgId);
            return;
          }
          if (dashAllPpl) {
            openAllPeoplePreview();
            return;
          }
          if (planBtn) {
            openPlanPreview(planBtn.dataset.planId, planBtn.dataset.serviceTypeId, planBtn.dataset.orgId);
          } else if (rosterBtn) {
            openRosterPreview(rosterBtn.dataset.serviceTypeId, rosterBtn.dataset.orgId);
          }
        });
        // Auto-scroll to today (or nearest upcoming day, or most recent
        // past if everything is past). Lets staff land on the current day
        // and scroll up for history -- same pattern as the PCO mobile app.
        setTimeout(function(){ scrollToTodayOrNext(host, dayKeys); }, 50);
      }

      function todayIsoLocal() {
        var d = new Date();
        return d.getFullYear() + '-'
          + ('0' + (d.getMonth() + 1)).slice(-2) + '-'
          + ('0' + d.getDate()).slice(-2);
      }

      function formatDayHeader(iso) {
        // iso is YYYY-MM-DD. Parse as local so we don't shift by timezone.
        var parts = iso.split('-');
        if (parts.length !== 3) return iso;
        var d = new Date(parseInt(parts[0],10), parseInt(parts[1],10)-1, parseInt(parts[2],10));
        var dow = ['SUNDAY','MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY'][d.getDay()];
        var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        var todayIso = todayIsoLocal();
        var suffix = '';
        if (iso === todayIso) suffix = ' (Today)';
        return dow + ' \xb7 ' + months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear() + suffix;
      }

      // Same target selection logic used by render to pick which day
      // section starts expanded. Kept as its own helper so render +
      // scroll always agree. Today preferred; if there's no plan
      // today, fall back to the MOST RECENT PAST day -- this is what
      // staff actually want for "what was the last service we ran",
      // not "what's the next one scheduled". Only if no past/today
      // days exist do we walk forward to the nearest future plan.
      function pickScrollTargetDay(dayKeys) {
        if (!dayKeys || !dayKeys.length) return null;
        var todayIso = todayIsoLocal();
        var mostRecentPast = null;
        for (var i = 0; i < dayKeys.length; i++) {
          if (dayKeys[i] <= todayIso) {
            mostRecentPast = dayKeys[i]; // ascending sort means each match overwrites
          } else {
            break;
          }
        }
        if (mostRecentPast) return mostRecentPast;
        return dayKeys[0]; // no past/today -> nearest future
      }

      function scrollToTodayOrNext(host, dayKeys) {
        if (!dayKeys || !dayKeys.length) return;
        var target = pickScrollTargetDay(dayKeys);
        if (!target) return;

        function doScroll() {
          var section = host.querySelector('.pco-day-section[data-date="' + target + '"]');
          if (!section) return;
          // Make sure the target day is expanded -- otherwise the user
          // lands on a collapsed header and has to click to see anything.
          var body = section.querySelector('.pco-day-body');
          var tog = section.querySelector('.pco-day-toggle');
          if (body && body.style.display === 'none') {
            body.style.display = 'flex';
            if (tog) tog.textContent = '▼';
          }
          // Compute the section's absolute Y in the document and scroll the
          // window. scrollIntoView is finicky inside TouchPoint's page (it
          // sometimes picks the wrong scroll container), so we do it
          // explicitly. Subtract a small margin so the day header isn't
          // jammed against the top.
          var rect = section.getBoundingClientRect();
          var top = rect.top + (window.pageYOffset || document.documentElement.scrollTop) - 16;
          try { window.scrollTo({top: top, behavior: 'smooth'}); }
          catch(e) { window.scrollTo(0, top); }
        }
        // Fire twice -- once now, once after layout settles. The first call
        // gets us close; the second corrects for any reflow from late-loading
        // content (web fonts, etc.) so we actually land on the day header.
        doScroll();
        setTimeout(doScroll, 250);
      }

      // Tiny "last synced X ago" pill. Server sends an ISO string from
      // the audit log walk (or null if never synced); we render relative.
      function lastSyncPill(isoStr) {
        if (!isoStr) {
          return '<span class="pco-pill" style="background:#eef0f2;color:#666;font-size:10px;">Never synced</span>';
        }
        var label = relativeFromIso(isoStr);
        return '<span class="pco-pill" style="background:#1f6b3a;color:#fff;font-size:10px;" title="' + escAttr(isoStr) + '">Synced ' + escHtml(label) + '</span>';
      }

      // v3.3: next-scheduled-run pill on dashboard cards. Picks the most
      // human-readable form: "Next: today 6 PM" / "Next: Sun 6 AM".
      function nextRunPill(scheduleLabel, nextRunIso) {
        if (!scheduleLabel) return '';
        var when = scheduleLabel;
        return '<span class="pco-pill" style="background:#1f4e79;color:#fff;font-size:10px;margin-left:4px;" title="' + escAttr(nextRunIso || '') + '">Next: ' + escHtml(when) + '</span>';
      }

      function relativeFromIso(isoStr) {
        // Server iso is local-naive (no Z). Parse as if local; if parsing
        // fails, fall back to the raw string so the user still sees
        // something meaningful.
        var ms = Date.parse(isoStr);
        if (isNaN(ms)) return isoStr;
        var diff = Math.max(0, Date.now() - ms);
        var sec = Math.floor(diff / 1000);
        if (sec < 45) return 'just now';
        var min = Math.floor(sec / 60);
        if (min < 60) return min + 'm ago';
        var hr = Math.floor(min / 60);
        if (hr < 24) return hr + 'h ago';
        var day = Math.floor(hr / 24);
        if (day < 7) return day + 'd ago';
        if (day < 30) return Math.floor(day / 7) + 'w ago';
        // Older than a month: show the date.
        try {
          var d = new Date(ms);
          return (d.getMonth() + 1) + '/' + d.getDate() + '/' + (d.getFullYear() % 100);
        } catch(e) { return isoStr.substring(0, 10); }
      }

      function planTimeLabel(sortDate) {
        if (!sortDate) return '';
        // PCO's sort_date is a "fudge" timestamp -- the time portion is the
        // service's local clock time (8:30 AM at the church), but PCO stamps
        // it with Z (UTC) anyway. If we use new Date(sortDate), the browser
        // shifts the time by its own timezone offset, which is wrong.
        // Parse just the HH:MM portion directly as written.
        var m = sortDate.match(/T(\d{1,2}):(\d{2})/);
        if (!m) return '';
        var h = parseInt(m[1], 10);
        var min = parseInt(m[2], 10);
        if (isNaN(h) || isNaN(min)) return '';
        var ampm = h >= 12 ? 'PM' : 'AM';
        h = h % 12; if (h === 0) h = 12;
        return h + ':' + (min < 10 ? '0' + min : min) + ' ' + ampm;
      }

      // Mode helpers for plan cards + preview modal. Returns:
      //   {label, color, btnText, verb}
      // - label  : short string for the card badge ("Attendance", "Add member", "Both")
      // - color  : background for the badge pill
      // - btnText: full text for the Sync button in the modal
      // - verb   : short verb used in the inline "Will Sync" pill per attendee row
      function pcoSyncMode(plan_or_info) {
        // Service Type Sync: PCO Service Type -> umbrella TP involvement
        // + team-as-subgroup. No per-plan, no attendance. (Internally
        // still called "people sync" / isPeopleSync for backwards
        // compatibility with state plumbing.)
        if (plan_or_info.isPeopleSync) {
          var t = !!plan_or_info.teamsAsSubgroups;
          return {
            label: t ? 'Service Type Sync + Teams' : 'Service Type Sync',
            color: '#1f4e79',
            btnText: t ? 'Sync Umbrella Roster + Teams' : 'Sync Umbrella Roster',
            verb: 'Join',
          };
        }
        // Team Sync is a distinct mode -- durable PCO team roster -> TP
        // involvement + per-position subgroups. No attendance ever.
        if (plan_or_info.isTeamSync) {
          var pos = !!plan_or_info.positionsAsSubgroups;
          return {
            label: pos ? 'Team Sync + Positions' : 'Team Sync',
            color: '#1f6b3a',
            btnText: pos ? 'Sync Team Roster + Positions' : 'Sync Team Roster',
            verb: 'Join',
          };
        }
        var att = !!plan_or_info.syncAttendance;
        var mem = !!plan_or_info.autoAddMember;
        // Note: button text is mode-only; the per-row "Will Sync" pill +
        // confirm dialog spell out the Confirmed-vs-all-team-members
        // split per the v1.9.2 rules.
        if (att && mem) return {label: 'Attendance + Member', color: '#1f6b3a', btnText: 'Sync Attendance + Members', verb: 'Sync'};
        if (att && !mem) return {label: 'Attendance only', color: '#1f4e79', btnText: 'Sync Confirmed Attendance', verb: 'Present'};
        if (!att && mem) return {label: 'Add member only', color: '#a85a00', btnText: 'Add Team Members', verb: 'Join'};
        // Both off shouldn't be reachable (server blocks) but be defensive.
        return {label: 'Disabled', color: '#777', btnText: 'No actions enabled', verb: ''};
      }

      // Dashboard cards for the durable-sync mapping types. Same shape
      // as the roster rollup card but with type-specific labels + click
      // wiring routed to the right preview.

      function renderDashAllPeopleCard(r) {
        return ''
          + '<div style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;border-left:3px solid #1f4e79;">'
          + '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;color:#1f4e79;font-size:15px;">PCO People directory</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">'
          +   '<strong>TP:</strong> ' + escHtml(r.tpOrgName) + ' <span class="pco-muted">(#' + r.tpOrgId + ')</span>'
          + '</div>'
          + '<div style="margin-top:4px;">' + lastSyncPill(r.lastSyncedAt) + nextRunPill(r.scheduleLabel, r.nextRunIso) + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<span class="pco-pill" style="background:#1f4e79;color:#fff;font-size:11px;">All People</span>'
          + (r.includeInactive ? '<span class="pco-pill pco-warn" style="font-size:11px;">incl. inactive</span>' : '')
          + '<button class="pco-btn pco-dash-allpeople-btn" style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
          + '</div>'
          + '</div>'
          + '</div>';
      }

      function renderDashPeopleCard(r) {
        return ''
          + '<div style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;border-left:3px solid #1f4e79;">'
          + '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;color:#1f4e79;font-size:15px;">' + escHtml(r.pcoServiceTypeName || ('Service Type ' + r.pcoServiceTypeId)) + '</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">'
          +   '<strong>TP:</strong> ' + escHtml(r.tpOrgName) + ' <span class="pco-muted">(#' + r.tpOrgId + ')</span>'
          + '</div>'
          + '<div style="margin-top:4px;">' + lastSyncPill(r.lastSyncedAt) + nextRunPill(r.scheduleLabel, r.nextRunIso) + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<span class="pco-pill" style="background:#1f4e79;color:#fff;font-size:11px;">People Sync</span>'
          + (r.teamsAsSubgroups ? '<span class="pco-pill" style="background:#1f6b3a;color:#fff;font-size:11px;">+ Teams</span>' : '')
          + '<button class="pco-btn pco-dash-people-btn"'
          + ' data-pco-st-id="' + escAttr(r.pcoServiceTypeId) + '"'
          + ' data-tp-org-id="' + r.tpOrgId + '"'
          + ' style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
          + '</div>'
          + '</div>'
          + '</div>';
      }

      function renderDashTeamCard(r) {
        return ''
          + '<div style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;border-left:3px solid #1f6b3a;">'
          + '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;color:#1f4e79;font-size:15px;">' + escHtml(r.pcoTeamName || ('Team ' + r.pcoTeamId))
          +   (r.pcoServiceTypeName ? ' <span class="pco-muted" style="font-weight:400;font-size:13px;">(' + escHtml(r.pcoServiceTypeName) + ')</span>' : '')
          + '</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">'
          +   '<strong>TP:</strong> ' + escHtml(r.tpOrgName) + ' <span class="pco-muted">(#' + r.tpOrgId + ')</span>'
          + '</div>'
          + '<div style="margin-top:4px;">' + lastSyncPill(r.lastSyncedAt) + nextRunPill(r.scheduleLabel, r.nextRunIso) + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<span class="pco-pill" style="background:#1f6b3a;color:#fff;font-size:11px;">Team Sync</span>'
          + (r.positionsAsSubgroups ? '<span class="pco-pill" style="background:#1f4e79;color:#fff;font-size:11px;">+ Positions</span>' : '')
          + '<button class="pco-btn pco-dash-team-btn"'
          + ' data-pco-team-id="' + escAttr(r.pcoTeamId) + '"'
          + ' data-tp-org-id="' + r.tpOrgId + '"'
          + ' style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
          + '</div>'
          + '</div>'
          + '</div>';
      }

      function renderRosterRollupCard(r) {
        var primary = r.serviceTypeName || '(Service Type unknown)';
        var dateLine = r.planCount + ' plan(s) in window'
                     + (r.latestPlanIso ? ' &middot; latest: ' + escHtml(formatDayHeader(r.latestPlanIso)) : '');
        return ''
          + '<div style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;border-left:3px solid #a85a00;">'
          + '<div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;font-size:15px;color:#1f4e79;">' + escHtml(primary) + '</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">'
          +   '<strong>TP:</strong> ' + escHtml(r.tpOrgName) + ' <span class="pco-muted">(#' + r.tpOrgId + ')</span>'
          + '</div>'
          + '<div style="margin-top:2px;font-size:12px;color:#888;">' + dateLine + '</div>'
          + '<div style="margin-top:4px;">' + lastSyncPill(r.lastSyncedAt) + nextRunPill(r.scheduleLabel, r.nextRunIso) + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<span class="pco-pill" style="background:#a85a00;color:#fff;font-size:11px;">Add member only</span>'
          + '<button class="pco-btn pco-roster-preview-btn"'
          + ' data-service-type-id="' + escAttr(r.serviceTypeId) + '"'
          + ' data-org-id="' + r.tpOrgId + '"'
          + ' style="font-size:13px;padding:5px 12px;">Preview Roster &amp; Sync</button>'
          + '</div>'
          + '</div>'
          + '</div>';
      }

      function renderPlanCard(plan) {
        var isMapped = !!(plan.tpOrgId);
        // PCO-style card: service type name is the bold primary title;
        // raw plan title (if any) becomes a subtitle.
        var primary = plan.serviceTypeName || '(Service Type unknown)';
        var planSubtitle = (plan.planTitle && plan.planTitle !== primary) ? plan.planTitle : '';
        var time = planTimeLabel(plan.sortDate);

        var rightSide = '';
        if (isMapped) {
          var mode = pcoSyncMode(plan);
          var modeBadge = '<span class="pco-pill" style="background:' + mode.color + ';color:#fff;font-size:11px;">' + escHtml(mode.label) + '</span>';
          rightSide = ''
            + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;">'
            + (time ? '<span class="pco-muted" style="font-size:13px;">' + escHtml(time) + '</span>' : '')
            + modeBadge
            + '<button class="pco-btn pco-plan-preview-btn" '
            + ' data-plan-id="' + escAttr(plan.planId) + '"'
            + ' data-service-type-id="' + escAttr(plan.serviceTypeId) + '"'
            + ' data-org-id="' + plan.tpOrgId + '"'
            + ' style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
            + '</div>';
        } else {
          rightSide = ''
            + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;">'
            + (time ? '<span class="pco-muted" style="font-size:13px;">' + escHtml(time) + '</span>' : '')
            + '<span class="pco-pill pco-warn">Unmapped</span>'
            + '</div>';
        }

        var subtitleLine = '';
        if (planSubtitle) {
          subtitleLine += '<span style="font-style:italic;">' + escHtml(planSubtitle) + '</span> &middot; ';
        }
        if (isMapped) {
          subtitleLine += '<strong>TP:</strong> ' + escHtml(plan.tpOrgName)
                       + ' <span class="pco-muted">(#' + plan.tpOrgId + ')</span>';
        } else {
          subtitleLine += '<span class="pco-muted">Map this service type to a TouchPoint involvement to enable sync.</span>';
        }

        return ''
          + '<div style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;border-left:3px solid ' + (isMapped ? '#1f6b3a' : '#e67e22') + ';">'
          + '<div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;font-size:15px;color:#1f4e79;">' + escHtml(primary) + '</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">' + subtitleLine + '</div>'
          + (isMapped ? '<div style="margin-top:4px;">' + lastSyncPill(plan.lastSyncedAt) + '</div>' : '')
          + '</div>'
          + rightSide
          + '</div>'
          + '</div>';
      }

      // ---- Per-plan review modal ----

      var _planPreviewState = null; // {planId, serviceTypeId, orgId, attendees: [...], planInfo}

      function openPlanPreview(planId, serviceTypeId, orgId) {
        renderPlanPreviewModal('<div class="pco-empty">Loading plan from PCO...</div>');
        ajax('load_plan_preview', {plan_id: planId, service_type_id: serviceTypeId}, function(err, d){
          if (err || !d || !d.success) {
            renderPlanPreviewModal('<div class="pco-pill pco-err">Error</div><div>' + escHtml((d && d.message) || err) + '</div>');
            return;
          }
          _planPreviewState = {
            planId: planId, serviceTypeId: serviceTypeId, orgId: parseInt(orgId, 10) || 0,
            attendees: d.attendees || [],
            planInfo: d.planInfo || {},
            summary: d.summary || {},
            isRosterSync: false,
          };
          renderPlanPreviewBody();
        });
      }

      function openRosterPreview(serviceTypeId, orgId) {
        renderPlanPreviewModal('<div class="pco-empty">Loading roster from PCO (this walks every plan in the window)...</div>');
        ajax('load_roster_preview', {service_type_id: serviceTypeId}, function(err, d){
          if (err || !d || !d.success) {
            renderPlanPreviewModal('<div class="pco-pill pco-err">Error</div><div>' + escHtml((d && d.message) || err) + '</div>');
            return;
          }
          _planPreviewState = {
            planId: '',
            serviceTypeId: serviceTypeId,
            orgId: parseInt(orgId, 10) || 0,
            attendees: d.attendees || [],
            planInfo: d.planInfo || {},
            summary: d.summary || {},
            isRosterSync: true,
          };
          renderPlanPreviewBody();
        });
      }

      function renderPlanPreviewModal(innerHtml) {
        var existing = document.getElementById('pcoPlanModal');
        if (existing) existing.parentNode.removeChild(existing);
        var modal = document.createElement('div');
        modal.id = 'pcoPlanModal';
        modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.4);z-index:500;display:flex;align-items:center;justify-content:center;padding:20px;';
        modal.innerHTML = ''
          + '<div style="background:#fff;border-radius:8px;max-width:920px;width:100%;max-height:90vh;display:flex;flex-direction:column;overflow:hidden;">'
          + '<div style="padding:14px 18px;border-bottom:1px solid #e1e4e8;display:flex;justify-content:space-between;align-items:center;">'
          + '  <h3 style="margin:0;color:#1f4e79;" id="pcoModalTitle">Plan Preview</h3>'
          + '  <button class="pco-btn pco-secondary" id="pcoModalClose" style="padding:4px 10px;font-size:13px;">&times; Close</button>'
          + '</div>'
          + '<div id="pcoModalBody" style="padding:14px 18px;overflow-y:auto;flex:1;">' + innerHtml + '</div>'
          + '<div id="pcoModalFooter" style="padding:12px 18px;border-top:1px solid #e1e4e8;display:none;justify-content:flex-end;gap:8px;background:#fafbfc;">'
          + '</div>'
          + '</div>';
        document.body.appendChild(modal);
        $('pcoModalClose').onclick = closePlanPreview;
        modal.addEventListener('click', function(ev){
          if (ev.target === modal) closePlanPreview();
        });
      }

      function closePlanPreview() {
        var m = document.getElementById('pcoPlanModal');
        if (m) m.parentNode.removeChild(m);
        _planPreviewState = null;
      }

      function renderPlanPreviewBody() {
        var s = _planPreviewState;
        var primary = s.planInfo.serviceTypeName || s.planInfo.title || 'Plan';
        var datePart = s.planInfo.shortDates || s.planInfo.planDateIso || '';
        var planSub = (s.planInfo.planTitle && s.planInfo.planTitle !== primary) ? s.planInfo.planTitle : '';
        var titleText = primary + (datePart ? ' — ' + datePart : '');
        if (planSub) titleText += ' (' + planSub + ')';
        $('pcoModalTitle').textContent = titleText;
        var summary = s.summary;
        var mode = pcoSyncMode(s.planInfo);
        var body = $('pcoModalBody');
        // Mode banner spells out what's about to happen, so an
        // attendance-only or add-only mapping doesn't surprise the user
        // when they hit Sync.
        var modeBanner = ''
          + '<div style="margin-bottom:10px;padding:8px 10px;border-radius:4px;background:' + mode.color + '0d;border-left:3px solid ' + mode.color + ';font-size:13px;">'
          + '  <strong style="color:' + mode.color + ';">Sync mode:</strong> ' + escHtml(mode.label) + '. '
          + (s.isPeopleSync
              ? ('Everyone on any team under this PCO Service Type will be added as a TouchPoint member of the umbrella involvement.'
                 + (s.planInfo.teamsAsSubgroups ? ' Each team they\'re on will be written as a subgroup membership.' : '')
                 + ' No attendance is written.')
              : '')
          + (s.isTeamSync
              ? ('Everyone on this PCO Team will be added as a TouchPoint member of the involvement.'
                 + (s.planInfo.positionsAsSubgroups ? ' Each Team Position they hold will be written as a subgroup membership inside the involvement.' : '')
                 + ' No attendance is written -- this is a durable roster sync.')
              : '')
          + (!s.isTeamSync && s.planInfo.syncAttendance && s.planInfo.autoAddMember ? '<strong>Confirmed</strong> people will be marked Present on the meeting; <strong>all matched team members</strong> (Confirmed, Unconfirmed, Declined) will be added as members of the involvement.' : '')
          + (!s.isTeamSync && s.planInfo.syncAttendance && !s.planInfo.autoAddMember ? '<strong>Confirmed</strong> people will be marked Present on the meeting. Membership will not be changed -- anyone not already on the roster will appear as a Visitor.' : '')
          + (!s.isTeamSync && !s.planInfo.syncAttendance && s.planInfo.autoAddMember ? '<strong>All matched team members</strong> (Confirmed, Unconfirmed, Declined) will be added as members of the involvement. Team membership is independent of weekly RSVP. No meeting will be created and no attendance written.' : '')
          + '</div>';
        // Team Sync position diagnostic: spells out what PCO returned so
        // "no chips" doesn't look like a bug when it's actually a PCO
        // setup issue (positions exist but nobody's assigned, etc.).
        var posDiag = '';
        if (s.isTeamSync && s.planInfo.positionsAsSubgroups) {
          var pc = s.planInfo.positionCount || 0;
          var pac = s.planInfo.positionAssignmentCount || 0;
          var pwpc = s.planInfo.peopleWithPositionsCount || 0;
          var diagStyle, diagMsg;
          if (pc === 0) {
            diagStyle = 'background:#fdecea;border-left:3px solid #c0392b;color:#8a2020;';
            diagMsg = '<strong>No Team Positions exist on this PCO team.</strong> Add positions in PCO (Services &rarr; Teams &rarr; this team &rarr; Positions tab) before subgroups can sync.';
          } else if (pac === 0) {
            diagStyle = 'background:#fff7e6;border-left:3px solid #f0c36d;color:#8a6d3b;';
            diagMsg = '<strong>' + pc + ' position(s) defined but no one is assigned</strong> to any of them in PCO. The roster will sync but no subgroups will be written. Assign people to positions in PCO to enable subgroup sync.';
          } else if (pwpc === 0) {
            diagStyle = 'background:#fff7e6;border-left:3px solid #f0c36d;color:#8a6d3b;';
            diagMsg = '<strong>' + pc + ' position(s), ' + pac + ' assignment(s)</strong> &mdash; but none of them matched a person on this team\'s roster. (Possible PCO data drift.)';
          } else {
            diagStyle = 'background:#e6f7e6;border-left:3px solid #1f6b3a;color:#1f6b3a;';
            diagMsg = '<strong>' + pc + ' position(s), ' + pac + ' assignment(s)</strong> across ' + pwpc + ' person/people. Subgroups will sync.';
          }
          posDiag = '<div style="margin-bottom:10px;padding:8px 10px;border-radius:4px;font-size:13px;' + diagStyle + '">' + diagMsg + '</div>';
        }
        // v3.2: mirror-drop warning. Sync removes TP members / subgroups
        // that no longer exist in PCO. Surface those counts before sync
        // so nothing happens silently.
        var dropWarn = '';
        if (s.isTeamSync || s.isPeopleSync) {
          var rdc = s.planInfo.rosterDropCount || 0;
          var sdc = s.planInfo.subgroupDropCount || 0;
          if (rdc || sdc) {
            var bits = [];
            if (rdc) bits.push('<strong>' + rdc + ' TP member(s)</strong> will be removed from the involvement (their PCO link is no longer in scope)');
            if (sdc) {
              var sgKind = s.isTeamSync ? 'position' : 'team';
              bits.push('<strong>' + sdc + ' stale subgroup association(s)</strong> will be dropped (' + sgKind + 's that no longer match PCO)');
            }
            dropWarn = '<div style="margin-bottom:10px;padding:8px 10px;border-radius:4px;font-size:13px;background:#fff7e6;border-left:3px solid #c0392b;color:#8a2020;">'
              + '<strong>Mirror removal:</strong> PCO is the source of truth -- ' + bits.join('; ') + '.'
              + '</div>';
          }
        }
        var html = ''
          + modeBanner
          + posDiag
          + dropWarn
          + '<div id="pcoSummaryPills" style="margin-bottom:12px;">'
          + '<span class="pco-pill pco-ok">' + (summary.matched || 0) + ' matched</span> '
          + '<span class="pco-pill pco-warn">' + (summary.unmatched || 0) + ' unmatched</span> '
          + (summary.ambiguousEmail ? '<span class="pco-pill pco-warn">' + summary.ambiguousEmail + ' ambiguous email</span> ' : '')
          + (((s.isTeamSync || s.isPeopleSync) && (s.planInfo.rosterDropCount || s.planInfo.subgroupDropCount))
              ? '<span class="pco-pill" style="background:#c0392b;color:#fff;">' + ((s.planInfo.rosterDropCount || 0) + (s.planInfo.subgroupDropCount || 0)) + ' will be removed</span> '
              : '')
          + '<span class="pco-muted">of ' + (summary.total || 0) + ' total attendees</span>'
          + '</div>'
          + '<div class="pco-help">'
          + (s.isPeopleSync
              ? 'People Sync aggregates everyone across every team in this service type into the umbrella involvement. '
                + (s.planInfo.teamsAsSubgroups ? 'Each team they\'re on becomes a subgroup membership. ' : '')
                + 'No attendance, no per-plan logic. '
              : '')
          + (s.isTeamSync
              ? 'Team Sync writes the durable PCO team roster -- everyone on the team becomes a member of the involvement. '
                + (s.planInfo.positionsAsSubgroups ? 'Each Team Position they hold writes a subgroup membership. ' : '')
                + 'No attendance, no per-plan logic. '
              : '')
          + (!s.isTeamSync && s.planInfo.syncAttendance && s.planInfo.autoAddMember
              ? '<strong>Attendance</strong>: only Confirmed (status C) attendees get marked Present. <strong>Membership</strong>: every matched team member (Confirmed, Unconfirmed, Declined) gets added. '
              : '')
          + (!s.isTeamSync && s.planInfo.syncAttendance && !s.planInfo.autoAddMember
              ? 'Only <strong>Confirmed</strong> (status C) attendees will be marked Present. Membership will not be changed. '
              : '')
          + (!s.isTeamSync && !s.planInfo.syncAttendance && s.planInfo.autoAddMember
              ? '<strong>All matched team members</strong> get added to the roster regardless of weekly RSVP (Confirmed, Unconfirmed, Declined). '
              : '')
          + 'Use the search box to manually match anyone unmatched -- the link is saved as a person extra value '
          + 'so they auto-match next time.'
          + '</div>';
        // Bulk matcher call-to-action when unmatched count is meaningful.
        // For previews with a handful (< 10) the per-row Search & Link
        // pattern is fine; above that, point staff at the dedicated
        // bulk matcher so they aren't stuck doing 50+ manual searches.
        var unmatchedCount = summary.unmatched || 0;
        if (unmatchedCount >= 10) {
          html += '<div style="margin-top:10px;padding:8px 12px;background:#0f7c840d;border-left:3px solid #0f7c84;border-radius:4px;font-size:13px;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">'
            + '<span><strong>' + unmatchedCount + ' unmatched here.</strong> The Proposed Matches review (People Matching tab) scores name + email signals against TP for these specific records. Apply / Skip per row, with bulk Apply for high-confidence matches.</span>'
            + '<button class="pco-btn" id="pcoOpenBulkMatcherInline" style="font-size:13px;padding:5px 12px;white-space:nowrap;">Open Proposed Matches &rarr;</button>'
            + '</div>';
        }
        html += '<div id="pcoAttendeeList" style="margin-top:10px;display:flex;flex-direction:column;gap:6px;"></div>';
        body.innerHTML = html;
        var bulkBtn = $('pcoOpenBulkMatcherInline');
        if (bulkBtn) bulkBtn.onclick = function(){
          // Bundle this preview's unmatched into a scoped subset so the
          // matcher only walks these records (not the full PCO directory).
          var unm = [];
          for (var ui = 0; ui < s.attendees.length; ui++) {
            var aa = s.attendees[ui];
            if (aa.tpPeopleId) continue; // already matched
            unm.push({
              pcoPersonId: aa.pcoPersonId,
              first_name: aa.pcoFirstName || '',
              last_name:  aa.pcoLastName || '',
              email: aa.email || '',
              name: aa.name || '',
            });
          }
          _proposedScope = unm.length ? {
            label: (s.planInfo.title || s.planInfo.serviceTypeName || 'preview'),
            peopleData: unm,
          } : null;
          _proposedState = {page: 1, tier: 'all', search: ''};
          closePlanPreview();
          selectTab('people');
          setTimeout(function(){
            var loadBtn = document.getElementById('pcoLoadProposedBtn');
            if (loadBtn) loadBtn.click();
          }, 100);
        };
        renderAttendeeRows();
        $('pcoModalFooter').style.display = 'flex';
        $('pcoModalFooter').innerHTML = ''
          + '<button class="pco-btn pco-secondary" onclick="window.__pcoCloseModal()">Cancel</button>'
          + '<button class="pco-btn" id="pcoSyncNowBtn">' + escHtml(mode.btnText) + '</button>';
        window.__pcoCloseModal = closePlanPreview;
        $('pcoSyncNowBtn').onclick = doSyncNow;
      }

      function renderAttendeeRows() {
        var list = $('pcoAttendeeList');
        var s = _planPreviewState;
        var html = '';
        for (var i = 0; i < s.attendees.length; i++) {
          html += renderAttendeeRow(s.attendees[i], i);
        }
        list.innerHTML = html;
        // Auto-trigger search for any prefilled inputs (unmatched/ambiguous
        // rows seed the input with email or name). Without this the staff
        // sees the prefilled value but no results.
        var prefilled = list.querySelectorAll('.pco-search-input');
        for (var j = 0; j < prefilled.length; j++) {
          var inp = prefilled[j];
          if (inp.value && inp.value.trim().length >= 2) {
            // Fire the same path the input listener uses.
            var evt;
            try { evt = new Event('input', {bubbles: true}); }
            catch(e) { evt = document.createEvent('Event'); evt.initEvent('input', true, true); }
            inp.dispatchEvent(evt);
          }
        }
      }

      function renderAttendeeRow(a, idx) {
        // Build position badges. One person can hold multiple positions on
        // the same plan -- show each as "Position (status)" badge so the
        // worship admin sees the full picture at a glance.
        var posBadges = '';
        var positions = a.positions || [];
        for (var pi = 0; pi < positions.length; pi++) {
          var pos = positions[pi];
          var pillKind = 'pco-warn';
          if (pos.status === 'C') pillKind = 'pco-ok';
          else if (pos.status === 'D') pillKind = 'pco-err';
          var label = pos.teamPosition || '(position)';
          if (pos.status === 'C') label += ' \xb7 Confirmed';
          else if (pos.status === 'U') label += ' \xb7 Unconfirmed';
          else if (pos.status === 'D') label += ' \xb7 Declined';
          else if (pos.status) label += ' \xb7 ' + pos.status;
          posBadges += ' <span class="pco-pill ' + pillKind + '" style="font-size:11px;">' + escHtml(label) + '</span>';
        }
        // Per-row "will sync" badge. Depends on both the mode and whether
        // this person is Confirmed for the plan:
        //   Add-only mapping -> everyone gets "Will Add as Member" (RSVP
        //     doesn't matter for membership)
        //   Attendance + Member -> Confirmed = "Will Mark Present + Member";
        //     Unconfirmed/Declined = "Will Add as Member" (still on the team)
        //   Attendance only -> Confirmed = "Will Mark Present"; others = "Not Confirmed"
        var _pi = (_planPreviewState && _planPreviewState.planInfo) || {};
        var _syncAtt = !!_pi.syncAttendance;
        var _addMem  = !!_pi.autoAddMember;
        var confirmedBadge;
        if (!a.tpPeopleId) {
          // Unmatched -- nothing will happen until they're linked.
          confirmedBadge = a.isConfirmed
            ? '<span class="pco-pill pco-warn" style="font-size:11px;">Confirmed (not linked)</span>'
            : '<span class="pco-pill" style="background:#eef0f2;color:#666;font-size:11px;">Not Confirmed</span>';
        } else if (a.isConfirmed) {
          if (_syncAtt && _addMem) confirmedBadge = '<span class="pco-pill pco-ok" style="font-size:11px;">Will Mark Present + Member</span>';
          else if (_syncAtt)       confirmedBadge = '<span class="pco-pill pco-ok" style="font-size:11px;">Will Mark Present</span>';
          else                     confirmedBadge = '<span class="pco-pill pco-ok" style="font-size:11px;">Will Add as Member</span>';
        } else {
          // Unconfirmed or Declined. Membership still applies if auto_add
          // is on; attendance never does.
          if (_addMem) confirmedBadge = '<span class="pco-pill pco-ok" style="font-size:11px;">Will Add as Member</span>';
          else         confirmedBadge = '<span class="pco-pill" style="background:#eef0f2;color:#666;font-size:11px;">Not Confirmed (skipped)</span>';
        }

        var matchInfo = '';
        if (a.tpPeopleId) {
          var src = a.matchSource === 'pco_id' ? 'PCO ID' : (a.matchSource === 'email' ? 'email' : 'manual');
          matchInfo = '<div style="margin-top:4px;font-size:13px;">'
            + '<span class="pco-pill pco-ok" style="font-size:11px;">Matched</span> '
            + '<strong>' + escHtml(a.tpName) + '</strong> '
            + '<span class="pco-muted">(TP #' + a.tpPeopleId + ', via ' + src + ')</span>'
            + '</div>';
        } else if (a.emailAmbiguous) {
          matchInfo = '<div style="margin-top:4px;font-size:13px;">'
            + '<span class="pco-pill pco-warn" style="font-size:11px;">Ambiguous email</span> '
            + '<span class="pco-muted">Multiple TP people match this email. Pick the correct one below.</span>'
            + '<div class="pco-attendee-search" style="margin-top:6px;">'
            + '  <input type="text" class="pco-input pco-search-input" data-idx="' + idx + '" placeholder="Search TouchPoint by name or email..." value="' + escAttr(a.email || a.name) + '">'
            + '  <div class="pco-search-results" style="margin-top:4px;"></div>'
            + '</div>'
            + '</div>';
        } else {
          matchInfo = '<div style="margin-top:4px;font-size:13px;">'
            + '<span class="pco-pill pco-warn" style="font-size:11px;">Unmatched</span> '
            + '<span class="pco-muted">No TouchPoint person linked. Search to match:</span>'
            + '<div class="pco-attendee-search" style="margin-top:6px;">'
            + '  <input type="text" class="pco-input pco-search-input" data-idx="' + idx + '" placeholder="Search TouchPoint by name or email...">'
            + '  <div class="pco-search-results" style="margin-top:4px;"></div>'
            + '</div>'
            + '</div>';
        }

        return ''
          + '<div class="pco-attendee-row" data-idx="' + idx + '" style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;">'
          + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">'
          + '<strong>' + escHtml(a.name) + '</strong>'
          + (a.email ? ' <span class="pco-muted">' + escHtml(a.email) + '</span>' : '')
          + ' ' + confirmedBadge
          + '</div>'
          + (posBadges ? '<div style="margin-top:4px;display:flex;flex-wrap:wrap;gap:4px;">' + posBadges + '</div>' : '')
          + matchInfo
          + '</div>';
      }

      var _attendeeSearchTimer = null;
      document.addEventListener('input', function(ev){
        if (!_planPreviewState) return;
        if (!ev.target.classList) return;
        if (!ev.target.classList.contains('pco-search-input')) return;
        clearTimeout(_attendeeSearchTimer);
        var inp = ev.target;
        var resultsEl = inp.parentElement.querySelector('.pco-search-results');
        var term = inp.value.trim();
        if (term.length < 2) {
          resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
          return;
        }
        resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
        var idx = parseInt(inp.dataset.idx, 10);
        _attendeeSearchTimer = setTimeout(function(){
          ajax('search_tp_people', {search_term: term}, function(err, d){
            if (err || !d || !d.success) {
              resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>';
              return;
            }
            var people = d.people || [];
            if (!people.length) {
              resultsEl.innerHTML = '<span class="pco-muted">No matches.</span>';
              return;
            }
            var html = '<div style="border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:200px;overflow-y:auto;">';
            for (var i = 0; i < people.length; i++) {
              var p = people[i];
              html += '<div class="pco-tp-pick" data-idx="' + idx + '" data-tp-id="' + p.peopleId + '" data-tp-name="' + escAttr(p.name) + '" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;">'
                + '<strong>' + escHtml(p.name) + '</strong> <span class="pco-muted">(TP #' + p.peopleId + (p.age != null ? ', ' + p.age + 'y' : '') + (p.gender ? ', ' + escHtml(p.gender) : '') + ')</span>';
              if (p.email) html += '<div class="pco-muted" style="font-size:12px;">' + escHtml(p.email) + '</div>';
              html += '</div>';
            }
            html += '</div>';
            resultsEl.innerHTML = html;
          });
        }, 250);
      });
      document.addEventListener('click', function(ev){
        if (!_planPreviewState) return;
        var t = ev.target.closest && ev.target.closest('.pco-tp-pick');
        if (!t) return;
        var idx = parseInt(t.dataset.idx, 10);
        var tpId = parseInt(t.dataset.tpId, 10) || 0;
        var tpName = t.dataset.tpName || '';
        if (!_planPreviewState.attendees[idx]) return;
        var a = _planPreviewState.attendees[idx];
        // Save the mapping (writes PCO_PersonId extra value on the TP person)
        // ONLY if the PCO person has an id. Some PCO team_members don't have
        // a linked Person -- in that case we just stage the match in memory
        // for this sync, without persisting.
        if (!a.pcoPersonId) {
          a.tpPeopleId = tpId;
          a.tpName = tpName;
          a.matchSource = 'manual_session';
          a.emailAmbiguous = false;
          renderAttendeeRows();
          updateSummary();
          return;
        }
        ajax('confirm_person_mapping', {
          pco_person_id: a.pcoPersonId,
          tp_people_id: tpId,
          pco_name: a.name,
        }, function(err, d){
          if (err || !d || !d.success) {
            toast('Map failed: ' + ((d && d.message) || err), 'err');
            return;
          }
          a.tpPeopleId = tpId;
          a.tpName = tpName;
          a.matchSource = 'manual';
          a.emailAmbiguous = false;
          renderAttendeeRows();
          updateSummary();
          toast('Linked. ' + a.name + ' will auto-match next time.', 'ok');
        });
      });

      function updateSummary() {
        var matched = 0, ambiguous = 0, unmatched = 0;
        var s = _planPreviewState;
        for (var i = 0; i < s.attendees.length; i++) {
          var a = s.attendees[i];
          if (a.tpPeopleId) matched++;
          else if (a.emailAmbiguous) ambiguous++;
          else unmatched++;
        }
        s.summary = {total: s.attendees.length, matched: matched, ambiguousEmail: ambiguous, unmatched: unmatched};
        // Reflect in pills at the top of the body. Target by ID -- the
        // body's firstChild is now the mode banner, not the summary,
        // since v1.7.2 added the banner above the pills.
        var pills = $('pcoSummaryPills');
        if (pills) {
          pills.innerHTML = '<span class="pco-pill pco-ok">' + matched + ' matched</span> '
            + '<span class="pco-pill pco-warn">' + unmatched + ' unmatched</span> '
            + (ambiguous ? '<span class="pco-pill pco-warn">' + ambiguous + ' ambiguous email</span> ' : '')
            + '<span class="pco-muted">of ' + s.attendees.length + ' total attendees</span>';
        }
      }

      function doSyncNow() {
        var s = _planPreviewState;
        // Sync any unique person whose isConfirmed = true (any of their
        // positions is Confirmed) AND has a TP match. Dedup is server-side
        // safe too, but trim here to keep the audit log clean.
        // Selection rules differ by mode:
        //   syncAttendance only       -> Confirmed only (RSVP'd yes)
        //   autoAddMember + sync_att  -> all team members (membership is
        //                                independent of weekly RSVP);
        //                                attendance still confined to
        //                                Confirmed server-side
        //   autoAddMember only        -> all team members
        // "Confirmed" in PCO is a per-plan response (C/U/D); team
        // membership is "appears on the worship team at all", which is
        // anyone in team_members regardless of status.
        var includeAllMembers = !!s.planInfo.autoAddMember || !!s.isTeamSync || !!s.isPeopleSync;
        var seen = {};
        var toSync = [];
        var confirmedCount = 0;
        // Bundle PCO field values + isConfirmed per person. Server uses
        // isConfirmed to gate attendance writes; membership applies to
        // all entries when auto_add_member is on.
        var peopleJsonArr = [];
        for (var i = 0; i < s.attendees.length; i++) {
          var a = s.attendees[i];
          if (!a.tpPeopleId) continue;
          if (!includeAllMembers && !a.isConfirmed) continue;
          if (seen[a.tpPeopleId]) continue;
          seen[a.tpPeopleId] = 1;
          toSync.push(a.tpPeopleId);
          if (a.isConfirmed) confirmedCount++;
          peopleJsonArr.push({
            tpPeopleId: a.tpPeopleId,
            isConfirmed: !!a.isConfirmed,
            pcoFirstName: a.pcoFirstName || '',
            pcoLastName: a.pcoLastName || '',
            pcoEmail: a.email || '',
          });
        }
        if (!toSync.length) {
          toast(includeAllMembers
                ? 'No matched team members to sync.'
                : 'No Confirmed + matched people to sync.', 'err');
          return;
        }
        var mode = pcoSyncMode(s.planInfo);
        // Confirm prompt tells the user exactly what mode is about to run.
        var confirmMsg = '';
        if (s.isPeopleSync) {
          var teamCount = 0;
          for (var pi = 0; pi < s.attendees.length; pi++) {
            var pa = s.attendees[pi];
            if (!pa.tpPeopleId) continue;
            for (var pj = 0; pj < (pa.positions || []).length; pj++) teamCount++;
          }
          var rdc1 = s.planInfo.rosterDropCount || 0;
          var sdc1 = s.planInfo.subgroupDropCount || 0;
          confirmMsg = 'Sync the umbrella roster across all teams in this service type?\n\n' +
                       '  - Add ' + toSync.length + ' people to the umbrella involvement (skip if already active)\n' +
                       (s.planInfo.teamsAsSubgroups ? '  - Write ~' + teamCount + ' team-subgroup association(s)\n' : '') +
                       (rdc1 ? '  - Remove ' + rdc1 + ' TP member(s) whose PCO link is no longer in scope\n' : '') +
                       (sdc1 ? '  - Remove ' + sdc1 + ' stale team-subgroup association(s)\n' : '') +
                       '\nPCO is the source of truth. No meeting, no attendance writes.';
        } else if (s.isTeamSync) {
          // Count distinct position subgroups about to be written.
          var posCount = 0;
          for (var pi = 0; pi < s.attendees.length; pi++) {
            var pa = s.attendees[pi];
            if (!pa.tpPeopleId) continue;
            for (var pj = 0; pj < (pa.positions || []).length; pj++) posCount++;
          }
          var rdc2 = s.planInfo.rosterDropCount || 0;
          var sdc2 = s.planInfo.subgroupDropCount || 0;
          confirmMsg = 'Sync the durable PCO team roster to TouchPoint?\n\n' +
                       '  - Add ' + toSync.length + ' team member(s) to the involvement (skip if already active)\n' +
                       (s.planInfo.positionsAsSubgroups ? '  - Write ~' + posCount + ' position-subgroup association(s)\n' : '') +
                       (rdc2 ? '  - Remove ' + rdc2 + ' TP member(s) whose PCO link is no longer on this team\n' : '') +
                       (sdc2 ? '  - Remove ' + sdc2 + ' stale position-subgroup association(s)\n' : '') +
                       '\nPCO is the source of truth. No meeting, no attendance writes.';
        } else if (s.isRosterSync) {
          confirmMsg = 'Add ' + toSync.length + ' team member(s) to the involvement roster?\n\n' +
                       'This pulls every matched team member (Confirmed, Unconfirmed, even Declined) across ' +
                       (s.planInfo.planCount || 0) +
                       ' plan(s) in the window -- team membership is independent of any single plan\'s RSVP. ' +
                       'No meeting will be created and no attendance written.';
        } else if (s.planInfo.syncAttendance && s.planInfo.autoAddMember) {
          confirmMsg = 'For this plan\'s ' + toSync.length + ' matched team member(s):\n\n' +
                       '  - Mark ' + confirmedCount + ' Confirmed as Present at the meeting\n' +
                       '  - Add all ' + toSync.length + ' as members of the involvement (if not already on the roster)\n\n' +
                       'A TouchPoint meeting will be created on ' + (s.planInfo.planDateIso || '?') +
                       ' at the involvement\'s scheduled time if one doesn\'t already exist.';
        } else if (s.planInfo.syncAttendance) {
          confirmMsg = 'Mark ' + toSync.length + ' Confirmed people Present at this plan\'s meeting?\n\n' +
                       'A TouchPoint meeting will be created on ' + (s.planInfo.planDateIso || '?') +
                       ' at the involvement\'s scheduled time if one doesn\'t already exist. ' +
                       'Membership will NOT be changed.';
        } else {
          confirmMsg = 'Add ' + toSync.length + ' team member(s) to the involvement roster?\n\n' +
                       'Includes Confirmed, Unconfirmed, and Declined -- they\'re all on the team. ' +
                       'No meeting will be created and no attendance will be written.';
        }
        if (!confirm(confirmMsg)) return;
        var btn = $('pcoSyncNowBtn');
        var origBtnText = mode.btnText;
        btn.disabled = true; btn.textContent = 'Syncing...';
        var syncAction;
        if (s.isPeopleSync) syncAction = 'sync_people';
        else if (s.isTeamSync) syncAction = 'sync_team';
        else if (s.isRosterSync) syncAction = 'sync_roster';
        else syncAction = 'sync_plan_attendance';
        var peopleJsonStr = JSON.stringify(peopleJsonArr);
        // For Team Sync / People Sync we bundle a per-person subgroup map
        // (positions or teams respectively) so the server can write
        // subgroup memberships.
        var subgroupJsonStr = '';
        var wantSubgroups = (s.isTeamSync && s.planInfo.positionsAsSubgroups) ||
                            (s.isPeopleSync && s.planInfo.teamsAsSubgroups);
        if (wantSubgroups) {
          var subgroupsByPid = {};
          for (var qi = 0; qi < s.attendees.length; qi++) {
            var qa = s.attendees[qi];
            if (!qa.tpPeopleId) continue;
            var names = [];
            for (var qj = 0; qj < (qa.positions || []).length; qj++) {
              var pos = qa.positions[qj];
              // Accept both keys -- per-plan + Team Sync previews now
              // both emit 'teamPosition'; old payloads used 'position'.
              var nm = (pos && (pos.teamPosition || pos.position)) || '';
              if (nm) names.push(nm);
            }
            if (names.length) subgroupsByPid[qa.tpPeopleId] = names;
          }
          subgroupJsonStr = JSON.stringify(subgroupsByPid);
        }
        var syncPayload;
        if (s.isPeopleSync) {
          syncPayload = {
            pco_service_type_id: s.serviceTypeId || '',
            tp_org_id: s.orgId,
            people_ids: toSync.join(','),
            people_json: peopleJsonStr,
            teams_json: subgroupJsonStr,
          };
        } else if (s.isTeamSync) {
          syncPayload = {
            pco_team_id: s.pcoTeamId || '',
            tp_org_id: s.orgId,
            people_ids: toSync.join(','),
            people_json: peopleJsonStr,
            positions_json: subgroupJsonStr,
          };
        } else if (s.isRosterSync) {
          syncPayload = {
            service_type_id: s.serviceTypeId || '',
            tp_org_id: s.orgId,
            people_ids: toSync.join(','),
            people_json: peopleJsonStr,
          };
        } else {
          syncPayload = {
            plan_id: s.planId,
            tp_org_id: s.orgId,
            service_type_id: s.serviceTypeId || '',
            plan_date_iso: s.planInfo.planDateIso || '',
            people_ids: toSync.join(','),
            people_json: peopleJsonStr,
          };
        }
        ajax(syncAction, syncPayload, function(err, d){
          btn.disabled = false; btn.textContent = origBtnText;
          if (err || !d || !d.success) {
            toast('Sync failed: ' + ((d && d.message) || err), 'err');
            return;
          }
          var msg = d.message || 'Sync complete.';
          if (d.verifyWarning) msg += ' Warning: ' + d.verifyWarning;
          toast(msg, d.verifyWarning ? 'info' : 'ok');
          setTimeout(closePlanPreview, 1500);
        });
      }

      function renderMappingsTab() {
        $('pcoContent').innerHTML = ''
          // Direction indicator banner -- v2.6+ makes the one-way nature
          // unmistakable. Every mapping pushes PCO state into TouchPoint;
          // TP changes never flow back to PCO (that's v3.x territory).
          + '<div style="margin-bottom:12px;padding:8px 12px;background:#0f7c840d;border-left:3px solid #0f7c84;border-radius:4px;font-size:13px;">'
          + '<strong>One-way sync:</strong> PCO &rarr; TouchPoint only. PCO is the source of truth; TouchPoint follows. '
          + 'Changes you make in TouchPoint will NOT flow back to PCO. <span class="pco-muted">(two-way sync is on the roadmap)</span>'
          + '</div>'
          // ----- All People Mapping (v2.2) -- singleton, broadest scope -----
          + '<div class="pco-card">'
          + '<h3>All People <span style="background:#0f7c84;color:#fff;font-size:10px;font-weight:700;letter-spacing:0.5px;padding:2px 6px;border-radius:3px;margin-left:6px;">PCO &rarr; TP</span> <span class="pco-muted" style="font-size:13px;font-weight:400;">(entire PCO directory &rarr; one TP involvement)</span></h3>'
          + '<div class="pco-help">'
          + '<strong>Singleton mapping.</strong> Every person in the PCO People app (the global directory) becomes a member of one designated TP involvement. '
          + '<strong>TouchPoint stays authoritative</strong> -- if a PCO record has no TP match by PCO_PersonId or email, '
          + 'it is reported as unmatched but a TP person is never created. Use when you want a single "PCO People" '
          + 'involvement in TP that mirrors who PCO knows about.'
          + '</div>'
          + '<div id="pcoAllPeopleMapping"></div>'
          + '</div>'
          // ----- Service Type Mappings (v3.0 unified) -----
          + '<div class="pco-card" style="margin-top:14px;">'
          + '<h3>Service Type Mappings <span style="background:#0f7c84;color:#fff;font-size:10px;font-weight:700;letter-spacing:0.5px;padding:2px 6px;border-radius:3px;margin-left:6px;">PCO &rarr; TP</span> <span class="pco-muted" style="font-size:13px;font-weight:400;">(roster + optional team subgroups + optional attendance)</span></h3>'
          + '<div class="pco-help">'
          + 'Each PCO Service Type maps to ONE TouchPoint involvement. The durable roster syncs automatically -- '
          + 'every person on any team in the service type becomes a member of the involvement. '
          + 'Two optional add-on layers per mapping: '
          + '<strong>Teams as subgroups</strong> (each Team becomes a TP subgroup like "Vocals" / "Production") and '
          + '<strong>Per-plan attendance</strong> (walks each PCO Plan in the window, writes Attendance for Confirmed people to the meeting on that plan\'s date).'
          + '</div>'
          + '<div style="display:flex;gap:8px;margin-bottom:10px;">'
          + '  <button class="pco-btn" id="pcoAddPeopleMappingBtn" style="font-size:13px;padding:5px 12px;">+ Add Service Type Mapping</button>'
          + '  <a href="#" id="pcoPeopleMappingsRefresh" style="align-self:center;font-size:13px;">Refresh</a>'
          + '</div>'
          + '<div id="pcoPeopleMappingPicker" style="display:none;border:1px solid #e1e4e8;border-radius:6px;padding:10px;margin-bottom:10px;background:#fff;"></div>'
          + '<div id="pcoPeopleMappingsList" class="pco-empty">Loading service type mappings...</div>'
          + '</div>'
          // ----- Team Mappings (v2.0) -----
          + '<div class="pco-card" style="margin-top:14px;">'
          + '<h3>Team Mappings <span style="background:#0f7c84;color:#fff;font-size:10px;font-weight:700;letter-spacing:0.5px;padding:2px 6px;border-radius:3px;margin-left:6px;">PCO &rarr; TP</span> <span class="pco-muted" style="font-size:13px;font-weight:400;">(durable roster + position subgroups)</span></h3>'
          + '<div class="pco-help">'
          + 'Each PCO Team (a recurring team like "Vocals" under a Service Type) maps to ONE TouchPoint '
          + 'involvement. Team Sync writes the durable roster -- people added to the PCO team become members '
          + 'of the TP involvement, and their position eligibility becomes subgroups inside the involvement. '
          + 'No attendance is written. Use this when you want the worship team\'s involvement roster to stay in '
          + 'lockstep with PCO without per-Sunday attendance entries.'
          + '</div>'
          + '<div style="display:flex;gap:8px;margin-bottom:10px;">'
          + '  <button class="pco-btn" id="pcoAddTeamMappingBtn" style="font-size:13px;padding:5px 12px;">+ Add Team Mapping</button>'
          + '  <a href="#" id="pcoTeamMappingsRefresh" style="align-self:center;font-size:13px;">Refresh</a>'
          + '</div>'
          + '<div id="pcoTeamMappingPicker" style="display:none;border:1px solid #e1e4e8;border-radius:6px;padding:10px;margin-bottom:10px;background:#fff;"></div>'
          + '<div id="pcoTeamMappingsList" class="pco-empty">Loading team mappings...</div>'
          + '</div>';
        // v3.0: Service Plan Mappings card removed (attendance is now a
        // toggle on Service Type Mappings). Stubs left for back-compat
        // with any cached DOM.
        $('pcoTeamMappingsRefresh').onclick = function(ev){ ev.preventDefault(); loadTeamMappings(); };
        $('pcoAddTeamMappingBtn').onclick = openTeamMappingPicker;
        $('pcoPeopleMappingsRefresh').onclick = function(ev){ ev.preventDefault(); loadPeopleMappings(); };
        $('pcoAddPeopleMappingBtn').onclick = openPeopleMappingPicker;
        loadAllPeopleMapping();
        loadPeopleMappings();
        loadTeamMappings();
      }

      // Add-mapping picker for Service Type Mappings. Mirrors the Team /
      // People Mappings UX -- pick from a dropdown of unmapped service
      // types, search a TP involvement, save. No inline unmapped list
      // is rendered in the main view, which keeps the section tight.
      function openServiceTypeMappingPicker() {
        var box = $('pcoMappingPicker');
        box.style.display = '';
        box.innerHTML = '<div class="pco-muted">Loading service types...</div>';
        ajax('list_service_types', {}, function(err, d){
          if (err || !d || !d.success) {
            box.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var sts = (d.serviceTypes || []).filter(function(s){ return !s.mappedOrgId; });
          // Server returns mapped-first then alpha. The picker is unmapped-only,
          // so just sort alpha by name to keep the dropdown predictable.
          sts.sort(function(a, b){
            return (a.pcoName || '').toLowerCase().localeCompare((b.pcoName || '').toLowerCase());
          });
          if (!sts.length) {
            box.innerHTML = '<div class="pco-muted">All PCO service types are already mapped. Nice work.</div>'
              + '<div style="margin-top:8px;"><button class="pco-btn pco-secondary" id="pcoStPickClose" style="font-size:13px;padding:4px 10px;">Close</button></div>';
            $('pcoStPickClose').onclick = function(){
              box.style.display = 'none';
              box.innerHTML = '';
            };
            return;
          }
          var opts = '<option value="">-- pick a service type --</option>';
          for (var i = 0; i < sts.length; i++) {
            opts += '<option value="' + escAttr(sts[i].pcoId) + '" data-name="' + escAttr(sts[i].pcoName) + '">' + escHtml(sts[i].pcoName) + '</option>';
          }
          box.innerHTML = ''
            + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">'
            + '  <label style="font-size:13px;font-weight:600;">Unmapped Service Type:</label>'
            + '  <select id="pcoStPickStSel" style="padding:4px 6px;">' + opts + '</select>'
            + '  <span class="pco-muted" style="font-size:12px;">' + sts.length + ' available</span>'
            + '  <button class="pco-btn pco-secondary" id="pcoStPickCancel" style="font-size:13px;padding:4px 10px;margin-left:auto;">Cancel</button>'
            + '</div>'
            + '<div id="pcoStPickOrgWrap" style="display:none;">'
            + '  <label class="pco-label">TouchPoint involvement</label>'
            + '  <input type="text" class="pco-input" id="pcoStPickOrgSearch" placeholder="Search by name or org id...">'
            + '  <div id="pcoStPickOrgResults" style="margin-top:6px;"></div>'
            + '</div>';
          $('pcoStPickCancel').onclick = function(){
            box.style.display = 'none';
            box.innerHTML = '';
          };
          $('pcoStPickStSel').onchange = function(){
            if (!this.value) { $('pcoStPickOrgWrap').style.display = 'none'; return; }
            $('pcoStPickOrgWrap').style.display = '';
            $('pcoStPickOrgSearch').focus();
          };
          $('pcoStPickOrgSearch').oninput = function(){
            var term = this.value.trim();
            var resultsEl = $('pcoStPickOrgResults');
            if (term.length < 2) {
              resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
              return;
            }
            resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
            clearTimeout(window.__pcoStPickSearchT);
            window.__pcoStPickSearchT = setTimeout(function(){
              ajax('search_involvements', {search_term: term}, function(err, d){
                if (err || !d || !d.success) { resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>'; return; }
                var orgs = d.orgs || [];
                if (!orgs.length) { resultsEl.innerHTML = '<span class="pco-muted">No matches.</span>'; return; }
                var html = '<div style="border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:240px;overflow-y:auto;">';
                for (var i = 0; i < orgs.length; i++) {
                  var o = orgs[i];
                  html += '<div class="pco-st-pick-org" data-org-id="' + o.orgId + '" data-org-name="' + escAttr(o.orgName) + '" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;">'
                    + '<strong>' + escHtml(o.orgName) + '</strong> '
                    + '<span class="pco-muted">(#' + o.orgId
                    + (o.divisionName ? ' &middot; ' + escHtml(o.divisionName) : '')
                    + (o.programName ? ' / ' + escHtml(o.programName) : '')
                    + ')</span>'
                    + '</div>';
                }
                html += '</div>';
                resultsEl.innerHTML = html;
                var picks = resultsEl.querySelectorAll('.pco-st-pick-org');
                for (var j = 0; j < picks.length; j++) {
                  picks[j].onclick = function(){
                    var orgId = this.dataset.orgId;
                    var orgName = this.dataset.orgName;
                    var stSel = $('pcoStPickStSel');
                    var pcoId = stSel.value;
                    var stName = stSel.options[stSel.selectedIndex].dataset.name || '';
                    saveOrgMapping(pcoId, parseInt(orgId, 10) || 0);
                    setTimeout(function(){
                      $('pcoMappingPicker').style.display = 'none';
                      $('pcoMappingPicker').innerHTML = '';
                    }, 100);
                  };
                }
              });
            }, 250);
          };
        });
      }

      // ----- All People Mapping logic -----------------------------

      function loadAllPeopleMapping() {
        var host = $('pcoAllPeopleMapping');
        host.innerHTML = '<div class="pco-muted">Loading...</div>';
        ajax('load_all_people_mapping', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var m = d.mapping || {orgId: 0};
          if (!m.orgId) {
            host.innerHTML = ''
              + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">'
              + '  <span class="pco-pill pco-warn">Not configured</span>'
              + '  <button class="pco-btn" id="pcoAllPeopleConfigBtn" style="font-size:13px;padding:5px 12px;">Configure</button>'
              + '</div>'
              + '<div id="pcoAllPeopleConfig" style="display:none;margin-top:10px;padding:10px;border:1px solid #e1e4e8;border-radius:6px;background:#fff;"></div>';
            $('pcoAllPeopleConfigBtn').onclick = openAllPeopleConfig;
            return;
          }
          var orgCtx = '';
          if (d.tpDivision) orgCtx = ' <span class="pco-muted">&middot; ' + escHtml(d.tpDivision) + (d.tpProgram ? ' / ' + escHtml(d.tpProgram) : '') + '</span>';
          host.innerHTML = ''
            + '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">'
            + '<div style="flex:1 1 320px;min-width:0;">'
            + '<div style="font-weight:700;color:#1f4e79;">PCO People directory &rarr; <span style="font-weight:400;">' + escHtml(d.tpOrgName) + '</span> <span class="pco-muted">(#' + m.orgId + ')</span>' + orgCtx + '</div>'
            + '<div style="margin-top:4px;font-size:12px;">'
            +   '<span class="pco-pill" style="background:#1f4e79;color:#fff;font-size:11px;">All People Sync</span> '
            +   (m.includeInactive ? '<span class="pco-pill pco-warn" style="font-size:11px;">incl. inactive</span>' : '<span class="pco-pill" style="background:#eef0f2;color:#666;font-size:11px;">active only</span>')
            + '</div>'
            + '</div>'
            + '<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;justify-content:flex-end;">'
            + '<button class="pco-btn" id="pcoAllPeoplePreviewBtn" style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
            + '<button class="pco-btn pco-secondary" id="pcoAllPeopleReconfigBtn" style="font-size:13px;padding:5px 12px;">Change</button>'
            + '<button class="pco-btn pco-secondary" id="pcoAllPeopleDelBtn" style="font-size:13px;padding:5px 12px;color:#a00;border-color:#a00;">Remove</button>'
            + '</div>'
            + '</div>'
            + '<div id="pcoAllPeopleConfig" style="display:none;margin-top:10px;padding:10px;border:1px solid #e1e4e8;border-radius:6px;background:#fff;"></div>';
          $('pcoAllPeoplePreviewBtn').onclick = openAllPeoplePreview;
          $('pcoAllPeopleReconfigBtn').onclick = openAllPeopleConfig;
          $('pcoAllPeopleDelBtn').onclick = function(){
            if (!confirm('Remove the All People mapping? Existing TP members will NOT be removed.')) return;
            ajax('delete_all_people_mapping', {}, function(err, d){
              if (err || !d || !d.success) { toast('Delete failed: ' + ((d && d.message) || err), 'err'); return; }
              toast('Removed.', 'ok');
              loadAllPeopleMapping();
            });
          };
        });
      }

      function openAllPeopleConfig() {
        var box = $('pcoAllPeopleConfig');
        if (!box) return;
        box.style.display = '';
        box.innerHTML = ''
          + '<label class="pco-label">TouchPoint involvement</label>'
          + '<input type="text" class="pco-input" id="pcoAllPeopleOrgSearch" placeholder="Search by name or org id...">'
          + '<div id="pcoAllPeopleOrgResults" style="margin-top:6px;"></div>'
          + '<div style="margin-top:8px;display:flex;align-items:center;gap:6px;">'
          + '  <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;">'
          + '    <input type="checkbox" id="pcoAllPeopleIncludeInactive">'
          + '    Include inactive PCO records'
          + '  </label>'
          + '  <button class="pco-btn pco-secondary" id="pcoAllPeopleCancelBtn" style="margin-left:auto;font-size:13px;padding:4px 10px;">Cancel</button>'
          + '</div>';
        $('pcoAllPeopleCancelBtn').onclick = function(){
          box.style.display = 'none';
          box.innerHTML = '';
        };
        $('pcoAllPeopleOrgSearch').oninput = function(){
          var term = this.value.trim();
          var resultsEl = $('pcoAllPeopleOrgResults');
          if (term.length < 2) {
            resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
            return;
          }
          resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
          clearTimeout(window.__pcoAllPeopleSearchT);
          window.__pcoAllPeopleSearchT = setTimeout(function(){
            ajax('search_involvements', {search_term: term}, function(err, d){
              if (err || !d || !d.success) { resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>'; return; }
              var orgs = d.orgs || [];
              if (!orgs.length) { resultsEl.innerHTML = '<span class="pco-muted">No matches.</span>'; return; }
              var html = '<div style="border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:240px;overflow-y:auto;">';
              for (var i = 0; i < orgs.length; i++) {
                var o = orgs[i];
                html += '<div class="pco-all-people-pick-org" data-org-id="' + o.orgId + '" data-org-name="' + escAttr(o.orgName) + '" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;">'
                  + '<strong>' + escHtml(o.orgName) + '</strong> '
                  + '<span class="pco-muted">(#' + o.orgId
                  + (o.divisionName ? ' &middot; ' + escHtml(o.divisionName) : '')
                  + (o.programName ? ' / ' + escHtml(o.programName) : '')
                  + ')</span>'
                  + '</div>';
              }
              html += '</div>';
              resultsEl.innerHTML = html;
              var picks = resultsEl.querySelectorAll('.pco-all-people-pick-org');
              for (var j = 0; j < picks.length; j++) {
                picks[j].onclick = function(){
                  var orgId = this.dataset.orgId;
                  var orgName = this.dataset.orgName;
                  var includeInactive = $('pcoAllPeopleIncludeInactive').checked ? '1' : '0';
                  ajax('save_all_people_mapping', {tp_org_id: orgId, include_inactive: includeInactive}, function(err, d){
                    if (err || !d || !d.success) { toast('Save failed: ' + ((d && d.message) || err), 'err'); return; }
                    toast('Mapped PCO People directory to ' + orgName + '.', 'ok');
                    loadAllPeopleMapping();
                  });
                };
              }
            });
          }, 250);
        };
      }

      // All People preview is a counts-only modal. Previewing 5K+ rows in
      // a list is impractical; the user sees totals + a sample of
      // unmatched names + a Sync Now button.
      function openAllPeoplePreview() {
        renderPlanPreviewModal('<div class="pco-empty">Walking the entire PCO People directory (this can take 10-30 seconds for large churches)...</div>');
        ajax('load_all_people_preview', {}, function(err, d){
          if (err || !d || !d.success) {
            renderPlanPreviewModal('<div class="pco-pill pco-err">Error</div><div>' + escHtml((d && d.message) || err) + '</div>');
            return;
          }
          $('pcoModalTitle').textContent = 'All People Sync — ' + (d.totalPco || 0) + ' PCO record(s)';
          var body = $('pcoModalBody');
          var sample = d.unmatchedSample || [];
          var sampleHtml = '';
          // Banner pointing users to the bulk matcher when there's a
          // meaningful unmatched residual. Without this they're stuck
          // looking at a static list with no way to act on it.
          var unmatchedTotal = d.unmatched || 0;
          if (unmatchedTotal > 0) {
            sampleHtml += '<div style="margin-top:12px;padding:8px 12px;background:#0f7c840d;border-left:3px solid #0f7c84;border-radius:4px;font-size:13px;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">'
              + '<span><strong>Need to resolve these?</strong> The Proposed Matches review (People Matching tab) scores TP candidates for every unmatched PCO record (name + email + birthdate signals). Apply or Skip Forever per row, with bulk Apply for high-confidence matches.</span>'
              + '<button class="pco-btn" id="pcoOpenBulkMatcher" style="font-size:13px;padding:5px 12px;white-space:nowrap;">Open Proposed Matches &rarr;</button>'
              + '</div>';
          }
          if (sample.length) {
            sampleHtml += '<div style="margin-top:10px;">'
              + '<strong style="font-size:13px;">Unmatched sample (first ' + sample.length + ')</strong>'
              + '<div style="margin-top:6px;border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:280px;overflow-y:auto;">';
            for (var i = 0; i < sample.length; i++) {
              var u = sample[i];
              sampleHtml += '<div style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;">'
                + '<strong>' + escHtml(u.name) + '</strong>'
                + (u.email ? ' <span class="pco-muted">' + escHtml(u.email) + '</span>' : '')
                + ' <span class="pco-muted">PCO #' + escHtml(u.pcoPersonId) + '</span>'
                + '</div>';
            }
            sampleHtml += '</div></div>';
          }
          var warnHtml = '';
          if (d.warnings && d.warnings.length) {
            warnHtml = '<div class="pco-card" style="background:#fff8e0;border-color:#f0c870;margin-top:10px;padding:8px 12px;font-size:13px;">'
              + '<strong>Warnings:</strong><ul style="margin:4px 0 0 20px;">';
            for (var w = 0; w < d.warnings.length; w++) warnHtml += '<li>' + escHtml(d.warnings[w]) + '</li>';
            warnHtml += '</ul></div>';
          }
          // v3.2: mirror-removal banner for All People Sync.
          var apDropCount = d.rosterDropCount || 0;
          var dropBanner = '';
          if (apDropCount > 0) {
            dropBanner = '<div style="margin-bottom:10px;padding:8px 10px;border-radius:4px;font-size:13px;background:#fff7e6;border-left:3px solid #c0392b;color:#8a2020;">'
              + '<strong>Mirror removal:</strong> PCO is the source of truth -- <strong>' + apDropCount + ' TP member(s)</strong> will be removed from the involvement (their PCO link is no longer in the PCO directory).'
              + '</div>';
          }
          body.innerHTML = ''
            + '<div style="margin-bottom:10px;padding:8px 10px;border-radius:4px;background:#1f4e790d;border-left:3px solid #1f4e79;font-size:13px;">'
            + '  <strong style="color:#1f4e79;">Sync mode:</strong> All People. Every matched PCO person will be added to the designated TP involvement. '
            + '  Unmatched PCO records will be reported but no TP person will be created (TP is authoritative).'
            + '</div>'
            + dropBanner
            + '<div style="margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap;">'
            + '  <span class="pco-pill pco-ok">' + (d.matched || 0) + ' matched</span>'
            + '  <span class="pco-pill pco-warn">' + (d.unmatched || 0) + ' unmatched</span>'
            + ((d.ambiguous || 0) > 0 ? '  <span class="pco-pill pco-warn">' + d.ambiguous + ' ambiguous email</span>' : '')
            + (apDropCount > 0 ? '  <span class="pco-pill" style="background:#c0392b;color:#fff;">' + apDropCount + ' will be removed</span>' : '')
            + '  <span class="pco-muted">of ' + (d.totalPco || 0) + ' total PCO people</span>'
            + '</div>'
            + sampleHtml
            + warnHtml;
          $('pcoModalFooter').style.display = 'flex';
          $('pcoModalFooter').innerHTML = ''
            + '<button class="pco-btn pco-secondary" onclick="window.__pcoCloseModal()">Cancel</button>'
            + '<button class="pco-btn" id="pcoAllPeopleSyncBtn">Sync ' + (d.matched || 0) + ' matched PCO people</button>';
          window.__pcoCloseModal = closePlanPreview;
          // Wire the "Open Proposed Matches" button if it rendered.
          // The All People preview goes for full directory mode (no
          // scope) since the user has clearly opted to see everything.
          var bulkBtn = $('pcoOpenBulkMatcher');
          if (bulkBtn) bulkBtn.onclick = function(){
            _proposedScope = null;
            _proposedState = {page: 1, tier: 'all', search: ''};
            closePlanPreview();
            selectTab('people');
            setTimeout(function(){
              var loadBtn = document.getElementById('pcoLoadProposedBtn');
              if (loadBtn) loadBtn.click();
            }, 100);
          };
          $('pcoAllPeopleSyncBtn').onclick = function(){
            var confirmMsg = 'Sync ' + (d.matched || 0) + ' matched PCO people into the TP involvement?\n\n'
              + 'Unmatched (' + (d.unmatched || 0) + ') will be reported, not created. ';
            if (apDropCount > 0) {
              confirmMsg += '\n\nMirror removal: ' + apDropCount + ' TP member(s) whose PCO link is no longer in the PCO directory will be REMOVED from the involvement. ';
            }
            confirmMsg += '\n\nContinue?';
            if (!confirm(confirmMsg)) return;
            var btn = this;
            btn.disabled = true; btn.textContent = 'Syncing...';
            ajax('sync_all_people', {}, function(err, d){
              btn.disabled = false; btn.textContent = 'Sync matched PCO people';
              if (err || !d || !d.success) {
                toast('Sync failed: ' + ((d && d.message) || err), 'err');
                return;
              }
              toast(d.message || 'Sync complete.', 'ok');
              setTimeout(closePlanPreview, 1500);
            });
          };
        });
      }

      // ----- People Mappings logic --------------------------------

      function loadPeopleMappings() {
        var host = $('pcoPeopleMappingsList');
        host.className = 'pco-empty';
        host.innerHTML = 'Loading people mappings...';
        ajax('list_people_mappings', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var rows = d.mappings || [];
          if (!rows.length) {
            host.innerHTML = '<div>No people mappings yet. Click <strong>+ Add People Mapping</strong> to map a PCO Service Type to a TouchPoint umbrella involvement.</div>';
            return;
          }
          host.className = '';
          var html = '<div class="pco-muted" style="margin-bottom:8px;">' + rows.length + ' people mapping(s)</div>'
                   + '<div style="display:flex;flex-direction:column;gap:8px;">';
          var schedInstalled = !!d.schedulerInstalled;
          for (var i = 0; i < rows.length; i++) {
            html += renderPeopleMappingRow(rows[i], schedInstalled);
          }
          html += '</div>';
          host.innerHTML = html;
          wireSchedulePanels(host);
          if (!host.__pcoPeopleWired) {
            host.__pcoPeopleWired = true;
            host.addEventListener('click', function(ev){
              var sync = ev.target.closest && ev.target.closest('.pco-people-sync-btn');
              var del  = ev.target.closest && ev.target.closest('.pco-people-del-btn');
              if (sync) {
                openPeopleSyncPreview(sync.dataset.pcoStId, sync.dataset.tpOrgId);
              } else if (del) {
                if (!confirm('Remove this people mapping? Existing TP members and subgroups will NOT be removed.')) return;
                ajax('delete_people_mapping', {pco_service_type_id: del.dataset.pcoStId}, function(err, d){
                  if (err || !d || !d.success) { toast('Delete failed: ' + ((d && d.message) || err), 'err'); return; }
                  toast('Removed.', 'ok');
                  loadPeopleMappings();
                });
              }
            });
            // v3.0 per-mapping toggles: teams-as-subgroups + per-plan
            // attendance. Save both on every flip; server preserves the
            // unchanged one.
            host.addEventListener('change', function(ev){
              var t = ev.target;
              if (!t.classList) return;
              var isSubg = t.classList.contains('pco-st-opt-subgroups');
              var isAtt  = t.classList.contains('pco-st-opt-attendance');
              if (!isSubg && !isAtt) return;
              var row = t.closest('.pco-people-map-row');
              if (!row) return;
              var subgBox = row.querySelector('.pco-st-opt-subgroups');
              var attBox  = row.querySelector('.pco-st-opt-attendance');
              var status  = row.querySelector('.pco-st-opt-status');
              if (status) { status.textContent = 'Saving...'; status.style.color = ''; }
              ajax('set_people_mapping_options', {
                pco_service_type_id: t.dataset.pcoStId,
                teams_as_subgroups: subgBox && subgBox.checked ? '1' : '0',
                per_plan_attendance: attBox && attBox.checked ? '1' : '0',
              }, function(err, d){
                if (err || !d || !d.success) {
                  if (status) { status.textContent = 'Save failed: ' + ((d && d.message) || err); status.style.color = '#a00'; }
                  return;
                }
                if (status) {
                  status.textContent = 'Saved.';
                  status.style.color = '#1f6b3a';
                  setTimeout(function(){ if (status.textContent === 'Saved.') status.textContent = ''; }, 1500);
                }
              });
            });
          }
        });
      }

      function renderPeopleMappingRow(r, schedulerInstalled) {
        var orgCtx = '';
        if (r.tpDivision) orgCtx = '<span class="pco-muted"> &middot; ' + escHtml(r.tpDivision) + (r.tpProgram ? ' / ' + escHtml(r.tpProgram) : '') + '</span>';
        var subgChk = r.teamsAsSubgroups ? ' checked' : '';
        var attChk  = r.perPlanAttendance ? ' checked' : '';
        return ''
          + '<div class="pco-people-map-row" style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;">'
          + '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;color:#1f4e79;">' + escHtml(r.pcoServiceTypeName || ('Service Type ' + r.pcoServiceTypeId)) + '</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">'
          +   '<strong>TP:</strong> ' + escHtml(r.tpOrgName) + ' <span class="pco-muted">(#' + r.tpOrgId + ')</span>' + orgCtx
          + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<span class="pco-pill" style="background:#1f4e79;color:#fff;font-size:11px;">Service Type Sync</span>'
          + '<button class="pco-btn pco-people-sync-btn" data-pco-st-id="' + escAttr(r.pcoServiceTypeId) + '" data-tp-org-id="' + r.tpOrgId + '" style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
          + '<button class="pco-btn pco-secondary pco-people-del-btn" data-pco-st-id="' + escAttr(r.pcoServiceTypeId) + '" style="font-size:13px;padding:5px 12px;color:#a00;border-color:#a00;">Remove</button>'
          + '</div>'
          + '</div>'
          // v3.0: per-mapping toggles inline. Roster sync is always on
          // (it's the whole point), but subgroups + attendance are
          // optional layers staff can flip on/off without re-saving the
          // involvement.
          + '<div style="margin-top:8px;padding-top:8px;border-top:1px dashed #e1e4e8;display:flex;gap:16px;flex-wrap:wrap;font-size:13px;">'
          + '  <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
          + '    <input type="checkbox" class="pco-st-opt-subgroups" data-pco-st-id="' + escAttr(r.pcoServiceTypeId) + '"' + subgChk + '>'
          + '    Teams as subgroups <span class="pco-muted">(Vocals, Production, etc.)</span>'
          + '  </label>'
          + '  <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
          + '    <input type="checkbox" class="pco-st-opt-attendance" data-pco-st-id="' + escAttr(r.pcoServiceTypeId) + '"' + attChk + '>'
          + '    Per-plan attendance <span class="pco-muted">(mark Confirmed as Present on each plan-date meeting)</span>'
          + '  </label>'
          + '  <span class="pco-st-opt-status pco-muted" data-pco-st-id="' + escAttr(r.pcoServiceTypeId) + '" style="font-size:12px;align-self:center;"></span>'
          + '</div>'
          // v3.3: schedule editor.
          + renderSchedulePanel('people', r.pcoServiceTypeId, r.schedule, r.scheduleLabel, r.nextRunIso, schedulerInstalled)
          + '</div>';
      }

      function openPeopleMappingPicker() {
        var box = $('pcoPeopleMappingPicker');
        box.style.display = '';
        box.innerHTML = '<div class="pco-muted">Loading service types...</div>';
        ajax('list_service_types', {}, function(err, d){
          if (err || !d || !d.success) {
            box.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var sts = (d.serviceTypes || []).slice();
          sts.sort(function(a, b){
            return (a.pcoName || '').toLowerCase().localeCompare((b.pcoName || '').toLowerCase());
          });
          var opts = '<option value="">-- pick a service type --</option>';
          for (var i = 0; i < sts.length; i++) {
            opts += '<option value="' + escAttr(sts[i].pcoId) + '" data-name="' + escAttr(sts[i].pcoName) + '">' + escHtml(sts[i].pcoName) + '</option>';
          }
          box.innerHTML = ''
            + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">'
            + '  <label style="font-size:13px;font-weight:600;">Service Type:</label>'
            + '  <select id="pcoPeoplePickStSel" style="padding:4px 6px;">' + opts + '</select>'
            + '  <button class="pco-btn pco-secondary" id="pcoPeoplePickCancel" style="font-size:13px;padding:4px 10px;margin-left:auto;">Cancel</button>'
            + '</div>'
            + '<div id="pcoPeoplePickOrgWrap" style="display:none;">'
            + '  <label class="pco-label">TouchPoint umbrella involvement</label>'
            + '  <input type="text" class="pco-input" id="pcoPeoplePickOrgSearch" placeholder="Search by name or org id...">'
            + '  <div id="pcoPeoplePickOrgResults" style="margin-top:6px;"></div>'
            + '</div>';
          $('pcoPeoplePickCancel').onclick = function(){
            box.style.display = 'none';
            box.innerHTML = '';
          };
          $('pcoPeoplePickStSel').onchange = function(){
            if (!this.value) { $('pcoPeoplePickOrgWrap').style.display = 'none'; return; }
            $('pcoPeoplePickOrgWrap').style.display = '';
            $('pcoPeoplePickOrgSearch').focus();
          };
          $('pcoPeoplePickOrgSearch').oninput = function(){
            var term = this.value.trim();
            var resultsEl = $('pcoPeoplePickOrgResults');
            if (term.length < 2) {
              resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
              return;
            }
            resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
            clearTimeout(window.__pcoPeoplePickSearchT);
            window.__pcoPeoplePickSearchT = setTimeout(function(){
              ajax('search_involvements', {search_term: term}, function(err, d){
                if (err || !d || !d.success) { resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>'; return; }
                var orgs = d.orgs || [];
                if (!orgs.length) { resultsEl.innerHTML = '<span class="pco-muted">No matches.</span>'; return; }
                var html = '<div style="border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:240px;overflow-y:auto;">';
                for (var i = 0; i < orgs.length; i++) {
                  var o = orgs[i];
                  html += '<div class="pco-people-pick-org" data-org-id="' + o.orgId + '" data-org-name="' + escAttr(o.orgName) + '" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;">'
                    + '<strong>' + escHtml(o.orgName) + '</strong> '
                    + '<span class="pco-muted">(#' + o.orgId
                    + (o.divisionName ? ' &middot; ' + escHtml(o.divisionName) : '')
                    + (o.programName ? ' / ' + escHtml(o.programName) : '')
                    + ')</span>'
                    + '</div>';
                }
                html += '</div>';
                resultsEl.innerHTML = html;
                var picks = resultsEl.querySelectorAll('.pco-people-pick-org');
                for (var j = 0; j < picks.length; j++) {
                  picks[j].onclick = function(){
                    var orgId = this.dataset.orgId;
                    var orgName = this.dataset.orgName;
                    var stSel = $('pcoPeoplePickStSel');
                    var stId = stSel.value;
                    var stName = stSel.options[stSel.selectedIndex].dataset.name || '';
                    ajax('save_people_mapping', {
                      pco_service_type_id: stId,
                      pco_service_type_name: stName,
                      tp_org_id: orgId,
                    }, function(err, d){
                      if (err || !d || !d.success) { toast('Save failed: ' + ((d && d.message) || err), 'err'); return; }
                      toast('Mapped ' + stName + ' to ' + orgName + '.', 'ok');
                      $('pcoPeopleMappingPicker').style.display = 'none';
                      $('pcoPeopleMappingPicker').innerHTML = '';
                      loadPeopleMappings();
                    });
                  };
                }
              });
            }, 250);
          };
        });
      }

      function openPeopleSyncPreview(pcoStId, tpOrgId) {
        renderPlanPreviewModal('<div class="pco-empty">Loading roster across teams from PCO...</div>');
        ajax('load_people_sync_preview', {pco_service_type_id: pcoStId}, function(err, d){
          if (err || !d || !d.success) {
            renderPlanPreviewModal('<div class="pco-pill pco-err">Error</div><div>' + escHtml((d && d.message) || err) + '</div>');
            return;
          }
          _planPreviewState = {
            planId: '',
            serviceTypeId: pcoStId,
            pcoTeamId: '',
            orgId: parseInt(tpOrgId, 10) || 0,
            attendees: d.attendees || [],
            planInfo: d.planInfo || {},
            summary: d.summary || {},
            isRosterSync: false,
            isTeamSync: false,
            isPeopleSync: true,
          };
          renderPlanPreviewBody();
        });
      }

      // ----- Team Mappings logic -----------------------------------

      function loadTeamMappings() {
        var host = $('pcoTeamMappingsList');
        host.className = 'pco-empty';
        host.innerHTML = 'Loading team mappings...';
        ajax('list_team_mappings', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var rows = d.mappings || [];
          if (!rows.length) {
            host.innerHTML = '<div>No team mappings yet. Click <strong>+ Add Team Mapping</strong> to wire a PCO Team to a TouchPoint involvement.</div>';
            return;
          }
          host.className = '';
          var html = '<div class="pco-muted" style="margin-bottom:8px;">' + rows.length + ' team mapping(s)</div>'
                   + '<div style="display:flex;flex-direction:column;gap:8px;">';
          var schedInstalled = !!d.schedulerInstalled;
          for (var i = 0; i < rows.length; i++) {
            html += renderTeamMappingRow(rows[i], schedInstalled);
          }
          html += '</div>';
          host.innerHTML = html;
          wireSchedulePanels(host);
          // Wire row buttons via delegation. Guard against re-attach.
          if (!host.__pcoTeamWired) {
            host.__pcoTeamWired = true;
            host.addEventListener('click', function(ev){
              var sync = ev.target.closest && ev.target.closest('.pco-team-sync-btn');
              var del  = ev.target.closest && ev.target.closest('.pco-team-del-btn');
              var chk  = ev.target.closest && ev.target.closest('.pco-team-check-pos-btn');
              if (sync) {
                openTeamSyncPreview(sync.dataset.pcoTeamId, sync.dataset.tpOrgId);
              } else if (del) {
                if (!confirm('Remove this team mapping? Existing TP members and subgroups will NOT be removed.')) return;
                ajax('delete_team_mapping', {pco_team_id: del.dataset.pcoTeamId}, function(err, d){
                  if (err || !d || !d.success) { toast('Delete failed: ' + ((d && d.message) || err), 'err'); return; }
                  toast('Removed.', 'ok');
                  loadTeamMappings();
                });
              } else if (chk) {
                var resultSpan = host.querySelector('.pco-team-pos-result[data-pco-team-id="' + chk.dataset.pcoTeamId + '"]');
                if (resultSpan) { resultSpan.textContent = 'Checking PCO...'; resultSpan.style.color = ''; }
                chk.disabled = true;
                ajax('check_pco_team_positions', {pco_team_id: chk.dataset.pcoTeamId}, function(err, d){
                  chk.disabled = false;
                  if (!resultSpan) return;
                  if (err || !d || !d.success) {
                    resultSpan.textContent = 'Check failed: ' + ((d && d.message) || err);
                    resultSpan.style.color = '#a00';
                    return;
                  }
                  var pc = d.positionCount || 0;
                  var pac = d.positionAssignmentCount || 0;
                  var pwpc = d.peopleWithPositionsCount || 0;
                  var names = (d.positions || []).map(function(p){ return p.positionName; }).filter(Boolean);
                  var sample = names.length ? ' [' + names.slice(0, 5).join(', ') + (names.length > 5 ? ', ...' : '') + ']' : '';
                  if (pc === 0) {
                    resultSpan.style.color = '#c0392b';
                    resultSpan.textContent = 'No positions defined in PCO for this team. Add them in PCO -> Services -> Teams -> Positions tab.';
                  } else if (pac === 0) {
                    resultSpan.style.color = '#8a6d3b';
                    resultSpan.textContent = pc + ' position(s)' + sample + ' defined, but no one is assigned in PCO.';
                  } else if (pwpc === 0) {
                    resultSpan.style.color = '#8a6d3b';
                    resultSpan.textContent = pc + ' position(s)' + sample + ', ' + pac + ' assignment(s), but none matched the team roster.';
                  } else {
                    resultSpan.style.color = '#1f6b3a';
                    resultSpan.textContent = pc + ' position(s)' + sample + ', ' + pac + ' assignment(s) across ' + pwpc + ' person/people. Subgroups will sync.';
                  }
                });
              }
            });
            // v3.0 per-mapping toggles for Team Mapping.
            host.addEventListener('change', function(ev){
              var t = ev.target;
              if (!t.classList) return;
              var isSubg = t.classList.contains('pco-team-opt-subgroups');
              var isAtt  = t.classList.contains('pco-team-opt-attendance');
              if (!isSubg && !isAtt) return;
              var row = t.closest('.pco-team-map-row');
              if (!row) return;
              var subgBox = row.querySelector('.pco-team-opt-subgroups');
              var attBox  = row.querySelector('.pco-team-opt-attendance');
              var status  = row.querySelector('.pco-team-opt-status');
              if (status) { status.textContent = 'Saving...'; status.style.color = ''; }
              ajax('set_team_mapping_options', {
                pco_team_id: t.dataset.pcoTeamId,
                positions_as_subgroups: subgBox && subgBox.checked ? '1' : '0',
                per_plan_attendance: attBox && attBox.checked ? '1' : '0',
              }, function(err, d){
                if (err || !d || !d.success) {
                  if (status) { status.textContent = 'Save failed: ' + ((d && d.message) || err); status.style.color = '#a00'; }
                  return;
                }
                if (status) {
                  status.textContent = 'Saved.';
                  status.style.color = '#1f6b3a';
                  setTimeout(function(){ if (status.textContent === 'Saved.') status.textContent = ''; }, 1500);
                }
              });
            });
          }
        });
      }

      // v3.3: shared collapsible schedule editor used by Team / People /
      // All People mapping rows. kind = 'team' | 'people' | 'all_people'.
      // key = PCO id ('' for all_people). sch is the schedule dict
      // shipped by the server (enabled, frequency, dayOfWeek, hour,
      // notifyUsername, includeIssues).
      function renderSchedulePanel(kind, key, sch, label, nextRunIso, schedulerInstalled) {
        sch = sch || {};
        // v3.3.1: if the global scheduler isn't installed, show a
        // disabled-state hint instead of the editor. Mappings can't
        // be enabled until install -- saves staff configuring something
        // that will never fire.
        if (!schedulerInstalled) {
          return ''
            + '<div class="pco-sched-row" style="margin-top:6px;padding-top:6px;border-top:1px dashed #e1e4e8;font-size:12px;color:#8a6d3b;">'
            + '<strong>Scheduler not installed.</strong> Open <em>Settings &rarr; Scheduled Sync</em> and click Install to enable scheduling on this mapping.'
            + '</div>';
        }
        var enabled = !!sch.enabled;
        var freq = sch.frequency || 'weekly';
        var dow = (sch.dayOfWeek != null) ? sch.dayOfWeek : 6;
        var hr = (sch.hour != null) ? sch.hour : 6;
        var user = sch.notifyUsername || '';
        var inc = (sch.includeIssues !== false);
        var dataAttrs = ' data-kind="' + escAttr(kind) + '" data-key="' + escAttr(key || '') + '"';
        var summaryText = enabled
          ? '<strong style="color:#1f4e79;">Scheduled:</strong> ' + escHtml(label || '') + (nextRunIso ? ' <span class="pco-muted" style="font-size:11px;">(next: ' + escHtml(nextRunIso.replace('T', ' ')) + ')</span>' : '')
          : '<span class="pco-muted">Schedule disabled.</span> <a href="#" class="pco-sched-open">Set up...</a>';
        var optHours = '';
        for (var h2 = 0; h2 < 24; h2++) {
          var lbl2 = (h2 === 0 ? '12 AM' : h2 < 12 ? h2 + ' AM' : h2 === 12 ? '12 PM' : (h2 - 12) + ' PM');
          optHours += '<option value="' + h2 + '"' + (h2 === hr ? ' selected' : '') + '>' + lbl2 + '</option>';
        }
        return ''
          + '<div class="pco-sched-row" style="margin-top:6px;padding-top:6px;border-top:1px dashed #e1e4e8;font-size:12px;">'
          + '  <div class="pco-sched-summary" style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;">'
          + '    <span class="pco-sched-summary-text">' + summaryText + '</span>'
          + '    <a href="#" class="pco-sched-toggle" style="font-size:11px;">Edit schedule &raquo;</a>'
          + '  </div>'
          + '  <div class="pco-sched-form" style="display:none;margin-top:8px;padding:8px 10px;background:#f4f7fa;border-radius:4px;"' + dataAttrs + '>'
          + '    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">'
          + '      <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
          + '        <input type="checkbox" class="pco-sched-enabled"' + (enabled ? ' checked' : '') + '>'
          + '        <strong>Enable scheduled sync</strong>'
          + '      </label>'
          + '    </div>'
          + '    <div class="pco-sched-fields" style="display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;">'
          + '      <div>'
          + '        <label style="display:block;font-size:11px;color:#666;">Frequency</label>'
          + '        <select class="pco-sched-freq" style="padding:3px 6px;">'
          + '          <option value="weekly"' + (freq === 'weekly' ? ' selected' : '') + '>Weekly</option>'
          + '          <option value="daily"' + (freq === 'daily' ? ' selected' : '') + '>Daily</option>'
          + '        </select>'
          + '      </div>'
          + '      <div class="pco-sched-dow-wrap"' + (freq === 'daily' ? ' style="display:none;"' : '') + '>'
          + '        <label style="display:block;font-size:11px;color:#666;">Day of week</label>'
          + '        <select class="pco-sched-dow" style="padding:3px 6px;">'
          + '          <option value="0"' + (dow === 0 ? ' selected' : '') + '>Mon</option>'
          + '          <option value="1"' + (dow === 1 ? ' selected' : '') + '>Tue</option>'
          + '          <option value="2"' + (dow === 2 ? ' selected' : '') + '>Wed</option>'
          + '          <option value="3"' + (dow === 3 ? ' selected' : '') + '>Thu</option>'
          + '          <option value="4"' + (dow === 4 ? ' selected' : '') + '>Fri</option>'
          + '          <option value="5"' + (dow === 5 ? ' selected' : '') + '>Sat</option>'
          + '          <option value="6"' + (dow === 6 ? ' selected' : '') + '>Sun</option>'
          + '        </select>'
          + '      </div>'
          + '      <div>'
          + '        <label style="display:block;font-size:11px;color:#666;">Hour</label>'
          + '        <select class="pco-sched-hour" style="padding:3px 6px;">' + optHours + '</select>'
          + '      </div>'
          + '      <div style="flex:1 1 240px;min-width:200px;position:relative;">'
          + '        <label style="display:block;font-size:11px;color:#666;">Notify TouchPoint user</label>'
          + '        <input type="text" class="pco-sched-username" value="' + escAttr(user) + '" placeholder="Type name, email, or username..." autocomplete="off" style="padding:3px 6px;width:100%;box-sizing:border-box;">'
          + '        <div class="pco-sched-user-dropdown" style="display:none;position:absolute;z-index:50;background:#fff;border:1px solid #ccc;border-radius:3px;box-shadow:0 2px 6px rgba(0,0,0,0.1);max-height:220px;overflow-y:auto;left:0;right:0;top:100%;"></div>'
          + '        <div class="pco-sched-user-info pco-muted" style="font-size:11px;margin-top:2px;"></div>'
          + '      </div>'
          + '    </div>'
          + '    <div style="margin-top:8px;display:flex;gap:14px;align-items:center;flex-wrap:wrap;">'
          + '      <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:12px;">'
          + '        <input type="checkbox" class="pco-sched-issues"' + (inc ? ' checked' : '') + '>'
          + '        Include issues list (unmatched, ambiguous, warnings)'
          + '      </label>'
          + '      <button class="pco-btn pco-sched-save" style="font-size:12px;padding:3px 12px;margin-left:auto;">Save schedule</button>'
          + '      <span class="pco-sched-status pco-muted" style="font-size:11px;"></span>'
          + '    </div>'
          + '  </div>'
          + '</div>';
      }

      // Wire schedule panel events on a host (delegated). Idempotent.
      function wireSchedulePanels(host) {
        if (!host || host.__pcoSchedWired) return;
        host.__pcoSchedWired = true;
        host.addEventListener('click', function(ev){
          var t = ev.target;
          if (!t.classList) return;
          if (t.classList.contains('pco-sched-toggle') || t.classList.contains('pco-sched-open')) {
            ev.preventDefault();
            var row = t.closest('.pco-sched-row');
            if (!row) return;
            var form = row.querySelector('.pco-sched-form');
            if (!form) return;
            var open = form.style.display === 'none' || !form.style.display;
            form.style.display = open ? 'block' : 'none';
            t.textContent = open ? 'Close schedule editor' : 'Edit schedule »';
            return;
          }
          if (t.classList.contains('pco-sched-save')) {
            ev.preventDefault();
            var form2 = t.closest('.pco-sched-form');
            if (!form2) return;
            var kind = form2.getAttribute('data-kind');
            var key  = form2.getAttribute('data-key');
            var data = {
              kind: kind,
              key:  key,
              enabled: form2.querySelector('.pco-sched-enabled').checked ? '1' : '0',
              frequency: form2.querySelector('.pco-sched-freq').value,
              day_of_week: form2.querySelector('.pco-sched-dow').value,
              hour: form2.querySelector('.pco-sched-hour').value,
              notify_username: form2.querySelector('.pco-sched-username').value,
              include_issues: form2.querySelector('.pco-sched-issues').checked ? '1' : '0',
            };
            var status = form2.querySelector('.pco-sched-status');
            if (status) { status.textContent = 'Saving...'; status.style.color = ''; }
            ajax('set_schedule_options', data, function(err, d){
              if (err || !d || !d.success) {
                if (status) { status.textContent = (d && d.message) || err || 'Save failed.'; status.style.color = '#a00'; }
                return;
              }
              if (status) { status.textContent = 'Saved.'; status.style.color = '#1f6b3a'; setTimeout(function(){ if (status.textContent === 'Saved.') status.textContent = ''; }, 1500); }
              // Update summary text.
              var summary = form2.parentNode.querySelector('.pco-sched-summary-text');
              if (summary) {
                if (d.schedule && d.schedule.enabled) {
                  summary.innerHTML = '<strong style="color:#1f4e79;">Scheduled:</strong> ' + escHtml(d.scheduleLabel || '') + (d.nextRunIso ? ' <span class="pco-muted" style="font-size:11px;">(next: ' + escHtml(d.nextRunIso.replace('T', ' ')) + ')</span>' : '');
                } else {
                  summary.innerHTML = '<span class="pco-muted">Schedule disabled.</span>';
                }
              }
            });
          }
        });
        // v3.3.2: typeahead picker for Notify TouchPoint user.
        // Debounced search on keystroke; click a result to fill the
        // input. Blur with delay so click on the dropdown still fires.
        var _schedUserDebounceTimer = null;
        host.addEventListener('input', function(ev){
          var t = ev.target;
          if (!t.classList || !t.classList.contains('pco-sched-username')) return;
          var form = t.closest('.pco-sched-form');
          if (!form) return;
          var dropdown = form.querySelector('.pco-sched-user-dropdown');
          var info = form.querySelector('.pco-sched-user-info');
          var u = (t.value || '').trim();
          if (info) info.textContent = '';
          if (_schedUserDebounceTimer) clearTimeout(_schedUserDebounceTimer);
          if (!u || u.length < 2) {
            if (dropdown) dropdown.style.display = 'none';
            return;
          }
          _schedUserDebounceTimer = setTimeout(function(){
            ajax('search_tp_users', {search_term: u}, function(err, d){
              if (!dropdown) return;
              if (err || !d || !d.success) {
                dropdown.innerHTML = '<div style="padding:6px 10px;color:#a00;font-size:12px;">Search failed.</div>';
                dropdown.style.display = 'block';
                return;
              }
              var users = d.users || [];
              if (!users.length) {
                dropdown.innerHTML = '<div style="padding:6px 10px;color:#666;font-size:12px;">No TouchPoint user matches (must have email on file).</div>';
                dropdown.style.display = 'block';
                return;
              }
              var inner = '';
              for (var i = 0; i < users.length; i++) {
                var u2 = users[i];
                inner += '<div class="pco-sched-user-pick" data-username="' + escAttr(u2.username) + '" data-name="' + escAttr(u2.name) + '" data-email="' + escAttr(u2.email) + '" style="padding:6px 10px;border-bottom:1px solid #f0f0f0;cursor:pointer;font-size:12px;">'
                  + '<strong>' + escHtml(u2.name) + '</strong> '
                  + '<span style="color:#666;">@' + escHtml(u2.username) + '</span>'
                  + '<div style="color:#888;font-size:11px;">' + escHtml(u2.email) + '</div>'
                  + '</div>';
              }
              dropdown.innerHTML = inner;
              dropdown.style.display = 'block';
            });
          }, 220);
        });
        // Pick from dropdown.
        host.addEventListener('mousedown', function(ev){
          var pick = ev.target.closest && ev.target.closest('.pco-sched-user-pick');
          if (!pick) return;
          ev.preventDefault();
          var form = pick.closest('.pco-sched-form');
          if (!form) return;
          var input = form.querySelector('.pco-sched-username');
          var dropdown = form.querySelector('.pco-sched-user-dropdown');
          var info = form.querySelector('.pco-sched-user-info');
          if (input) input.value = pick.getAttribute('data-username') || '';
          if (dropdown) { dropdown.style.display = 'none'; dropdown.innerHTML = ''; }
          if (info) {
            info.textContent = 'Selected: ' + (pick.getAttribute('data-name') || '') + ' · ' + (pick.getAttribute('data-email') || '');
            info.style.color = '#1f6b3a';
          }
        });
        // Close dropdown on outside click + blur (delayed so picks register).
        host.addEventListener('blur', function(ev){
          var t = ev.target;
          if (!t.classList || !t.classList.contains('pco-sched-username')) return;
          var form = t.closest('.pco-sched-form');
          if (!form) return;
          setTimeout(function(){
            var dd = form.querySelector('.pco-sched-user-dropdown');
            if (dd) dd.style.display = 'none';
          }, 180);
          // Validate non-empty input on blur.
          var u = (t.value || '').trim();
          if (!u) return;
          var info = form.querySelector('.pco-sched-user-info');
          if (info) info.textContent = 'Verifying...';
          ajax('resolve_username', {username: u}, function(err, d){
            if (!info) return;
            if (err || !d) { info.textContent = 'Check failed.'; info.style.color = '#a00'; return; }
            if (d.found) {
              info.textContent = 'Will notify: ' + d.name + ' · ' + d.email;
              info.style.color = '#1f6b3a';
            } else {
              info.textContent = 'No TouchPoint user with username "' + u + '" (or no email). Use the dropdown.';
              info.style.color = '#a00';
            }
          });
        }, true);
        // Frequency change -> show/hide day-of-week.
        host.addEventListener('change', function(ev){
          var t = ev.target;
          if (!t.classList || !t.classList.contains('pco-sched-freq')) return;
          var form = t.closest('.pco-sched-form');
          if (!form) return;
          var dowWrap = form.querySelector('.pco-sched-dow-wrap');
          if (!dowWrap) return;
          dowWrap.style.display = (t.value === 'daily') ? 'none' : '';
        });
      }

      function renderTeamMappingRow(r, schedulerInstalled) {
        var orgCtx = '';
        if (r.tpDivision) orgCtx = '<span class="pco-muted"> &middot; ' + escHtml(r.tpDivision) + (r.tpProgram ? ' / ' + escHtml(r.tpProgram) : '') + '</span>';
        var subgChk = r.positionsAsSubgroups ? ' checked' : '';
        var attChk  = r.perPlanAttendance ? ' checked' : '';
        return ''
          + '<div class="pco-team-map-row" style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;">'
          + '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">'
          + '<div style="flex:1 1 320px;min-width:0;">'
          + '<div style="font-weight:700;color:#1f4e79;">' + escHtml(r.pcoTeamName || ('Team ' + r.pcoTeamId))
          +   (r.pcoServiceTypeName ? ' <span class="pco-muted" style="font-weight:400;">(' + escHtml(r.pcoServiceTypeName) + ')</span>' : '')
          + '</div>'
          + '<div style="margin-top:2px;font-size:13px;color:#555;">'
          +   '<strong>TP:</strong> ' + escHtml(r.tpOrgName) + ' <span class="pco-muted">(#' + r.tpOrgId + ')</span>' + orgCtx
          + '</div>'
          + '</div>'
          + '<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;justify-content:flex-end;">'
          + '<span class="pco-pill" style="background:#1f6b3a;color:#fff;font-size:11px;">Team Sync</span>'
          + '<button class="pco-btn pco-team-sync-btn" data-pco-team-id="' + escAttr(r.pcoTeamId) + '" data-tp-org-id="' + r.tpOrgId + '" style="font-size:13px;padding:5px 12px;">Preview &amp; Sync</button>'
          + '<button class="pco-btn pco-secondary pco-team-del-btn" data-pco-team-id="' + escAttr(r.pcoTeamId) + '" style="font-size:13px;padding:5px 12px;color:#a00;border-color:#a00;">Remove</button>'
          + '</div>'
          + '</div>'
          // v3.0 per-mapping toggles (same UX as Service Type Mapping).
          + '<div style="margin-top:8px;padding-top:8px;border-top:1px dashed #e1e4e8;display:flex;gap:16px;flex-wrap:wrap;font-size:13px;">'
          + '  <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
          + '    <input type="checkbox" class="pco-team-opt-subgroups" data-pco-team-id="' + escAttr(r.pcoTeamId) + '"' + subgChk + '>'
          + '    Positions as subgroups <span class="pco-muted">(Lead Vocal, Backup, etc.)</span>'
          + '  </label>'
          + '  <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
          + '    <input type="checkbox" class="pco-team-opt-attendance" data-pco-team-id="' + escAttr(r.pcoTeamId) + '"' + attChk + '>'
          + '    Per-plan attendance <span class="pco-muted">(walks plans + marks Confirmed Present)</span>'
          + '  </label>'
          + '  <span class="pco-team-opt-status pco-muted" data-pco-team-id="' + escAttr(r.pcoTeamId) + '" style="font-size:12px;align-self:center;"></span>'
          + '</div>'
          // v3.1.1: inline "check positions" diagnostic so the user
          // can verify what PCO returns for this team without opening
          // the preview modal. Cheap: one call per Team Position.
          + '<div style="margin-top:6px;font-size:12px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">'
          + '  <button class="pco-btn pco-team-check-pos-btn" data-pco-team-id="' + escAttr(r.pcoTeamId) + '" style="font-size:11px;padding:2px 8px;">Check PCO positions</button>'
          + '  <span class="pco-team-pos-result pco-muted" data-pco-team-id="' + escAttr(r.pcoTeamId) + '"></span>'
          + '</div>'
          // v3.3: schedule editor.
          + renderSchedulePanel('team', r.pcoTeamId, r.schedule, r.scheduleLabel, r.nextRunIso, schedulerInstalled)
          + '</div>';
      }

      // Picker flow: pick Service Type -> pick Team -> search & link
      // involvement -> save. Two AJAX hops (list_service_types,
      // list_pco_teams_for_service_type) keep the experience interactive.
      function openTeamMappingPicker() {
        var box = $('pcoTeamMappingPicker');
        box.style.display = '';
        box.innerHTML = '<div class="pco-muted">Loading service types...</div>';
        ajax('list_service_types', {}, function(err, d){
          if (err || !d || !d.success) {
            box.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var sts = (d.serviceTypes || []).slice();
          sts.sort(function(a, b){
            return (a.pcoName || '').toLowerCase().localeCompare((b.pcoName || '').toLowerCase());
          });
          var opts = '<option value="">-- pick a service type --</option>';
          for (var i = 0; i < sts.length; i++) {
            opts += '<option value="' + escAttr(sts[i].pcoId) + '" data-name="' + escAttr(sts[i].pcoName) + '">' + escHtml(sts[i].pcoName) + '</option>';
          }
          box.innerHTML = ''
            + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">'
            + '  <label style="font-size:13px;font-weight:600;">Service Type:</label>'
            + '  <select id="pcoTeamPickStSel" style="padding:4px 6px;">' + opts + '</select>'
            + '  <button class="pco-btn pco-secondary" id="pcoTeamPickCancel" style="font-size:13px;padding:4px 10px;margin-left:auto;">Cancel</button>'
            + '</div>'
            + '<div id="pcoTeamPickTeams"></div>';
          $('pcoTeamPickCancel').onclick = function(){
            box.style.display = 'none';
            box.innerHTML = '';
          };
          $('pcoTeamPickStSel').onchange = function(){
            var sel = this;
            var stId = sel.value;
            var stName = sel.options[sel.selectedIndex].dataset.name || '';
            var teamsEl = $('pcoTeamPickTeams');
            if (!stId) { teamsEl.innerHTML = ''; return; }
            teamsEl.innerHTML = '<div class="pco-muted">Loading teams...</div>';
            ajax('list_pco_teams_for_service_type', {service_type_id: stId}, function(err, d){
              if (err || !d || !d.success) {
                teamsEl.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
                return;
              }
              var teams = d.teams || [];
              if (!teams.length) {
                teamsEl.innerHTML = '<div class="pco-muted">No teams in this service type.</div>';
                return;
              }
              var teamOpts = '<option value="">-- pick a team --</option>';
              for (var i = 0; i < teams.length; i++) {
                teamOpts += '<option value="' + escAttr(teams[i].teamId) + '" data-name="' + escAttr(teams[i].teamName) + '">' + escHtml(teams[i].teamName) + '</option>';
              }
              teamsEl.innerHTML = ''
                + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">'
                + '  <label style="font-size:13px;font-weight:600;">Team:</label>'
                + '  <select id="pcoTeamPickTeamSel" style="padding:4px 6px;">' + teamOpts + '</select>'
                + '</div>'
                + '<div id="pcoTeamPickOrgWrap" style="display:none;">'
                + '  <label class="pco-label">TouchPoint involvement</label>'
                + '  <input type="text" class="pco-input" id="pcoTeamPickOrgSearch" placeholder="Search by name or org id...">'
                + '  <div id="pcoTeamPickOrgResults" style="margin-top:6px;"></div>'
                + '</div>';
              $('pcoTeamPickTeamSel').onchange = function(){
                if (!this.value) { $('pcoTeamPickOrgWrap').style.display = 'none'; return; }
                $('pcoTeamPickOrgWrap').style.display = '';
                $('pcoTeamPickOrgSearch').focus();
              };
              $('pcoTeamPickOrgSearch').oninput = function(){
                var term = this.value.trim();
                var resultsEl = $('pcoTeamPickOrgResults');
                if (term.length < 2) {
                  resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
                  return;
                }
                resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
                clearTimeout(window.__pcoTeamPickSearchT);
                window.__pcoTeamPickSearchT = setTimeout(function(){
                  ajax('search_involvements', {search_term: term}, function(err, d){
                    if (err || !d || !d.success) { resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>'; return; }
                    var orgs = d.orgs || [];
                    if (!orgs.length) { resultsEl.innerHTML = '<span class="pco-muted">No matches.</span>'; return; }
                    var html = '<div style="border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:240px;overflow-y:auto;">';
                    for (var i = 0; i < orgs.length; i++) {
                      var o = orgs[i];
                      html += '<div class="pco-team-pick-org" data-org-id="' + o.orgId + '" data-org-name="' + escAttr(o.orgName) + '" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;">'
                        + '<strong>' + escHtml(o.orgName) + '</strong> '
                        + '<span class="pco-muted">(#' + o.orgId
                        + (o.divisionName ? ' &middot; ' + escHtml(o.divisionName) : '')
                        + (o.programName ? ' / ' + escHtml(o.programName) : '')
                        + ')</span>'
                        + '</div>';
                    }
                    html += '</div>';
                    resultsEl.innerHTML = html;
                    var picks = resultsEl.querySelectorAll('.pco-team-pick-org');
                    for (var j = 0; j < picks.length; j++) {
                      picks[j].onclick = function(){
                        var orgId = this.dataset.orgId;
                        var orgName = this.dataset.orgName;
                        var teamSel = $('pcoTeamPickTeamSel');
                        var pcoTeamId = teamSel.value;
                        var pcoTeamName = teamSel.options[teamSel.selectedIndex].dataset.name || '';
                        ajax('save_team_mapping', {
                          pco_team_id: pcoTeamId,
                          pco_team_name: pcoTeamName,
                          pco_service_type_id: stId,
                          pco_service_type_name: stName,
                          tp_org_id: orgId,
                        }, function(err, d){
                          if (err || !d || !d.success) { toast('Save failed: ' + ((d && d.message) || err), 'err'); return; }
                          toast('Team mapped to ' + orgName + '.', 'ok');
                          $('pcoTeamMappingPicker').style.display = 'none';
                          $('pcoTeamMappingPicker').innerHTML = '';
                          loadTeamMappings();
                        });
                      };
                    }
                  });
                }, 250);
              };
            });
          };
        });
      }

      function openTeamSyncPreview(pcoTeamId, tpOrgId) {
        renderPlanPreviewModal('<div class="pco-empty">Loading team roster from PCO...</div>');
        ajax('load_team_sync_preview', {pco_team_id: pcoTeamId}, function(err, d){
          if (err || !d || !d.success) {
            renderPlanPreviewModal('<div class="pco-pill pco-err">Error</div><div>' + escHtml((d && d.message) || err) + '</div>');
            return;
          }
          _planPreviewState = {
            planId: '',
            serviceTypeId: '',
            pcoTeamId: pcoTeamId,
            orgId: parseInt(tpOrgId, 10) || 0,
            attendees: d.attendees || [],
            planInfo: d.planInfo || {},
            summary: d.summary || {},
            isRosterSync: false,
            isTeamSync: true,
          };
          renderPlanPreviewBody();
        });
      }

      function loadMappings() {
        var host = $('pcoMappingsList');
        host.className = 'pco-empty';
        host.innerHTML = 'Loading service types from PCO...';
        ajax('list_service_types', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          var sts = d.serviceTypes || [];
          if (!sts.length) {
            host.innerHTML = 'No service types found in PCO. Make sure your PAT has access to Services.';
            return;
          }
          renderMappingsList(host, sts, d.mappedCount || 0);
        });
      }

      function renderMappingsList(host, sts, mappedCount) {
        host.className = '';
        var unmappedCount = sts.length - mappedCount;
        // v2.3+ shows mapped rows only. Unmapped service types live behind
        // the "+ Add Service Type Mapping" button so the page isn't full
        // of clutter staff don't act on.
        var html = ''
          + '<div class="pco-muted" style="margin-bottom:8px;">'
          + mappedCount + ' mapped &middot; '
          + unmappedCount + ' unmapped <span style="opacity:0.7;">(use + Add to wire one)</span>'
          + '</div>'
          + '<div style="display:flex;flex-direction:column;gap:8px;">';
        if (mappedCount === 0) {
          html += '<div class="pco-empty">No service types are mapped yet. Click <strong>+ Add Service Type Mapping</strong> above to wire one to a TouchPoint involvement.</div>';
        } else {
          html += renderMappingsSectionHeader('Mapped', mappedCount, '#1f6b3a');
          for (var i = 0; i < sts.length; i++) {
            var row = sts[i];
            if (!row.mappedOrgId) continue;
            html += renderMappingRow(row);
          }
        }
        html += '</div>';
        host.innerHTML = html;

        // Wire per-row handlers via delegation. Use .closest() so a click on
        // a nested element (e.g. <strong> inside a pick row) still resolves
        // to the right action -- without this, the first click misses the
        // outer class and the user had to double-click to register.
        // Guard so we only attach listeners once per host element. Without
        // this, every saveOrgMapping -> loadMappings -> renderMappingsList
        // cycle stacks an extra set of listeners, multiplying AJAX calls.
        if (host.__pcoWired) return;
        host.__pcoWired = true;
        host.addEventListener('click', function(ev){
          var unmapBtn = ev.target.closest && ev.target.closest('.pco-map-unmap-btn');
          var searchBtn = ev.target.closest && ev.target.closest('.pco-map-search-btn');
          var pickRow = ev.target.closest && ev.target.closest('.pco-map-pick');
          if (unmapBtn) {
            unmapServiceType(unmapBtn.dataset.pcoId);
          } else if (searchBtn) {
            var row = searchBtn.closest('.pco-map-row');
            var searchBox = row.querySelector('.pco-map-search-box');
            searchBox.style.display = searchBox.style.display === 'none' ? '' : 'none';
            if (searchBox.style.display !== 'none') searchBox.querySelector('input').focus();
          } else if (pickRow) {
            saveOrgMapping(pickRow.dataset.pcoId, parseInt(pickRow.dataset.orgId, 10) || 0);
          }
        });
        host.addEventListener('input', function(ev){
          if (ev.target.classList.contains('pco-map-search-input')) {
            doOrgSearch(ev.target);
          }
        });
        // Toggle handlers for the per-mapping option checkboxes. We send
        // both flags every time so the server doesn't have to merge state.
        host.addEventListener('change', function(ev){
          var t = ev.target;
          var isAtt = t.classList && t.classList.contains('pco-map-opt-attendance');
          var isMem = t.classList && t.classList.contains('pco-map-opt-member');
          if (!isAtt && !isMem) return;
          var pcoId = t.dataset.pcoId;
          var row = t.closest('.pco-map-row');
          if (!row) return;
          var attBox = row.querySelector('.pco-map-opt-attendance');
          var memBox = row.querySelector('.pco-map-opt-member');
          var status = row.querySelector('.pco-map-opt-status');
          // Block the user from turning OFF both at the same time (the row
          // would become a no-op). Snap the just-toggled one back on.
          if (attBox && memBox && !attBox.checked && !memBox.checked) {
            t.checked = true;
            if (status) {
              status.textContent = 'At least one option must stay on.';
              status.style.color = '#a85a00';
              setTimeout(function(){ if (status.textContent === 'At least one option must stay on.') status.textContent = ''; }, 2500);
            }
            return;
          }
          if (status) {
            status.textContent = 'Saving...';
            status.style.color = '';
          }
          ajax('set_mapping_options', {
            pco_service_type_id: pcoId,
            sync_attendance: attBox && attBox.checked ? '1' : '0',
            auto_add_member: memBox && memBox.checked ? '1' : '0',
          }, function(err, d){
            if (err || !d || !d.success) {
              // Revert visually on failure.
              if (attBox) attBox.checked = !!(d && d.syncAttendance);
              if (memBox) memBox.checked = !!(d && d.autoAddMember);
              if (status) {
                status.textContent = 'Save failed: ' + ((d && d.message) || err);
                status.style.color = '#a00';
              }
              return;
            }
            if (status) {
              status.textContent = 'Saved.';
              status.style.color = '#1f6b3a';
              setTimeout(function(){ if (status.textContent === 'Saved.') status.textContent = ''; }, 1500);
            }
          });
        });
      }

      function renderMappingsSectionHeader(label, count, color) {
        return ''
          + '<div style="margin-top:4px;padding:6px 4px 4px;'
          + 'border-bottom:1px solid #e1e4e8;'
          + 'display:flex;align-items:baseline;gap:8px;">'
          + '  <span style="font-size:11px;font-weight:700;letter-spacing:0.06em;'
          +     'text-transform:uppercase;color:' + color + ';">'
          +     escHtml(label)
          + '  </span>'
          + '  <span class="pco-muted" style="font-size:12px;">' + count + '</span>'
          + '</div>';
      }

      // Unmapped section header doubles as the toggle button so a long
      // list of unmapped service types doesn't bury the Team Mappings
      // section below.
      function renderUnmappedCollapseHeader(count, defaultExpanded) {
        var label = defaultExpanded ? 'Hide ' + count + ' unmapped' : 'Show ' + count + ' unmapped';
        return ''
          + '<div style="margin-top:4px;padding:6px 4px 4px;'
          + 'border-bottom:1px solid #e1e4e8;'
          + 'display:flex;align-items:baseline;gap:8px;justify-content:space-between;">'
          + '  <span style="font-size:11px;font-weight:700;letter-spacing:0.06em;'
          +     'text-transform:uppercase;color:#a85a00;">UNMAPPED</span>'
          + '  <button class="pco-btn pco-secondary pco-unmapped-toggle" style="font-size:12px;padding:2px 10px;">' + escHtml(label) + '</button>'
          + '</div>';
      }

      function renderMappingRow(st) {
        var pcoIdEsc = escAttr(st.pcoId);
        var pcoNameEsc = escHtml(st.pcoName);
        var head = ''
          + '<div class="pco-map-row" style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;">'
          + '  <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'
          + '    <div style="flex:1 1 240px;font-weight:600;">' + pcoNameEsc + '</div>';
        if (st.mappedOrgId) {
          head += ''
            + '    <span class="pco-pill pco-ok">Mapped</span>'
            + '    <div style="flex:2 1 320px;">'
            + '      <span style="font-weight:600;">' + escHtml(st.mappedOrgName) + '</span>'
            + (st.mappedDivision ? ' <span class="pco-muted">&middot; ' + escHtml(st.mappedDivision) + (st.mappedProgram ? ' / ' + escHtml(st.mappedProgram) : '') + '</span>' : '')
            + '    </div>'
            + '    <button class="pco-btn pco-secondary pco-map-search-btn" data-pco-id="' + pcoIdEsc + '" style="padding:4px 10px;font-size:13px;">Change</button>'
            + '    <button class="pco-btn pco-secondary pco-map-unmap-btn" data-pco-id="' + pcoIdEsc + '" style="padding:4px 10px;font-size:13px;color:#a00;border-color:#a00;">Remove</button>';
        } else {
          head += ''
            + '    <span class="pco-pill pco-warn">Unmapped</span>'
            + '    <div style="flex:2 1 320px;color:#888;font-style:italic;">No TouchPoint involvement chosen yet.</div>'
            + '    <button class="pco-btn pco-map-search-btn" data-pco-id="' + pcoIdEsc + '" style="padding:4px 10px;font-size:13px;">Pick Involvement</button>';
        }
        head += '  </div>';
        // Per-mapping sync options. Only render the toggle row when the
        // service type is actually mapped -- toggles don't mean anything
        // without an org pointer.
        if (st.mappedOrgId) {
          var syncChk = st.syncAttendance ? ' checked' : '';
          var memChk = st.autoAddMember ? ' checked' : '';
          head += ''
            + '  <div class="pco-map-options" style="margin-top:10px;padding-top:8px;border-top:1px dashed #e1e4e8;display:flex;gap:18px;flex-wrap:wrap;align-items:center;font-size:13px;">'
            + '    <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
            + '      <input type="checkbox" class="pco-map-opt-attendance" data-pco-id="' + pcoIdEsc + '"' + syncChk + '>'
            + '      <span><strong>Sync attendance</strong> <span class="pco-muted">(mark Confirmed as Present)</span></span>'
            + '    </label>'
            + '    <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">'
            + '      <input type="checkbox" class="pco-map-opt-member" data-pco-id="' + pcoIdEsc + '"' + memChk + '>'
            + '      <span><strong>Auto-add as member</strong> <span class="pco-muted">(JoinOrg if not already active)</span></span>'
            + '    </label>'
            + '    <span class="pco-map-opt-status pco-muted" data-pco-id="' + pcoIdEsc + '" style="font-size:12px;"></span>'
            + '  </div>';
        }
        head += ''
          + '  <div class="pco-map-search-box" style="display:none;margin-top:10px;padding:10px;background:#fff;border:1px solid #e1e4e8;border-radius:4px;">'
          + '    <input type="text" class="pco-input pco-map-search-input" data-pco-id="' + pcoIdEsc + '" placeholder="Search TouchPoint involvements by name or ID...">'
          + '    <div class="pco-map-search-results" style="margin-top:8px;"></div>'
          + '  </div>'
          + '</div>';
        return head;
      }

      var _searchTimer = null;
      function doOrgSearch(inp) {
        clearTimeout(_searchTimer);
        var pcoId = inp.dataset.pcoId;
        var term = inp.value.trim();
        var resultsEl = inp.parentElement.querySelector('.pco-map-search-results');
        if (term.length < 2) {
          resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
          return;
        }
        resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
        _searchTimer = setTimeout(function(){
          ajax('search_involvements', {search_term: term}, function(err, d){
            if (err || !d || !d.success) {
              resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>';
              return;
            }
            var orgs = d.orgs || [];
            if (!orgs.length) {
              resultsEl.innerHTML = '<span class="pco-muted">No active involvements match "' + escHtml(term) + '".</span>';
              return;
            }
            var html = '';
            for (var i = 0; i < orgs.length; i++) {
              var o = orgs[i];
              // Use data attributes so apostrophes in org names don't break the
              // onclick handler (see RollSheet v1.2.2 changelog).
              html += '<div class="pco-map-pick" style="padding:6px 10px;cursor:pointer;border-radius:4px;border-bottom:1px solid #f0f0f0;" '
                   + 'data-pco-id="' + escAttr(pcoId) + '" data-org-id="' + o.orgId + '">'
                   + '<div><strong>' + escHtml(o.orgName) + '</strong> '
                   + '<span class="pco-muted">(ID: ' + o.orgId + ', ' + (o.memberCount || 0) + ' members)</span></div>';
              if (o.programName || o.divisionName) {
                html += '<div class="pco-muted" style="font-size:12px;">' + escHtml(o.programName) + (o.divisionName ? ' / ' + escHtml(o.divisionName) : '') + '</div>';
              }
              html += '</div>';
            }
            resultsEl.innerHTML = html;
          });
        }, 250);
      }

      function saveOrgMapping(pcoId, orgId) {
        ajax('save_org_mapping', {pco_service_type_id: pcoId, tp_org_id: orgId}, function(err, d){
          if (err || !d || !d.success) {
            toast('Save failed: ' + ((d && d.message) || err), 'err');
            return;
          }
          toast(d.message || 'Mapping saved.', 'ok');
          loadMappings(); // Re-render the full list to reflect the change
        });
      }

      function unmapServiceType(pcoId) {
        if (!confirm('Remove this mapping? Sync for this service type will be disabled until you remap.')) return;
        saveOrgMapping(pcoId, 0);
      }

      function escHtml(s) {
        if (s == null) return '';
        var d = document.createElement('div');
        d.appendChild(document.createTextNode(String(s)));
        return d.innerHTML;
      }
      function escAttr(s) {
        return escHtml(s).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
      }

      // ---- People Sync tab (combined Unmatched + Pending Reviews) ----
      // Two sections side by side under one tab. Unmatched on top because
      // those need manual action before any sync can flow through; field
      // diffs below for matched people. Both refresh together so the
      // counts stay consistent.
      function renderPeopleSyncTab() {
        var host = $('pcoContent');
        host.innerHTML = ''
          + '<div class="pco-card">'
          + '<h3>People Matching</h3>'
          + '<div class="pco-help">'
          + 'Two review queues for resolving PCO &harr; TP people links. '
          + '<strong>Proposed Matches</strong> (below) is the main tool -- it scores TP candidates for unmatched PCO records using name + email + birthdate, with per-row Apply / Skip Forever and bulk Apply for high-confidence matches. '
          + '<strong>Pending Data Reviews</strong> handles field diffs (email changes, name corrections) flagged by your Person Data Sync rules.'
          + ' <a href="#" id="pcoPeopleRefresh">Refresh</a>'
          + '</div>'
          + '<div id="pcoPeopleSyncSummary" class="pco-muted" style="margin-bottom:10px;font-size:13px;"></div>'
          // PROPOSED MATCHES section -- the main matching tool.
          + '<div class="pco-day-header" style="margin:8px 0 6px 0;padding:8px 12px;background:#0f7c84;color:#fff;border-radius:4px;font-size:14px;font-weight:700;letter-spacing:0.5px;">PROPOSED MATCHES <span style="font-weight:400;font-size:12px;opacity:0.85;">(score-based, supports bulk Apply)</span></div>'
          + '<div id="pcoProposedMatches" class="pco-empty" style="margin-bottom:14px;">Click <strong>Load proposed matches</strong> to score TP candidates for every unmatched PCO record. Walks the entire PCO People directory (takes a few seconds), unless you opened this from a specific preview &mdash; then it scopes to those records only.</div>'
          + '<div style="margin-top:8px;margin-bottom:18px;display:flex;gap:8px;align-items:center;">'
          + '  <button class="pco-btn" id="pcoLoadProposedBtn" style="font-size:13px;padding:5px 12px;">Load proposed matches</button>'
          + '  <span class="pco-muted" id="pcoProposedHint" style="font-size:12px;">Walks PCO + scores name/email/birthdate signals against TP.</span>'
          + '</div>'
          // VERIFY LINK section (v3.1) -- reactive tool. Search a TP person ->
          // see their current PCO link side-by-side -> Unlink or Replace it.
          + '<div class="pco-day-header" style="margin:8px 0 6px 0;padding:6px 10px;background:#5a3e7a;color:#fff;border-radius:4px;font-size:13px;font-weight:700;letter-spacing:0.5px;">VERIFY PERSON LINK <span style="font-weight:400;font-size:12px;opacity:0.85;">(check or repair a specific TP &harr; PCO match)</span></div>'
          + '<div class="pco-help" style="margin-bottom:6px;">Use this when something looks off ("why is Alice\'s phone wrong?"). Find the TP person, compare their stored PCO link side-by-side, then Unlink or Replace.</div>'
          + '<div style="margin-bottom:8px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">'
          + '  <input type="text" id="pcoVerifyTpSearch" placeholder="Search TP person by name or email..." style="flex:1;min-width:240px;padding:6px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">'
          + '  <button class="pco-btn" id="pcoVerifyTpSearchBtn" style="font-size:13px;padding:5px 12px;">Search</button>'
          + '</div>'
          + '<div id="pcoVerifyTpResults" style="margin-bottom:10px;"></div>'
          + '<div id="pcoVerifyPanel" style="margin-bottom:18px;"></div>'
          + '<div class="pco-day-header" style="margin:8px 0 6px 0;padding:6px 10px;background:#1f4e79;color:#fff;border-radius:4px;font-size:13px;font-weight:700;letter-spacing:0.5px;">PENDING DATA REVIEWS (matched people, field diffs)</div>'
          + '<div id="pcoReviewsList" class="pco-empty">Loading pending changes...</div>'
          + '</div>';
        $('pcoPeopleRefresh').onclick = function(ev){
          ev.preventDefault();
          loadPendingReviews();
          refreshPeopleSyncSummary();
        };
        $('pcoLoadProposedBtn').onclick = function(){
          _proposedState = {page: 1, tier: 'all', search: ''};
          loadProposedMatches();
        };
        // Verify Link search hookup (v3.1).
        var doVerifyTpSearch = function(){
          var t = ($('pcoVerifyTpSearch').value || '').trim();
          if (!t || t.length < 2) {
            $('pcoVerifyTpResults').innerHTML = '<div class="pco-muted" style="font-size:12px;">Type at least 2 characters.</div>';
            return;
          }
          $('pcoVerifyTpResults').innerHTML = '<div class="pco-muted" style="font-size:12px;">Searching...</div>';
          ajax('search_tp_people', {search_term: t}, function(err, d){
            if (err) {
              $('pcoVerifyTpResults').innerHTML = '<div class="pco-help" style="color:#c00;">' + escHtml(err) + '</div>';
              return;
            }
            renderVerifyTpResults(d.people || []);
          });
        };
        $('pcoVerifyTpSearchBtn').onclick = doVerifyTpSearch;
        $('pcoVerifyTpSearch').onkeypress = function(ev){
          if (ev.keyCode === 13) { ev.preventDefault(); doVerifyTpSearch(); }
        };
        loadPendingReviews();
        refreshPeopleSyncSummary();
      }

      // ----- Verify Person Link (v3.1) ---------------------------

      // Shared state for the Verify panel so the Replace flow can find
      // the TP person currently being inspected.
      var _verifyState = { tpPeopleId: 0, tpName: '' };

      function renderVerifyTpResults(people) {
        var host = $('pcoVerifyTpResults');
        if (!people || !people.length) {
          host.innerHTML = '<div class="pco-help">No TP people matched. Try a different name or email.</div>';
          return;
        }
        var html = '<table class="pco-table" style="font-size:13px;"><thead><tr>'
          + '<th>TP Person</th><th>Email</th><th>PCO Link?</th><th></th>'
          + '</tr></thead><tbody>';
        for (var i = 0; i < people.length; i++) {
          var p = people[i];
          var hasLink = !!(p.pcoPersonId && String(p.pcoPersonId).length > 0);
          html += '<tr>'
            + '<td>' + escHtml(p.name || '') + ' <span class="pco-muted" style="font-size:11px;">#' + (p.peopleId || 0) + '</span></td>'
            + '<td>' + escHtml(p.email || '') + '</td>'
            + '<td>' + (hasLink
                ? '<span class="pco-pill pco-ok" title="PCO_PersonId=' + escHtml(p.pcoPersonId) + '">Linked</span>'
                : '<span class="pco-pill" style="background:#eef0f2;color:#666;">Unlinked</span>') + '</td>'
            + '<td><button class="pco-btn" data-tpid="' + (p.peopleId || 0) + '" data-tpname="' + escHtml(p.name || '').replace(/"/g, '&quot;') + '" style="font-size:12px;padding:3px 10px;">Inspect</button></td>'
            + '</tr>';
        }
        html += '</tbody></table>';
        host.innerHTML = html;
        // Wire Inspect buttons.
        var btns = host.getElementsByTagName('button');
        for (var b = 0; b < btns.length; b++) {
          btns[b].onclick = (function(btn){
            return function(){
              var tpid = parseInt(btn.getAttribute('data-tpid'), 10) || 0;
              var nm = btn.getAttribute('data-tpname') || '';
              loadVerifyDetails(tpid, nm);
            };
          })(btns[b]);
        }
      }

      function loadVerifyDetails(tpPeopleId, tpName) {
        _verifyState.tpPeopleId = tpPeopleId;
        _verifyState.tpName = tpName;
        var host = $('pcoVerifyPanel');
        host.innerHTML = '<div class="pco-empty">Loading link details for <strong>' + escHtml(tpName) + '</strong>...</div>';
        ajax('load_verify_details', {tp_people_id: tpPeopleId}, function(err, d){
          if (err) {
            host.innerHTML = '<div class="pco-help" style="color:#c00;">' + escHtml(err) + '</div>';
            return;
          }
          renderVerifyPanel(d);
        });
      }

      // Compare two strings for visual highlighting. Case-insensitive,
      // trims whitespace. Returns true if they look like the same value.
      function _vfMatch(a, b) {
        var na = (a == null ? '' : String(a)).trim().toLowerCase();
        var nb = (b == null ? '' : String(b)).trim().toLowerCase();
        if (!na && !nb) return null; // both blank -- don't color
        return na === nb;
      }
      function _vfCell(val, match) {
        var bg = '';
        if (match === true)  bg = 'background:#e6f7e6;';
        if (match === false) bg = 'background:#fdecea;';
        var text = (val == null || val === '') ? '<span class="pco-muted" style="font-style:italic;">(empty)</span>' : escHtml(String(val));
        return '<td style="' + bg + '">' + text + '</td>';
      }

      function renderVerifyPanel(d) {
        var host = $('pcoVerifyPanel');
        var tp = d.tp || {};
        var pco = d.pco || null;
        var pcoId = d.pcoPersonId || '';
        var hasLink = !!(pcoId && pcoId.length);

        var headerHtml = '<div class="pco-day-header" style="margin:8px 0 6px 0;padding:8px 12px;background:#5a3e7a;color:#fff;border-radius:4px;font-size:13px;font-weight:700;">'
          + 'LINK DETAIL: ' + escHtml(tp.name || ('TP #' + tp.peopleId))
          + (hasLink ? ' <span style="font-weight:400;opacity:0.85;">&rarr; PCO #' + escHtml(pcoId) + '</span>' : '')
          + '</div>';

        if (!hasLink) {
          host.innerHTML = headerHtml
            + '<div class="pco-card" style="background:#f8f8f8;border:1px solid #ddd;padding:12px;">'
            + '<div style="margin-bottom:10px;"><strong>' + escHtml(tp.name || '') + '</strong> has no PCO_PersonId extra value. Nothing linked.</div>'
            + '<button class="pco-btn pco-btn-primary" id="pcoVerifyRelinkBtn" style="font-size:13px;padding:5px 12px;">Link to a PCO person...</button>'
            + '</div>';
          $('pcoVerifyRelinkBtn').onclick = function(){ showRelinkSearch(tp); };
          return;
        }

        if (d.pcoFetchError) {
          host.innerHTML = headerHtml
            + '<div class="pco-card" style="background:#fff7e6;border:1px solid #f0c36d;padding:12px;">'
            + '<div style="margin-bottom:8px;color:#8a6d3b;"><strong>Could not fetch PCO record</strong> for ID ' + escHtml(pcoId) + ': ' + escHtml(d.pcoFetchError) + '</div>'
            + '<div class="pco-muted" style="font-size:12px;margin-bottom:10px;">The stored PCO_PersonId may point to a person that no longer exists, or PCO returned an error. Unlink or replace.</div>'
            + '<button class="pco-btn" id="pcoVerifyUnlinkBtn" style="font-size:13px;padding:5px 12px;margin-right:6px;background:#c0392b;color:#fff;">Unlink</button>'
            + '<button class="pco-btn" id="pcoVerifyRelinkBtn" style="font-size:13px;padding:5px 12px;">Replace with another PCO person...</button>'
            + '</div>';
          $('pcoVerifyUnlinkBtn').onclick = function(){ doUnlink(tp); };
          $('pcoVerifyRelinkBtn').onclick = function(){ showRelinkSearch(tp); };
          return;
        }

        // Side-by-side comparison table. Highlight cells red where TP
        // and PCO disagree on a non-empty value -- that's what would
        // tell a staffer "yeah, this link is wrong."
        var tpFirstDisplay = tp.nickName || tp.firstName || '';
        var pcoFirstDisplay = pco.nickname || pco.firstName || '';
        var fnameMatch = _vfMatch(tpFirstDisplay, pcoFirstDisplay);
        var lnameMatch = _vfMatch(tp.lastName, pco.lastName);
        var emailMatch = _vfMatch(tp.email, pco.email);
        var bdayMatch  = _vfMatch(tp.birthdate, pco.birthdate);

        var rowsHtml = ''
          + '<tr><th style="text-align:left;">First / Nick</th>' + _vfCell(tpFirstDisplay, fnameMatch) + _vfCell(pcoFirstDisplay, fnameMatch) + '</tr>'
          + '<tr><th style="text-align:left;">Last</th>'         + _vfCell(tp.lastName, lnameMatch)   + _vfCell(pco.lastName, lnameMatch) + '</tr>'
          + '<tr><th style="text-align:left;">Email</th>'        + _vfCell(tp.email, emailMatch)      + _vfCell(pco.email, emailMatch) + '</tr>'
          + '<tr><th style="text-align:left;">Birthdate</th>'    + _vfCell(tp.birthdate, bdayMatch)   + _vfCell(pco.birthdate, bdayMatch) + '</tr>'
          + '<tr><th style="text-align:left;">PCO Status</th>'   + '<td class="pco-muted" style="font-style:italic;">&mdash;</td>' + _vfCell(pco.status, null) + '</tr>';

        // Verdict summary so staff don't have to eyeball every row.
        var mismatches = 0;
        if (fnameMatch === false) mismatches++;
        if (lnameMatch === false) mismatches++;
        if (emailMatch === false) mismatches++;
        if (bdayMatch === false)  mismatches++;
        var verdictHtml = '';
        if (mismatches === 0) {
          verdictHtml = '<div style="background:#e6f7e6;border:1px solid #b6dfb6;padding:8px 12px;border-radius:4px;margin-bottom:10px;font-size:13px;"><strong>Looks consistent.</strong> All filled fields match between TP and PCO.</div>';
        } else {
          verdictHtml = '<div style="background:#fdecea;border:1px solid #f5c2bd;padding:8px 12px;border-radius:4px;margin-bottom:10px;font-size:13px;"><strong>' + mismatches + ' field' + (mismatches === 1 ? '' : 's') + ' disagree.</strong> Review the red cells -- if this is the wrong person, Unlink or Replace below.</div>';
        }

        host.innerHTML = headerHtml
          + '<div class="pco-card" style="background:#fff;border:1px solid #ddd;padding:12px;">'
          + verdictHtml
          + '<table class="pco-table" style="font-size:13px;width:100%;">'
          + '<thead><tr><th></th><th>TouchPoint (#' + tp.peopleId + ')</th><th>Planning Center (#' + escHtml(pcoId) + ')</th></tr></thead>'
          + '<tbody>' + rowsHtml + '</tbody>'
          + '</table>'
          + '<div style="margin-top:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">'
          + '  <button class="pco-btn" id="pcoVerifyUnlinkBtn" style="font-size:13px;padding:5px 12px;background:#c0392b;color:#fff;">Unlink</button>'
          + '  <button class="pco-btn" id="pcoVerifyRelinkBtn" style="font-size:13px;padding:5px 12px;">Replace with another PCO person...</button>'
          + '  <a href="https://people.planningcenteronline.com/people/' + escHtml(pcoId) + '" target="_blank" rel="noopener" class="pco-muted" style="font-size:12px;">Open in PCO &rarr;</a>'
          + '</div>'
          + '</div>'
          + '<div id="pcoVerifyRelinkPanel" style="margin-top:10px;"></div>';

        $('pcoVerifyUnlinkBtn').onclick = function(){ doUnlink(tp); };
        $('pcoVerifyRelinkBtn').onclick = function(){ showRelinkSearch(tp); };
      }

      function doUnlink(tp) {
        if (!confirm('Unlink ' + (tp.name || ('TP #' + tp.peopleId)) + ' from PCO?\n\nThis clears the PCO_PersonId extra value. Next sync that walks PCO will re-evaluate them as unmatched.')) {
          return;
        }
        ajax('unlink_tp_person', {tp_people_id: tp.peopleId}, function(err, d){
          if (err) { alert('Unlink failed: ' + err); return; }
          // Refresh the panel with the now-unlinked state.
          loadVerifyDetails(tp.peopleId, tp.name);
        });
      }

      function showRelinkSearch(tp) {
        var host = $('pcoVerifyRelinkPanel');
        if (!host) {
          // Unlinked case: relink UI replaces the whole verify panel.
          host = $('pcoVerifyPanel');
          host.innerHTML = '<div class="pco-day-header" style="margin:8px 0 6px 0;padding:8px 12px;background:#5a3e7a;color:#fff;border-radius:4px;font-size:13px;font-weight:700;">'
            + 'LINK ' + escHtml(tp.name || '') + ' &rarr; PCO PERSON</div>'
            + '<div id="pcoVerifyRelinkPanel" style="margin-top:10px;"></div>';
          host = $('pcoVerifyRelinkPanel');
        }
        host.innerHTML = ''
          + '<div class="pco-card" style="background:#f4f0fa;border:1px solid #c9b8e1;padding:12px;">'
          + '<div style="margin-bottom:8px;font-size:13px;"><strong>Pick a PCO person to link to ' + escHtml(tp.name || '') + ':</strong></div>'
          + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">'
          + '  <input type="text" id="pcoVerifyPcoSearch" placeholder="Search PCO by name or email..." style="flex:1;min-width:240px;padding:6px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;">'
          + '  <button class="pco-btn" id="pcoVerifyPcoSearchBtn" style="font-size:13px;padding:5px 12px;">Search PCO</button>'
          + '  <button class="pco-btn" id="pcoVerifyRelinkCancel" style="font-size:13px;padding:5px 12px;">Cancel</button>'
          + '</div>'
          + '<div id="pcoVerifyPcoResults"></div>'
          + '</div>';
        var doSearch = function(){
          var t = ($('pcoVerifyPcoSearch').value || '').trim();
          if (!t || t.length < 2) {
            $('pcoVerifyPcoResults').innerHTML = '<div class="pco-muted" style="font-size:12px;">Type at least 2 characters.</div>';
            return;
          }
          $('pcoVerifyPcoResults').innerHTML = '<div class="pco-muted" style="font-size:12px;">Searching PCO...</div>';
          ajax('search_pco_people_by_name', {search_term: t}, function(err, d){
            if (err) { $('pcoVerifyPcoResults').innerHTML = '<div class="pco-help" style="color:#c00;">' + escHtml(err) + '</div>'; return; }
            renderRelinkResults(tp, d.people || []);
          });
        };
        $('pcoVerifyPcoSearchBtn').onclick = doSearch;
        $('pcoVerifyPcoSearch').onkeypress = function(ev){
          if (ev.keyCode === 13) { ev.preventDefault(); doSearch(); }
        };
        $('pcoVerifyRelinkCancel').onclick = function(){
          $('pcoVerifyRelinkPanel').innerHTML = '';
        };
        // Pre-fill search with the TP person's name to save a keystroke.
        var seed = (tp.firstName || '') + ' ' + (tp.lastName || '');
        seed = seed.trim();
        if (seed) { $('pcoVerifyPcoSearch').value = seed; doSearch(); }
      }

      function renderRelinkResults(tp, people) {
        var host = $('pcoVerifyPcoResults');
        if (!people || !people.length) {
          host.innerHTML = '<div class="pco-help">No PCO people matched that search.</div>';
          return;
        }
        var html = '<table class="pco-table" style="font-size:13px;"><thead><tr>'
          + '<th>PCO Person</th><th>Email</th><th>Birthdate</th><th>Status</th><th></th>'
          + '</tr></thead><tbody>';
        for (var i = 0; i < people.length; i++) {
          var p = people[i];
          html += '<tr>'
            + '<td>' + escHtml(p.name || '') + ' <span class="pco-muted" style="font-size:11px;">#' + escHtml(p.pcoPersonId) + '</span></td>'
            + '<td>' + escHtml(p.email || '') + '</td>'
            + '<td>' + escHtml(p.birthdate || '') + '</td>'
            + '<td>' + escHtml(p.status || '') + '</td>'
            + '<td><button class="pco-btn pco-btn-primary" data-pcoid="' + escHtml(p.pcoPersonId) + '" data-pconame="' + escHtml(p.name || '').replace(/"/g, '&quot;') + '" style="font-size:12px;padding:3px 10px;">Link this</button></td>'
            + '</tr>';
        }
        html += '</tbody></table>';
        host.innerHTML = html;
        var btns = host.getElementsByTagName('button');
        for (var b = 0; b < btns.length; b++) {
          btns[b].onclick = (function(btn){
            return function(){
              var pcoId = btn.getAttribute('data-pcoid');
              var pcoName = btn.getAttribute('data-pconame');
              if (!confirm('Link ' + (tp.name || '') + ' (TP) to ' + pcoName + ' (PCO #' + pcoId + ')?\n\nThis overwrites any existing PCO_PersonId extra value.')) return;
              // Reuse the existing confirm_person_mapping handler.
              ajax('confirm_person_mapping', {
                tp_people_id: tp.peopleId,
                pco_person_id: pcoId
              }, function(err, d){
                if (err) { alert('Link failed: ' + err); return; }
                // Refresh the verify panel with the new link.
                loadVerifyDetails(tp.peopleId, tp.name);
              });
            };
          })(btns[b]);
        }
      }

      // ----- All People proposed matches (v2.5) ------------------

      var _proposedState = {page: 1, tier: 'all', search: ''};
      // Optional: when set, the matcher walks only this subset of PCO
      // people instead of the entire directory. Populated when opening
      // from a Preview & Sync modal so staff stay focused.
      var _proposedScope = null; // {label, peopleData: [...]}
      // Client-side cache of the last server response. Tier / search /
      // page changes all operate on this; only Reload + Clear Scope +
      // initial load actually hit the server (which re-walks PCO).
      var _proposedCache = null; // {rows, counts, totalPco, lockedCount, unmatchedCount, scoped}

      // Apply tier + search filters and paginate from cached rows.
      function _proposedFilterAndPaginate() {
        if (!_proposedCache || !_proposedCache.rows) return null;
        var rows = _proposedCache.rows;
        var tier = _proposedState.tier;
        var srch = (_proposedState.search || '').toLowerCase();
        var filtered = [];
        for (var i = 0; i < rows.length; i++) {
          var r = rows[i];
          // Tier filter (also handles skipped + unmatched specials).
          var keep = false;
          if (tier === 'all')      keep = !r.skipped;
          else if (tier === 'skipped')   keep = !!r.skipped;
          else if (tier === 'unmatched') keep = !r.skipped && r.tier === 'none';
          else                     keep = !r.skipped && r.tier === tier;
          if (!keep) continue;
          if (srch) {
            var hay = ((r.pcoName || '') + ' ' + (r.pcoEmail || '')).toLowerCase();
            if (hay.indexOf(srch) === -1) continue;
          }
          filtered.push(r);
        }
        var perPage = 50;
        var totalPages = Math.max(1, Math.ceil(filtered.length / perPage));
        if (_proposedState.page > totalPages) _proposedState.page = totalPages;
        if (_proposedState.page < 1) _proposedState.page = 1;
        var start = (_proposedState.page - 1) * perPage;
        return {
          rows: filtered.slice(start, start + perPage),
          totalFiltered: filtered.length,
          page: _proposedState.page,
          perPage: perPage,
          totalPages: totalPages,
        };
      }

      // Re-render from cached data without hitting the server. Used
      // by tier chips, search Enter / Reload, pagination, and after
      // local row removal (Apply / Skip Forever).
      function rerenderProposedFromCache() {
        var host = $('pcoProposedMatches');
        if (!host || !_proposedCache) return;
        var page = _proposedFilterAndPaginate();
        if (!page) return;
        renderProposedMatches(host, {
          success: true,
          totalPco: _proposedCache.totalPco,
          lockedCount: _proposedCache.lockedCount,
          unmatchedCount: _proposedCache.unmatchedCount,
          counts: _proposedCache.counts,
          rows: page.rows,
          page: page.page,
          perPage: page.perPage,
          totalFiltered: page.totalFiltered,
          totalPages: page.totalPages,
          scoped: _proposedCache.scoped,
        });
      }

      function loadProposedMatches() {
        var host = $('pcoProposedMatches');
        host.className = '';
        // Visible spinner + ticking elapsed counter. Server-side walk
        // typically runs 5-15 seconds for medium churches; without
        // feedback users assume the page froze. We can't show real
        // progress (one synchronous call) but the seconds counter
        // reassures the page is alive.
        var scopeLabel = (_proposedScope && _proposedScope.label) ? _proposedScope.label : '';
        var titleText = _proposedScope
          ? 'Scoring ' + (_proposedScope.peopleData.length || 0) + ' record(s) from "' + scopeLabel + '"...'
          : 'Walking the full PCO People directory and scoring TP candidates...';
        host.innerHTML = ''
          + '<div class="pco-loading-block">'
          + '  <div class="pco-spinner"></div>'
          + '  <div class="pco-loading-title">' + escHtml(titleText) + '</div>'
          + '  <div class="pco-loading-elapsed" id="pcoProposedElapsed">Elapsed: 0s</div>'
          + '  <div class="pco-muted" style="margin-top:4px;font-size:11px;">Each row needs a PCO record + a TP match check, so this scales with directory size. Usually 5-15s for medium churches; up to 30s+ for 5000+ records.</div>'
          + '</div>';
        var elapsedEl = document.getElementById('pcoProposedElapsed');
        var startMs = Date.now();
        var elapsedTimer = setInterval(function(){
          var el = document.getElementById('pcoProposedElapsed');
          if (!el) { clearInterval(elapsedTimer); return; }
          var sec = Math.floor((Date.now() - startMs) / 1000);
          el.textContent = 'Elapsed: ' + sec + 's';
        }, 1000);
        var params = {};
        if (_proposedScope && _proposedScope.peopleData && _proposedScope.peopleData.length) {
          params.subset_json = JSON.stringify(_proposedScope.peopleData);
        }
        ajax('load_all_people_proposed_matches', params, function(err, d){
          clearInterval(elapsedTimer);
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            return;
          }
          // Cache the full server response. Subsequent tier/search/page
          // changes operate on this without re-walking PCO.
          _proposedCache = {
            rows: d.rows || [],
            counts: d.counts || {},
            totalPco: d.totalPco || 0,
            lockedCount: d.lockedCount || 0,
            unmatchedCount: d.unmatchedCount || 0,
            scoped: !!d.scoped,
          };
          _proposedState.page = 1;
          rerenderProposedFromCache();
        });
      }

      function renderProposedMatches(host, d) {
        host.className = '';
        var counts = d.counts || {};
        var totalPco = d.totalPco || 0;
        var locked = d.lockedCount || 0;
        var unmatched = d.unmatchedCount || 0;
        var rows = d.rows || [];
        var s = _proposedState;
        var tierFilters = [
          ['all', 'All', unmatched - (counts.skipped || 0)],
          ['strong', 'Strong', counts.strong || 0],
          ['medium', 'Medium', counts.medium || 0],
          ['weak', 'Weak', counts.weak || 0],
          ['unmatched', 'No candidates', counts.none || 0],
          ['skipped', 'Skipped', counts.skipped || 0],
        ];
        var filterChips = '';
        for (var i = 0; i < tierFilters.length; i++) {
          var t = tierFilters[i];
          var active = (s.tier === t[0]);
          filterChips += '<button class="pco-btn ' + (active ? '' : 'pco-secondary') + ' pco-proposed-tier" data-tier="' + t[0] + '" style="font-size:12px;padding:3px 10px;">' + escHtml(t[1]) + ' (' + t[2] + ')</button>';
        }
        // Scope banner: if matcher was opened from a specific preview,
        // show what we're focused on + a button to clear back to the
        // full directory walk.
        var scopeBanner = '';
        if (_proposedScope) {
          scopeBanner = '<div style="margin-bottom:8px;padding:8px 12px;background:#0f7c840d;border-left:3px solid #0f7c84;border-radius:4px;font-size:13px;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">'
            + '<span><strong>Scoped:</strong> Matching only the ' + (_proposedScope.peopleData.length) + ' record(s) from "' + escHtml(_proposedScope.label) + '"</span>'
            + '<button class="pco-btn pco-secondary" id="pcoClearScopeBtn" style="font-size:12px;padding:3px 10px;">Clear scope (walk full directory)</button>'
            + '</div>';
        }
        var summary = scopeBanner
          + '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;align-items:center;">'
          + filterChips
          + '<span style="margin-left:auto;display:flex;gap:6px;align-items:center;">'
          + '  <input type="text" id="pcoProposedSearch" placeholder="Search name or email..." value="' + escAttr(s.search) + '" class="pco-input" style="font-size:13px;padding:3px 8px;max-width:240px;">'
          + '  <button class="pco-btn pco-secondary" id="pcoProposedReload" style="font-size:12px;padding:3px 10px;">Reload</button>'
          + '</span>'
          + '</div>'
          + '<div class="pco-muted" style="margin-bottom:8px;font-size:12px;">'
          + (_proposedScope
              ? 'Scoped subset: ' + totalPco + ' record(s) &middot; Locked (PCO_PersonId already set): ' + locked + ' &middot; Eligible for review: ' + unmatched
              : 'Total PCO: ' + totalPco + ' &middot; Locked (PCO_PersonId already set): ' + locked + ' &middot; Eligible for review: ' + unmatched)
          + '</div>';

        // Bulk-action bar.
        var bulkBar = '<div style="display:flex;gap:8px;margin-bottom:10px;align-items:center;">'
          + '<label style="font-size:13px;display:flex;align-items:center;gap:6px;cursor:pointer;"><input type="checkbox" id="pcoProposedSelectAll"> Select all on this page</label>'
          + '<button class="pco-btn" id="pcoBulkApplyBtn" style="font-size:13px;padding:5px 12px;">Apply selected</button>'
          + '<span class="pco-muted" id="pcoBulkSelectedCount" style="font-size:12px;">0 selected</span>'
          + '</div>';

        if (!rows.length) {
          host.innerHTML = summary + bulkBar + '<div>No rows match the current filter.</div>';
          wireProposedHandlers();
          return;
        }

        var tableRows = '';
        for (var r = 0; r < rows.length; r++) {
          tableRows += renderProposedRow(rows[r]);
        }
        var pageNav = '';
        if ((d.totalPages || 1) > 1) {
          pageNav = '<div style="margin-top:10px;display:flex;justify-content:space-between;align-items:center;font-size:13px;">'
            + '<span class="pco-muted">Page ' + d.page + ' of ' + d.totalPages + ' &middot; ' + d.totalFiltered + ' row(s) total</span>'
            + '<span style="display:flex;gap:6px;">'
            + (d.page > 1 ? '<button class="pco-btn pco-secondary" id="pcoProposedPrev" style="font-size:12px;padding:3px 10px;">&larr; Prev</button>' : '')
            + (d.page < d.totalPages ? '<button class="pco-btn pco-secondary" id="pcoProposedNext" style="font-size:12px;padding:3px 10px;">Next &rarr;</button>' : '')
            + '</span>'
            + '</div>';
        }
        host.innerHTML = summary + bulkBar
          + '<div style="border:1px solid #e1e4e8;border-radius:6px;overflow:hidden;">'
          + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
          + '<thead><tr style="background:#f0f3f6;text-align:left;">'
          + '<th style="padding:6px 8px;width:30px;"></th>'
          + '<th style="padding:6px 8px;">PCO Person</th>'
          + '<th style="padding:6px 8px;">Tier</th>'
          + '<th style="padding:6px 8px;">Best TP candidate(s)</th>'
          + '<th style="padding:6px 8px;width:240px;">Actions</th>'
          + '</tr></thead>'
          + '<tbody>' + tableRows + '</tbody>'
          + '</table></div>'
          + pageNav;
        wireProposedHandlers();
      }

      function renderProposedRow(r) {
        var tierColor = r.tier === 'strong' ? '#1f6b3a' : (r.tier === 'medium' ? '#1f4e79' : (r.tier === 'weak' ? '#a85a00' : '#666'));
        var pcoBlock = '<strong>' + escHtml(r.pcoName) + '</strong>'
          + ' <span class="pco-muted" style="font-size:11px;">PCO #' + escHtml(r.pcoPersonId) + '</span>'
          + (r.pcoEmail ? '<div class="pco-muted" style="font-size:12px;">' + escHtml(r.pcoEmail) + '</div>' : '')
          + (r.pcoBirthdate ? '<div class="pco-muted" style="font-size:12px;">bday ' + escHtml(r.pcoBirthdate) + '</div>' : '');
        var cands = r.candidates || [];
        var candBlock = '';
        if (!cands.length) {
          candBlock = '<span class="pco-muted">No TP candidates found.</span>';
        } else {
          for (var i = 0; i < Math.min(cands.length, 3); i++) {
            var c = cands[i];
            var tierSelf = c.score >= 90 ? 'strong' : (c.score >= 70 ? 'medium' : 'weak');
            var tierBg = tierSelf === 'strong' ? '#dff3e0' : (tierSelf === 'medium' ? '#dde6f1' : '#fde7d3');
            candBlock += '<div style="margin-bottom:4px;padding:4px 6px;background:' + tierBg + ';border-radius:3px;">'
              + '<strong>' + escHtml(c.tpName) + '</strong> <span class="pco-muted">(TP #' + c.tpPeopleId + ')</span>'
              + ' <span style="color:#333;font-size:11px;">score ' + c.score + '</span>'
              + (c.tpEmail ? '<div style="font-size:11px;color:#555;">' + escHtml(c.tpEmail) + '</div>' : '')
              + (c.tpBirthdate ? '<div style="font-size:11px;color:#555;">bday ' + escHtml(c.tpBirthdate) + '</div>' : '')
              + (c.signals && c.signals.length ? '<div style="font-size:11px;color:#555;">' + c.signals.join(', ') + '</div>' : '')
              + ' <button class="pco-btn pco-proposed-apply" data-pco-id="' + escAttr(r.pcoPersonId) + '" data-tp-id="' + c.tpPeopleId + '" style="font-size:11px;padding:2px 8px;margin-top:2px;">Apply</button>'
              + '</div>';
          }
        }
        // Selectable checkbox is only meaningful when there's at least
        // one candidate; bulk Apply targets the top candidate.
        var checkbox = cands.length
          ? '<input type="checkbox" class="pco-proposed-check" data-pco-id="' + escAttr(r.pcoPersonId) + '" data-tp-id="' + cands[0].tpPeopleId + '">'
          : '';
        // Tier stored as data attribute so client-side row removal can
        // decrement the right counter without a full reload.
        return '<tr class="pco-proposed-row" data-tier="' + escAttr(r.tier) + '" data-pco-id="' + escAttr(r.pcoPersonId) + '" style="border-top:1px solid #e1e4e8;vertical-align:top;">'
          + '<td style="padding:6px 8px;">' + checkbox + '</td>'
          + '<td style="padding:6px 8px;">' + pcoBlock + '</td>'
          + '<td style="padding:6px 8px;"><span class="pco-pill" style="background:' + tierColor + ';color:#fff;font-size:11px;">' + escHtml((r.tier === 'none' ? 'No match' : r.tier)) + '</span></td>'
          + '<td style="padding:6px 8px;">' + candBlock + '</td>'
          + '<td style="padding:6px 8px;">'
          +   '<button class="pco-btn pco-secondary pco-proposed-skip" data-pco-id="' + escAttr(r.pcoPersonId) + '" style="font-size:11px;padding:3px 10px;color:#a00;border-color:#a00;">Skip Forever</button>'
          + '</td>'
          + '</tr>';
      }

      // Remove a row from the cached dataset + re-render. Used by
      // per-row Apply / Skip Forever -- the full reload took 10+
      // seconds because the server re-walks the entire PCO directory
      // each time. With the v2.5.6+ client-side cache we just splice
      // the row out and refresh the table from cache.
      function _removeProposedRowAndDecrement(pcoId, decrementTier) {
        if (_proposedCache && _proposedCache.rows) {
          var rows = _proposedCache.rows;
          for (var i = 0; i < rows.length; i++) {
            if (rows[i].pcoPersonId === pcoId) {
              var wasSkipped = !!rows[i].skipped;
              var rowTier = rows[i].tier;
              rows.splice(i, 1);
              // Keep counts in sync so the chips don't lie.
              if (wasSkipped) {
                _proposedCache.counts.skipped = Math.max(0, (_proposedCache.counts.skipped || 0) - 1);
              } else if (rowTier && _proposedCache.counts.hasOwnProperty(rowTier)) {
                _proposedCache.counts[rowTier] = Math.max(0, _proposedCache.counts[rowTier] - 1);
              }
              _proposedCache.unmatchedCount = Math.max(0, (_proposedCache.unmatchedCount || 0) - 1);
              break;
            }
          }
        }
        rerenderProposedFromCache();
      }

      function wireProposedHandlers() {
        var host = $('pcoProposedMatches');
        // Clear-scope button (only present when a scoped subset is active).
        // This DOES re-fetch since the data source changes (scoped subset
        // vs full directory walk).
        var clearScope = $('pcoClearScopeBtn');
        if (clearScope) clearScope.onclick = function(){
          _proposedScope = null;
          _proposedState = {page: 1, tier: 'all', search: ''};
          _proposedCache = null;
          loadProposedMatches();
        };
        // Tier chip clicks -- pure client-side filter on the cached data.
        var chips = host.querySelectorAll('.pco-proposed-tier');
        for (var i = 0; i < chips.length; i++) {
          chips[i].onclick = function(){
            _proposedState.tier = this.dataset.tier;
            _proposedState.page = 1;
            rerenderProposedFromCache();
          };
        }
        var srchEl = $('pcoProposedSearch');
        if (srchEl) {
          srchEl.onkeydown = function(ev){
            if (ev.key === 'Enter') {
              _proposedState.search = this.value.trim();
              _proposedState.page = 1;
              rerenderProposedFromCache();
            }
          };
        }
        // Reload IS the "give me fresh PCO data" button -- always re-fetches.
        var reload = $('pcoProposedReload');
        if (reload) reload.onclick = function(){
          var s = $('pcoProposedSearch');
          _proposedState.search = s ? s.value.trim() : '';
          _proposedState.page = 1;
          _proposedCache = null;
          loadProposedMatches();
        };
        // Pagination is client-side too.
        var prev = $('pcoProposedPrev');
        if (prev) prev.onclick = function(){ _proposedState.page = Math.max(1, _proposedState.page - 1); rerenderProposedFromCache(); };
        var next = $('pcoProposedNext');
        if (next) next.onclick = function(){ _proposedState.page = _proposedState.page + 1; rerenderProposedFromCache(); };
        // Apply buttons (per-row, top candidate). Removes the row
        // locally instead of re-walking PCO -- that walk was taking 10+
        // seconds for a single click.
        var applies = host.querySelectorAll('.pco-proposed-apply');
        for (var a = 0; a < applies.length; a++) {
          applies[a].onclick = function(){
            var pcoId = this.dataset.pcoId;
            var tpId = parseInt(this.dataset.tpId, 10) || 0;
            if (!pcoId || !tpId) return;
            var btn = this; btn.disabled = true; btn.textContent = 'Applying...';
            ajax('apply_proposed_match', {pco_person_id: pcoId, tp_people_id: tpId}, function(err, d){
              if (err || !d || !d.success) { toast('Apply failed: ' + ((d && d.message) || err), 'err'); btn.disabled = false; btn.textContent = 'Apply'; return; }
              toast('Linked.', 'ok');
              var row = document.querySelector('.pco-proposed-row[data-pco-id="' + pcoId + '"]');
              var tier = row ? row.dataset.tier : 'all';
              _removeProposedRowAndDecrement(pcoId, tier);
            });
          };
        }
        // Skip Forever buttons -- same local-removal pattern.
        var skips = host.querySelectorAll('.pco-proposed-skip');
        for (var k = 0; k < skips.length; k++) {
          skips[k].onclick = function(){
            if (!confirm('Mark this PCO person as known-unmatched? They will stop appearing in this review until you unskip them.')) return;
            var pcoId = this.dataset.pcoId;
            ajax('skip_pco_person_forever', {pco_person_id: pcoId}, function(err, d){
              if (err || !d || !d.success) { toast('Skip failed: ' + ((d && d.message) || err), 'err'); return; }
              toast('Skipped.', 'ok');
              var row = document.querySelector('.pco-proposed-row[data-pco-id="' + pcoId + '"]');
              var tier = row ? row.dataset.tier : 'all';
              _removeProposedRowAndDecrement(pcoId, tier);
            });
          };
        }
        // Bulk select.
        var selAll = $('pcoProposedSelectAll');
        var checks = host.querySelectorAll('.pco-proposed-check');
        var updateCount = function(){
          var n = 0;
          for (var i = 0; i < checks.length; i++) if (checks[i].checked) n++;
          var lbl = $('pcoBulkSelectedCount');
          if (lbl) lbl.textContent = n + ' selected';
        };
        if (selAll) selAll.onchange = function(){
          for (var i = 0; i < checks.length; i++) checks[i].checked = selAll.checked;
          updateCount();
        };
        for (var ci = 0; ci < checks.length; ci++) checks[ci].onchange = updateCount;
        var bulk = $('pcoBulkApplyBtn');
        if (bulk) bulk.onclick = function(){
          var pairs = [];
          for (var i = 0; i < checks.length; i++) {
            if (checks[i].checked) pairs.push([checks[i].dataset.pcoId, parseInt(checks[i].dataset.tpId, 10) || 0]);
          }
          if (!pairs.length) { toast('Nothing selected.', 'err'); return; }
          if (!confirm('Apply ' + pairs.length + ' match(es)? This writes PCO_PersonId on each TP person.')) return;
          bulk.disabled = true; bulk.textContent = 'Applying...';
          ajax('bulk_apply_proposed_matches', {pairs_json: JSON.stringify(pairs)}, function(err, d){
            bulk.disabled = false; bulk.textContent = 'Apply selected';
            if (err || !d || !d.success) { toast('Bulk apply failed: ' + ((d && d.message) || err), 'err'); return; }
            toast('Applied ' + (d.applied || 0) + ', skipped ' + (d.skipped || 0) + '.', 'ok');
            // Splice each successfully-applied pair from the cache so
            // we don't re-walk PCO. Only "applied" entries are removed;
            // skipped / error rows stay so the user can retry.
            if (_proposedCache && _proposedCache.rows && d.perRow) {
              for (var i = 0; i < d.perRow.length; i++) {
                var rec = d.perRow[i];
                if (rec.status !== 'applied') continue;
                _removeProposedRowAndDecrement(rec.pcoPersonId, null);
              }
            } else {
              // Fallback if server didn't enumerate.
              rerenderProposedFromCache();
            }
          });
        };
      }

      // Updates the count line at the top after each loader finishes.
      function refreshPeopleSyncSummary() {
        var r = $('pcoReviewsList');
        var summary = $('pcoPeopleSyncSummary');
        if (!summary) return;
        var reviewCount = r && r.dataset && r.dataset.count ? parseInt(r.dataset.count, 10) || 0 : 0;
        summary.textContent = reviewCount + ' pending data review(s)';
      }

      function loadPendingReviews() {
        var host = $('pcoReviewsList');
        host.className = 'pco-empty';
        host.innerHTML = 'Loading pending changes...';
        ajax('list_pending_person_changes', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            host.dataset.count = '0';
            refreshPeopleSyncSummary();
            return;
          }
          var entries = d.entries || [];
          host.dataset.count = String(entries.length);
          refreshPeopleSyncSummary();
          if (!entries.length) {
            host.innerHTML = '<div>No pending person-data changes. Either nothing has been flagged for review, or you have no review-mode rules active.</div>';
            return;
          }
          host.className = '';
          var html = '<div class="pco-muted" style="margin-bottom:8px;">' + entries.length + ' pending change(s)</div>'
            + '<div style="display:flex;flex-direction:column;gap:8px;">';
          for (var i = 0; i < entries.length; i++) {
            var e = entries[i];
            var entryId = escAttr(e.id || '');
            html += '<div class="pco-review-row" data-entry-id="' + entryId + '" style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;">'
              + '<div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap;">'
              + '<div style="flex:1 1 320px;min-width:0;">'
              + '<div style="font-weight:700;color:#1f4e79;">' + escHtml(e.tpName || ('TP #' + e.tpPeopleId))
              +   ' <span class="pco-muted" style="font-weight:400;">&middot; TP #' + e.tpPeopleId + '</span></div>'
              + '<div style="margin-top:4px;font-size:13px;">'
              +   '<strong>' + escHtml(e.fieldLabel || e.field) + '</strong>: '
              +   '<span style="color:#a00;text-decoration:line-through;">' + escHtml(e.tpValue || '(empty)') + '</span>'
              +   ' <span class="pco-muted">&rarr;</span> '
              +   '<span style="color:#1f6b3a;font-weight:600;">' + escHtml(e.pcoValue || '(empty)') + '</span>'
              + '</div>'
              + (e.note ? '<div class="pco-muted" style="font-size:12px;margin-top:2px;">' + escHtml(e.note) + '</div>' : '')
              + '<div class="pco-muted" style="font-size:11px;margin-top:2px;">Queued ' + escHtml(e.queuedAt || '') + '</div>'
              + '</div>'
              + '<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;justify-content:flex-end;">'
              + '<button class="pco-btn pco-review-apply" data-entry-id="' + entryId + '" style="padding:4px 12px;font-size:13px;">Apply</button>'
              + '<button class="pco-btn pco-secondary pco-review-skip" data-entry-id="' + entryId + '" style="padding:4px 12px;font-size:13px;">Skip</button>'
              + '<button class="pco-btn pco-secondary pco-review-skip-forever" data-entry-id="' + entryId + '" style="padding:4px 12px;font-size:13px;color:#a00;border-color:#a00;">Skip Forever</button>'
              + '</div>'
              + '</div>'
              + '</div>';
          }
          html += '</div>';
          host.innerHTML = html;
          host.addEventListener('click', function(ev){
            var apply = ev.target.closest && ev.target.closest('.pco-review-apply');
            var skip  = ev.target.closest && ev.target.closest('.pco-review-skip');
            var sf    = ev.target.closest && ev.target.closest('.pco-review-skip-forever');
            if (apply) {
              doReviewAction('apply_person_change', {entry_id: apply.dataset.entryId});
            } else if (sf) {
              if (!confirm('Skip Forever stores a flag on this person so this field will never re-queue. Continue?')) return;
              doReviewAction('skip_person_change', {entry_id: sf.dataset.entryId, forever: '1'});
            } else if (skip) {
              doReviewAction('skip_person_change', {entry_id: skip.dataset.entryId, forever: '0'});
            }
          });
        });
      }

      function doReviewAction(action, payload) {
        ajax(action, payload, function(err, d){
          if (err || !d || !d.success) {
            toast('Action failed: ' + ((d && d.message) || err), 'err');
            return;
          }
          toast(action === 'apply_person_change' ? 'Applied.' : 'Skipped.', 'ok');
          loadPendingReviews();
        });
      }

      function loadUnmatchedPeople() {
        var host = $('pcoUnmatchedList');
        host.className = 'pco-empty';
        host.innerHTML = 'Scanning recent plans...';
        ajax('list_unmatched_people', {}, function(err, d){
          if (err || !d || !d.success) {
            host.innerHTML = '<span class="pco-pill pco-err">Error</span> ' + ((d && d.message) || err);
            host.dataset.count = '0';
            refreshPeopleSyncSummary();
            return;
          }
          var people = d.people || [];
          host.dataset.count = String(people.length);
          refreshPeopleSyncSummary();
          if (!people.length) {
            host.innerHTML = '<div>' + (d.message || 'Everyone in recent plans is already matched. Nice.') + '</div>'
              + (d.plansScanned ? '<div class="pco-muted" style="margin-top:6px;">Plans scanned: ' + d.plansScanned + '</div>' : '');
            return;
          }
          renderUnmatchedList(host, people, d);
        });
      }

      function renderUnmatchedList(host, people, meta) {
        host.className = '';
        var html = '';
        if (meta.warnings && meta.warnings.length) {
          html += '<div class="pco-card" style="background:#fff8e0;border-color:#f0c870;margin-bottom:10px;padding:8px 12px;">'
            + '<strong>Warnings:</strong><ul style="margin:4px 0 0 20px;font-size:13px;">';
          for (var w = 0; w < meta.warnings.length; w++) html += '<li>' + escHtml(meta.warnings[w]) + '</li>';
          html += '</ul></div>';
        }
        html += '<div class="pco-muted" style="margin-bottom:8px;">'
             + meta.unmatchedCount + ' unmatched person/people across ' + (meta.plansScanned || 0) + ' plan(s)'
             + '</div>'
             + '<div style="display:flex;flex-direction:column;gap:6px;">';
        for (var i = 0; i < people.length; i++) {
          html += renderUnmatchedRow(people[i], i);
        }
        html += '</div>';
        host.innerHTML = html;
        host.addEventListener('click', function(ev){
          var confirmBtn = ev.target.closest && ev.target.closest('.pco-confirm-suggestion');
          var pickerBtn = ev.target.closest && ev.target.closest('.pco-unmatched-picker-btn');
          var pickRow = ev.target.closest && ev.target.closest('.pco-unmatched-tp-pick');
          if (confirmBtn) {
            confirmUnmatchedMatch(confirmBtn.dataset.pcoId, parseInt(confirmBtn.dataset.tpId, 10) || 0, confirmBtn.dataset.pcoName, confirmBtn.dataset.tpName, confirmBtn.closest('.pco-unmatched-row'));
          } else if (pickerBtn) {
            var row = pickerBtn.closest('.pco-unmatched-row');
            var box = row.querySelector('.pco-unmatched-picker-box');
            if (box) {
              box.style.display = box.style.display === 'none' ? '' : 'none';
              if (box.style.display !== 'none') {
                var inp = box.querySelector('input');
                if (inp) {
                  inp.focus();
                  inp.select();
                  // Auto-trigger search if the seed value is long enough.
                  // Without this, the staff sees the prefilled name but
                  // gets "Type to search..." and has to edit to fire.
                  if (inp.value.trim().length >= 2) {
                    doUnmatchedSearch(inp);
                  }
                }
              }
            }
          } else if (pickRow) {
            confirmUnmatchedMatch(pickRow.dataset.pcoId, parseInt(pickRow.dataset.tpId, 10) || 0, pickRow.dataset.pcoName, pickRow.dataset.tpName, pickRow.closest('.pco-unmatched-row'));
          }
        });
        // Per-row inline search (delegated via input event).
        host.addEventListener('input', function(ev){
          if (ev.target.classList && ev.target.classList.contains('pco-unmatched-search-input')) {
            doUnmatchedSearch(ev.target);
          }
        });
      }

      var _unmatchedSearchTimer = null;
      function doUnmatchedSearch(inp) {
        clearTimeout(_unmatchedSearchTimer);
        var row = inp.closest('.pco-unmatched-row');
        var resultsEl = row.querySelector('.pco-unmatched-search-results');
        var term = inp.value.trim();
        var pcoId = inp.dataset.pcoId;
        var pcoName = inp.dataset.pcoName;
        if (term.length < 2) {
          resultsEl.innerHTML = '<span class="pco-muted">Type at least 2 characters...</span>';
          return;
        }
        resultsEl.innerHTML = '<span class="pco-muted">Searching...</span>';
        _unmatchedSearchTimer = setTimeout(function(){
          ajax('search_tp_people', {search_term: term}, function(err, d){
            if (err || !d || !d.success) {
              resultsEl.innerHTML = '<span class="pco-pill pco-err">Search failed</span>';
              return;
            }
            var people = d.people || [];
            if (!people.length) {
              resultsEl.innerHTML = '<span class="pco-muted">No matches.</span>';
              return;
            }
            var html = '<div style="border:1px solid #e1e4e8;border-radius:4px;background:#fff;max-height:240px;overflow-y:auto;">';
            for (var i = 0; i < people.length; i++) {
              var pp = people[i];
              html += '<div class="pco-unmatched-tp-pick" '
                + 'data-pco-id="' + escAttr(pcoId) + '" '
                + 'data-pco-name="' + escAttr(pcoName) + '" '
                + 'data-tp-id="' + pp.peopleId + '" '
                + 'data-tp-name="' + escAttr(pp.name) + '" '
                + 'style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;">'
                + '<strong>' + escHtml(pp.name) + '</strong> <span class="pco-muted">(TP #' + pp.peopleId + (pp.age != null ? ', ' + pp.age + 'y' : '') + (pp.gender ? ', ' + escHtml(pp.gender) : '') + ')</span>';
              if (pp.email) html += '<div class="pco-muted" style="font-size:12px;">' + escHtml(pp.email) + '</div>';
              html += '</div>';
            }
            html += '</div>';
            resultsEl.innerHTML = html;
          });
        }, 250);
      }

      function renderUnmatchedRow(p, idx) {
        var planList = '';
        if (p.plans && p.plans.length) {
          var slice = p.plans.slice(0, 3);
          var names = slice.map(function(pl){ return escHtml(pl.title) + ' (' + escHtml(pl.dateIso) + ')'; }).join(', ');
          if (p.plans.length > 3) names += ' +' + (p.plans.length - 3) + ' more';
          planList = '<div class="pco-muted" style="font-size:12px;margin-top:2px;">On: ' + names + '</div>';
        }
        var suggestion = '';
        if (p.suggestion) {
          suggestion = ''
            + '<div style="margin-top:8px;padding:8px;background:#e8f4ea;border:1px solid #b8e0c4;border-radius:4px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">'
            + '<span class="pco-pill pco-ok" style="font-size:11px;">Email match</span> '
            + '<strong>' + escHtml(p.suggestion.tpName) + '</strong> <span class="pco-muted">(TP #' + p.suggestion.tpPeopleId + ')</span>'
            + '<button class="pco-btn pco-confirm-suggestion" '
            + ' data-pco-id="' + escAttr(p.pcoPersonId) + '"'
            + ' data-tp-id="' + p.suggestion.tpPeopleId + '"'
            + ' data-pco-name="' + escAttr(p.name) + '"'
            + ' data-tp-name="' + escAttr(p.suggestion.tpName) + '"'
            + ' style="padding:4px 12px;font-size:13px;margin-left:auto;">Confirm</button>'
            + '</div>';
        } else if (p.suggestionAmbiguous) {
          suggestion = '<div style="margin-top:8px;padding:8px;background:#fff4d6;border:1px solid #f0c870;border-radius:4px;font-size:13px;">'
            + '<span class="pco-pill pco-warn" style="font-size:11px;">Multiple email matches</span> '
            + 'Use the picker below to pick the right one.'
            + '</div>';
        }
        // Manual picker -- always available so unmatched rows can resolve
        // without bouncing to the Sync Dashboard.
        var pickerSeed = p.email || p.name || '';
        var picker = ''
          + '<div style="margin-top:8px;">'
          + '  <button class="pco-btn pco-secondary pco-unmatched-picker-btn" style="padding:4px 12px;font-size:13px;">Search TouchPoint &amp; Link</button>'
          + '  <div class="pco-unmatched-picker-box" style="display:none;margin-top:8px;padding:8px;background:#fff;border:1px solid #e1e4e8;border-radius:4px;">'
          + '    <input type="text" class="pco-input pco-unmatched-search-input"'
          + '      data-pco-id="' + escAttr(p.pcoPersonId) + '"'
          + '      data-pco-name="' + escAttr(p.name) + '"'
          + '      value="' + escAttr(pickerSeed) + '"'
          + '      placeholder="Search by name or email...">'
          + '    <div class="pco-unmatched-search-results" style="margin-top:6px;"><span class="pco-muted">Type to search...</span></div>'
          + '  </div>'
          + '</div>';
        return ''
          + '<div class="pco-unmatched-row" style="border:1px solid #e1e4e8;border-radius:6px;padding:10px 12px;background:#fafbfc;">'
          + '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">'
          + '<strong>' + escHtml(p.name) + '</strong>'
          + (p.email ? ' <span class="pco-muted">' + escHtml(p.email) + '</span>' : '')
          + ' <span class="pco-muted">PCO #' + escHtml(p.pcoPersonId) + ' &middot; ' + (p.plansSeen || 0) + ' plan(s)</span>'
          + '</div>'
          + planList
          + suggestion
          + picker
          + '</div>';
      }

      function confirmUnmatchedMatch(pcoId, tpId, pcoName, tpName, rowEl) {
        ajax('confirm_person_mapping', {
          pco_person_id: pcoId,
          tp_people_id: tpId,
          pco_name: pcoName,
        }, function(err, d){
          if (err || !d || !d.success) {
            toast('Confirm failed: ' + ((d && d.message) || err), 'err');
            return;
          }
          toast('Linked ' + pcoName + ' to ' + tpName + '.', 'ok');
          // Visually mark this row as done and disable the button.
          if (rowEl) {
            rowEl.style.opacity = '0.5';
            rowEl.style.pointerEvents = 'none';
            var btn = rowEl.querySelector('.pco-confirm-suggestion');
            if (btn) { btn.disabled = true; btn.textContent = 'Linked'; }
          }
        });
      }

      // ---- Boot --------------------------------------------------

      function bind() {
        var tabs = document.querySelectorAll('.pco-tab');
        for (var i = 0; i < tabs.length; i++) {
          tabs[i].addEventListener('click', function(ev){
            selectTab(ev.target.getAttribute('data-tab'));
          });
        }
      }

      bind();
      // Land on Sync Dashboard by default; if PAT isn't configured the
      // Settings tab is the obvious next click. (Was 'settings' in v1.0.0.)
      selectTab('sync');
    })();
    </script>
    """

    if not _needs_redirect:
        # Inject Python APP_VERSION into the JS via placeholder substitution.
        # JS lives in r"""...""" so we can't break out for inline concat.
        css = css.replace('__APP_VERSION__', APP_VERSION)
        body = body.replace('__APP_VERSION__', APP_VERSION)
        js = js.replace('__APP_VERSION__', APP_VERSION)
        model.Form = css + body + js
