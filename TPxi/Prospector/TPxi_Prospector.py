#Roles=Edit
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
TPxi Prospector
===========================
A configurable prospect management tool that combines:
- Named configs with content storage persistence
- Multiple source types (involvement, tag, saved query)
- Cross-query family relationship flags
- Contact effort tracking with keyword badges
- Multi-view workspace (list, single, batch)
- Session save/resume for work continuity
- Ad-hoc display fields (person, family, medical, extra values, reg questions)

Dashboard: /PyScript/TPxi_ProspectBuilder

INSTALLATION
1. Admin > Advanced > Special Content > Python
2. Click "Add New", name it: TPxi_ProspectBuilder
3. Paste this entire file and Save
4. Navigate to /PyScriptForm/TPxi_ProspectBuilder

CONTENT STORAGE
  ProspectBuilder_Configs   - Named prospect configurations
  ProspectBuilder_Settings  - Global settings (sender, contact methods)
  ProspectBuilder_Sessions  - Saved work sessions

Written By: Ben Swaby
Version: 1.2.2
Date: June 2026

CHANGELOG
- 1.2.2 (June 2026): UX -- Action menu now shows "as <MemberType>" next to
        each involvement (green) or "no role set" (red) when no role is
        configured. Label-derived value preferred over the cached id so
        the indicator matches what the server actually does, even on
        churches with customized lookup.MemberType ids.
- 1.2.1 (June 2026): ROOT-CAUSE FIX -- "as Prospect" target actions were
        landing people as the WRONG member type on churches that
        customized lookup.MemberType. Three converging bugs, fixed
        together:
          a) HTML template hardcoded
             <option value="230">Prospect</option> (and similar for
             Group Management). 230 is the default-seed id for Prospect
             but on FBCH (and any church that renamed lookup rows) 230
             is "InActive" and Prospect lives at 311.
             FIX: dropped the hardcoded options; both dropdowns start
             with "Loading..." and get rebuilt from lookup.MemberType.
          b) pbLoadMemberTypes() skipped the rebuild on its cached path
             (modal re-open), so the hardcoded options came back.
             FIX: always call pbUpdateActionMemberTypeDropdown(),
             cached or not.
          c) Server `process_action` was calling
             `model.SetMemberType(pid, target_org, int(mt_id))` -- but
             SetMemberType takes a Description *string*, not an int.
             Passing an int (e.g. 230) makes FetchOrCreateMemberType
             match on Description="230", find nothing, and CREATE a
             new junk MemberType row with AttendanceTypeId=Member.
             Verified against bvcms-develop CmsData/API/PythonModel/
             PythonModel.Organizations.cs:235 and
             CmsData/Organization/Organization.cs:575.
             FIX: switched to LABEL-FIRST resolution. The action label
             (e.g. "Foo (#123) as Prospect") is what staff saw when
             configuring -- it's the source of truth. We parse it via
             `lookup.MemberType.Description` first and only fall back
             to the cached `memberTypeId` if the label is unparseable.
             Then we pass the resolved Description string (not the
             id) to SetMemberType.
        Diagnosis aided by PBDumpConfigs.py (read-only) which surfaces
        per-action mismatches between memberTypeId and label.
- 1.2 (June 2026): Dashboard polish + in-app auto-update via DisplayCache.
        * Conversion-rate trend chart switched from single-month buckets to
          90-day rolling per month-end, so the latest chart point matches the
          Card 6 (90-day) headline. Added "Show numbers" toggle exposing
          per-anchor conversion + prospect-involvement counts.
        * Relabeled chart help text: "pairs" -> "per involvement". Added scope
          note pointing at Group Management.
        * New action `apply_update` and DisplayCache banner — version check
          on every page load, one-click update preserving all saved data.
- 1.1 (May 2026): Added Dashboard tab (KPI cards + Chart.js trends) - outreach
        health, not email-queue health. New constant OVERDUE_DAYS controls the
        threshold for "no touch in N days". New POST action `load_dashboard_data`.
- 1.0 (March 2026): Initial release.
"""

import json
import datetime
import traceback

# ============================================================
# CONFIGURATION
# ============================================================
APP_VERSION = "1.2.2"
# --- Auto-update wiring (see TPxi/AutoUpdate/README.md) ----------------
# DisplayCache hosts a manifest at scripts.displaycache.com that lists the
# latest published version for each script. On every page load the browser
# pings it; if a newer version is published, a banner offers an in-place
# update via the workers.dev mirror (bypasses CF Bot Fight Mode for the
# server-side fetch). Content storage is preserved across updates.
DC_SCRIPT_ID = "TPxi_Prospector"
DC_API_BASE = "https://scripts.displaycache.com/api/touchpoint"
DC_API_WORKER = "https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint"
CONFIGS_KEY = "ProspectBuilder_Configs"
SETTINGS_KEY = "ProspectBuilder_Settings"
SESSIONS_KEY = "ProspectBuilder_Sessions"
GROUPS_KEY = "ProspectBuilder_Groups"
SENDERS_KEY = "ProspectBuilder_Senders"
SENDER_LOG_KEY = "ProspectBuilder_SenderLog"
TITLE = "Prospector"
CSS_PREFIX = "pb"

# Dashboard tuning. A prospect is "overdue" when neither a TaskNote contact
# effort nor a sender_log inclusion has touched them within this many days.
# Bump this if your team has a longer follow-up cadence.
OVERDUE_DAYS = 14

# Giving visibility in Journey timeline
SHOW_GIVING_IN_JOURNEY = False

# Default message body shown above the prospect list in scheduler emails.
# Falls back to this when neither sender.message_body nor settings.default_message_body is set.
DEFAULT_MESSAGE_BODY = (
    "We just wanted you to be aware that the following prospect(s) have been added to your group. "
    "We are thankful for the ways that you follow up and pursue prospects for your group."
)

# Legacy default that shipped with the first release. Used by the one-time migration
# to detect saved senders that still hold the old "Please take the following steps..." text
# so they can pick up the new central default. Match must be exact (after strip()).
_LEGACY_DEFAULT_MESSAGE_BODY = (
    "The following {Count} new prospect(s) were added as of {Date}.\n\n"
    "Please take the following steps:\n"
    "1. Review each person below\n"
    "2. Make personal contact within 48 hours (phone, text, or visit)\n"
    "3. Log your contact effort in the Prospector tool\n"
    "4. If you are unable to reach them after 3 attempts, mark as unresponsive"
)

# ============================================================
# FIELD CATALOG
# ============================================================
FIELD_CATALOG = {
    'person': [
        {'sourceField': 'Name2', 'label': 'Full Name'},
        {'sourceField': 'FirstName', 'label': 'First Name'},
        {'sourceField': 'LastName', 'label': 'Last Name'},
        {'sourceField': 'NickName', 'label': 'Nickname'},
        {'sourceField': 'PreferredName', 'label': 'Preferred Name'},
        {'sourceField': 'EmailAddress', 'label': 'Email'},
        {'sourceField': 'CellPhone', 'label': 'Cell Phone'},
        {'sourceField': 'HomePhone', 'label': 'Home Phone'},
        {'sourceField': 'Age', 'label': 'Age'},
        {'sourceField': 'BDate', 'label': 'Date of Birth'},
        {'sourceField': 'Gender', 'label': 'Gender'},
        {'sourceField': 'MaritalStatus', 'label': 'Marital Status'},
        {'sourceField': 'MemberStatus', 'label': 'Member Status'},
        {'sourceField': 'JoinDate', 'label': 'Join Date'},
        {'sourceField': 'FullAddress', 'label': 'Full Address'},
        {'sourceField': 'CampusName', 'label': 'Campus'},
    ],
    'family': [
        {'sourceField': 'FamilyDetail', 'label': 'Family Summary'},
        {'sourceField': 'SpouseName', 'label': 'Spouse Name'},
        {'sourceField': 'SpouseDetail', 'label': 'Spouse Detail'},
        {'sourceField': 'Parents', 'label': 'Parent(s)'},
        {'sourceField': 'ParentPhones', 'label': 'Parent Phone(s)'},
        {'sourceField': 'ParentEmails', 'label': 'Parent Email(s)'},
        {'sourceField': 'Children', 'label': 'Children'},
        {'sourceField': 'FamilyMembers', 'label': 'All Family Members'},
    ],
    'involvement': [
        {'sourceField': 'CurrentInvolvements', 'label': 'Current Involvements'},
    ],
    'medical': [
        {'sourceField': 'emcontact', 'label': 'Emergency Contact'},
        {'sourceField': 'emphone', 'label': 'Emergency Phone'},
        {'sourceField': 'doctor', 'label': 'Doctor'},
        {'sourceField': 'docphone', 'label': 'Doctor Phone'},
        {'sourceField': 'insurance', 'label': 'Insurance'},
        {'sourceField': 'policy', 'label': 'Policy #'},
        {'sourceField': 'MedAllergy', 'label': 'Allergies'},
        {'sourceField': 'CustodyIssue', 'label': 'Custody Issue'},
    ],
    'extravalue': [],
    'regquestion': [],
}

# ============================================================
# UNICODE SAFETY (IronPython)
# ============================================================
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

def html_escape(val):
    s = safe_str(val)
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def get_form_data(attr_name, default_value):
    try:
        value = getattr(model.Data, attr_name, None)
        if value is None or str(value).strip() == '':
            return default_value
        return str(value).strip()
    except:
        return default_value

def load_content(key, default=None):
    try:
        content = model.TextContent(key)
        if content:
            return json.loads(content)
    except:
        pass
    return default if default is not None else {}

def save_content(key, data):
    model.WriteContentText(key, json.dumps(sanitize_for_json(data), indent=2), "")

def load_configs():
    return load_content(CONFIGS_KEY, [])

def save_configs(configs):
    save_content(CONFIGS_KEY, configs)

def load_settings():
    return load_content(SETTINGS_KEY, {})

def save_settings(settings):
    save_content(SETTINGS_KEY, settings)


# ============================================================
# SCHEDULED TASKS MANAGEMENT
# ============================================================
# Markers used to safely add/remove our block in the ScheduledTasks
# content without disturbing other scripts sharing the slot. Mirrors the
# pattern in TPxi_OpsChecklists.
_SCHED_MARKER_START = "# >>> TPxi_ProspectBuilder schedule start (managed by app, do not edit) >>>"
_SCHED_MARKER_END   = "# <<< TPxi_ProspectBuilder schedule end <<<"
_SCHED_CONTENT_SLOT = "ScheduledTasks"


def get_pb_script_name():
    """Detect the install name with a three-tier resolution. Lets the
    install/uninstall flow and auto-update point ScheduledTasks /
    WriteContentPython at whatever the admin actually named the script.

    Resolution order:
      1. Posted `script_name` form param (most reliable — set by the
         browser from `window.location.pathname` and forwarded by pbAjax).
      2. `model.URL` regex parse.
      3. Hardcoded default `TPxi_ProspectBuilder` (DC_SCRIPT_ID).
    """
    try:
        if hasattr(model.Data, 'script_name'):
            sn = str(getattr(model.Data, 'script_name', '') or '').strip()
            if sn:
                return sn
    except:
        pass
    try:
        url = str(getattr(model, 'URL', '') or '')
        import re as _re
        m = _re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return DC_SCRIPT_ID


def is_scheduling_enabled():
    """Top-level kill switch for scheduled sends. Defaults to True so an
    install with no setting yet behaves the way the script always has.
    """
    s = load_settings()
    return bool(s.get('sched_enabled', True))

def load_sessions():
    return load_content(SESSIONS_KEY, [])

def save_sessions(sessions):
    save_content(SESSIONS_KEY, sessions)

def load_groups_data():
    return load_content(GROUPS_KEY, {"groups": [], "assignments": {}, "efforts": [], "changeLog": []})

def save_groups_data(data):
    # Prune efforts to prevent unbounded growth
    if len(data.get('efforts', [])) > 2000:
        data['efforts'] = data['efforts'][:1500]
    if len(data.get('changeLog', [])) > 500:
        data['changeLog'] = data['changeLog'][:400]
    save_content(GROUPS_KEY, data)

def now_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def make_id(prefix):
    return prefix + '_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')[:20]

# ============================================================
# PROSPECT SENDER ENGINE
# ============================================================

def load_senders():
    return load_content(SENDERS_KEY, [])

def save_senders(senders):
    save_content(SENDERS_KEY, senders)

def load_sender_log():
    return load_content(SENDER_LOG_KEY, [])

def save_sender_log(log):
    # Keep last 500 entries
    if len(log) > 500:
        log = log[-500:]
    save_content(SENDER_LOG_KEY, log)

def append_sender_log(entry):
    log = load_sender_log()
    log.append(entry)
    save_sender_log(log)

def resolve_message_body(sender):
    """Return the message body for a sender, falling back through:
    1. Per-sender override (sender.message_body)
    2. Central default in settings (settings.default_message_body)
    3. Hard-coded DEFAULT_MESSAGE_BODY constant
    """
    body = (sender.get('message_body') or '').strip()
    if body:
        return sender.get('message_body')
    settings = load_settings()
    body = (settings.get('default_message_body') or '').strip()
    if body:
        return settings.get('default_message_body')
    return DEFAULT_MESSAGE_BODY


def get_org_url(org_id):
    """Return an absolute URL to TouchPoint's involvement (org) page.

    Format: https://<CmsHost>/Org/<org_id>

    model.CmsHost typically returns the bare hostname (e.g. "mychurch.tpsdb.com"),
    so we prepend the scheme. Returns '' when org_id is missing.
    """
    if not org_id:
        return ''
    host = ''
    try:
        host = str(getattr(model, 'CmsHost', '') or '').strip()
    except:
        host = ''
    # Strip any scheme the host may have already (paranoia for installs that
    # set CmsHost to a full URL).
    host = host.replace('https://', '').replace('http://', '').rstrip('/')
    if host:
        return 'https://' + host + '/Org/' + str(org_id)
    # Fallback: site-relative path. Works inside TouchPoint, may not work
    # in some mail clients that don't rebase relative URLs.
    return '/Org/' + str(org_id)


def apply_merge_fields(msg, count, sender_name, org_name='', org_id=None, date_str=None):
    """Single source of truth for body merge-field substitution.

    Supported tokens:
        {Count}, {SenderName}, {OrgName}, {Date}
        {ProspectsURL}  - absolute URL to the leader's involvement (/Org/<id>)
        {ProspectsLink} - full <a> tag to that involvement

    When org_id is missing (e.g. tag/query-source sender that doesn't email
    by-org), both tokens substitute to empty string -- a broken/stub link
    is worse than no link.
    """
    if not msg:
        return msg
    if date_str is None:
        date_str = now_str()[:10]
    out = msg.replace('{Count}', str(count))
    out = out.replace('{SenderName}', sender_name or '')
    out = out.replace('{OrgName}', org_name or '')
    out = out.replace('{Date}', date_str)
    org_url = get_org_url(org_id)
    if org_url:
        link_html = ('<a href="' + org_url + '" '
                     'style="color:#2980b9;font-weight:600;text-decoration:underline">'
                     'Open ' + (org_name or 'your involvement') + ' &raquo;</a>')
        out = out.replace('{ProspectsURL}', org_url)
        out = out.replace('{ProspectsLink}', link_html)
    else:
        # No org context -- drop the tokens silently so a leader email
        # doesn't ship with a broken or pointless link.
        out = out.replace('{ProspectsURL}', '')
        out = out.replace('{ProspectsLink}', '')
    out = out.replace('\\n', '<br>').replace('\n', '<br>')
    return out

def get_sender_prospects(sender):
    """Query prospects for a sender config using its source settings.

    Returns list of {people_id, name, email, ...} for people who:
    1. Match the source query (involvement, tag, saved search)
    2. Joined/were added within the lookback window since last send
    """
    source = sender.get('source', {})
    source_type = source.get('source_type', '')
    lookback = sender.get('lookback', 'since_last')
    last_sent = sender.get('last_sent', '')

    # Determine date cutoff
    if lookback == 'since_last' and last_sent:
        try:
            cutoff = last_sent[:10]  # Use date portion of ISO string
        except:
            cutoff = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    elif lookback == 'yesterday':
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    elif lookback == 'last_7_days':
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
    elif lookback == 'last_30_days':
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
    else:
        # Default: since last send or yesterday
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')

    prospects = []

    # Build org filter based on scope
    scope = source.get('scope', source.get('source_type', 'org'))
    member_types = source.get('member_types', '')
    mt_filter = ''
    if member_types:
        mt_filter = "AND om.MemberTypeId IN ({0})".format(member_types)

    org_filter = ''
    needs_os_join = False
    if scope == 'org' or scope == 'involvement':
        org_id = source.get('org_id', '')
        if not org_id:
            return []
        org_filter = "AND om.OrganizationId = {0}".format(org_id)
    elif scope == 'program':
        prog_id = source.get('program_id', '')
        if not prog_id:
            return []
        needs_os_join = True
        org_filter = "AND os.ProgId = {0}".format(prog_id)
    elif scope == 'division':
        prog_id = source.get('program_id', '')
        div_id = source.get('division_id', '')
        needs_os_join = True
        if prog_id and div_id:
            org_filter = "AND os.ProgId = {0} AND os.DivId = {1}".format(prog_id, div_id)
        elif prog_id:
            org_filter = "AND os.ProgId = {0}".format(prog_id)
        else:
            return []
    elif scope == 'all':
        org_filter = ''
    else:
        return []

    os_join = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId" if needs_os_join else ""

    sql = '''
        SELECT DISTINCT p.PeopleId, p.Name, p.Name2, p.EmailAddress,
               p.FirstName, p.NickName, p.LastName,
               p.CellPhone, p.HomePhone,
               p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
               p.Age, p.MaritalStatusId,
               ms.Description as MemberStatus,
               om.EnrollmentDate,
               o.OrganizationName
        FROM People p
        JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
        LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
        {0}
        WHERE om.EnrollmentDate >= '{1}'
            AND o.OrganizationStatusId = 30
            AND p.IsDeceased = 0
            {2}
            {3}
        ORDER BY om.EnrollmentDate DESC
    '''.format(os_join, cutoff, org_filter, mt_filter)

    global _sender_debug_sql
    _sender_debug_sql = sql

    try:
        results = list(q.QuerySql(sql))
        for r in results:
            prospects.append({
                'people_id': r.PeopleId,
                'name': safe_str(r.Name),
                'name2': safe_str(r.Name2),
                'email': safe_str(r.EmailAddress),
                'first_name': safe_str(r.FirstName),
                'nick_name': safe_str(r.NickName),
                'last_name': safe_str(r.LastName),
                'cell_phone': safe_str(r.CellPhone) if hasattr(r, 'CellPhone') and r.CellPhone else '',
                'home_phone': safe_str(r.HomePhone) if hasattr(r, 'HomePhone') and r.HomePhone else '',
                'address': safe_str(r.PrimaryAddress) if hasattr(r, 'PrimaryAddress') and r.PrimaryAddress else '',
                'city': safe_str(r.PrimaryCity) if hasattr(r, 'PrimaryCity') and r.PrimaryCity else '',
                'state': safe_str(r.PrimaryState) if hasattr(r, 'PrimaryState') and r.PrimaryState else '',
                'zip': safe_str(r.PrimaryZip) if hasattr(r, 'PrimaryZip') and r.PrimaryZip else '',
                'age': safe_str(r.Age) if hasattr(r, 'Age') and r.Age else '',
                'member_status': safe_str(r.MemberStatus) if hasattr(r, 'MemberStatus') and r.MemberStatus else '',
                'enrollment_date': safe_str(r.EnrollmentDate) if r.EnrollmentDate else '',
                'org_name': safe_str(r.OrganizationName) if hasattr(r, 'OrganizationName') and r.OrganizationName else ''
            })
    except:
        pass

    return prospects

def get_role_recipients(role_name):
    """Get people with a specific TouchPoint role who have email addresses."""
    sql = '''
        SELECT DISTINCT p.PeopleId, p.Name, p.EmailAddress
        FROM People p
        JOIN Users u ON p.PeopleId = u.PeopleId
        JOIN UserRole ur ON u.UserId = ur.UserId
        JOIN Roles r ON ur.RoleId = r.RoleId
        WHERE r.RoleName = '{0}'
            AND p.EmailAddress IS NOT NULL
            AND p.EmailAddress != ''
            AND u.IsApproved = 1
    '''.format(role_name.replace("'", "''"))

    recipients = []
    try:
        for r in q.QuerySql(sql):
            recipients.append({
                'people_id': r.PeopleId,
                'name': safe_str(r.Name),
                'email': safe_str(r.EmailAddress)
            })
    except:
        pass
    return recipients

def _execute_per_org_sender(sender, prospects, results, dry_run, triggered_by='manual'):
    """Send per-org emails: each leader gets only THEIR group's prospects."""
    from_email = sender.get('from_email', '')
    from_name = sender.get('from_name', '')
    subject_template = sender.get('subject', '')
    sender_name = sender.get('name', 'Unnamed')
    sender_id = sender.get('id', '')
    custom_message = resolve_message_body(sender)
    recip_member_types = sender.get('recipient_member_types', '')
    source = sender.get('source', {})

    # Group prospects by org
    org_prospects = {}
    for p in prospects:
        org = p.get('org_name', 'Unknown')
        if org not in org_prospects:
            org_prospects[org] = []
        org_prospects[org].append(p)

    results['orgs_with_prospects'] = len(org_prospects)
    results['org_details'] = []

    # We need org IDs to find leaders. Query the prospect data with org IDs
    scope = source.get('scope', '')
    member_types = source.get('member_types', '')
    mt_filter = ''
    if member_types:
        mt_filter = "AND om.MemberTypeId IN ({0})".format(member_types)

    lookback = sender.get('lookback', 'since_last')
    last_sent = sender.get('last_sent', '')
    if lookback == 'since_last' and last_sent:
        cutoff = last_sent[:10]
    elif lookback == 'last_7_days':
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
    elif lookback == 'last_30_days':
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
    else:
        cutoff = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')

    # Build the org filter
    needs_os_join = False
    org_filter = ''
    if scope == 'org' or scope == 'involvement':
        org_filter = "AND om.OrganizationId = {0}".format(source.get('org_id', 0))
    elif scope == 'program':
        needs_os_join = True
        org_filter = "AND os.ProgId = {0}".format(source.get('program_id', 0))
    elif scope == 'division':
        needs_os_join = True
        prog_id = source.get('program_id', '')
        div_id = source.get('division_id', '')
        if prog_id and div_id:
            org_filter = "AND os.ProgId = {0} AND os.DivId = {1}".format(prog_id, div_id)
        elif prog_id:
            org_filter = "AND os.ProgId = {0}".format(prog_id)

    os_join = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId" if needs_os_join else ""

    # Get prospects grouped by org with org IDs
    recip_mt_filter = ''
    if recip_member_types:
        recip_mt_filter = "AND om2.MemberTypeId IN ({0})".format(recip_member_types)

    # Query: for each org with new prospects, find the leaders
    orgs_sql = '''
        SELECT DISTINCT o.OrganizationId, o.OrganizationName
        FROM OrganizationMembers om
        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
        {0}
        WHERE om.EnrollmentDate >= '{1}'
            AND o.OrganizationStatusId = 30
            {2}
            {3}
    '''.format(os_join, cutoff, org_filter, mt_filter)

    total_recipients = 0
    total_emails = 0

    try:
        org_rows = list(q.QuerySql(orgs_sql))
    except:
        org_rows = []

    # Determine queued_by once
    queued_by = None
    if from_email:
        try:
            sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}'".format(from_email.replace("'", "''"))
            result = list(q.QuerySql(sql))
            if result:
                queued_by = result[0].PeopleId
        except:
            pass
    if not queued_by:
        try:
            queued_by = model.UserPeopleId
        except:
            pass

    for org_row in org_rows:
        org_id = org_row.OrganizationId
        org_name = safe_str(org_row.OrganizationName)

        # Get prospects for this org
        org_p = [p for p in prospects if p.get('org_name', '') == org_name]
        if not org_p:
            continue

        # Get leaders/recipients for this org
        leader_sql = '''
            SELECT DISTINCT p.PeopleId, p.Name, p.EmailAddress
            FROM People p
            JOIN OrganizationMembers om2 ON p.PeopleId = om2.PeopleId
            WHERE om2.OrganizationId = {0}
                AND p.EmailAddress IS NOT NULL AND p.EmailAddress != ''
                AND p.IsDeceased = 0
                {1}
        '''.format(org_id, recip_mt_filter)

        try:
            leaders = list(q.QuerySql(leader_sql))
        except:
            leaders = []

        if not leaders:
            continue

        org_detail = {
            'org_name': org_name,
            'org_id': org_id,
            'prospects': len(org_p),
            'leaders': len(leaders),
            'leader_names': [safe_str(l.Name) for l in leaders]
        }
        results['org_details'].append(org_detail)

        # Build email for this org's leaders
        ef = sender.get('email_fields', {})
        # Defaults if not configured
        if not ef:
            ef = {'email': True, 'cell_phone': True, 'home_phone': False, 'address': True, 'age': True, 'member_status': False, 'enrollment_date': True, 'person_link': True}

        td = 'style="padding:6px 8px;border-bottom:1px solid #eee;font-size:13px"'
        th = 'style="padding:8px;text-align:left;border-bottom:2px solid #ddd;font-size:13px;background:#f5f5f5"'

        # Build dynamic header
        headers = ['Name']
        has_contact = ef.get('email') or ef.get('cell_phone') or ef.get('home_phone')
        has_details = ef.get('age') or ef.get('member_status') or ef.get('address')
        if has_contact:
            headers.append('Contact')
        if has_details:
            headers.append('Details')
        if ef.get('enrollment_date'):
            headers.append('Added')

        prospect_list_html = '<table style="width:100%;border-collapse:collapse;font-family:Segoe UI,sans-serif">'
        prospect_list_html += '<tr>' + ''.join('<th {0}>{1}</th>'.format(th, h) for h in headers) + '</tr>'

        for p in org_p:
            # Name column
            if ef.get('person_link'):
                name_cell = '<a href="/Person2/{0}" style="font-weight:600;color:#2c3e50">{1}</a>'.format(p['people_id'], p['name'])
            else:
                name_cell = '<span style="font-weight:600">{0}</span>'.format(p['name'])

            # Contact column
            contact = ''
            if has_contact:
                parts = []
                if ef.get('email') and p.get('email'):
                    parts.append('<a href="mailto:{0}">{0}</a>'.format(p['email']))
                if ef.get('cell_phone') and p.get('cell_phone'):
                    parts.append('<a href="tel:{0}">{0}</a> <span style="color:#999;font-size:11px">cell</span>'.format(p['cell_phone']))
                if ef.get('home_phone') and p.get('home_phone') and p.get('home_phone') != p.get('cell_phone'):
                    parts.append('{0} <span style="color:#999;font-size:11px">home</span>'.format(p['home_phone']))
                contact = '<br>'.join(parts) if parts else '<span style="color:#999">-</span>'

            # Details column
            details = ''
            if has_details:
                detail_parts = []
                if ef.get('age') and p.get('age'):
                    detail_parts.append('Age {0}'.format(p['age']))
                if ef.get('member_status') and p.get('member_status'):
                    detail_parts.append(p['member_status'])
                detail_line = ' &middot; '.join(detail_parts) if detail_parts else ''
                if ef.get('address') and p.get('address'):
                    addr = p['address']
                    if p.get('city'):
                        addr += ', {0}'.format(p['city'])
                    if p.get('state'):
                        addr += ', {0}'.format(p['state'])
                    if detail_line:
                        detail_line += '<br>'
                    detail_line += '<span style="color:#666;font-size:12px">{0}</span>'.format(addr)
                details = detail_line or '-'

            # Build row
            row = '<td {0}>{1}</td>'.format(td, name_cell)
            if has_contact:
                row += '<td {0}>{1}</td>'.format(td, contact)
            if has_details:
                row += '<td {0}>{1}</td>'.format(td, details)
            if ef.get('enrollment_date'):
                row += '<td {0}>{1}</td>'.format(td, p['enrollment_date'] or '-')

            prospect_list_html += '<tr>{0}</tr>'.format(row)
        prospect_list_html += '</table>'

        email_subject = subject_template or '{0} - {1} New Prospect(s)'.format(org_name, len(org_p))
        email_subject = email_subject.replace('{Count}', str(len(org_p))).replace('{OrgName}', org_name).replace('{SenderName}', sender_name)

        email_body = '<div style="font-family:Segoe UI,sans-serif;max-width:700px">'
        email_body += '<h2 style="color:#333;margin-bottom:4px">{0}</h2>'.format(org_name)
        email_body += '<p style="color:#666;margin-top:0">{0} new prospect(s) as of {1}</p>'.format(len(org_p), now_str()[:10])

        if custom_message:
            msg = apply_merge_fields(custom_message, len(org_p), sender_name,
                                     org_name=org_name, org_id=org_id)
            email_body += '<div style="background:#f8f9fa;border-left:4px solid #3498db;padding:12px 16px;margin:16px 0;color:#333;line-height:1.6">{0}</div>'.format(msg)

        email_body += prospect_list_html
        email_body += '</div>'

        # Capture first org's email as sample for preview
        if 'sample_email_html' not in results:
            results['sample_email_html'] = email_body
            results['sample_email_subject'] = email_subject
            results['sample_email_from'] = '{0} <{1}>'.format(from_name, from_email) if from_name else from_email
            results['sample_email_to'] = ', '.join([safe_str(l.Name) + ' <' + safe_str(l.EmailAddress) + '>' for l in leaders[:3]])
            if len(leaders) > 3:
                results['sample_email_to'] += ' (+{0} more)'.format(len(leaders) - 3)

        # Send to each leader
        if not dry_run:
            for leader in leaders:
                try:
                    qb = queued_by or leader.PeopleId
                    model.Email(
                        "PeopleId={0}".format(leader.PeopleId),
                        qb,
                        from_email or '',
                        from_name or '',
                        email_subject,
                        email_body,
                        ""
                    )
                    total_emails += 1
                except Exception as e:
                    results['errors'].append('Failed to email {0} for {1}: {2}'.format(safe_str(leader.Name), org_name, safe_str(e)))

        total_recipients += len(leaders)

    results['recipients'] = total_recipients
    results['emails_sent'] = total_emails if not dry_run else 0
    results['message'] = '{0} org(s) with prospects, {1} leader(s) to receive emails about {2} prospect(s)'.format(
        len(org_rows), total_recipients, results['prospects_found'])

    if not dry_run:
        results['message'] = 'Sent {0} email(s) to {1} leader(s) across {2} org(s)'.format(
            total_emails, total_recipients, len(org_rows))
        senders = load_senders()
        for s in senders:
            if s.get('id') == sender_id:
                s['last_sent'] = datetime.datetime.now().isoformat()
                s['last_sent_count'] = results['prospects_found']
                break
        save_senders(senders)

    append_sender_log({
        'timestamp': now_str(),
        'sender_id': sender_id,
        'sender_name': sender_name,
        'dry_run': dry_run,
        'triggered_by': triggered_by,
        'frequency': sender.get('frequency', ''),
        'prospects': results['prospects_found'],
        'recipients': total_recipients,
        'emails_sent': total_emails if not dry_run else 0,
        'errors': len(results['errors'])
    })

    return results

def execute_sender(sender, dry_run=False, triggered_by='manual'):
    """Execute a prospect sender — find new prospects and email role recipients.

    triggered_by: 'manual' (user clicked Send Now), 'batch' (scheduler),
                  'preview' (dry-run preview), 'oneoff' (manual one-off send).
    Returns dict with results.
    """
    sender_id = sender.get('id', '')
    sender_name = sender.get('name', 'Unnamed')
    template_name = sender.get('template_name', '')
    template_title = sender.get('template_title', '')
    from_email = sender.get('from_email', '')
    from_name = sender.get('from_name', '')
    subject = sender.get('subject', '')
    roles = sender.get('roles', [])
    send_to_mode = sender.get('send_to_mode', 'roles')  # 'roles' or 'specific_people'
    specific_people = sender.get('specific_people', [])

    results = {
        'sender_id': sender_id,
        'sender_name': sender_name,
        'dry_run': dry_run,
        'timestamp': now_str(),
        'prospects_found': 0,
        'recipients': 0,
        'emails_sent': 0,
        'errors': [],
        'prospect_names': []
    }

    # Step 1: Find new prospects
    prospects = get_sender_prospects(sender)
    results['prospects_found'] = len(prospects)
    results['prospect_names'] = [p['name'] for p in prospects[:20]]
    try:
        results['debug_sql'] = _sender_debug_sql
    except:
        results['debug_sql'] = 'N/A'
    results['debug_source'] = sender.get('source', {})

    if len(prospects) == 0:
        results['message'] = 'No new prospects found since last send'
        return results

    # For involvement_members mode: group prospects by org and send per-org emails
    if send_to_mode == 'involvement_members':
        return _execute_per_org_sender(sender, prospects, results, dry_run, triggered_by)

    # Step 2: Get recipients
    recipients = []
    if send_to_mode == 'roles':
        for role in roles:
            recipients.extend(get_role_recipients(role))
    elif send_to_mode == 'involvement_members':
        # Recipients are members of the same involvement(s) with specific member types
        source = sender.get('source', {})
        scope = source.get('scope', source.get('source_type', ''))
        recip_member_types = sender.get('recipient_member_types', '')
        recip_mt_filter = ''
        if recip_member_types:
            recip_mt_filter = "AND om.MemberTypeId IN ({0})".format(recip_member_types)

        # Build org filter matching the same scope as prospects
        recip_org_filter = ''
        recip_os_join = ''
        if scope == 'org' or scope == 'involvement':
            org_id = source.get('org_id', '')
            if org_id:
                recip_org_filter = "AND om.OrganizationId = {0}".format(org_id)
        elif scope == 'program':
            prog_id = source.get('program_id', '')
            if prog_id:
                recip_os_join = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId"
                recip_org_filter = "AND os.ProgId = {0}".format(prog_id)
        elif scope == 'division':
            prog_id = source.get('program_id', '')
            div_id = source.get('division_id', '')
            recip_os_join = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId"
            if prog_id and div_id:
                recip_org_filter = "AND os.ProgId = {0} AND os.DivId = {1}".format(prog_id, div_id)
            elif prog_id:
                recip_org_filter = "AND os.ProgId = {0}".format(prog_id)

        recip_sql = '''
            SELECT DISTINCT p.PeopleId, p.Name, p.EmailAddress
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            {0}
            WHERE o.OrganizationStatusId = 30
                AND p.EmailAddress IS NOT NULL AND p.EmailAddress != ''
                AND p.IsDeceased = 0
                {1}
                {2}
        '''.format(recip_os_join, recip_org_filter, recip_mt_filter)
        try:
            for r in q.QuerySql(recip_sql):
                recipients.append({
                    'people_id': r.PeopleId,
                    'name': safe_str(r.Name),
                    'email': safe_str(r.EmailAddress)
                })
        except:
            pass
    elif send_to_mode == 'specific_people':
        for pid in specific_people:
            try:
                person = model.GetPerson(int(pid))
                if person and person.EmailAddress:
                    recipients.append({
                        'people_id': person.PeopleId,
                        'name': person.Name or '',
                        'email': person.EmailAddress or ''
                    })
            except:
                pass

    # Deduplicate recipients by PeopleId
    seen = set()
    unique_recipients = []
    for r in recipients:
        if r['people_id'] not in seen:
            seen.add(r['people_id'])
            unique_recipients.append(r)
    recipients = unique_recipients
    results['recipients'] = len(recipients)

    if len(recipients) == 0:
        results['errors'].append('No recipients found for configured roles')
        return results

    # Step 3: Build email content
    custom_message = resolve_message_body(sender)

    # Build prospect table
    prospect_list_html = '<table style="width:100%;border-collapse:collapse;font-family:Segoe UI,sans-serif">'
    prospect_list_html += '<tr style="background:#f5f5f5"><th style="padding:8px;text-align:left;border-bottom:2px solid #ddd">Name</th><th style="padding:8px;text-align:left;border-bottom:2px solid #ddd">Email</th><th style="padding:8px;text-align:left;border-bottom:2px solid #ddd">Involvement</th><th style="padding:8px;text-align:left;border-bottom:2px solid #ddd">Date</th></tr>'
    for p in prospects:
        prospect_list_html += '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee"><a href="/Person2/{0}">{1}</a></td><td style="padding:6px 8px;border-bottom:1px solid #eee">{2}</td><td style="padding:6px 8px;border-bottom:1px solid #eee">{3}</td><td style="padding:6px 8px;border-bottom:1px solid #eee">{4}</td></tr>'.format(
            p['people_id'], p['name'], p['email'], p.get('org_name', ''), p['enrollment_date'] or '-')
    prospect_list_html += '</table>'

    email_subject = subject or '{0} - {1} New Prospect(s)'.format(sender_name, len(prospects))
    email_body = '<div style="font-family:Segoe UI,sans-serif;max-width:700px">'
    email_body += '<h2 style="color:#333;margin-bottom:4px">{0}</h2>'.format(sender_name)
    email_body += '<p style="color:#666;margin-top:0">{0} new prospect(s) found as of {1}</p>'.format(len(prospects), now_str())

    # Custom message (instructions/next steps)
    if custom_message:
        # If this sender pulls from a single involvement, surface its
        # OrgId so {ProspectsLink} resolves to /Org/<id>. For tag/query
        # sources we have no single org -- the helper drops the token.
        src = sender.get('source', {}) or {}
        merge_org_id = src.get('orgId') if (src.get('pb_type') == 'involvement') else None
        merge_org_name = src.get('orgName', '') if merge_org_id else ''
        msg = apply_merge_fields(custom_message, len(prospects), sender_name,
                                 org_name=merge_org_name, org_id=merge_org_id)
        email_body += '<div style="background:#f8f9fa;border-left:4px solid #3498db;padding:12px 16px;margin:16px 0;color:#333;line-height:1.6">{0}</div>'.format(msg)

    email_body += prospect_list_html
    email_body += '</div>'

    # Step 4: Send emails
    if not dry_run:
        # Determine queued_by
        queued_by = None
        if from_email:
            try:
                sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}'".format(
                    from_email.replace("'", "''"))
                result = list(q.QuerySql(sql))
                if result:
                    queued_by = result[0].PeopleId
            except:
                pass
        if not queued_by:
            try:
                queued_by = model.UserPeopleId
            except:
                pass
        if not queued_by and recipients:
            queued_by = recipients[0]['people_id']

        for recipient in recipients:
            try:
                query_str = "PeopleId={0}".format(recipient['people_id'])
                model.Email(
                    query_str,
                    queued_by,
                    from_email or '',
                    from_name or '',
                    email_subject,
                    email_body,
                    ""
                )
                results['emails_sent'] += 1
            except Exception as e:
                results['errors'].append('Failed to email {0}: {1}'.format(recipient['name'], safe_str(e)))

        # Update last_sent timestamp on the sender
        senders = load_senders()
        for s in senders:
            if s.get('id') == sender_id:
                s['last_sent'] = datetime.datetime.now().isoformat()
                s['last_sent_count'] = len(prospects)
                break
        save_senders(senders)

    results['message'] = 'Sent {0} email(s) to {1} recipient(s) about {2} prospect(s)'.format(
        results['emails_sent'], results['recipients'], results['prospects_found'])

    # Log the execution
    append_sender_log({
        'timestamp': now_str(),
        'sender_id': sender_id,
        'sender_name': sender_name,
        'dry_run': dry_run,
        'triggered_by': triggered_by,
        'frequency': sender.get('frequency', ''),
        'prospects': results['prospects_found'],
        'recipients': results['recipients'],
        'emails_sent': results['emails_sent'],
        'errors': len(results['errors'])
    })

    return results

def run_scheduled_senders():
    """Called from ScheduledTasks. Checks each sender's schedule and runs if due.

    Smart scheduling: uses last_sent timestamp + frequency to decide if a sender
    should run. Handles TouchPoint's unreliable scheduler timing by checking
    if the scheduled window has been missed and running anyway.
    """
    senders = load_senders()
    now = datetime.datetime.now()
    results = []

    for sender in senders:
        if not sender.get('enabled', False):
            continue

        frequency = sender.get('frequency', 'daily')
        target_hour = int(sender.get('target_hour', 7))
        last_sent = sender.get('last_sent', '')

        # Determine if this sender should run now
        should_run = False

        if not last_sent:
            # Never run before — run now
            should_run = True
        else:
            try:
                last_dt = datetime.datetime.strptime(last_sent[:19], '%Y-%m-%dT%H:%M:%S')
            except:
                try:
                    last_dt = datetime.datetime.strptime(last_sent[:19], '%Y-%m-%d %H:%M:%S')
                except:
                    last_dt = now - datetime.timedelta(days=999)

            if frequency == 'daily':
                # Run if: last sent was before today's target hour AND it's now past target hour
                today_target = now.replace(hour=target_hour, minute=0, second=0, microsecond=0)
                if last_dt < today_target and now >= today_target:
                    should_run = True

            elif frequency == 'weekly':
                target_day = int(sender.get('target_day', 1))  # 0=Mon, 6=Sun
                # Find this week's target datetime
                days_until_target = (target_day - now.weekday()) % 7
                if days_until_target == 0 and now.hour >= target_hour:
                    # Today is the target day and we're past the hour
                    this_week_target = now.replace(hour=target_hour, minute=0, second=0, microsecond=0)
                elif days_until_target == 0:
                    # Today is target day but too early — check last week
                    this_week_target = (now - datetime.timedelta(days=7)).replace(hour=target_hour, minute=0, second=0, microsecond=0)
                else:
                    this_week_target = (now - datetime.timedelta(days=(7 - days_until_target))).replace(hour=target_hour, minute=0, second=0, microsecond=0)

                if last_dt < this_week_target and now >= this_week_target:
                    should_run = True

            elif frequency == 'monthly':
                target_dom = int(sender.get('target_day_of_month', 1))
                # This month's target
                try:
                    this_month_target = now.replace(day=target_dom, hour=target_hour, minute=0, second=0, microsecond=0)
                except:
                    # Day doesn't exist this month (e.g., 31st in Feb)
                    this_month_target = now.replace(day=28, hour=target_hour, minute=0, second=0, microsecond=0)

                if last_dt < this_month_target and now >= this_month_target:
                    should_run = True

        if should_run:
            try:
                result = execute_sender(sender, triggered_by='batch')
                results.append(result)
            except Exception as e:
                results.append({
                    'sender_id': sender.get('id'),
                    'sender_name': sender.get('name'),
                    'error': safe_str(e)
                })

    return results

def migrate_legacy_default_message():
    """One-time migration: clear message_body on saved senders that still hold
    the original 'Please take the following steps...' default. Cleared senders
    will then pick up the new central default from settings (or DEFAULT_MESSAGE_BODY).
    Idempotent: settings flag prevents re-running.
    """
    settings = load_settings()
    if settings.get('message_body_migrated_v1'):
        return 0

    legacy = _LEGACY_DEFAULT_MESSAGE_BODY.strip()
    senders = load_senders()
    cleared = 0
    for s in senders:
        body = s.get('message_body') or ''
        if body.strip() == legacy:
            s['message_body'] = ''
            cleared += 1

    if cleared:
        save_senders(senders)
    settings['message_body_migrated_v1'] = True
    save_settings(settings)
    return cleared

# ============================================================
# GROUP METRICS (attendance/conversion stats per involvement)
# ============================================================
# Cached in content storage; computed by the batch job (weekly full refresh
# Monday >= 3am, daily fill-in for new groups missing metrics). Manual refresh
# is available per-group from the Health view.

GROUP_METRICS_KEY = "ProspectBuilder_GroupMetrics"
METRIC_WINDOWS = [90, 180, 365]  # days

def load_group_metrics_all():
    return load_content(GROUP_METRICS_KEY, {})

def save_group_metrics_all(data):
    save_content(GROUP_METRICS_KEY, data)

def get_group_metrics(group_id):
    return load_group_metrics_all().get(group_id, {})

def set_group_metrics(group_id, metrics):
    data = load_group_metrics_all()
    data[group_id] = metrics
    save_group_metrics_all(data)

def _safe_int_csv(items):
    """Coerce a list of mixed-type IDs to a comma-separated SQL-safe int list."""
    out = []
    for x in (items or []):
        try:
            out.append(str(int(x)))
        except:
            pass
    return ','.join(out) if out else '0'

def _group_scope_org_filter(group, om_alias='om', o_alias='o', os_alias='os'):
    """Return (org_filter_sql, needs_os_join) for the scope of a saved group.

    Mirrors the patterns in get_sender_prospects/_execute_per_org_sender so
    metric queries see the same set of involvements the Health view does.
    """
    level = group.get('level', 'program')
    prog_id = group.get('programId') or 0
    div_id = group.get('divisionId') or 0
    org_id = group.get('orgId') or 0

    if level == 'involvement' and org_id:
        return ("AND {0}.OrganizationId = {1}".format(om_alias, int(org_id)), False)
    if level == 'division' and prog_id and div_id:
        return ("AND {0}.ProgId = {1} AND {0}.DivId = {2}".format(os_alias, int(prog_id), int(div_id)), True)
    if level == 'program' and prog_id:
        return ("AND {0}.ProgId = {1}".format(os_alias, int(prog_id)), True)
    return ('', False)

def compute_group_metrics(group, windows=None):
    """Compute prospect/converted/attended counts for each involvement in the
    group's scope, across the requested rolling windows (in days).

    Returns:
        {
          'computedAt': ISO timestamp,
          'byOrg': {
            '<orgId>': {
              'orgName': str,
              'windows': {'90': {prospects, converted, attended, conversionRate}, ...}
            }
          }
        }

    Definitions:
      - Prospects in window: distinct people with an EnrollmentTransaction in
        this org with MemberTypeId in the group's prospect types, whose
        membership period overlaps the window.
      - Attended: of the prospects in window, people with AttendanceFlag=1
        attend rows for this org within the window.
      - Converted: of the prospects in window, people who attended this org
        in window with AttendanceFlag=1 AND AttendanceTypeId in the
        "converted" set (configurable via settings.converted_attend_type_ids,
        default [30] = Member). Detecting the AttendanceTypeId transition
        is more responsive than waiting on a MemberType change because the
        type is set at each meeting.
    """
    if windows is None:
        windows = METRIC_WINDOWS

    prospect_types_csv = _safe_int_csv(group.get('memberTypes', []) or [311])
    # Per-group override falls back to settings, then to [30] (Member).
    converted_types = group.get('convertedAttendTypeIds') or load_settings().get('converted_attend_type_ids') or [30]
    converted_types_csv = _safe_int_csv(converted_types)
    org_filter, needs_os_join = _group_scope_org_filter(group, 'om', 'o', 'os')

    # Find orgs in scope (active only). Use OrganizationMembers as a discovery
    # anchor so we only pay for orgs that have ever had relevant memberships.
    os_join_for_discovery = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId" if needs_os_join else ""
    orgs_sql = """
        SELECT DISTINCT o.OrganizationId, o.OrganizationName
        FROM Organizations o
        {os_join}
        WHERE o.OrganizationStatusId = 30
          {scope_filter}
        ORDER BY o.OrganizationName
    """.format(
        os_join=os_join_for_discovery,
        scope_filter=org_filter.replace('om.', 'o.') if 'om.' in org_filter else org_filter
    )

    org_rows = []
    try:
        for r in q.QuerySql(orgs_sql):
            org_rows.append((int(r.OrganizationId), safe_str(r.OrganizationName)))
    except:
        pass

    if not org_rows:
        return {
            'computedAt': datetime.datetime.now().isoformat(),
            'byOrg': {}
        }

    org_ids_csv = ','.join(str(oid) for oid, _ in org_rows)

    by_org = {}
    for oid, oname in org_rows:
        by_org[str(oid)] = {
            'orgName': oname,
            'windows': {str(w): {'prospects': 0, 'converted': 0, 'engaged': 0,
                                 'noShow': 0, 'dropped': 0, 'conversionRate': 0.0}
                        for w in windows}
        }

    # One query per window — returns the 4-state funnel counts (Converted,
    # Engaged, No-show, Dropped) per org. Definitions:
    #   - Converted: attended in window with AttendanceTypeId in converted set
    #   - Engaged:   attended in window, but never with a converted AT
    #   - No-show:   no attendance in window, still on the current roster
    #   - Dropped:   no attendance in window, no longer on the current roster
    # These four are mutually exclusive and sum to Prospects.
    for w in windows:
        metric_sql = """
            ;WITH PiW AS (
                SELECT et.OrganizationId, et.PeopleId
                FROM EnrollmentTransaction et
                WHERE et.OrganizationId IN ({org_ids})
                  AND et.MemberTypeId IN ({prospect_types})
                  AND et.TransactionStatus = 0
                  AND et.EnrollmentDate <= GETDATE()
                  AND (et.InactiveDate IS NULL OR et.InactiveDate >= DATEADD(day, -{window}, GETDATE()))
                GROUP BY et.OrganizationId, et.PeopleId
            ),
            AnyAttend AS (
                SELECT DISTINCT piw.OrganizationId, piw.PeopleId
                FROM PiW piw
                JOIN Attend a WITH (NOLOCK)
                  ON a.OrganizationId = piw.OrganizationId
                 AND a.PeopleId = piw.PeopleId
                 AND a.AttendanceFlag = 1
                 AND a.MeetingDate >= DATEADD(day, -{window}, GETDATE())
            ),
            ConvAttend AS (
                SELECT DISTINCT piw.OrganizationId, piw.PeopleId
                FROM PiW piw
                JOIN Attend a WITH (NOLOCK)
                  ON a.OrganizationId = piw.OrganizationId
                 AND a.PeopleId = piw.PeopleId
                 AND a.AttendanceFlag = 1
                 AND a.AttendanceTypeId IN ({converted_types})
                 AND a.MeetingDate >= DATEADD(day, -{window}, GETDATE())
            ),
            CurrentRoster AS (
                SELECT om.OrganizationId, om.PeopleId
                FROM OrganizationMembers om
                WHERE om.OrganizationId IN ({org_ids})
            )
            SELECT
                piw.OrganizationId,
                COUNT(*) AS prospects,
                SUM(CASE WHEN ca.PeopleId IS NOT NULL THEN 1 ELSE 0 END) AS converted,
                SUM(CASE WHEN aa.PeopleId IS NOT NULL AND ca.PeopleId IS NULL THEN 1 ELSE 0 END) AS engaged,
                SUM(CASE WHEN aa.PeopleId IS NULL AND cr.PeopleId IS NOT NULL THEN 1 ELSE 0 END) AS no_show,
                SUM(CASE WHEN aa.PeopleId IS NULL AND cr.PeopleId IS NULL THEN 1 ELSE 0 END) AS dropped
            FROM PiW piw
            LEFT JOIN AnyAttend aa
              ON aa.OrganizationId = piw.OrganizationId AND aa.PeopleId = piw.PeopleId
            LEFT JOIN ConvAttend ca
              ON ca.OrganizationId = piw.OrganizationId AND ca.PeopleId = piw.PeopleId
            LEFT JOIN CurrentRoster cr
              ON cr.OrganizationId = piw.OrganizationId AND cr.PeopleId = piw.PeopleId
            GROUP BY piw.OrganizationId
        """.format(
            org_ids=org_ids_csv,
            prospect_types=prospect_types_csv,
            converted_types=converted_types_csv,
            window=int(w)
        )

        try:
            for r in q.QuerySql(metric_sql):
                key = str(int(r.OrganizationId))
                if key not in by_org:
                    continue
                prospects = int(r.prospects or 0)
                converted = int(r.converted or 0)
                engaged = int(r.engaged or 0)
                no_show = int(r.no_show or 0)
                dropped = int(r.dropped or 0)
                rate = round((float(converted) / prospects) * 100.0, 1) if prospects > 0 else 0.0
                by_org[key]['windows'][str(w)] = {
                    'prospects': prospects,
                    'converted': converted,
                    'engaged': engaged,
                    'noShow': no_show,
                    'dropped': dropped,
                    'conversionRate': rate
                }
        except Exception as e:
            # Log but don't abort other windows; record an error marker on the
            # group result so the UI can surface it.
            by_org['_error_window_' + str(w)] = safe_str(e)

    return {
        'computedAt': datetime.datetime.now().isoformat(),
        'byOrg': by_org
    }

def run_group_metrics_batch():
    """Daily fill-in + weekly full refresh of group metrics.

    Daily: compute metrics for any saved group missing them.
    Weekly: on Monday >= 3am, recompute ALL groups (always overwrites cache).

    Tracks last_metrics_full_run + last_metrics_daily_run in settings.
    Returns a list of result dicts for logging.
    """
    settings = load_settings()
    groups = load_groups_data().get('groups', [])
    if not groups:
        return []

    now = datetime.datetime.now()
    today_str = now.strftime('%Y-%m-%d')
    results = []

    # Determine if a weekly full refresh is due.
    full_due = False
    if now.weekday() == 0 and now.hour >= 3:  # Monday >= 3am
        last_full = settings.get('last_metrics_full_run', '')
        this_mon_target = now.replace(hour=3, minute=0, second=0, microsecond=0)
        last_dt = None
        if last_full:
            try:
                last_dt = datetime.datetime.strptime(last_full[:19], '%Y-%m-%dT%H:%M:%S')
            except:
                pass
        if last_dt is None or last_dt < this_mon_target:
            full_due = True

    # Daily fill-in: skip if we already ran today AND not doing full refresh.
    daily_due = settings.get('last_metrics_daily_run', '')[:10] != today_str

    if not full_due and not daily_due:
        return []

    cache = load_group_metrics_all()
    for g in groups:
        gid = g.get('id', '')
        if not gid:
            continue
        if full_due or gid not in cache:
            try:
                metrics = compute_group_metrics(g)
                cache[gid] = metrics
                results.append({
                    'group_id': gid,
                    'group_name': g.get('name', ''),
                    'org_count': len(metrics.get('byOrg', {})),
                    'trigger': 'full' if full_due else 'fill_in'
                })
            except Exception as e:
                results.append({
                    'group_id': gid,
                    'group_name': g.get('name', ''),
                    'error': safe_str(e)
                })

    if results:
        save_group_metrics_all(cache)

    if full_due:
        settings['last_metrics_full_run'] = now.isoformat()
    settings['last_metrics_daily_run'] = now.isoformat()
    save_settings(settings)

    return results

# ============================================================
# DASHBOARD HELPERS (outreach-health KPIs + chart series)
# ============================================================
# These power the Dashboard tab. They are pure aggregate reads on big tables
# (Attend, EnrollmentTransaction, TaskNote) and the sender_log content blob.
# All SQL uses WITH (NOLOCK) and is filtered to a small set of orgs (the
# union of every saved group's scope). Returns dicts that ship straight to
# the browser via load_dashboard_data.

def _dash_parse_date(s, default):
    """Parse 'YYYY-MM-DD' from form input. Falls back to default on bad input."""
    if not s:
        return default
    try:
        return datetime.datetime.strptime(str(s)[:10], '%Y-%m-%d')
    except:
        return default


def _dash_group_org_ids():
    """Return (org_ids_csv, all_groups) for the union of all saved groups'
    scopes. Org discovery mirrors compute_group_metrics so the dashboard sees
    the same involvements the Health view does.
    """
    groups = load_groups_data().get('groups', [])
    if not groups:
        return ('0', [])

    org_ids = set()
    for g in groups:
        for oid in _dash_org_ids_for_group(g):
            org_ids.add(oid)

    if not org_ids:
        return ('0', groups)
    return (','.join(str(o) for o in org_ids), groups)


def _dash_org_ids_for_group(group):
    """Return set of OrganizationIds in scope for a single group. Used by
    Chart 3 to compute per-group metrics with the same SQL the dashboard's
    cards use, so the per-group rate and the overall card rate stay in
    the same units (distinct people, not per-org-membership)."""
    org_ids = set()
    org_filter, needs_os_join = _group_scope_org_filter(group, 'om', 'o', 'os')
    os_join = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId" if needs_os_join else ""
    scope_filter = org_filter.replace('om.', 'o.') if 'om.' in org_filter else org_filter
    sql = """
        SELECT DISTINCT o.OrganizationId
        FROM Organizations o
        {os_join}
        WHERE o.OrganizationStatusId = 30
          {scope_filter}
    """.format(os_join=os_join, scope_filter=scope_filter)
    try:
        for r in q.QuerySql(sql):
            try:
                org_ids.add(int(r.OrganizationId))
            except:
                pass
    except:
        pass
    return org_ids


def _dash_active_prospect_memberships(org_ids_csv, prospect_types_csv):
    """Count of (PeopleId, OrgId) prospect memberships right now -- mirrors
    the Group Management page's 'Prospects' total. A person who is a
    prospect in 3 orgs counts as 3 here. Used as the secondary number under
    the Active Prospects card so users can reconcile the two views."""
    if org_ids_csv == '0' or not prospect_types_csv:
        return 0
    sql = """
        SELECT COUNT(*) AS Cnt
        FROM OrganizationMembers om WITH (NOLOCK)
        WHERE om.OrganizationId IN ({org_ids})
          AND om.MemberTypeId IN ({prospect_types})
          AND (om.InactiveDate IS NULL)
    """.format(org_ids=org_ids_csv, prospect_types=prospect_types_csv)
    try:
        row = q.QuerySqlTop1(sql)
        return int(row.Cnt) if row and row.Cnt else 0
    except:
        return 0


def _dash_active_prospects_pids(org_ids_csv, prospect_types_csv):
    """Currently active prospects (point-in-time): on the org's current roster
    with a prospect MemberType. Used by the Active Prospects KPI and as the
    universe for the Overdue KPI.
    """
    if org_ids_csv == '0' or not prospect_types_csv:
        return set()
    sql = """
        SELECT DISTINCT om.PeopleId
        FROM OrganizationMembers om WITH (NOLOCK)
        WHERE om.OrganizationId IN ({org_ids})
          AND om.MemberTypeId IN ({prospect_types})
          AND (om.InactiveDate IS NULL)
    """.format(org_ids=org_ids_csv, prospect_types=prospect_types_csv)
    pids = set()
    try:
        for r in q.QuerySql(sql):
            try:
                pids.add(int(r.PeopleId))
            except:
                pass
    except:
        pass
    return pids


def _dash_all_prospect_types_csv(groups):
    """Union of memberTypes across all groups. Each group may have its own
    prospect-type list (default [311]); we keep the dashboard aligned with
    however groups are configured."""
    types = set()
    for g in groups:
        for t in (g.get('memberTypes') or [311]):
            try:
                types.add(int(t))
            except:
                pass
    if not types:
        return '311'
    return ','.join(str(t) for t in types)


def _dash_all_converted_types_csv(groups):
    """Union of convertedAttendTypeIds across groups, falling back to settings
    then [30]."""
    types = set()
    settings_types = load_settings().get('converted_attend_type_ids') or [30]
    for g in groups:
        chosen = g.get('convertedAttendTypeIds') or settings_types
        for t in chosen:
            try:
                types.add(int(t))
            except:
                pass
    if not types:
        return '30'
    return ','.join(str(t) for t in types)


def _dash_contact_method_keyword_ids():
    """Return the configured Contact Method KeywordIds (shared with ProgramPulse).
    A 'touch' on the dashboard means a TaskNote tagged with one of these keywords."""
    try:
        pp = load_content("ProgramPulse_Settings", {}) or {}
        methods = pp.get('contact_methods', []) or []
        ids = []
        for m in methods:
            try:
                kid = int(m.get('keywordId') or 0)
                if kid > 0:
                    ids.append(kid)
            except:
                pass
        return ids
    except:
        return []


def _dash_touches_in_range(start_dt, end_dt):
    """Count of contact-method NOTES between start_dt and end_dt.

    A "touch" = a NOTE (IsNote=1) whose KeywordId is in the configured
    contact-method set. Tasks (IsNote=0) are intentions to do something;
    only the resulting Note counts as a touch -- otherwise the dashboard
    would credit you for things that haven't happened yet.
    """
    if start_dt is None or end_dt is None:
        return 0
    kw_ids = _dash_contact_method_keyword_ids()
    if not kw_ids:
        return 0
    kw_csv = ','.join(str(k) for k in kw_ids)
    end_plus = (end_dt + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    sql = """
        SELECT COUNT(DISTINCT tn.TaskNoteId) AS Cnt
        FROM TaskNote tn WITH (NOLOCK)
        JOIN TaskNoteKeyword tnk WITH (NOLOCK) ON tnk.TaskNoteId = tn.TaskNoteId
        WHERE tnk.KeywordId IN ({kw})
          AND tn.IsNote = 1
          AND tn.CreatedDate >= '{start}'
          AND tn.CreatedDate <  '{end_plus}'
    """.format(
        kw=kw_csv,
        start=start_dt.strftime('%Y-%m-%d'),
        end_plus=end_plus
    )
    try:
        row = q.QuerySqlTop1(sql)
        return int(row.Cnt) if row and row.Cnt else 0
    except:
        return 0


def _dash_distinct_converted_pids(org_ids_csv, converted_types_csv, start_dt, end_dt):
    """Distinct PeopleIds who CONVERTED in [start, end] across the saved orgs.

    A conversion is detected via the Attend table -- the right place per
    compute_group_metrics's design ("AttendanceTypeId transition is more
    responsive than waiting on a MemberType change"). Rule:

      For a given (PeopleId, OrganizationId), find the FIRST attendance with
      AttendanceTypeId in the converted set. If that first-converted-date
      falls in the window AND the person previously attended the same org
      with a NON-converted AttendanceTypeId, that's a conversion event.

    The "had a prior non-converted attendance" guard is what prevents
    long-time Members showing up on Sunday from counting as conversions --
    their first converted-type attendance is years ago, not in this window.
    """
    if org_ids_csv == '0' or start_dt is None or end_dt is None:
        return set()
    if not converted_types_csv or converted_types_csv == '0':
        return set()
    sql = """
        SELECT DISTINCT fc.PeopleId
        FROM (
            SELECT a.PeopleId, a.OrganizationId, MIN(a.MeetingDate) AS FirstConvertedDate
            FROM Attend a WITH (NOLOCK)
            WHERE a.OrganizationId IN ({org_ids})
              AND a.AttendanceFlag = 1
              AND a.AttendanceTypeId IN ({converted_types})
            GROUP BY a.PeopleId, a.OrganizationId
        ) fc
        WHERE fc.FirstConvertedDate >= '{start}'
          AND fc.FirstConvertedDate <  '{end_plus}'
          AND EXISTS (
              SELECT 1 FROM Attend a2 WITH (NOLOCK)
              WHERE a2.OrganizationId = fc.OrganizationId
                AND a2.PeopleId       = fc.PeopleId
                AND a2.AttendanceFlag = 1
                AND a2.AttendanceTypeId NOT IN ({converted_types})
                AND a2.MeetingDate < fc.FirstConvertedDate
          )
    """.format(
        org_ids=org_ids_csv,
        converted_types=converted_types_csv,
        start=start_dt.strftime('%Y-%m-%d'),
        end_plus=(end_dt + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )
    pids = set()
    try:
        for r in q.QuerySql(sql):
            try:
                pids.add(int(r.PeopleId))
            except:
                pass
    except:
        pass
    return pids


def _dash_distinct_prospects_in_range(org_ids_csv, prospect_types_csv, start_dt, end_dt):
    """Distinct PeopleIds who were prospects at some point during the range.
    Mirrors the PiW CTE in compute_group_metrics: enrollment overlaps the window.
    """
    if org_ids_csv == '0' or start_dt is None or end_dt is None:
        return set()
    sql = """
        SELECT DISTINCT et.PeopleId
        FROM EnrollmentTransaction et WITH (NOLOCK)
        WHERE et.OrganizationId IN ({org_ids})
          AND et.MemberTypeId IN ({prospect_types})
          AND et.TransactionStatus = 0
          AND et.EnrollmentDate <= '{end_plus}'
          AND (et.InactiveDate IS NULL OR et.InactiveDate >= '{start}')
    """.format(
        org_ids=org_ids_csv,
        prospect_types=prospect_types_csv,
        start=start_dt.strftime('%Y-%m-%d'),
        end_plus=(end_dt + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )
    pids = set()
    try:
        for r in q.QuerySql(sql):
            try:
                pids.add(int(r.PeopleId))
            except:
                pass
    except:
        pass
    return pids


def _dash_count_prospect_pairs(org_ids_csv, prospect_types_csv, start_dt, end_dt):
    """Distinct (PeopleId, OrgId) PROSPECT pairs whose membership overlaps
    the window. Each involvement represents its own conversion 'scenario'
    -- a prospect in 3 orgs counts as 3 pairs in this denominator."""
    if org_ids_csv == '0' or start_dt is None or end_dt is None:
        return 0
    if not prospect_types_csv or prospect_types_csv == '0':
        return 0
    sql = """
        SELECT COUNT(*) AS Cnt FROM (
            SELECT DISTINCT et.PeopleId, et.OrganizationId
            FROM dbo.EnrollmentTransaction et WITH (NOLOCK)
            WHERE et.OrganizationId IN ({org_ids})
              AND et.MemberTypeId   IN ({prospect_types})
              AND et.TransactionStatus = 0
              AND et.EnrollmentDate <= '{end_plus}'
              AND (et.InactiveDate IS NULL OR et.InactiveDate >= '{start}')
        ) x
    """.format(
        org_ids=org_ids_csv,
        prospect_types=prospect_types_csv,
        start=start_dt.strftime('%Y-%m-%d'),
        end_plus=(end_dt + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )
    try:
        row = q.QuerySqlTop1(sql)
        return int(row.Cnt) if row and row.Cnt else 0
    except:
        return 0


def _dash_count_conversion_pairs(org_ids_csv, converted_types_csv, start_dt, end_dt):
    """Distinct (PeopleId, OrgId) CONVERSION pairs in the window. Same
    transition rule as _dash_distinct_converted_pids but counts pairs
    instead of collapsing to distinct people. Matches the per-involvement
    'each org is its own scenario' model."""
    if org_ids_csv == '0' or start_dt is None or end_dt is None:
        return 0
    if not converted_types_csv or converted_types_csv == '0':
        return 0
    sql = """
        SELECT COUNT(*) AS Cnt FROM (
            SELECT fc.PeopleId, fc.OrganizationId
            FROM (
                SELECT a.PeopleId, a.OrganizationId, MIN(a.MeetingDate) AS FirstConvertedDate
                FROM dbo.Attend a WITH (NOLOCK)
                WHERE a.OrganizationId IN ({org_ids})
                  AND a.AttendanceFlag = 1
                  AND a.AttendanceTypeId IN ({converted_types})
                GROUP BY a.PeopleId, a.OrganizationId
            ) fc
            WHERE fc.FirstConvertedDate >= '{start}'
              AND fc.FirstConvertedDate <  '{end_plus}'
              AND EXISTS (
                  SELECT 1 FROM dbo.Attend a2 WITH (NOLOCK)
                  WHERE a2.OrganizationId = fc.OrganizationId
                    AND a2.PeopleId       = fc.PeopleId
                    AND a2.AttendanceFlag = 1
                    AND a2.AttendanceTypeId NOT IN ({converted_types})
                    AND a2.MeetingDate < fc.FirstConvertedDate
              )
        ) x
    """.format(
        org_ids=org_ids_csv,
        converted_types=converted_types_csv,
        start=start_dt.strftime('%Y-%m-%d'),
        end_plus=(end_dt + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )
    try:
        row = q.QuerySqlTop1(sql)
        return int(row.Cnt) if row and row.Cnt else 0
    except:
        return 0


def _dash_avg_orgs_before_conversion(org_ids_csv, prospect_types_csv,
                                     converted_types_csv, start_dt, end_dt):
    """Behavioral metric: of the people who converted in the window, how
    many distinct involvements were they a prospect in (within scope)?
    Returns a tuple (avg_float, sample_size_int).

    Tells you whether prospects typically try one room and stick, or
    bounce through several before finding a home.
    """
    if (org_ids_csv == '0' or start_dt is None or end_dt is None
            or not prospect_types_csv or not converted_types_csv):
        return (0.0, 0)
    sql = """
        WITH ConvertedPeople AS (
            SELECT DISTINCT fc.PeopleId
            FROM (
                SELECT a.PeopleId, a.OrganizationId, MIN(a.MeetingDate) AS FirstConvertedDate
                FROM dbo.Attend a WITH (NOLOCK)
                WHERE a.OrganizationId IN ({org_ids})
                  AND a.AttendanceFlag = 1
                  AND a.AttendanceTypeId IN ({converted_types})
                GROUP BY a.PeopleId, a.OrganizationId
            ) fc
            WHERE fc.FirstConvertedDate >= '{start}'
              AND fc.FirstConvertedDate <  '{end_plus}'
              AND EXISTS (
                  SELECT 1 FROM dbo.Attend a2 WITH (NOLOCK)
                  WHERE a2.OrganizationId = fc.OrganizationId
                    AND a2.PeopleId       = fc.PeopleId
                    AND a2.AttendanceFlag = 1
                    AND a2.AttendanceTypeId NOT IN ({converted_types})
                    AND a2.MeetingDate < fc.FirstConvertedDate
              )
        ),
        PerPersonOrgs AS (
            SELECT et.PeopleId, COUNT(DISTINCT et.OrganizationId) AS OrgsTried
            FROM dbo.EnrollmentTransaction et WITH (NOLOCK)
            JOIN ConvertedPeople cp ON cp.PeopleId = et.PeopleId
            WHERE et.OrganizationId IN ({org_ids})
              AND et.MemberTypeId   IN ({prospect_types})
              AND et.TransactionStatus = 0
            GROUP BY et.PeopleId
        )
        SELECT COUNT(*) AS Sample, ISNULL(AVG(CAST(OrgsTried AS FLOAT)), 0) AS Avg_
        FROM PerPersonOrgs
    """.format(
        org_ids=org_ids_csv,
        prospect_types=prospect_types_csv,
        converted_types=converted_types_csv,
        start=start_dt.strftime('%Y-%m-%d'),
        end_plus=(end_dt + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    )
    try:
        row = q.QuerySqlTop1(sql)
        if not row: return (0.0, 0)
        sample = int(getattr(row, 'Sample', 0) or 0)
        avg = float(getattr(row, 'Avg_', 0) or 0.0)
        return (avg, sample)
    except:
        return (0.0, 0)


def _dash_overdue_count(active_pids):
    """Number of active prospects whose most recent contact-method TaskNote
    is older than OVERDUE_DAYS, OR who have never been touched.

    "Touch" here matches Touches Sent + Daily Touches: a TaskNote tagged
    with one of the configured Contact Method Keywords (Phone Call, Email,
    Visit, etc.). A TaskNote without a contact-method keyword does NOT count
    as a touch -- internal admin notes shouldn't reset the overdue clock.
    """
    if not active_pids:
        return 0

    kw_ids = _dash_contact_method_keyword_ids()
    if not kw_ids:
        # No contact methods configured -> every active prospect is "overdue".
        # Surface that loudly so admins know to configure the setting.
        return len(active_pids)
    kw_csv = ','.join(str(k) for k in kw_ids)

    cutoff_dt = datetime.datetime.now() - datetime.timedelta(days=OVERDUE_DAYS)
    cutoff_date = cutoff_dt.strftime('%Y-%m-%d')

    # People with a recent contact-method TaskNote (within OVERDUE_DAYS).
    # Chunked to keep the IN(...) clause sane on huge prospect sets.
    recent_touched = set()
    CHUNK = 1500
    pid_list = list(active_pids)
    for i in range(0, len(pid_list), CHUNK):
        chunk = pid_list[i:i + CHUNK]
        sql = """
            SELECT DISTINCT tn.AboutPersonId
            FROM TaskNote tn WITH (NOLOCK)
            JOIN TaskNoteKeyword tnk WITH (NOLOCK) ON tnk.TaskNoteId = tn.TaskNoteId
            WHERE tn.AboutPersonId IN ({pids})
              AND tnk.KeywordId IN ({kw})
              AND tn.IsNote = 1
              AND tn.CreatedDate >= '{cutoff}'
        """.format(
            pids=','.join(str(p) for p in chunk),
            kw=kw_csv,
            cutoff=cutoff_date
        )
        try:
            for r in q.QuerySql(sql):
                try:
                    recent_touched.add(int(r.AboutPersonId))
                except:
                    pass
        except:
            pass

    overdue = 0
    for pid in active_pids:
        if pid not in recent_touched:
            overdue += 1
    return overdue


def dashboard_compute_kpis(start_iso, end_iso):
    """Compute all six KPI numbers. start/end are inclusive 'YYYY-MM-DD' strings.

    Cards 2/3/4/6 respect the range. Cards 1 ("Active Prospects") and 5
    ("Overdue") are point-in-time and ignore the range.
    """
    now = datetime.datetime.now()
    today = datetime.datetime(now.year, now.month, now.day)
    default_start = today - datetime.timedelta(days=29)
    start_dt = _dash_parse_date(start_iso, default_start)
    end_dt = _dash_parse_date(end_iso, today)

    org_ids_csv, groups = _dash_group_org_ids()
    prospect_types_csv = _dash_all_prospect_types_csv(groups)
    converted_types_csv = _dash_all_converted_types_csv(groups)

    # Card 1: Active Prospects (point-in-time). Surface BOTH
    # "distinct people" (the headline value) and "per-org memberships"
    # (so it reconciles with Group Management's 'Prospects' total).
    active_pids = _dash_active_prospects_pids(org_ids_csv, prospect_types_csv)
    active_memberships = _dash_active_prospect_memberships(org_ids_csv, prospect_types_csv)

    # Card 2 / 3: Touches Sent (7d / 30d) from sender_log
    touches_7d = _dash_touches_in_range(today - datetime.timedelta(days=7), today)
    touches_30d = _dash_touches_in_range(today - datetime.timedelta(days=30), today)
    touches_range = _dash_touches_in_range(start_dt, end_dt)

    # Card 4: New Conversions (30d + range). Per the Attend-based definition
    # in _dash_distinct_converted_pids, the SQL already requires both:
    # (a) FIRST converted-type attendance is in the window, AND
    # (b) the person previously attended the SAME org with a non-converted
    #     AttendanceTypeId.
    # That captures the prospect -> member transition properly without
    # needing a Python set-intersection.
    conv_30d_pids = _dash_distinct_converted_pids(
        org_ids_csv, converted_types_csv,
        today - datetime.timedelta(days=30), today
    )
    conv_range_pids = _dash_distinct_converted_pids(
        org_ids_csv, converted_types_csv, start_dt, end_dt
    )

    # Card 5: Overdue (point-in-time)
    overdue = _dash_overdue_count(active_pids)

    # Card 6: Conversion Rate (90d + range) -- PER INVOLVEMENT.
    # Denominator = distinct (PeopleId, OrgId) prospect pairs in window.
    # Numerator   = distinct (PeopleId, OrgId) conversion pairs in window.
    # Each involvement represents its own conversion 'scenario' (the user's
    # mental model); a prospect in 3 orgs is 3 prospect pairs.
    prospect_pairs_90d = _dash_count_prospect_pairs(
        org_ids_csv, prospect_types_csv,
        today - datetime.timedelta(days=90), today
    )
    conv_pairs_90d = _dash_count_conversion_pairs(
        org_ids_csv, converted_types_csv,
        today - datetime.timedelta(days=90), today
    )
    rate_90d = (float(conv_pairs_90d) / prospect_pairs_90d * 100.0) if prospect_pairs_90d else 0.0

    prospect_pairs_range = _dash_count_prospect_pairs(
        org_ids_csv, prospect_types_csv, start_dt, end_dt
    )
    conv_pairs_range = _dash_count_conversion_pairs(
        org_ids_csv, converted_types_csv, start_dt, end_dt
    )
    rate_range = (float(conv_pairs_range) / prospect_pairs_range * 100.0) if prospect_pairs_range else 0.0

    # Card 7 (new): Avg involvements tried as prospect before conversion (90d)
    avg_orgs_90d, avg_orgs_sample = _dash_avg_orgs_before_conversion(
        org_ids_csv, prospect_types_csv, converted_types_csv,
        today - datetime.timedelta(days=90), today
    )

    # Empty-state signal: did the user actually configure any groups?
    # If groups is empty, the dashboard should show a friendly "get started"
    # instead of a wall of zeros.
    has_groups = len(groups) > 0
    has_org_scope = (org_ids_csv != '0')

    return {
        'hasGroups': has_groups,
        'hasOrgScope': has_org_scope,
        'activeProspects': len(active_pids),
        'activeProspectMemberships': active_memberships,
        'touches7d': touches_7d,
        'touches30d': touches_30d,
        'touchesRange': touches_range,
        'newConversions30d': len(conv_30d_pids),
        'newConversionsRange': len(conv_range_pids),
        'overdueCount': overdue,
        'conversionRate90d': round(rate_90d, 1),
        'conversionRateRange': round(rate_range, 1),
        # Per-pair counts surfaced for verification + future use.
        'prospectPairs90d': prospect_pairs_90d,
        'conversionPairs90d': conv_pairs_90d,
        # New behavioral metric:
        'avgOrgsBeforeConv90d': round(avg_orgs_90d, 2),
        'avgOrgsBeforeConvSample90d': avg_orgs_sample,
    }


def dashboard_compute_charts(start_iso, end_iso):
    """Three chart payloads: dailyTouches, dailyConversions, groupConversion."""
    now = datetime.datetime.now()
    today = datetime.datetime(now.year, now.month, now.day)
    default_start = today - datetime.timedelta(days=29)
    start_dt = _dash_parse_date(start_iso, default_start)
    end_dt = _dash_parse_date(end_iso, today)

    org_ids_csv, groups = _dash_group_org_ids()
    converted_types_csv = _dash_all_converted_types_csv(groups)

    # ---- Chart 1: Weekly Touches (last ~13 weeks), split Notes vs Tasks ----
    # Contact-method TaskNotes split by IsNote so the bar shows both:
    #   * Notes (IsNote=1) = completed touches -- the canonical signal
    #   * Tasks (IsNote=0) = intentions, not yet acted on
    # Stacked bar in the JS keeps the total height representing all activity
    # while the color split shows planning-vs-doing.
    NUM_TOUCH_WEEKS = 13
    touch_window_start = today - datetime.timedelta(weeks=NUM_TOUCH_WEEKS - 1)
    touch_window_start = touch_window_start - datetime.timedelta(days=touch_window_start.isoweekday() % 7)
    touch_labels = []
    touch_notes = []
    touch_tasks = []
    touch_notes_buckets = {}
    touch_tasks_buckets = {}
    kw_ids = _dash_contact_method_keyword_ids()
    if kw_ids:
        kw_csv = ','.join(str(k) for k in kw_ids)
        window_end_plus = (today + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        sql = """
            SELECT CONVERT(VARCHAR(10),
                           DATEADD(day, 1 - DATEPART(weekday, tn.CreatedDate),
                                   CAST(tn.CreatedDate AS DATE)),
                           120) AS WeekKey,
                   CASE WHEN ISNULL(tn.IsNote, 0) = 1 THEN 1 ELSE 0 END AS IsNoteFlag,
                   COUNT(DISTINCT tn.TaskNoteId) AS Cnt
            FROM TaskNote tn WITH (NOLOCK)
            JOIN TaskNoteKeyword tnk WITH (NOLOCK) ON tnk.TaskNoteId = tn.TaskNoteId
            WHERE tnk.KeywordId IN ({kw})
              AND tn.CreatedDate >= '{start}'
              AND tn.CreatedDate <  '{end_plus}'
            GROUP BY DATEADD(day, 1 - DATEPART(weekday, tn.CreatedDate),
                             CAST(tn.CreatedDate AS DATE)),
                     CASE WHEN ISNULL(tn.IsNote, 0) = 1 THEN 1 ELSE 0 END
        """.format(
            kw=kw_csv,
            start=touch_window_start.strftime('%Y-%m-%d'),
            end_plus=window_end_plus
        )
        try:
            for r in q.QuerySql(sql):
                try:
                    key = safe_str(r.WeekKey)
                    n = int(r.Cnt) if r.Cnt else 0
                    is_note = int(getattr(r, 'IsNoteFlag', 0) or 0)
                    if is_note == 1:
                        touch_notes_buckets[key] = touch_notes_buckets.get(key, 0) + n
                    else:
                        touch_tasks_buckets[key] = touch_tasks_buckets.get(key, 0) + n
                except:
                    pass
        except:
            pass
    for i in range(NUM_TOUCH_WEEKS):
        d = touch_window_start + datetime.timedelta(weeks=i)
        key = d.strftime('%Y-%m-%d')
        touch_labels.append(key)
        touch_notes.append(touch_notes_buckets.get(key, 0))
        touch_tasks.append(touch_tasks_buckets.get(key, 0))

    # ---- Chart 2: Weekly Conversions (last ~13 weeks) ----
    # Same Attend-based "first converted in org, with prior non-converted
    # attendance in same org" definition as the card. Bucketed by
    # week-starting-Sunday so the per-day noise smooths into something
    # readable.
    NUM_WEEKS = 13
    conv_window_start = today - datetime.timedelta(weeks=NUM_WEEKS - 1)
    # Align to the Sunday of that week so labels are consistent with the
    # SQL bucket keys.
    conv_window_start = conv_window_start - datetime.timedelta(days=conv_window_start.isoweekday() % 7)
    conv_labels = []
    conv_data = []
    conv_buckets = {}
    if org_ids_csv != '0' and converted_types_csv and converted_types_csv != '0':
        sql = """
            SELECT CONVERT(VARCHAR(10),
                           DATEADD(day, 1 - DATEPART(weekday, fc.FirstConvertedDate),
                                   CAST(fc.FirstConvertedDate AS DATE)),
                           120) AS WeekKey,
                   COUNT(DISTINCT fc.PeopleId) AS Cnt
            FROM (
                SELECT a.PeopleId, a.OrganizationId, MIN(a.MeetingDate) AS FirstConvertedDate
                FROM Attend a WITH (NOLOCK)
                WHERE a.OrganizationId IN ({org_ids})
                  AND a.AttendanceFlag = 1
                  AND a.AttendanceTypeId IN ({converted_types})
                GROUP BY a.PeopleId, a.OrganizationId
            ) fc
            WHERE fc.FirstConvertedDate >= '{start}'
              AND fc.FirstConvertedDate <  '{end_plus}'
              AND EXISTS (
                  SELECT 1 FROM Attend a2 WITH (NOLOCK)
                  WHERE a2.OrganizationId = fc.OrganizationId
                    AND a2.PeopleId       = fc.PeopleId
                    AND a2.AttendanceFlag = 1
                    AND a2.AttendanceTypeId NOT IN ({converted_types})
                    AND a2.MeetingDate < fc.FirstConvertedDate
              )
            GROUP BY DATEADD(day, 1 - DATEPART(weekday, fc.FirstConvertedDate),
                             CAST(fc.FirstConvertedDate AS DATE))
        """.format(
            org_ids=org_ids_csv,
            converted_types=converted_types_csv,
            start=conv_window_start.strftime('%Y-%m-%d'),
            end_plus=(today + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        )
        try:
            for r in q.QuerySql(sql):
                try:
                    conv_buckets[safe_str(r.WeekKey)] = int(r.Cnt or 0)
                except:
                    pass
        except:
            pass
    for i in range(NUM_WEEKS):
        d = conv_window_start + datetime.timedelta(weeks=i)
        key = d.strftime('%Y-%m-%d')
        conv_labels.append(key)
        conv_data.append(conv_buckets.get(key, 0))

    # ---- Chart 3: Conversion Rate by Group -- 90-DAY ROLLING ----
    # Multi-line chart. X = last 6 month-ends. Y = per-involvement
    # conversion rate (P x O pairs, same math as Card 6) computed over the
    # 90-day window ending at that month-end. One line per group.
    #
    # Why 90-day rolling instead of single-month buckets: keeps the chart
    # and Card 6 (90-day conversion rate) on the same definition. The most
    # recent chart point should equal the card value, and trend points
    # smooth out single-month volatility while still showing real change
    # over time.
    NUM_MONTHS = 6
    ROLLING_DAYS = 90
    # Build month-end anchors from oldest -> newest. For each calendar
    # month in the lookback, the anchor is the LAST DAY of that month.
    # This gives consistent month-to-month spacing on the x-axis while
    # making the rolling window land on natural calendar boundaries.
    first_of_this_month = datetime.datetime(today.year, today.month, 1)

    def _month_end_exclusive(m_start):
        if m_start.month == 12:
            return datetime.datetime(m_start.year + 1, 1, 1)
        return datetime.datetime(m_start.year, m_start.month + 1, 1)

    month_starts = []
    cur = first_of_this_month
    for _ in range(NUM_MONTHS):
        month_starts.append(cur)
        if cur.month == 1:
            cur = datetime.datetime(cur.year - 1, 12, 1)
        else:
            cur = datetime.datetime(cur.year, cur.month - 1, 1)
    month_starts.reverse()   # oldest first
    month_labels = [m.strftime('%Y-%m') for m in month_starts]

    # For the rolling window, each chart point's end-date is the LAST day
    # of its month — except for the current (in-progress) month, where we
    # use `today` so the rightmost point reflects live state instead of
    # extrapolating to a future date.
    anchor_dates = []
    for m_start in month_starts:
        m_end_exclusive = _month_end_exclusive(m_start)
        anchor = min(m_end_exclusive - datetime.timedelta(days=1), today)
        anchor_dates.append(anchor)

    # Per-anchor window labels (e.g. "2026-01-02 to 2026-04-01") let the
    # debug table show the user exactly what date range each chart point
    # measured. Surfaces the "real spike vs config drift vs denominator
    # collapse" diagnostic without requiring a SQL re-run.
    window_labels = []
    for anchor in anchor_dates:
        window_start = anchor - datetime.timedelta(days=ROLLING_DAYS - 1)
        window_labels.append('{0} to {1}'.format(
            window_start.strftime('%Y-%m-%d'), anchor.strftime('%Y-%m-%d')))

    grp_series = []   # [{name, data, prospectPairs, conversionPairs}, ...]
    for g in groups:
        g_org_ids = _dash_org_ids_for_group(g)
        if not g_org_ids:
            grp_series.append({
                'name': g.get('name') or '(unnamed)',
                'data': [0.0] * NUM_MONTHS,
                'prospectPairs': [0] * NUM_MONTHS,
                'conversionPairs': [0] * NUM_MONTHS,
            })
            continue
        g_org_ids_csv = ','.join(str(o) for o in g_org_ids)
        g_prospect_types_csv = ','.join(str(int(t)) for t in (g.get('memberTypes') or [311])) or '311'
        g_converted_types = g.get('convertedAttendTypeIds') or load_settings().get('converted_attend_type_ids') or [30]
        g_converted_types_csv = ','.join(str(int(t)) for t in g_converted_types) or '30'

        data_points = []
        prospect_counts = []
        conversion_counts = []
        for anchor in anchor_dates:
            window_start = anchor - datetime.timedelta(days=ROLLING_DAYS - 1)
            p_pairs = _dash_count_prospect_pairs(g_org_ids_csv, g_prospect_types_csv, window_start, anchor)
            c_pairs = _dash_count_conversion_pairs(g_org_ids_csv, g_converted_types_csv, window_start, anchor)
            rate = (float(c_pairs) / p_pairs * 100.0) if p_pairs else 0.0
            data_points.append(round(rate, 1))
            prospect_counts.append(p_pairs)
            conversion_counts.append(c_pairs)
        grp_series.append({
            'name': g.get('name') or '(unnamed)',
            'data': data_points,
            'prospectPairs': prospect_counts,
            'conversionPairs': conversion_counts,
        })

    return {
        'dailyTouches': {
            'labels': touch_labels,
            'notes': touch_notes,
            'tasks': touch_tasks,
            # Keep 'data' for back-compat = combined total.
            'data': [touch_notes[i] + touch_tasks[i] for i in range(len(touch_labels))],
        },
        'dailyConversions': {'labels': conv_labels, 'data': conv_data},
        'groupConversion': {
            'months': month_labels,
            'windows': window_labels,
            'series': grp_series,
        },
    }


# ============================================================
# PAGE HEADER
# ============================================================
model.Header = ''

# One-time migration of legacy default message body on saved senders.
# Cheap when already migrated (single settings read).
try:
    migrate_legacy_default_message()
except:
    pass

# ============================================================
# SCHEDULED TASK ENTRY POINT
# ============================================================
# Called from ScheduledTasks via:
#   Data.run_batch = "true"
#   model.CallScript("TPxi_ProspectBuilder")
#
# Legacy alias `Data.run_senders = "true"` is still honored for any existing
# ScheduledTasks scripts written against the prior name.

_run_senders = False
try:
    if hasattr(Data, 'run_batch') and str(Data.run_batch).lower() == 'true':
        _run_senders = True
    elif hasattr(Data, 'run_senders') and str(Data.run_senders).lower() == 'true':
        _run_senders = True
except:
    pass

if _run_senders:
    # Honor the Settings -> Scheduled Tasks "Enable" toggle. Lets an admin
    # pause sends without ripping the block out of ScheduledTasks (safe
    # when other scripts share the slot).
    if not is_scheduling_enabled():
        print "Prospect Builder scheduled sends are DISABLED in Settings; skipping."
        _run_senders = False

if _run_senders:
    results = run_scheduled_senders()
    for r in results:
        if r.get('errors'):
            print "Sender '{0}': {1} prospects, {2} emails, {3} errors".format(
                r.get('sender_name', '?'), r.get('prospects_found', 0),
                r.get('emails_sent', 0), len(r.get('errors', [])))
        else:
            print "Sender '{0}': {1} prospects, {2} emails sent to {3} recipients".format(
                r.get('sender_name', '?'), r.get('prospects_found', 0),
                r.get('emails_sent', 0), r.get('recipients', 0))
    if not results:
        print "No senders were due to run"

    # Group metrics: daily fill-in for missing groups + weekly full refresh.
    try:
        metric_results = run_group_metrics_batch()
        for mr in metric_results:
            if mr.get('error'):
                print "Group metrics '{0}': ERROR {1}".format(mr.get('group_name', '?'), mr.get('error'))
            else:
                print "Group metrics '{0}': {1} orgs ({2})".format(
                    mr.get('group_name', '?'), mr.get('org_count', 0), mr.get('trigger', '?'))
        if not metric_results:
            print "Group metrics: nothing due"
    except Exception as e:
        print "Group metrics batch FAILED: {0}".format(safe_str(e))

# ============================================================
# AJAX HANDLER
# ============================================================
if model.HttpMethod == "post":
    action = get_form_data('action', '')
    response = {'success': False, 'message': 'Unknown action'}

    try:
        # ==========================================================
        # AUTO-UPDATE: fetch latest version from DisplayCache worker
        # mirror and overwrite this script's PythonContent slot. Saved
        # configurations, groups, sessions, and senders all live in
        # separate Content slots, so they survive the rewrite. Always
        # checked first since it has the lightest dependency surface.
        # ==========================================================
        if action == 'apply_update':
            new_code = ''
            try:
                fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
                new_code = str(model.RestGet(fetch_url, {}))
            except Exception as fe:
                response = {'success': False, 'message': 'Failed to fetch update: ' + str(fe)}
            else:
                if not new_code or len(new_code) < 200:
                    response = {'success': False, 'message': 'Invalid or empty script code received'}
                else:
                    target_name = get_pb_script_name() or DC_SCRIPT_ID
                    try:
                        model.WriteContentPython(target_name, new_code)
                        response = {'success': True, 'message': 'Updated ' + target_name + '. Reload the page.'}
                    except Exception as we:
                        response = {'success': False, 'message': 'Write failed: ' + str(we)}

        # ==========================================================
        # DASHBOARD (KPI cards + Chart.js series for the Dashboard tab)
        # ==========================================================
        if action == 'load_dashboard_data':
            start_iso = get_form_data('start', '')
            end_iso = get_form_data('end', '')
            # `parts`: 'kpis' (cards only), 'charts' (chart payloads only), or
            # '' / 'all' (both -- back-compat). The UI now requests them in
            # parallel so cards render in ~3s while Chart 3 finishes.
            parts = (get_form_data('parts', '') or 'all').lower()
            now = datetime.datetime.now()
            today = datetime.datetime(now.year, now.month, now.day)
            resolved_start = _dash_parse_date(start_iso, today - datetime.timedelta(days=29))
            resolved_end = _dash_parse_date(end_iso, today)
            response = {
                'success': True,
                'asOf': now.isoformat(),
                'rangeStart': resolved_start.strftime('%Y-%m-%d'),
                'rangeEnd': resolved_end.strftime('%Y-%m-%d'),
                'overdueDays': OVERDUE_DAYS,
                'parts': parts,
            }
            if parts in ('kpis', 'all'):
                response['kpis'] = dashboard_compute_kpis(start_iso, end_iso)
            if parts in ('charts', 'all'):
                response['charts'] = dashboard_compute_charts(start_iso, end_iso)

        # ==========================================================
        # CONFIG CRUD
        # ==========================================================
        elif action == 'load_configs':
            configs = load_configs()
            response = {'success': True, 'configs': configs}

        elif action == 'save_config':
            config_json = get_form_data('config_data', '{}')
            config = json.loads(config_json)
            configs = load_configs()

            config_id = config.get('id', '')
            if not config_id:
                config['id'] = make_id('cfg')
                config['createdAt'] = now_str()
                configs.append(config)
            else:
                for i, c in enumerate(configs):
                    if c.get('id') == config_id:
                        config['createdAt'] = c.get('createdAt', now_str())
                        configs[i] = config
                        break
                else:
                    config['createdAt'] = now_str()
                    configs.append(config)

            config['updatedAt'] = now_str()
            save_configs(configs)
            response = {'success': True, 'config': config, 'message': 'Configuration saved'}

        elif action == 'delete_config':
            config_id = get_form_data('config_id', '')
            configs = load_configs()
            configs = [c for c in configs if c.get('id') != config_id]
            save_configs(configs)
            response = {'success': True, 'message': 'Configuration deleted'}

        elif action == 'get_config_stats':
            config_json = get_form_data('config_data', '{}')
            config = json.loads(config_json)
            src = config.get('source', {})
            src_type = src.get('pb_type', '')
            member_types = config.get('memberTypes', [])
            max_prospects = config.get('maxProspects', 0) or 2000

            prospect_count = 0
            contacted_count = 0
            no_contact_count = 0

            if True:
                # Build WHERE clause based on source type
                if src_type == 'involvement' and src.get('orgId'):
                    mt_filter = ''
                    if member_types:
                        mt_filter = ' AND om.MemberTypeId IN ({0})'.format(','.join(str(mt) for mt in member_types))
                    count_sql = """
                        SELECT COUNT(DISTINCT om.PeopleId) as Cnt
                        FROM OrganizationMembers om WITH (NOLOCK)
                        JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
                        WHERE om.OrganizationId = {0}
                          AND p.IsDeceased = 0
                          {1}
                    """.format(int(src['orgId']), mt_filter)
                    for row in q.QuerySql(count_sql):
                        prospect_count = int(row.Cnt) if row.Cnt else 0

                    # Count contacted (have TaskNote in last 30 days)
                    contact_sql = """
                        SELECT COUNT(DISTINCT om.PeopleId) as Cnt
                        FROM OrganizationMembers om WITH (NOLOCK)
                        JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
                        WHERE om.OrganizationId = {0}
                          AND p.IsDeceased = 0
                          {1}
                          AND om.PeopleId IN (
                              SELECT DISTINCT tn.AboutPersonId FROM TaskNote tn WITH (NOLOCK)
                              WHERE tn.CreatedDate >= DATEADD(day, -30, GETDATE())
                          )
                    """.format(int(src['orgId']), mt_filter)
                    for row in q.QuerySql(contact_sql):
                        contacted_count = int(row.Cnt) if row.Cnt else 0

                elif src_type == 'tag' and src.get('tagName'):
                    tag_name_safe = safe_str(src['tagName']).replace("'", "''")
                    count_sql = """
                        SELECT COUNT(DISTINCT tp.PeopleId) as Cnt
                        FROM TagPerson tp WITH (NOLOCK)
                        JOIN Tag t WITH (NOLOCK) ON tp.Id = t.Id
                        JOIN People p WITH (NOLOCK) ON tp.PeopleId = p.PeopleId
                        WHERE t.Name = '{0}'
                          AND p.IsDeceased = 0
                    """.format(tag_name_safe)
                    for row in q.QuerySql(count_sql):
                        prospect_count = int(row.Cnt) if row.Cnt else 0

                    contact_sql = """
                        SELECT COUNT(DISTINCT tp.PeopleId) as Cnt
                        FROM TagPerson tp WITH (NOLOCK)
                        JOIN Tag t WITH (NOLOCK) ON tp.Id = t.Id
                        JOIN People p WITH (NOLOCK) ON tp.PeopleId = p.PeopleId
                        WHERE t.Name = '{0}'
                          AND p.IsDeceased = 0
                          AND tp.PeopleId IN (
                              SELECT DISTINCT tn.AboutPersonId FROM TaskNote tn WITH (NOLOCK)
                              WHERE tn.CreatedDate >= DATEADD(day, -30, GETDATE())
                          )
                    """.format(tag_name_safe)
                    for row in q.QuerySql(contact_sql):
                        contacted_count = int(row.Cnt) if row.Cnt else 0

                elif src_type == 'query' and src.get('queryId'):
                    # Saved query tag - query the tag directly via SQL
                    query_code = str(src.get('queryId', ''))
                    # In TouchPoint, saved search builder queries store results as tags
                    # The tag name matches the query name
                    tag_count_sql = """
                        SELECT COUNT(DISTINCT tp.PeopleId) as Cnt
                        FROM TagPerson tp WITH (NOLOCK)
                        JOIN Tag t WITH (NOLOCK) ON tp.Id = t.Id
                        JOIN People p WITH (NOLOCK) ON tp.PeopleId = p.PeopleId
                        WHERE t.Name = '{0}'
                          AND p.IsDeceased = 0
                    """.format(query_code.replace("'", "''"))
                    for row in q.QuerySql(tag_count_sql):
                        prospect_count = int(row.Cnt) if row.Cnt else 0

                    # If tag approach returned 0, the query may not use TagPerson
                    # Fall back to q.QueryCount
                    if prospect_count == 0:
                        prospect_count = q.QueryCount(query_code)

                    # Contact check using tag people
                    if prospect_count > 0:
                        tag_contact_sql = """
                            SELECT COUNT(DISTINCT tp.PeopleId) as Cnt
                            FROM TagPerson tp WITH (NOLOCK)
                            JOIN Tag t WITH (NOLOCK) ON tp.Id = t.Id
                            JOIN People p WITH (NOLOCK) ON tp.PeopleId = p.PeopleId
                            WHERE t.Name = '{0}'
                              AND p.IsDeceased = 0
                              AND tp.PeopleId IN (
                                  SELECT DISTINCT tn.AboutPersonId
                                  FROM TaskNote tn WITH (NOLOCK)
                                  WHERE tn.CreatedDate >= DATEADD(day, -30, GETDATE())
                              )
                        """.format(query_code.replace("'", "''"))
                        for row in q.QuerySql(tag_contact_sql):
                            contacted_count = int(row.Cnt) if row.Cnt else 0

                no_contact_count = max(0, prospect_count - contacted_count)

            # Fallback: if still 0 and there's a queryId, try q.QueryCount directly
            if prospect_count == 0 and src.get('queryId'):
                try:
                    prospect_count = q.QueryCount(str(src['queryId']))
                    no_contact_count = prospect_count
                except:
                    pass

            # Get processed count from work state
            processed_count = 0
            config_id = config.get('id', '')
            if config_id:
                work_states = load_content("ProspectBuilder_WorkStates", {})
                state = work_states.get(config_id, {})
                processed_count = len(state.get('processedMap', {}))

            response = {'success': True, 'stats': {
                'prospectCount': prospect_count,
                'contactedCount': contacted_count,
                'noContactCount': no_contact_count,
                'processedCount': processed_count
            }}

        # ==========================================================
        # SETTINGS
        # ==========================================================
        elif action == 'load_settings':
            # Load PB settings, but pull contact_methods from ProgramPulse_Settings
            settings = load_settings()
            pp_settings = load_content("ProgramPulse_Settings", {})
            settings['contact_methods'] = pp_settings.get('contact_methods', [])
            response = {'success': True, 'settings': settings}

        elif action == 'save_settings':
            settings_json = get_form_data('settings_data', '{}')
            settings = json.loads(settings_json)
            # Save contact_methods back to ProgramPulse_Settings (shared)
            contact_methods = settings.pop('contact_methods', None)
            if contact_methods is not None:
                pp_settings = load_content("ProgramPulse_Settings", {})
                pp_settings['contact_methods'] = contact_methods
                save_content("ProgramPulse_Settings", pp_settings)
            save_settings(settings)
            response = {'success': True, 'message': 'Settings saved'}

        elif action == 'check_sched_install':
            # Report whether our managed block is in ScheduledTasks + whether
            # any other reference to this script exists outside the block.
            sn = get_pb_script_name()
            try:
                existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ""
            except:
                existing = ""
            installed = (_SCHED_MARKER_START in existing)
            ref_outside = False
            try:
                stripped = existing
                if installed:
                    import re as _re
                    pat = _re.escape(_SCHED_MARKER_START) + r".*?" + _re.escape(_SCHED_MARKER_END)
                    stripped = _re.sub(pat, "", existing, flags=_re.DOTALL)
                if (('CallScript("' + sn + '")' in stripped)
                    or ("CallScript('" + sn + "')" in stripped)):
                    ref_outside = True
            except:
                pass
            response = {
                'success': True,
                'installed': installed,
                'referencedOutsideBlock': ref_outside,
                'scriptName': sn,
                'contentSlot': _SCHED_CONTENT_SLOT,
            }

        elif action == 'install_sched':
            sn = get_pb_script_name()
            try:
                existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ""
            except Exception as _e_read:
                response = {'success': False,
                            'message': 'Could not read ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_e_read)}
            else:
                if _SCHED_MARKER_START in existing:
                    response = {'success': True,
                                'message': 'Already installed in ' + _SCHED_CONTENT_SLOT + '.',
                                'alreadyInstalled': True}
                else:
                    block = (
                        _SCHED_MARKER_START + "\n"
                        "try:\n"
                        "    Data.run_batch = 'true'\n"
                        "    model.CallScript('" + sn + "')\n"
                        "except Exception as _pb_e:\n"
                        "    print 'Prospect Builder batch error: ' + str(_pb_e)\n"
                        + _SCHED_MARKER_END + "\n"
                    )
                    new_content = (existing.rstrip()
                                   + ("\n\n" if existing.strip() else "")
                                   + block)
                    try:
                        model.WriteContentPython(_SCHED_CONTENT_SLOT, new_content)
                        response = {'success': True,
                                    'message': 'Added to ' + _SCHED_CONTENT_SLOT
                                               + '. Sends will fire on the next scheduled run.'}
                    except Exception as _e_w:
                        response = {'success': False,
                                    'message': 'Could not write ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_e_w)}

        elif action == 'uninstall_sched':
            try:
                existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ""
            except Exception as _e_r2:
                response = {'success': False,
                            'message': 'Could not read ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_e_r2)}
            else:
                if _SCHED_MARKER_START not in existing:
                    response = {'success': True,
                                'message': 'Not currently installed in ' + _SCHED_CONTENT_SLOT + '.',
                                'notInstalled': True}
                else:
                    import re as _re
                    pat = _re.escape(_SCHED_MARKER_START) + r".*?" + _re.escape(_SCHED_MARKER_END) + r"\n?"
                    new_content = _re.sub(pat, "", existing, flags=_re.DOTALL)
                    new_content = _re.sub(r"\n{3,}", "\n\n", new_content).rstrip() + "\n"
                    try:
                        model.WriteContentPython(_SCHED_CONTENT_SLOT, new_content)
                        response = {'success': True,
                                    'message': 'Removed from ' + _SCHED_CONTENT_SLOT + '.'}
                    except Exception as _e_w2:
                        response = {'success': False,
                                    'message': 'Could not write ' + _SCHED_CONTENT_SLOT + ': ' + safe_str(_e_w2)}

        elif action == 'get_keywords':
            sql = """
                SELECT KeywordId, Description
                FROM dbo.Keyword
                WHERE IsActive = 1
                ORDER BY Description
            """
            keywords = []
            for row in q.QuerySql(sql):
                keywords.append({'keywordId': int(row.KeywordId), 'description': safe_str(row.Description)})
            response = {'success': True, 'keywords': keywords}

        elif action == 'get_field_catalog':
            response = {'success': True, 'catalog': FIELD_CATALOG}

        # ==========================================================
        # DATA LOADING - Phase 1: Core prospect data
        # ==========================================================
        elif action == 'load_prospects':
            config_json = get_form_data('config_data', '{}')
            config = json.loads(config_json)
            source = config.get('source', {})
            source_type = source.get('pb_type', 'involvement')
            member_types = config.get('memberTypes', [])
            max_prospects = int(config.get('maxProspects', 0)) or 2000

            prospects = []
            mt_filter = ','.join(str(m) for m in member_types) if member_types else ''

            if source_type == 'involvement':
                org_id = int(source.get('orgId', 0))
                if org_id > 0:
                    sql = """
                        SELECT TOP {2}
                            p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
                            p.PreferredName, p.EmailAddress, p.CellPhone, p.HomePhone,
                            p.Age, CONVERT(VARCHAR(10), p.BDate, 120) as BDate,
                            g.Description as Gender,
                            ms.Description as MemberStatus,
                            mar.Description as MaritalStatus,
                            CONVERT(VARCHAR(10), p.JoinDate, 120) as JoinDate,
                            c.Description as CampusName,
                            p.FamilyId, p.PositionInFamilyId,
                            pic.ThumbUrl as PhotoUrl,
                            om.MemberTypeId,
                            mt.Description as MemberType,
                            ISNULL(p.PrimaryAddress, '') + ' ' + ISNULL(p.PrimaryCity, '') + ', ' + ISNULL(p.PrimaryState, '') + ' ' + ISNULL(p.PrimaryZip, '') as FullAddress,
                            om.EnrollmentDate,
                            CASE WHEN mar.Description = 'Widowed' THEN
                                COALESCE(
                                    (SELECT TOP 1 CONVERT(VARCHAR(10), sp.DeceasedDate, 120)
                                     FROM People sp WHERE sp.FamilyId = p.FamilyId
                                     AND sp.PeopleId != p.PeopleId
                                     AND sp.DeceasedDate IS NOT NULL
                                     ORDER BY sp.DeceasedDate DESC),
                                    (SELECT TOP 1 'date unknown'
                                     FROM People sp WHERE sp.FamilyId = p.FamilyId
                                     AND sp.PeopleId != p.PeopleId AND sp.IsDeceased = 1)
                                )
                            END AS SpouseDeceasedDate
                        FROM OrganizationMembers om
                        JOIN People p ON om.PeopleId = p.PeopleId
                        LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                        LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                        LEFT JOIN lookup.MaritalStatus mar ON p.MaritalStatusId = mar.Id
                        LEFT JOIN lookup.Campus c ON p.CampusId = c.Id
                        LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                        LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                        WHERE om.OrganizationId = {0}
                          {1}
                          AND p.IsDeceased = 0
                        ORDER BY p.Name2
                    """.format(org_id, 'AND om.MemberTypeId IN (' + mt_filter + ')' if mt_filter else '', max_prospects)
                    for r in q.QuerySql(sql):
                        prospects.append({
                            'PeopleId': r.PeopleId,
                            'Name2': safe_str(r.Name2),
                            'FirstName': safe_str(r.FirstName),
                            'LastName': safe_str(r.LastName),
                            'NickName': safe_str(r.NickName),
                            'PreferredName': safe_str(r.PreferredName),
                            'EmailAddress': safe_str(r.EmailAddress),
                            'CellPhone': safe_str(r.CellPhone),
                            'HomePhone': safe_str(r.HomePhone),
                            'Age': r.Age if r.Age else None,
                            'BDate': safe_str(r.BDate),
                            'Gender': safe_str(r.Gender),
                            'MemberStatus': safe_str(r.MemberStatus),
                            'MaritalStatus': safe_str(r.MaritalStatus),
                            'JoinDate': safe_str(r.JoinDate),
                            'CampusName': safe_str(r.CampusName),
                            'FamilyId': r.FamilyId,
                            'PositionInFamilyId': r.PositionInFamilyId,
                            'PhotoUrl': safe_str(r.PhotoUrl) if r.PhotoUrl else '',
                            'MemberTypeId': r.MemberTypeId,
                            'MemberType': safe_str(r.MemberType),
                            'FullAddress': safe_str(r.FullAddress),
                            'EnrollmentDate': safe_str(r.EnrollmentDate),
                            'SpouseDeceasedDate': safe_str(r.SpouseDeceasedDate) if r.SpouseDeceasedDate else '',
                        })

            elif source_type == 'tag':
                tag_name = source.get('tagName', '')
                if tag_name:
                    safe_tag = tag_name.replace("'", "''")
                    sql = """
                        SELECT TOP {1}
                            p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
                            p.PreferredName, p.EmailAddress, p.CellPhone, p.HomePhone,
                            p.Age, CONVERT(VARCHAR(10), p.BDate, 120) as BDate,
                            g.Description as Gender,
                            ms.Description as MemberStatus,
                            mar.Description as MaritalStatus,
                            CONVERT(VARCHAR(10), p.JoinDate, 120) as JoinDate,
                            c.Description as CampusName,
                            p.FamilyId, p.PositionInFamilyId,
                            pic.ThumbUrl as PhotoUrl,
                            0 as MemberTypeId, '' as MemberType,
                            ISNULL(p.PrimaryAddress, '') + ' ' + ISNULL(p.PrimaryCity, '') + ', ' + ISNULL(p.PrimaryState, '') + ' ' + ISNULL(p.PrimaryZip, '') as FullAddress,
                            NULL as EnrollmentDate,
                            CASE WHEN mar.Description = 'Widowed' THEN
                                COALESCE(
                                    (SELECT TOP 1 CONVERT(VARCHAR(10), sp.DeceasedDate, 120)
                                     FROM People sp WHERE sp.FamilyId = p.FamilyId
                                     AND sp.PeopleId != p.PeopleId
                                     AND sp.DeceasedDate IS NOT NULL
                                     ORDER BY sp.DeceasedDate DESC),
                                    (SELECT TOP 1 'date unknown'
                                     FROM People sp WHERE sp.FamilyId = p.FamilyId
                                     AND sp.PeopleId != p.PeopleId AND sp.IsDeceased = 1)
                                )
                            END AS SpouseDeceasedDate
                        FROM People p
                        JOIN TagPerson tp ON tp.PeopleId = p.PeopleId
                        JOIN Tag t ON tp.Id = t.Id
                        LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                        LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                        LEFT JOIN lookup.MaritalStatus mar ON p.MaritalStatusId = mar.Id
                        LEFT JOIN lookup.Campus c ON p.CampusId = c.Id
                        LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                        WHERE t.Name = '{0}'
                          AND p.IsDeceased = 0
                        ORDER BY p.Name2
                    """.format(safe_tag, max_prospects)
                    for r in q.QuerySql(sql):
                        prospects.append({
                            'PeopleId': r.PeopleId,
                            'Name2': safe_str(r.Name2),
                            'FirstName': safe_str(r.FirstName),
                            'LastName': safe_str(r.LastName),
                            'NickName': safe_str(r.NickName),
                            'PreferredName': safe_str(r.PreferredName),
                            'EmailAddress': safe_str(r.EmailAddress),
                            'CellPhone': safe_str(r.CellPhone),
                            'HomePhone': safe_str(r.HomePhone),
                            'Age': r.Age if r.Age else None,
                            'BDate': safe_str(r.BDate),
                            'Gender': safe_str(r.Gender),
                            'MemberStatus': safe_str(r.MemberStatus),
                            'MaritalStatus': safe_str(r.MaritalStatus),
                            'JoinDate': safe_str(r.JoinDate),
                            'CampusName': safe_str(r.CampusName),
                            'FamilyId': r.FamilyId,
                            'PositionInFamilyId': r.PositionInFamilyId,
                            'PhotoUrl': safe_str(r.PhotoUrl) if r.PhotoUrl else '',
                            'MemberTypeId': 0,
                            'MemberType': '',
                            'FullAddress': safe_str(r.FullAddress),
                            'EnrollmentDate': '',
                            'SpouseDeceasedDate': safe_str(r.SpouseDeceasedDate) if r.SpouseDeceasedDate else '',
                        })

            elif source_type == 'query':
                query_code = source.get('queryId', '')
                if query_code:
                    # Use q.QueryList to execute saved query, then get full person data
                    try:
                        query_people = q.QueryList(query_code)
                        query_pids = []
                        for qp in query_people:
                            query_pids.append(str(qp.PeopleId))
                        if query_pids:
                            pid_list = ','.join(query_pids[:max_prospects])
                            sql = """
                                SELECT
                                    p.PeopleId, p.Name2, p.FirstName, p.LastName, p.NickName,
                                    p.PreferredName, p.EmailAddress, p.CellPhone, p.HomePhone,
                                    p.Age, CONVERT(VARCHAR(10), p.BDate, 120) as BDate,
                                    g.Description as Gender,
                                    ms.Description as MemberStatus,
                                    mar.Description as MaritalStatus,
                                    CONVERT(VARCHAR(10), p.JoinDate, 120) as JoinDate,
                                    c.Description as CampusName,
                                    p.FamilyId, p.PositionInFamilyId,
                            pic.ThumbUrl as PhotoUrl,
                                    0 as MemberTypeId, '' as MemberType,
                                    ISNULL(p.PrimaryAddress, '') + ' ' + ISNULL(p.PrimaryCity, '') + ', ' + ISNULL(p.PrimaryState, '') + ' ' + ISNULL(p.PrimaryZip, '') as FullAddress,
                                    NULL as EnrollmentDate,
                                    CASE WHEN p.MaritalStatusId = 40 THEN
                                        (SELECT TOP 1 CONVERT(VARCHAR(10), sp.DeceasedDate, 120)
                                         FROM People sp WHERE sp.FamilyId = p.FamilyId
                                         AND sp.PeopleId != p.PeopleId AND sp.IsDeceased = 1
                                         AND sp.DeceasedDate IS NOT NULL
                                         ORDER BY sp.DeceasedDate DESC)
                                    END AS SpouseDeceasedDate
                                FROM People p
                                LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                                LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                                LEFT JOIN lookup.MaritalStatus mar ON p.MaritalStatusId = mar.Id
                                LEFT JOIN lookup.Campus c ON p.CampusId = c.Id
                                LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                                WHERE p.PeopleId IN ({0})
                                  AND p.IsDeceased = 0
                                ORDER BY p.Name2
                            """.format(pid_list)
                            for r in q.QuerySql(sql):
                                prospects.append({
                                    'PeopleId': r.PeopleId,
                                    'Name2': safe_str(r.Name2),
                                    'FirstName': safe_str(r.FirstName),
                                    'LastName': safe_str(r.LastName),
                                    'NickName': safe_str(r.NickName),
                                    'PreferredName': safe_str(r.PreferredName),
                                    'EmailAddress': safe_str(r.EmailAddress),
                                    'CellPhone': safe_str(r.CellPhone),
                                    'HomePhone': safe_str(r.HomePhone),
                                    'Age': r.Age if r.Age else None,
                                    'BDate': safe_str(r.BDate),
                                    'Gender': safe_str(r.Gender),
                                    'MemberStatus': safe_str(r.MemberStatus),
                                    'MaritalStatus': safe_str(r.MaritalStatus),
                                    'JoinDate': safe_str(r.JoinDate),
                                    'CampusName': safe_str(r.CampusName),
                                    'FamilyId': r.FamilyId,
                                    'PositionInFamilyId': r.PositionInFamilyId,
                                    'PhotoUrl': safe_str(r.PhotoUrl) if r.PhotoUrl else '',
                                    'MemberTypeId': 0,
                                    'MemberType': '',
                                    'FullAddress': safe_str(r.FullAddress),
                                    'EnrollmentDate': '',
                                    'SpouseDeceasedDate': safe_str(r.SpouseDeceasedDate) if r.SpouseDeceasedDate else '',
                                })
                    except Exception as e:
                        response = {'success': False, 'message': 'Query error: ' + safe_str(e)}
                        print json.dumps(sanitize_for_json(response))

            response = {'success': True, 'prospects': sanitize_for_json(prospects), 'total': len(prospects)}

        # ==========================================================
        # DATA LOADING - Phase 2: Family data
        # ==========================================================
        elif action == 'load_family_data':
            pids_str = get_form_data('people_ids', '')
            if not pids_str:
                response = {'success': True, 'familyData': {}}
            else:
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                family_map = {}

                if pids:
                    pid_list = ','.join(str(p) for p in pids)
                    # Get family IDs for our prospects
                    fam_sql = """
                        SELECT PeopleId, FamilyId FROM People
                        WHERE PeopleId IN ({0})
                    """.format(pid_list)
                    pid_to_fam = {}
                    fam_ids = set()
                    for r in q.QuerySql(fam_sql):
                        pid_to_fam[r.PeopleId] = r.FamilyId
                        fam_ids.add(r.FamilyId)

                    if fam_ids:
                        fam_list = ','.join(str(f) for f in fam_ids)
                        # Get all family members with member status
                        members_sql = """
                            SELECT p.PeopleId, p.FamilyId, p.Name2, p.FirstName,
                                   p.PositionInFamilyId, p.Age, p.EmailAddress, p.CellPhone,
                                   g.Description as Gender,
                                   ms.Description as MemberStatus,
                                   p.MemberStatusId
                            FROM People p
                            LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                            WHERE p.FamilyId IN ({0})
                              AND p.IsDeceased = 0
                            ORDER BY p.FamilyId, p.PositionInFamilyId
                        """.format(fam_list)

                        fam_members = {}
                        all_fam_pids = []
                        for r in q.QuerySql(members_sql):
                            fid = r.FamilyId
                            if fid not in fam_members:
                                fam_members[fid] = []
                            all_fam_pids.append(r.PeopleId)
                            fam_members[fid].append({
                                'PeopleId': r.PeopleId,
                                'Name2': safe_str(r.Name2),
                                'FirstName': safe_str(r.FirstName),
                                'Position': r.PositionInFamilyId,
                                'Age': r.Age if r.Age else None,
                                'Email': safe_str(r.EmailAddress),
                                'Phone': safe_str(r.CellPhone),
                                'Gender': safe_str(r.Gender),
                                'MemberStatus': safe_str(r.MemberStatus),
                                'MemberStatusId': r.MemberStatusId,
                            })

                        # Get involvement details per family member
                        fam_inv_data = {}  # pid -> [{name, memberType}]
                        fam_last_attend = {}  # pid -> last attendance date string
                        if all_fam_pids:
                            fam_pid_list = ','.join(str(p) for p in all_fam_pids)
                            inv_detail_sql = """
                                SELECT om.PeopleId, o.OrganizationName,
                                       mt.Description as MemberType
                                FROM OrganizationMembers om
                                JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                                LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                                WHERE om.PeopleId IN ({0})
                                  AND o.OrganizationStatusId = 30
                                ORDER BY om.PeopleId, o.OrganizationName
                            """.format(fam_pid_list)
                            for r in q.QuerySql(inv_detail_sql):
                                pid = r.PeopleId
                                if pid not in fam_inv_data:
                                    fam_inv_data[pid] = []
                                fam_inv_data[pid].append({
                                    'name': safe_str(r.OrganizationName),
                                    'type': safe_str(r.MemberType),
                                })
                            # Last attendance per family member
                            attend_sql = """
                                SELECT a.PeopleId,
                                       CONVERT(VARCHAR(10), MAX(m.MeetingDate), 120) as LastAttend
                                FROM Attend a
                                JOIN Meetings m ON a.MeetingId = m.MeetingId
                                WHERE a.PeopleId IN ({0})
                                  AND a.AttendanceFlag = 1
                                  AND m.MeetingDate >= DATEADD(year, -1, GETDATE())
                                GROUP BY a.PeopleId
                            """.format(fam_pid_list)
                            try:
                                for r in q.QuerySql(attend_sql):
                                    fam_last_attend[r.PeopleId] = safe_str(r.LastAttend)
                            except:
                                pass

                        # Build family data per prospect
                        def _member_line(m):
                            """Build a rich single-line description for a family member."""
                            parts = [safe_str(m['Name2'])]
                            if m.get('Age'):
                                parts.append('age ' + str(m['Age']))
                            if m.get('MemberStatus'):
                                parts.append(safe_str(m['MemberStatus']))
                            invs = fam_inv_data.get(m['PeopleId'], [])
                            if invs:
                                parts.append(str(len(invs)) + ' involvement' + ('s' if len(invs) != 1 else ''))
                            la = fam_last_attend.get(m['PeopleId'], '')
                            if la:
                                parts.append('last attended ' + la)
                            return ' | '.join(parts)

                        def _member_detail(m):
                            """Build a detail dict for a family member."""
                            invs = fam_inv_data.get(m['PeopleId'], [])
                            inv_names = [i['name'] + ' (' + i['type'] + ')' for i in invs] if invs else []
                            pos_map = {10: 'Adult', 20: 'Adult', 30: 'Child'}
                            return {
                                'pid': m['PeopleId'],
                                'name': safe_str(m['Name2']),
                                'age': m.get('Age'),
                                'gender': m.get('Gender', ''),
                                'posLabel': pos_map.get(m.get('Position', 0), 'Other'),
                                'status': safe_str(m.get('MemberStatus', '')),
                                'statusId': m.get('MemberStatusId', 0),
                                'phone': m.get('Phone', ''),
                                'email': m.get('Email', ''),
                                'invCount': len(invs),
                                'invNames': inv_names,
                                'lastAttend': fam_last_attend.get(m['PeopleId'], ''),
                                'position': m.get('Position', 0),
                            }

                        for pid in pids:
                            fid = pid_to_fam.get(pid)
                            if not fid or fid not in fam_members:
                                continue
                            members = fam_members[fid]
                            me = None
                            spouse = None
                            parents = []
                            children = []
                            all_members = []
                            all_detail = []
                            for m in members:
                                if m['PeopleId'] == pid:
                                    me = m
                                    continue
                                all_members.append(safe_str(m['Name2']))
                                all_detail.append(_member_detail(m))
                                if m['Position'] == 10 or m['Position'] == 20:
                                    if me and me.get('Position') in (10, 20) and m['Position'] in (10, 20):
                                        spouse = m
                                    else:
                                        parents.append(m)
                                elif m['Position'] == 30:
                                    children.append(m)
                                else:
                                    if m.get('Age') and m['Age'] >= 18:
                                        parents.append(m)
                                    else:
                                        children.append(m)

                            # Build rich FamilyDetail lines
                            detail_lines = []
                            if spouse:
                                detail_lines.append('Spouse: ' + _member_line(spouse))
                            for p in parents:
                                detail_lines.append('Parent: ' + _member_line(p))
                            for c in children:
                                detail_lines.append('Child: ' + _member_line(c))

                            # Spouse detail for single-view
                            spouse_detail = ''
                            if spouse:
                                sd_parts = []
                                if spouse.get('MemberStatus'):
                                    sd_parts.append(safe_str(spouse['MemberStatus']))
                                if spouse.get('Phone'):
                                    sd_parts.append(safe_str(spouse['Phone']))
                                if spouse.get('Email'):
                                    sd_parts.append(safe_str(spouse['Email']))
                                sc = fam_inv_counts.get(spouse['PeopleId'], 0)
                                if sc > 0:
                                    sd_parts.append(str(sc) + ' involvement' + ('s' if sc != 1 else ''))
                                spouse_detail = ' | '.join(sd_parts)

                            family_map[str(pid)] = {
                                'SpouseName': safe_str(spouse['Name2']) if spouse else '',
                                'SpouseDetail': spouse_detail,
                                'Parents': ', '.join(safe_str(p['Name2']) for p in parents),
                                'ParentPhones': ', '.join(safe_str(p['Phone']) for p in parents if p.get('Phone')),
                                'ParentEmails': ', '.join(safe_str(p['Email']) for p in parents if p.get('Email')),
                                'Children': ', '.join(_member_line(c) for c in children),
                                'FamilyMembers': ', '.join(all_members),
                                'FamilyDetail': '\n'.join(detail_lines),
                                'FamilyDetailList': all_detail,
                            }

                response = {'success': True, 'familyData': sanitize_for_json(family_map)}

        # ==========================================================
        # DATA LOADING - Phase 3: Extra values
        # ==========================================================
        elif action == 'load_extra_values':
            pids_str = get_form_data('people_ids', '')
            ev_fields_str = get_form_data('ev_fields', '')
            ev_map = {}

            if pids_str and ev_fields_str:
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                ev_fields = [f.strip() for f in ev_fields_str.split('|') if f.strip()]

                if pids and ev_fields:
                    pid_list = ','.join(str(p) for p in pids)
                    safe_names = ["'" + n.replace("'", "''") + "'" for n in ev_fields]
                    ev_sql = """
                        SELECT PeopleId, Field,
                               StrValue, Data,
                               CAST(DateValue AS NVARCHAR(50)) as DateVal,
                               IntValue, BitValue
                        FROM PeopleExtra
                        WHERE PeopleId IN ({0})
                          AND Field IN ({1})
                    """.format(pid_list, ','.join(safe_names))

                    for r in q.QuerySql(ev_sql):
                        pid_key = str(r.PeopleId)
                        if pid_key not in ev_map:
                            ev_map[pid_key] = {}
                        val = ''
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
                        ev_map[pid_key][safe_str(r.Field)] = val

            response = {'success': True, 'extraValues': sanitize_for_json(ev_map)}

        # ==========================================================
        # DATA LOADING - Phase 4: Contact efforts (TaskNote keywords)
        # ==========================================================
        elif action == 'load_contact_efforts':
            pids_str = get_form_data('people_ids', '')
            contact_map = {}

            if pids_str:
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                if pids:
                    pid_list = ','.join(str(p) for p in pids)

                    # Load contact methods from ProgramPulse shared settings
                    pp_settings = load_content("ProgramPulse_Settings", {})
                    methods = pp_settings.get('contact_methods', [])

                    keyword_id_to_code = {}
                    all_keyword_ids = []
                    for m in methods:
                        kid = m.get('keywordId')
                        if kid:
                            keyword_id_to_code[int(kid)] = m.get('code', '?')
                            all_keyword_ids.append(int(kid))

                    # Step 1: Get total TaskNote counts per person
                    total_sql = """
                        SELECT tn.AboutPersonId as PeopleId,
                               COUNT(*) as TotalCount,
                               MAX(tn.CreatedDate) as LastContactDate
                        FROM TaskNote tn WITH (NOLOCK)
                        WHERE tn.AboutPersonId IN ({0})
                          AND tn.CreatedDate >= DATEADD(week, -26, GETDATE())
                        GROUP BY tn.AboutPersonId
                    """.format(pid_list)
                    for cr in q.QuerySql(total_sql):
                        pid_key = str(cr.PeopleId)
                        contact_map[pid_key] = {
                            'methods': {},
                            'total': int(cr.TotalCount) if cr.TotalCount else 0,
                            'lastDate': safe_str(cr.LastContactDate)
                        }

                    # Step 2: Get per-keyword counts via TaskNoteKeyword join
                    if all_keyword_ids and contact_map:
                        contacted_pids = ','.join(str(p) for p in contact_map.keys())
                        kid_list = ','.join(str(k) for k in all_keyword_ids)
                        method_sql = """
                            SELECT tn.AboutPersonId as PeopleId, tnk.KeywordId, COUNT(*) as Cnt
                            FROM TaskNote tn WITH (NOLOCK)
                            JOIN TaskNoteKeyword tnk ON tnk.TaskNoteId = tn.TaskNoteId
                            WHERE tn.AboutPersonId IN ({0})
                              AND tnk.KeywordId IN ({1})
                              AND tn.CreatedDate >= DATEADD(week, -26, GETDATE())
                            GROUP BY tn.AboutPersonId, tnk.KeywordId
                        """.format(contacted_pids, kid_list)
                        try:
                            for row in q.QuerySql(method_sql):
                                pid_key = str(row.PeopleId)
                                if pid_key in contact_map:
                                    code = keyword_id_to_code.get(int(row.KeywordId), '')
                                    if code:
                                        contact_map[pid_key]['methods'][code] = int(row.Cnt) if row.Cnt else 0
                        except:
                            pass

                    # Load other_weight setting
                    other_weight = 1.0
                    try:
                        pb_settings = load_content(SETTINGS_KEY, {})
                        ow = pb_settings.get('other_weight', 1)
                        if ow is not None:
                            other_weight = float(ow)
                    except:
                        pass

                    # Calculate Other counts (total minus matched keywords)
                    for pid_key in contact_map:
                        matched_total = sum(contact_map[pid_key]['methods'].values())
                        other = contact_map[pid_key]['total'] - matched_total
                        if other > 0:
                            contact_map[pid_key]['methods']['O'] = other
                        # Weighted total: matched efforts at full weight, Other at configured weight
                        contact_map[pid_key]['weightedTotal'] = int(matched_total + (max(other, 0) * other_weight))

            response = {'success': True, 'contactMap': sanitize_for_json(contact_map)}

        # ==========================================================
        # DATA LOADING - Phase 5: Cross-query flags
        # ==========================================================
        elif action == 'evaluate_cross_flags':
            pids_str = get_form_data('people_ids', '')
            flags_json = get_form_data('flags_data', '[]')
            flags = json.loads(flags_json)
            flag_results = {}

            if pids_str and flags:
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                pid_list = ','.join(str(p) for p in pids)

                for flag in flags:
                    ftype = flag.get('pb_type', '')
                    fid = flag.get('id', ftype)
                    matching_pids = []

                    try:
                        if ftype == 'children_attending':
                            prog_id = int(flag.get('programId', 0))
                            if prog_id > 0:
                                sql = """
                                    SELECT DISTINCT parent.PeopleId
                                    FROM People parent
                                    JOIN People child ON child.FamilyId = parent.FamilyId
                                        AND child.PositionInFamilyId = 30
                                    JOIN OrganizationMembers om ON om.PeopleId = child.PeopleId
                                    JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                                    JOIN Division d ON o.DivisionId = d.Id
                                    WHERE parent.PeopleId IN ({0})
                                      AND d.ProgId = {1}
                                      AND o.OrganizationStatusId = 30
                                """.format(pid_list, prog_id)
                                for r in q.QuerySql(sql):
                                    matching_pids.append(r.PeopleId)

                        elif ftype == 'parents_not_attending':
                            prog_id = int(flag.get('programId', 0))
                            if prog_id > 0:
                                sql = """
                                    SELECT DISTINCT parent.PeopleId
                                    FROM People parent
                                    JOIN People child ON child.FamilyId = parent.FamilyId
                                        AND child.PositionInFamilyId = 30
                                    JOIN OrganizationMembers om_child ON om_child.PeopleId = child.PeopleId
                                    JOIN Organizations o_child ON o_child.OrganizationId = om_child.OrganizationId
                                    JOIN Division d ON o_child.DivisionId = d.Id
                                    LEFT JOIN OrganizationMembers om_parent ON om_parent.PeopleId = parent.PeopleId
                                        AND om_parent.OrganizationId IN (
                                            SELECT o2.OrganizationId FROM Organizations o2
                                            JOIN Division d2 ON o2.DivisionId = d2.Id
                                            WHERE d2.ProgId = {1} AND o2.OrganizationStatusId = 30
                                        )
                                    WHERE parent.PeopleId IN ({0})
                                      AND d.ProgId = {1}
                                      AND o_child.OrganizationStatusId = 30
                                      AND parent.PositionInFamilyId IN (10, 20)
                                      AND om_parent.PeopleId IS NULL
                                """.format(pid_list, prog_id)
                                for r in q.QuerySql(sql):
                                    matching_pids.append(r.PeopleId)

                        elif ftype == 'spouse_in_org':
                            org_id = int(flag.get('orgId', 0))
                            if org_id > 0:
                                sql = """
                                    SELECT DISTINCT prospect.PeopleId
                                    FROM People prospect
                                    JOIN People spouse ON spouse.FamilyId = prospect.FamilyId
                                        AND spouse.PeopleId != prospect.PeopleId
                                        AND spouse.PositionInFamilyId IN (10, 20)
                                    JOIN OrganizationMembers om ON om.PeopleId = spouse.PeopleId
                                    WHERE prospect.PeopleId IN ({0})
                                      AND prospect.PositionInFamilyId IN (10, 20)
                                      AND om.OrganizationId = {1}
                                """.format(pid_list, org_id)
                                for r in q.QuerySql(sql):
                                    matching_pids.append(r.PeopleId)

                        elif ftype == 'has_extra_value':
                            ev_field = flag.get('evField', '')
                            if ev_field:
                                safe_field = ev_field.replace("'", "''")
                                sql = """
                                    SELECT DISTINCT PeopleId
                                    FROM PeopleExtra
                                    WHERE PeopleId IN ({0})
                                      AND Field = '{1}'
                                      AND (StrValue IS NOT NULL AND StrValue != ''
                                           OR Data IS NOT NULL
                                           OR DateValue IS NOT NULL
                                           OR IntValue IS NOT NULL
                                           OR BitValue = 1)
                                """.format(pid_list, safe_field)
                                for r in q.QuerySql(sql):
                                    matching_pids.append(r.PeopleId)

                    except Exception as e:
                        pass  # Skip failed flags silently

                    flag_results[fid] = matching_pids

            response = {'success': True, 'flagResults': sanitize_for_json(flag_results)}

        # ==========================================================
        # DATA LOADING - Phase 5b: Prospect Scorecard
        # ==========================================================
        elif action == 'compute_scores':
            pids_str = get_form_data('people_ids', '')
            sc_json = get_form_data('scorecard_config', '{}')
            member_json = get_form_data('member_data', '{}')
            scorecard = json.loads(sc_json)
            member_data = json.loads(member_json)  # pid -> {status, weightedTotal}
            scores = {}

            if pids_str and scorecard.get('enabled'):
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                pid_list = ','.join(str(p) for p in pids)
                factors = scorecard.get('factors', {})

                # Gather factor data per person
                factor_data = {}
                for p in pids:
                    factor_data[str(p)] = member_data.get(str(p), {})

                # SQL 1: Attendance metrics
                if factors.get('attend_recency', {}).get('enabled') or factors.get('attend_frequency', {}).get('enabled'):
                    try:
                        att_sql = """
                            SELECT a.PeopleId,
                                   DATEDIFF(day, MAX(a.MeetingDate), GETDATE()) AS DaysSinceLast,
                                   COUNT(CASE WHEN a.MeetingDate >= DATEADD(day, -90, GETDATE()) THEN 1 END) AS Attend90
                            FROM Attend a WITH (NOLOCK)
                            WHERE a.PeopleId IN ({0})
                              AND a.AttendanceFlag = 1
                              AND a.MeetingDate >= DATEADD(day, -180, GETDATE())
                            GROUP BY a.PeopleId
                        """.format(pid_list)
                        for r in q.QuerySql(att_sql):
                            pk = str(r.PeopleId)
                            if pk in factor_data:
                                factor_data[pk]['daysSinceLast'] = int(r.DaysSinceLast) if r.DaysSinceLast is not None else None
                                factor_data[pk]['attend90'] = int(r.Attend90) if r.Attend90 else 0
                    except:
                        pass

                # SQL 2: Involvements + enrollment recency
                if factors.get('involvements', {}).get('enabled') or factors.get('enrollment_recency', {}).get('enabled'):
                    try:
                        inv_sql = """
                            SELECT om.PeopleId,
                                   COUNT(*) AS InvCount,
                                   MIN(DATEDIFF(day, om.EnrollmentDate, GETDATE())) AS NewestEnrollDays
                            FROM OrganizationMembers om WITH (NOLOCK)
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            WHERE om.PeopleId IN ({0})
                              AND om.InactiveDate IS NULL
                              AND o.OrganizationStatusId = 30
                            GROUP BY om.PeopleId
                        """.format(pid_list)
                        for r in q.QuerySql(inv_sql):
                            pk = str(r.PeopleId)
                            if pk in factor_data:
                                factor_data[pk]['invCount'] = int(r.InvCount) if r.InvCount else 0
                                factor_data[pk]['newestEnrollDays'] = int(r.NewestEnrollDays) if r.NewestEnrollDays is not None else None
                    except:
                        pass

                # SQL 3: Family engagement
                if factors.get('family_engaged', {}).get('enabled'):
                    try:
                        fam_sql = """
                            SELECT DISTINCT p.PeopleId
                            FROM People p WITH (NOLOCK)
                            WHERE p.PeopleId IN ({0})
                              AND EXISTS (
                                SELECT 1 FROM People fam WITH (NOLOCK)
                                JOIN OrganizationMembers om WITH (NOLOCK) ON fam.PeopleId = om.PeopleId
                                WHERE fam.FamilyId = p.FamilyId
                                  AND fam.PeopleId != p.PeopleId
                                  AND fam.IsDeceased = 0
                                  AND om.InactiveDate IS NULL
                                  AND om.MemberTypeId IN (220, 140, 310, 710)
                              )
                        """.format(pid_list)
                        for r in q.QuerySql(fam_sql):
                            pk = str(r.PeopleId)
                            if pk in factor_data:
                                factor_data[pk]['familyEngaged'] = True
                    except:
                        pass

                # SQL 4: Serving roles (leader, volunteer, assistant leader)
                if factors.get('serving_roles', {}).get('enabled'):
                    try:
                        srv_sql = """
                            SELECT om.PeopleId, COUNT(*) AS ServingCount
                            FROM OrganizationMembers om WITH (NOLOCK)
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            WHERE om.PeopleId IN ({0})
                              AND om.MemberTypeId IN (140, 310, 320, 710)
                              AND om.InactiveDate IS NULL
                              AND o.OrganizationStatusId = 30
                            GROUP BY om.PeopleId
                        """.format(pid_list)
                        for r in q.QuerySql(srv_sql):
                            pk = str(r.PeopleId)
                            if pk in factor_data:
                                factor_data[pk]['servingCount'] = int(r.ServingCount) if r.ServingCount else 0
                    except:
                        pass

                # Compute sub-scores and weighted composite
                def score_factor(name, d):
                    if name == 'contact_efforts':
                        wt = d.get('weightedTotal', 0)
                        if wt >= 10: return 100
                        if wt >= 7: return 80
                        if wt >= 4: return 60
                        if wt >= 2: return 40
                        if wt >= 1: return 20
                        return 0
                    elif name == 'attend_recency':
                        days = d.get('daysSinceLast')
                        if days is None: return 0
                        if days <= 7: return 100
                        if days <= 14: return 85
                        if days <= 30: return 70
                        if days <= 60: return 40
                        if days <= 90: return 20
                        return 5
                    elif name == 'attend_frequency':
                        c = d.get('attend90', 0)
                        return min(100, int(float(c) / 13 * 100))
                    elif name == 'involvements':
                        c = d.get('invCount', 0)
                        if c >= 4: return 100
                        if c >= 3: return 80
                        if c >= 2: return 60
                        if c >= 1: return 40
                        return 0
                    elif name == 'serving_roles':
                        c = d.get('servingCount', 0)
                        if c >= 3: return 100
                        if c >= 2: return 80
                        if c >= 1: return 50
                        return 0
                    elif name == 'enrollment_recency':
                        days = d.get('newestEnrollDays')
                        if days is None: return 0
                        if days <= 30: return 100
                        if days <= 90: return 75
                        if days <= 180: return 50
                        if days <= 365: return 25
                        return 10
                    elif name == 'family_engaged':
                        return 100 if d.get('familyEngaged') else 0
                    elif name == 'member_status':
                        s = (d.get('memberStatus') or '').lower()
                        if 'prospect' in s: return 100
                        if 'visitor' in s: return 80
                        if 'not member' in s: return 60
                        if 'just added' in s: return 90
                        if 'member' in s: return 30
                        return 50
                    elif name == 'tasknote_activity':
                        t = d.get('taskNoteTotal', 0)
                        if t >= 5: return 100
                        if t >= 3: return 70
                        if t >= 1: return 40
                        return 0
                    return 0

                enabled = [(n, f) for n, f in factors.items() if f.get('enabled')]
                total_weight = sum(f.get('weight', 0) for _, f in enabled) or 1

                for p in pids:
                    pk = str(p)
                    d = factor_data.get(pk, {})
                    composite = 0
                    breakdown = {}
                    for fname, fcfg in enabled:
                        sub = score_factor(fname, d)
                        breakdown[fname] = sub
                        composite += sub * (float(fcfg.get('weight', 0)) / total_weight)
                    scores[pk] = {
                        'score': int(round(composite)),
                        'breakdown': breakdown
                    }

            response = {'success': True, 'scores': sanitize_for_json(scores)}

        # ==========================================================
        # DATA LOADING - Phase 6: Registration data
        # ==========================================================
        elif action == 'load_registration_data':
            org_id = int(get_form_data('org_id', '0'))
            pids_str = get_form_data('people_ids', '')
            reg_map = {}

            if org_id > 0 and pids_str:
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                pid_list = ','.join(str(p) for p in pids)

                # Try new RegQuestion/RegAnswer format first
                qa_sql = """
                    SELECT ra.PeopleId, rq.Question, ra.Answer
                    FROM RegAnswer ra
                    JOIN RegQuestion rq ON ra.QuestionId = rq.Id
                    WHERE ra.OrganizationId = {0}
                      AND ra.PeopleId IN ({1})
                """.format(org_id, pid_list)
                found_new = False
                try:
                    for r in q.QuerySql(qa_sql):
                        found_new = True
                        pid_key = str(r.PeopleId)
                        if pid_key not in reg_map:
                            reg_map[pid_key] = {}
                        reg_map[pid_key][safe_str(r.Question)] = safe_str(r.Answer)
                except:
                    pass

                # Fallback to RegistrationData XML
                if not found_new:
                    try:
                        rd_sql = """
                            SELECT PeopleId, Data
                            FROM RegistrationData
                            WHERE OrganizationId = {0}
                              AND PeopleId IN ({1})
                        """.format(org_id, pid_list)
                        for r in q.QuerySql(rd_sql):
                            pid_key = str(r.PeopleId)
                            if pid_key not in reg_map:
                                reg_map[pid_key] = {}
                            # Parse simple XML-like answers
                            data = safe_str(r.Data)
                            if data:
                                import re
                                pairs = re.findall(r'<(\w+)>(.*?)</\1>', data)
                                for k, v in pairs:
                                    reg_map[pid_key][k] = v
                    except:
                        pass

            # Also get question list for this org
            questions = []
            if org_id > 0:
                try:
                    q_sql = """
                        SELECT DISTINCT rq.Question
                        FROM RegQuestion rq
                        WHERE rq.OrganizationId = {0}
                        ORDER BY rq.Question
                    """.format(org_id)
                    for r in q.QuerySql(q_sql):
                        questions.append(safe_str(r.Question))
                except:
                    pass

            response = {'success': True, 'regData': sanitize_for_json(reg_map), 'questions': questions}

        # ==========================================================
        # DATA LOADING - Phase 7: Involvement associations
        # ==========================================================
        elif action == 'load_involvement_data':
            pids_str = get_form_data('people_ids', '')
            inv_map = {}

            if pids_str:
                pids = [int(x) for x in pids_str.split(',') if x.strip()]
                if pids:
                    pid_list = ','.join(str(p) for p in pids)
                    inv_sql = """
                        SELECT om.PeopleId, o.OrganizationId, o.OrganizationName,
                               mt.Description as MemberType,
                               p.Name as ProgramName, d.Name as DivisionName,
                               CONVERT(VARCHAR(10), om.EnrollmentDate, 120) as EnrollDate
                        FROM OrganizationMembers om
                        JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                        LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        LEFT JOIN Program p ON d.ProgId = p.Id
                        WHERE om.PeopleId IN ({0})
                          AND o.OrganizationStatusId = 30
                        ORDER BY om.PeopleId, p.Name, o.OrganizationName
                    """.format(pid_list)

                    # Build base involvement list
                    pid_org_pairs = []
                    for r in q.QuerySql(inv_sql):
                        pid_key = str(r.PeopleId)
                        if pid_key not in inv_map:
                            inv_map[pid_key] = []
                        inv_map[pid_key].append({
                            'orgId': r.OrganizationId,
                            'name': safe_str(r.OrganizationName),
                            'memberType': safe_str(r.MemberType),
                            'program': safe_str(r.ProgramName),
                            'division': safe_str(r.DivisionName),
                            'enrollDate': safe_str(r.EnrollDate),
                            'attended': 0,
                            'totalMeetings': 0,
                            'pct': 0,
                            'lastAttend': '',
                        })
                        pid_org_pairs.append((r.PeopleId, r.OrganizationId))

                    # Attendance stats per person+org (last 90 days)
                    if pid_org_pairs:
                        attend_sql = """
                            SELECT a.PeopleId, a.OrganizationId,
                                   COUNT(CASE WHEN a.AttendanceFlag = 1 THEN 1 END) as Attended,
                                   COUNT(*) as TotalMeetings,
                                   CONVERT(VARCHAR(10), MAX(CASE WHEN a.AttendanceFlag = 1 THEN m.MeetingDate END), 120) as LastAttend
                            FROM Attend a
                            JOIN Meetings m ON a.MeetingId = m.MeetingId
                            WHERE a.PeopleId IN ({0})
                              AND m.MeetingDate >= DATEADD(day, -90, GETDATE())
                              AND m.DidNotMeet = 0
                            GROUP BY a.PeopleId, a.OrganizationId
                        """.format(pid_list)
                        attend_map = {}
                        try:
                            for r in q.QuerySql(attend_sql):
                                key = (r.PeopleId, r.OrganizationId)
                                attended = int(r.Attended) if r.Attended else 0
                                total = int(r.TotalMeetings) if r.TotalMeetings else 0
                                attend_map[key] = {
                                    'attended': attended,
                                    'total': total,
                                    'pct': int(round(float(attended) / total * 100)) if total > 0 else 0,
                                    'lastAttend': safe_str(r.LastAttend),
                                }
                        except:
                            pass

                        # Merge attendance into involvement data
                        for pid_key in inv_map:
                            for inv in inv_map[pid_key]:
                                key = (int(pid_key), inv['orgId'])
                                if key in attend_map:
                                    inv['attended'] = attend_map[key]['attended']
                                    inv['totalMeetings'] = attend_map[key]['total']
                                    inv['pct'] = attend_map[key]['pct']
                                    inv['lastAttend'] = attend_map[key]['lastAttend']

            response = {'success': True, 'involvementData': sanitize_for_json(inv_map)}

        # ==========================================================
        # LAZY DETAIL LOAD - Single person detail
        # ==========================================================
        elif action == 'load_person_detail':
            pid = int(get_form_data('people_id', '0'))
            org_id = int(get_form_data('org_id', '0'))
            ev_fields_str = get_form_data('ev_fields', '')
            detail = {'family': {}, 'involvements': [], 'extraValues': {}, 'regData': {}, 'profile': {}, 'milestones': []}

            if pid > 0:
                pid_str = str(pid)
                pid_list = str(pid)

                # Extended profile + milestones + larger photo
                profile_sql = """
                    SELECT p.PeopleId, p.FamilyId,
                           p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip,
                           p.HomePhone, p.WorkPhone,
                           p.EmployerOther, p.OccupationOther, p.SchoolOther,
                           CONVERT(VARCHAR(10), p.BDate, 120) as BirthDate,
                           CONVERT(VARCHAR(10), p.JoinDate, 120) as JoinDate,
                           CONVERT(VARCHAR(10), p.BaptismDate, 120) as BaptismDate,
                           CONVERT(VARCHAR(10), p.WeddingDate, 120) as WeddingDate,
                           CONVERT(VARCHAR(10), p.NewMemberClassDate, 120) as NewMemberClassDate,
                           CONVERT(VARCHAR(10), p.DropDate, 120) as DropDate,
                           CONVERT(VARCHAR(10), p.CreatedDate, 120) as RecordCreated,
                           dt.Description as DecisionType,
                           CONVERT(VARCHAR(10), p.DecisionDate, 120) as DecisionDate,
                           pic.LargeUrl as LargePhotoUrl,
                           pic.MediumUrl as MediumPhotoUrl,
                           pic.SmallUrl as SmallPhotoUrl
                    FROM People p
                    LEFT JOIN lookup.DecisionType dt ON p.DecisionTypeId = dt.Id
                    LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                    WHERE p.PeopleId = {0}
                """.format(pid)
                fam_id = None
                try:
                    for r in q.QuerySql(profile_sql):
                        fam_id = r.FamilyId
                        detail['profile'] = {
                            'address': safe_str((safe_str(r.PrimaryAddress) + ', ' + safe_str(r.PrimaryCity) + ', ' + safe_str(r.PrimaryState) + ' ' + safe_str(r.PrimaryZip)).strip(', ')),
                            'homePhone': safe_str(r.HomePhone),
                            'workPhone': safe_str(r.WorkPhone),
                            'employer': safe_str(r.EmployerOther),
                            'occupation': safe_str(r.OccupationOther),
                            'school': safe_str(r.SchoolOther),
                            'largePhoto': safe_str(r.LargePhotoUrl) if r.LargePhotoUrl else (safe_str(r.MediumPhotoUrl) if r.MediumPhotoUrl else ''),
                            'mediumPhoto': safe_str(r.MediumPhotoUrl) if r.MediumPhotoUrl else (safe_str(r.SmallPhotoUrl) if r.SmallPhotoUrl else ''),
                        }
                        # Build milestones
                        milestones = []
                        if r.BirthDate:
                            milestones.append({'label': 'Birth Date', 'date': safe_str(r.BirthDate), 'icon': '&#127874;'})
                        if r.RecordCreated:
                            milestones.append({'label': 'Record Created', 'date': safe_str(r.RecordCreated), 'icon': '&#128221;'})
                        if r.DecisionDate:
                            milestones.append({'label': safe_str(r.DecisionType) if r.DecisionType else 'Decision', 'date': safe_str(r.DecisionDate), 'icon': '&#10013;'})
                        if r.BaptismDate:
                            milestones.append({'label': 'Baptism', 'date': safe_str(r.BaptismDate), 'icon': '&#128167;'})
                        if r.NewMemberClassDate:
                            milestones.append({'label': 'New Member Class', 'date': safe_str(r.NewMemberClassDate), 'icon': '&#127891;'})
                        if r.JoinDate:
                            milestones.append({'label': 'Joined Church', 'date': safe_str(r.JoinDate), 'icon': '&#127969;'})
                        if r.WeddingDate:
                            milestones.append({'label': 'Wedding', 'date': safe_str(r.WeddingDate), 'icon': '&#128141;'})
                        if r.DropDate:
                            milestones.append({'label': 'Dropped', 'date': safe_str(r.DropDate), 'icon': '&#128308;'})
                        detail['milestones'] = milestones
                except:
                    # Fallback: just get FamilyId
                    for r in q.QuerySql("SELECT FamilyId FROM People WHERE PeopleId = {0}".format(pid)):
                        fam_id = r.FamilyId
                if fam_id:
                    fam_members_sql = """
                        SELECT p.PeopleId, p.Name2, p.FirstName, p.PositionInFamilyId,
                               p.Age, p.EmailAddress, p.CellPhone,
                               g.Description as Gender,
                               ms.Description as MemberStatus, p.MemberStatusId,
                               pic.ThumbUrl as PhotoUrl,
                               pic.MediumUrl as MediumPhotoUrl
                        FROM People p
                        LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                        LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                        LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                        WHERE p.FamilyId = {0} AND p.IsDeceased = 0
                        ORDER BY p.PositionInFamilyId
                    """.format(fam_id)
                    fam_list = []
                    fam_pids = []
                    for r in q.QuerySql(fam_members_sql):
                        if r.PeopleId == pid:
                            continue
                        fam_pids.append(r.PeopleId)
                        fam_list.append({
                            'PeopleId': r.PeopleId,
                            'Name2': safe_str(r.Name2),
                            'Position': r.PositionInFamilyId,
                            'Age': r.Age if r.Age else None,
                            'Email': safe_str(r.EmailAddress),
                            'Phone': safe_str(r.CellPhone),
                            'Gender': safe_str(r.Gender),
                            'MemberStatus': safe_str(r.MemberStatus),
                            'MemberStatusId': r.MemberStatusId,
                            'PhotoUrl': safe_str(r.PhotoUrl) if r.PhotoUrl else '',
                            'MediumPhotoUrl': safe_str(r.MediumPhotoUrl) if r.MediumPhotoUrl else '',
                        })

                    # Family member involvements + last attendance
                    fam_inv_data = {}
                    fam_last_attend = {}
                    if fam_pids:
                        fp_list = ','.join(str(p) for p in fam_pids)
                        for r in q.QuerySql("""
                            SELECT om.PeopleId, o.OrganizationName, mt.Description as MemberType
                            FROM OrganizationMembers om
                            JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                            WHERE om.PeopleId IN ({0}) AND o.OrganizationStatusId = 30
                            ORDER BY om.PeopleId, o.OrganizationName
                        """.format(fp_list)):
                            fpid = r.PeopleId
                            if fpid not in fam_inv_data:
                                fam_inv_data[fpid] = []
                            fam_inv_data[fpid].append({'name': safe_str(r.OrganizationName), 'type': safe_str(r.MemberType)})
                        try:
                            for r in q.QuerySql("""
                                SELECT a.PeopleId,
                                       CONVERT(VARCHAR(10), MAX(m.MeetingDate), 120) as LastAttend,
                                       DATEDIFF(day, MAX(m.MeetingDate), GETDATE()) as DaysSince
                                FROM Attend a JOIN Meetings m ON a.MeetingId = m.MeetingId
                                WHERE a.PeopleId IN ({0}) AND a.AttendanceFlag = 1
                                  AND m.MeetingDate >= DATEADD(year, -2, GETDATE())
                                GROUP BY a.PeopleId
                            """.format(fp_list)):
                                fam_last_attend[r.PeopleId] = {
                                    'date': safe_str(r.LastAttend),
                                    'days': int(r.DaysSince) if r.DaysSince else 999,
                                }
                        except:
                            pass

                    pos_map = {10: 'Adult', 20: 'Adult', 30: 'Child'}
                    detail_list = []
                    for m in fam_list:
                        invs = fam_inv_data.get(m['PeopleId'], [])
                        la = fam_last_attend.get(m['PeopleId'], {})
                        detail_list.append({
                            'pid': m['PeopleId'],
                            'name': safe_str(m['Name2']),
                            'age': m.get('Age'),
                            'gender': m.get('Gender', ''),
                            'posLabel': pos_map.get(m.get('Position', 0), 'Other'),
                            'status': safe_str(m.get('MemberStatus', '')),
                            'statusId': m.get('MemberStatusId', 0),
                            'phone': m.get('Phone', ''),
                            'email': m.get('Email', ''),
                            'photoUrl': m.get('PhotoUrl', ''),
                            'bigPhotoUrl': m.get('MediumPhotoUrl', '') or m.get('PhotoUrl', ''),
                            'invCount': len(invs),
                            'invNames': [i['name'] + ' (' + i['type'] + ')' for i in invs],
                            'lastAttend': la.get('date', '') if isinstance(la, dict) else '',
                            'daysSince': la.get('days', None) if isinstance(la, dict) else None,
                        })
                    detail['family'] = {'FamilyDetailList': detail_list}

                # Involvements with attendance
                inv_sql = """
                    SELECT om.PeopleId, o.OrganizationId, o.OrganizationName,
                           mt.Description as MemberType,
                           p.Name as ProgramName, d.Name as DivisionName,
                           CONVERT(VARCHAR(10), om.EnrollmentDate, 120) as EnrollDate
                    FROM OrganizationMembers om
                    JOIN Organizations o ON o.OrganizationId = om.OrganizationId
                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                    LEFT JOIN Division d ON o.DivisionId = d.Id
                    LEFT JOIN Program p ON d.ProgId = p.Id
                    WHERE om.PeopleId = {0} AND o.OrganizationStatusId = 30
                    ORDER BY p.Name, o.OrganizationName
                """.format(pid)
                inv_list = []
                org_ids = []
                for r in q.QuerySql(inv_sql):
                    org_ids.append(r.OrganizationId)
                    inv_list.append({
                        'orgId': r.OrganizationId,
                        'name': safe_str(r.OrganizationName),
                        'memberType': safe_str(r.MemberType),
                        'program': safe_str(r.ProgramName),
                        'division': safe_str(r.DivisionName),
                        'enrollDate': safe_str(r.EnrollDate),
                        'attended': 0, 'totalMeetings': 0, 'pct': 0, 'lastAttend': '',
                    })
                # Attendance stats
                if org_ids:
                    try:
                        attend_sql = """
                            SELECT a.OrganizationId,
                                   COUNT(CASE WHEN a.AttendanceFlag = 1 THEN 1 END) as Attended,
                                   COUNT(*) as TotalMeetings,
                                   CONVERT(VARCHAR(10), MAX(CASE WHEN a.AttendanceFlag = 1 THEN m.MeetingDate END), 120) as LastAttend
                            FROM Attend a JOIN Meetings m ON a.MeetingId = m.MeetingId
                            WHERE a.PeopleId = {0}
                              AND m.MeetingDate >= DATEADD(day, -90, GETDATE()) AND m.DidNotMeet = 0
                            GROUP BY a.OrganizationId
                        """.format(pid)
                        att_map = {}
                        for r in q.QuerySql(attend_sql):
                            attended = int(r.Attended) if r.Attended else 0
                            total = int(r.TotalMeetings) if r.TotalMeetings else 0
                            att_map[r.OrganizationId] = {
                                'attended': attended, 'total': total,
                                'pct': int(round(float(attended) / total * 100)) if total > 0 else 0,
                                'lastAttend': safe_str(r.LastAttend),
                            }
                        for inv in inv_list:
                            if inv['orgId'] in att_map:
                                inv.update(att_map[inv['orgId']])
                    except:
                        pass
                detail['involvements'] = inv_list

                # Past involvements from EnrollmentTransaction
                try:
                    # Get current org IDs to exclude them from past list
                    current_org_ids = set(org_ids)
                    past_sql = """
                        SELECT TOP 20 et.OrganizationId, o.OrganizationName,
                               mt.Description as MemberType,
                               p.Name as ProgramName, d.Name as DivisionName,
                               CONVERT(VARCHAR(10), et.EnrollmentDate, 120) as EnrollDate,
                               CONVERT(VARCHAR(10), et.InactiveDate, 120) as DroppedDate,
                               DATEDIFF(day, et.EnrollmentDate, et.InactiveDate) as DaysEnrolled
                        FROM EnrollmentTransaction et
                        JOIN Organizations o ON o.OrganizationId = et.OrganizationId
                        LEFT JOIN lookup.MemberType mt ON et.MemberTypeId = mt.Id
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        LEFT JOIN Program p ON d.ProgId = p.Id
                        WHERE et.PeopleId = {0}
                          AND et.InactiveDate IS NOT NULL
                          AND et.TransactionStatus = 0
                        ORDER BY et.InactiveDate DESC
                    """.format(pid)
                    past_list = []
                    seen_orgs = set()
                    for r in q.QuerySql(past_sql):
                        # Skip if they're currently in this org, or already listed
                        if r.OrganizationId in current_org_ids:
                            continue
                        if r.OrganizationId in seen_orgs:
                            continue
                        seen_orgs.add(r.OrganizationId)
                        past_list.append({
                            'orgId': r.OrganizationId,
                            'name': safe_str(r.OrganizationName),
                            'memberType': safe_str(r.MemberType),
                            'program': safe_str(r.ProgramName),
                            'division': safe_str(r.DivisionName),
                            'enrollDate': safe_str(r.EnrollDate),
                            'droppedDate': safe_str(r.DroppedDate),
                            'daysEnrolled': r.DaysEnrolled if r.DaysEnrolled else 0,
                        })
                    detail['pastInvolvements'] = past_list
                except:
                    detail['pastInvolvements'] = []

                # Extra values
                if ev_fields_str:
                    ev_fields = [f.strip() for f in ev_fields_str.split('|') if f.strip()]
                    if ev_fields:
                        safe_names = ["'" + n.replace("'", "''") + "'" for n in ev_fields]
                        ev_sql = """
                            SELECT Field, StrValue, Data,
                                   CAST(DateValue AS NVARCHAR(50)) as DateVal,
                                   IntValue, BitValue
                            FROM PeopleExtra
                            WHERE PeopleId = {0} AND Field IN ({1})
                        """.format(pid, ','.join(safe_names))
                        evs = {}
                        for r in q.QuerySql(ev_sql):
                            val = ''
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
                            evs[safe_str(r.Field)] = val
                        detail['extraValues'] = evs

                # Registration data
                if org_id > 0:
                    try:
                        for r in q.QuerySql("""
                            SELECT rq.Question, ra.Answer
                            FROM RegAnswer ra JOIN RegQuestion rq ON ra.QuestionId = rq.Id
                            WHERE ra.OrganizationId = {0} AND ra.PeopleId = {1}
                        """.format(org_id, pid)):
                            detail['regData'][safe_str(r.Question)] = safe_str(r.Answer)
                    except:
                        pass

                # Engagement stats (lightweight - single person subqueries)
                try:
                    eng_sql = """
                    SELECT
                        (SELECT TOP 1 CONVERT(varchar, a.MeetingDate, 101) + ' - ' + o.OrganizationName
                         FROM Attend a JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                         WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                         ORDER BY a.MeetingDate DESC) AS LastAttendance,
                        (SELECT COUNT(*) FROM Attend a
                         WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                         AND a.MeetingDate >= DATEADD(day, -90, GETDATE())) AS Attend90,
                        (SELECT DATEDIFF(day, MAX(a.MeetingDate), GETDATE())
                         FROM Attend a WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1) AS DaysSince,
                        (SELECT COUNT(DISTINCT om.OrganizationId)
                         FROM OrganizationMembers om
                         JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                         WHERE om.PeopleId = {0}
                         AND om.MemberTypeId IN (140, 310, 320, 710)
                         AND om.InactiveDate IS NULL
                         AND o.OrganizationStatusId = 30) AS ServingCount,
                        (SELECT COUNT(DISTINCT om.OrganizationId)
                         FROM OrganizationMembers om
                         JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                         WHERE om.PeopleId = {0}
                         AND o.OrganizationStatusId = 30
                         AND om.InactiveDate IS NULL) AS GroupCount
                    """.format(pid)
                    eng = q.QuerySqlTop1(eng_sql)
                    a90 = int(eng.Attend90) if eng and eng.Attend90 else 0
                    ds = int(eng.DaysSince) if eng and eng.DaysSince is not None else None
                    srv = int(eng.ServingCount) if eng and eng.ServingCount else 0
                    grp = int(eng.GroupCount) if eng and eng.GroupCount else 0

                    # Score each factor 0-100
                    def _score_recency(d):
                        if d is None: return 0
                        if d <= 7: return 100
                        if d <= 14: return 85
                        if d <= 30: return 70
                        if d <= 60: return 40
                        if d <= 90: return 20
                        return 5

                    def _score_frequency(n):
                        if n >= 10: return 100
                        if n >= 7: return 80
                        if n >= 4: return 60
                        if n >= 2: return 40
                        if n >= 1: return 20
                        return 0

                    def _score_groups(n):
                        if n >= 4: return 100
                        if n >= 3: return 80
                        if n >= 2: return 60
                        if n >= 1: return 40
                        return 0

                    def _score_serving(n):
                        if n >= 3: return 100
                        if n >= 2: return 80
                        if n >= 1: return 60
                        return 0

                    factors = {
                        'attend_recency': {'score': _score_recency(ds), 'weight': 30},
                        'attend_frequency': {'score': _score_frequency(a90), 'weight': 30},
                        'group_involvement': {'score': _score_groups(grp), 'weight': 20},
                        'serving': {'score': _score_serving(srv), 'weight': 20},
                    }

                    total_weight = sum(f['weight'] for f in factors.values())
                    weighted_sum = sum(f['score'] * f['weight'] for f in factors.values())
                    score = int(round(weighted_sum / total_weight)) if total_weight > 0 else 0
                    score = max(0, min(100, score))

                    if score >= 80: level = 'Highly engaged'
                    elif score >= 60: level = 'Engaged'
                    elif score >= 40: level = 'Moderately engaged'
                    elif score >= 20: level = 'Low engagement'
                    else: level = 'Not engaged'

                    if ds is not None and ds <= 30: status = 'Active'
                    elif ds is not None and ds > 90: status = 'Inactive'
                    else: status = 'Occasional'

                    detail['engagement'] = {
                        'score': score,
                        'level': level,
                        'status': status,
                        'attend90': a90,
                        'daysSince': ds,
                        'servingCount': srv,
                        'groupCount': grp,
                        'lastAttendance': safe_str(eng.LastAttendance) if eng and eng.LastAttendance else None,
                        'factors': factors,
                    }
                except:
                    detail['engagement'] = None

                # Journey timeline - comprehensive with dates and program
                try:
                    j_icon_map = {
                        'system': {'icon': 'fa-user-plus', 'color': '#6c757d'},
                        'joined': {'icon': 'fa-sign-in', 'color': '#ffc107'},
                        'attendance': {'icon': 'fa-calendar-check-o', 'color': '#007bff'},
                        'smallgroup': {'icon': 'fa-users', 'color': '#28a745'},
                        'serving': {'icon': 'fa-hands-helping', 'color': '#17a2b8'},
                        'leader': {'icon': 'fa-star', 'color': '#fd7e14'},
                        'giving': {'icon': 'fa-gift', 'color': '#e83e8c'}
                    }

                    # Current involvements with last attendance
                    curr_sql = """
                        SELECT
                            CONVERT(VARCHAR(10), om.EnrollmentDate, 120) AS EnrollmentDate,
                            o.OrganizationName,
                            mt.Description AS MemberType,
                            p.Name AS ProgramName,
                            om.MemberTypeId,
                            d.ProgId,
                            (SELECT TOP 1 CONVERT(VARCHAR(10), a.MeetingDate, 120)
                             FROM Attend a WHERE a.PeopleId = {0}
                             AND a.OrganizationId = om.OrganizationId AND a.AttendanceFlag = 1
                             ORDER BY a.MeetingDate DESC) AS LastAttended
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        LEFT JOIN Program p ON d.ProgId = p.Id
                        WHERE om.PeopleId = {0} AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate
                    """.format(pid)

                    # Past involvements with last attendance
                    past_sql = """
                        SELECT
                            CONVERT(VARCHAR(10), et.EnrollmentDate, 120) AS EnrollmentDate,
                            CONVERT(VARCHAR(10), et.InactiveDate, 120) AS InactiveDate,
                            o.OrganizationName,
                            mt.Description AS MemberType,
                            p.Name AS ProgramName,
                            et.MemberTypeId,
                            d.ProgId,
                            (SELECT TOP 1 CONVERT(VARCHAR(10), a.MeetingDate, 120)
                             FROM Attend a WHERE a.PeopleId = {0}
                             AND a.OrganizationId = et.OrganizationId AND a.AttendanceFlag = 1
                             ORDER BY a.MeetingDate DESC) AS LastAttended
                        FROM EnrollmentTransaction et
                        JOIN Organizations o ON et.OrganizationId = o.OrganizationId
                        LEFT JOIN lookup.MemberType mt ON et.MemberTypeId = mt.Id
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        LEFT JOIN Program p ON d.ProgId = p.Id
                        WHERE et.PeopleId = {0}
                          AND et.InactiveDate IS NOT NULL
                          AND et.TransactionStatus = 0
                          AND et.EnrollmentDate IS NOT NULL
                          AND et.OrganizationId NOT IN (
                              SELECT OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}
                          )
                        ORDER BY et.EnrollmentDate
                    """.format(pid)

                    # First attendance
                    first_att_sql = """
                        SELECT TOP 1 CONVERT(VARCHAR(10), a.MeetingDate, 120) AS EventDate, o.OrganizationName
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                        ORDER BY a.MeetingDate
                    """.format(pid)

                    # First serving via attend type
                    first_serve_sql = """
                        SELECT TOP 1 CONVERT(VARCHAR(10), a.MeetingDate, 120) AS EventDate, o.OrganizationName
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                        AND a.AttendanceTypeId IN (10, 20, 145)
                        ORDER BY a.MeetingDate
                    """.format(pid)

                    # System entry
                    sys_sql = "SELECT CONVERT(VARCHAR(10), CreatedDate, 120) AS EventDate FROM People WHERE PeopleId = {0}".format(pid)

                    j_events = []

                    # Added to system
                    sys_r = q.QuerySqlTop1(sys_sql)
                    if sys_r and sys_r.EventDate:
                        j_events.append({
                            'date': safe_str(sys_r.EventDate),
                            'event': 'Added to System',
                            'description': 'First contact/registration',
                            'type': 'system', 'icon': 'fa-user-plus', 'color': '#6c757d',
                            'program': '', 'enrolled': '', 'lastAttended': '', 'inactive': ''
                        })

                    # First attendance
                    fa_r = q.QuerySqlTop1(first_att_sql)
                    if fa_r and fa_r.EventDate:
                        j_events.append({
                            'date': safe_str(fa_r.EventDate),
                            'event': 'First Attendance',
                            'description': safe_str(fa_r.OrganizationName),
                            'type': 'attendance', 'icon': 'fa-calendar-check-o', 'color': '#007bff',
                            'program': '', 'enrolled': '', 'lastAttended': '', 'inactive': ''
                        })

                    # First serving
                    fs_r = q.QuerySqlTop1(first_serve_sql)
                    if fs_r and fs_r.EventDate:
                        j_events.append({
                            'date': safe_str(fs_r.EventDate),
                            'event': 'Started Serving',
                            'description': safe_str(fs_r.OrganizationName) + ' (Volunteer)',
                            'type': 'serving', 'icon': 'fa-hands-helping', 'color': '#17a2b8',
                            'program': '', 'enrolled': '', 'lastAttended': '', 'inactive': ''
                        })

                    def j_category(member_type_id, prog_id):
                        if prog_id == 1128:
                            return 'smallgroup'
                        if member_type_id in (140, 145, 310, 320) or (member_type_id and member_type_id > 300):
                            return 'leader'
                        if member_type_id == 710:
                            return 'serving'
                        return 'joined'

                    # Current involvements
                    for r in q.QuerySql(curr_sql):
                        cat = j_category(r.MemberTypeId, r.ProgId)
                        ic = j_icon_map.get(cat, {'icon': 'fa-circle', 'color': '#6c757d'})
                        edate = safe_str(r.EnrollmentDate or '')
                        j_events.append({
                            'date': edate,
                            'event': 'Joined',
                            'description': safe_str(r.OrganizationName) + ' (' + safe_str(r.MemberType or 'Member') + ')',
                            'type': cat, 'icon': ic['icon'], 'color': ic['color'],
                            'program': safe_str(r.ProgramName or ''),
                            'enrolled': edate,
                            'lastAttended': safe_str(r.LastAttended or ''),
                            'inactive': ''
                        })

                    # Past involvements
                    seen_past = set()
                    for r in q.QuerySql(past_sql):
                        org_key = safe_str(r.OrganizationName)
                        if org_key in seen_past:
                            continue
                        seen_past.add(org_key)
                        cat = j_category(r.MemberTypeId, r.ProgId)
                        ic = j_icon_map.get(cat, {'icon': 'fa-circle', 'color': '#6c757d'})
                        edate = safe_str(r.EnrollmentDate or '')
                        j_events.append({
                            'date': edate,
                            'event': 'Joined (Past)',
                            'description': safe_str(r.OrganizationName) + ' (' + safe_str(r.MemberType or 'Member') + ')',
                            'type': cat, 'icon': ic['icon'], 'color': ic['color'],
                            'program': safe_str(r.ProgramName or ''),
                            'enrolled': edate,
                            'lastAttended': safe_str(r.LastAttended or ''),
                            'inactive': safe_str(r.InactiveDate or '')
                        })

                    # Sort by date
                    j_events.sort(key=lambda x: x['date'] or '0000')

                    detail['journeyEvents'] = j_events
                except:
                    detail['journeyEvents'] = []

            response = {'success': True, 'detail': sanitize_for_json(detail)}

        # ==========================================================
        # PROCESSING ACTIONS
        # ==========================================================
        # DESTINATION HEALTH - Intelligence on target involvements
        # ==========================================================
        elif action == 'get_destination_health':
            org_ids_str = get_form_data('org_ids', '')
            no_contact_days = int(get_form_data('no_contact_days', '90'))
            health = {}
            # Always use description-based matching for prospect identification
            # This is database-agnostic regardless of what IDs your church uses
            prospect_names_set = set(['prospect', 'new guest', 'visitor', 'visiting member'])

            if org_ids_str:
                org_ids = [int(x) for x in org_ids_str.split(',') if x.strip()]
                for oid in org_ids:
                    try:
                        h = {'orgId': oid, 'name': '', 'leader': '', 'leaderEmail': '',
                             'orgCreated': '', 'orgAgeDays': 0, 'typeCounts': {},
                             'totalMembers': 0, 'prospectCount': 0, 'staleProspects': 0,
                             'graduated': 0, 'dropped': 0,
                             'meetings90d': 0, 'avgAttendance': 0, 'lastMeeting': ''}
                        # Basic org info + leader
                        org_sql = """
                            SELECT o.OrganizationName, o.LeaderId,
                                   leader.Name2 as LeaderName,
                                   leader.EmailAddress as LeaderEmail,
                                   CONVERT(VARCHAR(10), o.CreatedDate, 120) as OrgCreated,
                                   DATEDIFF(day, o.CreatedDate, GETDATE()) as OrgAgeDays
                            FROM Organizations o
                            LEFT JOIN People leader ON o.LeaderId = leader.PeopleId
                            WHERE o.OrganizationId = {0}
                        """.format(oid)
                        for r in q.QuerySql(org_sql):
                            h['name'] = safe_str(r.OrganizationName)
                            h['leader'] = safe_str(r.LeaderName)
                            h['leaderEmail'] = safe_str(r.LeaderEmail)
                            h['orgCreated'] = safe_str(r.OrgCreated)
                            h['orgAgeDays'] = int(r.OrgAgeDays) if r.OrgAgeDays else 0

                        # Member type breakdown
                        mt_sql = """
                            SELECT mt.Description as MemberType, om.MemberTypeId, COUNT(*) as Cnt
                            FROM OrganizationMembers om
                            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                            WHERE om.OrganizationId = {0}
                            GROUP BY mt.Description, om.MemberTypeId
                        """.format(oid)
                        type_counts = {}
                        type_ids_by_name = {}  # lowercase desc -> list of IDs
                        total_members = 0
                        prospect_count = 0
                        for r in q.QuerySql(mt_sql):
                            cnt = int(r.Cnt)
                            desc = safe_str(r.MemberType)
                            type_counts[desc] = cnt
                            total_members += cnt
                            dl = desc.lower()
                            if dl not in type_ids_by_name:
                                type_ids_by_name[dl] = []
                            type_ids_by_name[dl].append(r.MemberTypeId)
                            if dl in prospect_names_set:
                                prospect_count += cnt
                        h['typeCounts'] = type_counts
                        h['totalMembers'] = total_members
                        h['prospectCount'] = prospect_count
                        # Collect the actual prospect MemberTypeIds for stale query
                        prospect_ids_for_org = []
                        for pn in prospect_names_set:
                            prospect_ids_for_org.extend(type_ids_by_name.get(pn, []))

                        # Stale prospects (prospect member type, no attendance in X days)
                        try:
                            if prospect_ids_for_org:
                                stale_mt_filter = "AND om.MemberTypeId IN ({0})".format(','.join(str(m) for m in prospect_ids_for_org))
                            else:
                                stale_mt_filter = "AND 1=0"  # No prospect types found, skip
                            stale_sql = """
                            SELECT COUNT(DISTINCT om.PeopleId) as StaleCount
                            FROM OrganizationMembers om
                            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                            LEFT JOIN (
                                SELECT a.PeopleId, MAX(m.MeetingDate) as LastAttend
                                FROM Attend a
                                JOIN Meetings m ON a.MeetingId = m.MeetingId
                                WHERE a.OrganizationId = {0}
                                  AND a.AttendanceFlag = 1
                                GROUP BY a.PeopleId
                            ) att ON att.PeopleId = om.PeopleId
                            WHERE om.OrganizationId = {0}
                              {2}
                              AND om.EnrollmentDate < DATEADD(day, -{1}, GETDATE())
                              AND (att.LastAttend IS NULL OR att.LastAttend < DATEADD(day, -{1}, GETDATE()))
                        """.format(oid, no_contact_days, stale_mt_filter)
                            h['staleProspects'] = 0
                            for r in q.QuerySql(stale_sql):
                                h['staleProspects'] = int(r.StaleCount) if r.StaleCount else 0
                        except:
                            pass

                        # Graduated: currently non-prospect but had attend records as prospect type
                        try:
                            grad_sql = """
                                SELECT COUNT(DISTINCT om.PeopleId) as Graduated
                                FROM OrganizationMembers om
                                LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                                WHERE om.OrganizationId = {0}
                                  AND LOWER(mt.Description) NOT IN ('prospect','new guest','visitor','visiting member')
                                  AND om.InactiveDate IS NULL
                                  AND om.PeopleId IN (
                                      SELECT DISTINCT a.PeopleId FROM Attend a
                                      LEFT JOIN lookup.MemberType amt ON a.MemberTypeId = amt.Id
                                      WHERE a.OrganizationId = {0}
                                        AND LOWER(amt.Description) IN ('prospect','new guest','visitor','visiting member')
                                  )
                            """.format(oid)
                            for r in q.QuerySql(grad_sql):
                                h['graduated'] = int(r.Graduated) if r.Graduated else 0
                        except:
                            pass

                        # Dropped: check both OrganizationMembers.InactiveDate AND EnrollmentTransaction
                        try:
                            drop_sql = """
                                SELECT COUNT(DISTINCT PeopleId) as Dropped FROM (
                                    -- Currently inactive in OrganizationMembers
                                    SELECT om.PeopleId
                                    FROM OrganizationMembers om
                                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                                    WHERE om.OrganizationId = {0}
                                      AND om.InactiveDate IS NOT NULL
                                      AND (LOWER(mt.Description) IN ('prospect','new guest','visitor','visiting member')
                                           OR om.PeopleId IN (
                                               SELECT DISTINCT a.PeopleId FROM Attend a
                                               LEFT JOIN lookup.MemberType amt ON a.MemberTypeId = amt.Id
                                               WHERE a.OrganizationId = {0}
                                                 AND LOWER(amt.Description) IN ('prospect','new guest','visitor','visiting member')
                                           ))
                                    UNION
                                    -- Dropped via EnrollmentTransaction (removed from org)
                                    SELECT et.PeopleId
                                    FROM EnrollmentTransaction et
                                    LEFT JOIN lookup.MemberType mt ON et.MemberTypeId = mt.Id
                                    WHERE et.OrganizationId = {0}
                                      AND et.InactiveDate IS NOT NULL
                                      AND LOWER(mt.Description) IN ('prospect','new guest','visitor','visiting member')
                                      AND et.PeopleId NOT IN (
                                          SELECT om2.PeopleId FROM OrganizationMembers om2
                                          WHERE om2.OrganizationId = {0}
                                      )
                                ) dropped
                            """.format(oid)
                            for r in q.QuerySql(drop_sql):
                                h['dropped'] = int(r.Dropped) if r.Dropped else 0
                        except:
                            pass

                        # Recent meetings (last 90 days)
                        try:
                            meeting_sql = """
                                SELECT COUNT(*) as MeetingCount,
                                       AVG(CAST(NumPresent as FLOAT)) as AvgAttendance,
                                       CONVERT(VARCHAR(10), MAX(MeetingDate), 120) as LastMeeting,
                                       SUM(CASE WHEN HeadCount IS NOT NULL AND HeadCount > 0
                                                AND HeadCount > NumPresent THEN 1 ELSE 0 END) as HeadCountMeetings,
                                       SUM(CASE WHEN HeadCount IS NOT NULL AND HeadCount > 0
                                                AND HeadCount > NumPresent THEN HeadCount ELSE 0 END) as TotalHeadCount
                                FROM Meetings
                                WHERE OrganizationId = {0}
                                  AND MeetingDate >= DATEADD(day, -90, GETDATE())
                                  AND DidNotMeet = 0
                            """.format(oid)
                            for r in q.QuerySql(meeting_sql):
                                h['meetings90d'] = int(r.MeetingCount) if r.MeetingCount else 0
                                h['avgAttendance'] = round(float(r.AvgAttendance), 1) if r.AvgAttendance else 0
                                h['lastMeeting'] = safe_str(r.LastMeeting)
                                hc_meetings = int(r.HeadCountMeetings) if r.HeadCountMeetings else 0
                                if hc_meetings > 0:
                                    h['headCountWarning'] = True
                                    h['headCountMeetings'] = hc_meetings
                        except:
                            pass

                    except Exception as e:
                        h['error'] = safe_str(e)
                    health[str(oid)] = h

            response = {'success': True, 'health': sanitize_for_json(health)}

        elif action == 'get_destination_people':
            org_id = int(get_form_data('org_id', '0'))
            drill_type = get_form_data('drill_type', '')
            prospect_mt_str = get_form_data('prospect_types', '')
            no_contact_days = int(get_form_data('no_contact_days', '90'))
            people = []

            if org_id > 0:
                # Build prospect filter: by ID if configured, else by description name
                # Always use description-based matching for prospect drill-down
                mt_where = "AND LOWER(mt.Description) IN ('prospect','new guest','visitor','visiting member')"

                # Base SELECT with enrollment days and attendance
                base_select = """
                    SELECT TOP 200 p.PeopleId, p.Name2, p.EmailAddress, p.CellPhone, p.Age,
                           mt.Description as MemberType, pic.ThumbUrl as PhotoUrl,
                           CONVERT(VARCHAR(10), om.EnrollmentDate, 120) as EnrollDate,
                           DATEDIFF(day, om.EnrollmentDate, GETDATE()) as DaysEnrolled,
                           att.AttendCount, att.LastAttend,
                           DATEDIFF(day, att.LastAttend, GETDATE()) as DaysSinceAttend
                    FROM OrganizationMembers om
                    JOIN People p ON om.PeopleId = p.PeopleId
                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                    LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                    LEFT JOIN (
                        SELECT a.PeopleId,
                               COUNT(CASE WHEN a.AttendanceFlag = 1 THEN 1 END) as AttendCount,
                               CONVERT(VARCHAR(10), MAX(CASE WHEN a.AttendanceFlag = 1 THEN m.MeetingDate END), 120) as LastAttend
                        FROM Attend a JOIN Meetings m ON a.MeetingId = m.MeetingId
                        WHERE a.OrganizationId = {0}
                        GROUP BY a.PeopleId
                    ) att ON att.PeopleId = om.PeopleId
                """.format(org_id)

                sql = None
                if drill_type == 'total':
                    sql = base_select + " WHERE om.OrganizationId = {0} AND p.IsDeceased = 0 ORDER BY p.Name2".format(org_id)
                elif drill_type == 'prospects':
                    sql = base_select + " WHERE om.OrganizationId = {0} {1} AND p.IsDeceased = 0 ORDER BY om.EnrollmentDate".format(org_id, mt_where)
                elif drill_type == 'stale':
                    sql = base_select + " WHERE om.OrganizationId = {0} {1} AND p.IsDeceased = 0 AND om.EnrollmentDate < DATEADD(day, -{2}, GETDATE()) AND (att.LastAttend IS NULL OR att.LastAttend < CONVERT(VARCHAR(10), DATEADD(day, -{2}, GETDATE()), 120)) ORDER BY att.LastAttend".format(org_id, mt_where, no_contact_days)
                elif drill_type == 'graduated':
                    # Currently non-prospect but attended as prospect (proved by Attend.MemberTypeId)
                    sql = base_select + """ WHERE om.OrganizationId = {0}
                        AND LOWER(mt.Description) NOT IN ('prospect','new guest','visitor','visiting member')
                        AND om.InactiveDate IS NULL AND p.IsDeceased = 0
                        AND om.PeopleId IN (
                            SELECT DISTINCT a2.PeopleId FROM Attend a2
                            LEFT JOIN lookup.MemberType mt2 ON a2.MemberTypeId = mt2.Id
                            WHERE a2.OrganizationId = {0}
                              AND LOWER(mt2.Description) IN ('prospect','new guest','visitor','visiting member')
                        )
                        ORDER BY p.Name2""".format(org_id)
                elif drill_type == 'dropped':
                    # Dropped prospects from EnrollmentTransaction
                    sql = """
                        SELECT TOP 200 p.PeopleId, p.Name2, p.EmailAddress, p.CellPhone, p.Age,
                               mt.Description as MemberType, pic.ThumbUrl as PhotoUrl,
                               CONVERT(VARCHAR(10), et.EnrollmentDate, 120) as EnrollDate,
                               DATEDIFF(day, et.EnrollmentDate, et.InactiveDate) as DaysEnrolled,
                               CONVERT(VARCHAR(10), et.InactiveDate, 120) as DroppedDate,
                               (SELECT COUNT(*) FROM Attend a WHERE a.PeopleId = et.PeopleId
                                AND a.OrganizationId = {0} AND a.AttendanceFlag = 1) as AttendCount,
                               (SELECT CONVERT(VARCHAR(10), MAX(m.MeetingDate), 120) FROM Attend a
                                JOIN Meetings m ON a.MeetingId = m.MeetingId
                                WHERE a.PeopleId = et.PeopleId AND a.OrganizationId = {0}
                                AND a.AttendanceFlag = 1) as LastAttend,
                               NULL as DaysSinceAttend
                        FROM EnrollmentTransaction et
                        JOIN People p ON et.PeopleId = p.PeopleId
                        LEFT JOIN lookup.MemberType mt ON et.MemberTypeId = mt.Id
                        LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
                        WHERE et.OrganizationId = {0}
                          AND et.InactiveDate IS NOT NULL
                          AND LOWER(mt.Description) IN ('prospect','new guest','visitor','visiting member')
                          AND et.PeopleId NOT IN (SELECT om2.PeopleId FROM OrganizationMembers om2 WHERE om2.OrganizationId = {0})
                          AND p.IsDeceased = 0
                        ORDER BY et.InactiveDate DESC
                    """.format(org_id)

                if sql:
                    try:
                        for r in q.QuerySql(sql):
                            person = {
                                'pid': r.PeopleId,
                                'name': safe_str(r.Name2),
                                'email': safe_str(r.EmailAddress),
                                'phone': safe_str(r.CellPhone),
                                'age': r.Age if r.Age else None,
                                'memberType': safe_str(r.MemberType),
                                'photoUrl': safe_str(r.PhotoUrl) if r.PhotoUrl else '',
                                'enrollDate': safe_str(r.EnrollDate),
                                'daysEnrolled': int(r.DaysEnrolled) if r.DaysEnrolled else None,
                                'attendCount': int(r.AttendCount) if r.AttendCount else 0,
                                'lastAttend': safe_str(r.LastAttend),
                                'daysSince': int(r.DaysSinceAttend) if r.DaysSinceAttend else None,
                            }
                            if hasattr(r, 'DroppedDate'):
                                person['droppedDate'] = safe_str(r.DroppedDate)
                            people.append(person)
                    except:
                        pass

            response = {'success': True, 'people': sanitize_for_json(people), 'drillType': drill_type}

        # ==========================================================
        elif action == 'process_action':
            target_json = get_form_data('target_data', '{}')
            target = json.loads(target_json)
            pids_str = get_form_data('people_ids', '')
            pids = [int(x) for x in pids_str.split(',') if x.strip()]
            ttype = target.get('pb_type', '')
            processed = 0
            errors = []

            for pid in pids:
                try:
                    person = model.GetPerson(pid)
                    if not person:
                        errors.append('Person {0} not found'.format(pid))
                        continue

                    if ttype == 'tag':
                        tag_name = target.get('tagName', '')
                        if tag_name:
                            model.AddTag(pid, tag_name)
                            processed += 1

                    elif ttype == 'involvement':
                        target_org = int(target.get('orgId', 0))
                        if target_org > 0:
                            if not model.InOrg(pid, target_org):
                                model.JoinOrg(target_org, person)
                            # MEMBER TYPE RESOLUTION (label-first, v1.3.3):
                            # `model.SetMemberType` takes the member-type
                            # *Description string* (e.g. "Prospect"), not an
                            # int id. Passing the int makes FetchOrCreateMemberType
                            # match on Description="230" -- no row exists, so a
                            # new junk MemberType is created with AttendanceTypeId=Member.
                            # Verified in bvcms-develop CmsData/API/PythonModel/
                            # PythonModel.Organizations.cs:235 and
                            # CmsData/Organization/Organization.cs:575.
                            # We resolve via the action LABEL first because
                            # earlier PB versions hardcoded memberTypeId=230
                            # for "Prospect" -- correct on default-seed
                            # installs, but on churches that renamed
                            # lookup.MemberType (e.g. FBCH where Prospect=311
                            # and 230='InActive') the cached id is WRONG.
                            # The label is what staff saw and is the source
                            # of truth.
                            mt_id = None
                            _label = target.get('label', '') or ''
                            if ' as ' in _label:
                                _trailing = _label.rsplit(' as ', 1)[-1].strip()
                                if _trailing:
                                    try:
                                        _rows2 = q.QuerySql(
                                            "SELECT TOP 1 Id FROM lookup.MemberType WITH (NOLOCK) "
                                            "WHERE Description = '"
                                            + _trailing.replace("'", "''") + "'"
                                        )
                                        for _r2 in _rows2:
                                            mt_id = _r2.Id
                                            break
                                    except:
                                        pass
                            # Fall back to cached memberTypeId if label
                            # is unparseable (tag actions, custom labels).
                            if not mt_id:
                                mt_id = target.get('memberTypeId')
                            if mt_id:
                                mt_desc = ''
                                try:
                                    _rows = q.QuerySql(
                                        "SELECT TOP 1 Description "
                                        "FROM lookup.MemberType WITH (NOLOCK) "
                                        "WHERE Id = " + str(int(mt_id))
                                    )
                                    for _r in _rows:
                                        mt_desc = safe_str(_r.Description)
                                        break
                                except:
                                    mt_desc = ''
                                if mt_desc:
                                    model.SetMemberType(pid, target_org, mt_desc)
                                else:
                                    errors.append(
                                        'Skipped SetMemberType for {0}: '
                                        'unknown MemberTypeId {1}'.format(pid, mt_id)
                                    )
                            else:
                                errors.append(
                                    'No member type set for action -- {0} landed as default. '
                                    'Edit the target action and re-pick the role.'.format(pid)
                                )
                            # Also add to any selected subgroups
                            subgroups = target.get('subgroups', [])
                            for sg_name in subgroups:
                                if sg_name:
                                    model.AddSubGroup(pid, target_org, sg_name)
                            processed += 1

                except Exception as e:
                    errors.append('Error processing {0}: {1}'.format(pid, safe_str(e)))

            response = {'success': True, 'processed': processed, 'errors': errors}

        elif action == 'log_contact':
            pid = int(get_form_data('people_id', '0'))
            note_text = get_form_data('note_text', '')
            keyword = get_form_data('keyword', '')
            user_id = model.UserPeopleId

            if pid > 0 and note_text:
                full_note = note_text
                if keyword:
                    full_note = '[' + keyword + '] ' + note_text
                model.AddTaskNote(0, full_note)
                # Also update LastProspectContact EV
                model.AddExtraValueDate(pid, 'LastProspectContact', datetime.datetime.now())
                response = {'success': True, 'message': 'Contact logged'}
            else:
                response = {'success': False, 'message': 'Missing people_id or note_text'}

        # ==========================================================
        # SESSION MANAGEMENT
        # ==========================================================
        elif action == 'log_activity':
            log_json = get_form_data('log_data', '{}')
            log_entry = json.loads(log_json)
            log_entry['timestamp'] = now_str()
            log_entry['userId'] = model.UserPeopleId
            if not log_entry.get('source'):
                log_entry['source'] = 'workspace'
            try:
                user = model.GetPerson(model.UserPeopleId)
                log_entry['userName'] = safe_str(user.Name2) if user else 'Unknown'
            except:
                log_entry['userName'] = 'Unknown'

            activity_log = load_content("ProspectBuilder_ActivityLog", [])
            activity_log.insert(0, log_entry)
            # Keep last 500 entries
            if len(activity_log) > 500:
                activity_log = activity_log[:500]
            save_content("ProspectBuilder_ActivityLog", activity_log)
            response = {'success': True}

        elif action == 'load_activity_log':
            config_filter = get_form_data('config_id', '')
            source_filter = get_form_data('source_filter', '')
            group_filter = get_form_data('group_id', '')
            activity_log = load_content("ProspectBuilder_ActivityLog", [])
            if config_filter:
                activity_log = [e for e in activity_log if e.get('configId') == config_filter]
            if source_filter:
                activity_log = [e for e in activity_log if e.get('source', 'workspace') == source_filter]
            if group_filter:
                activity_log = [e for e in activity_log if e.get('groupId') == group_filter]
            response = {'success': True, 'log': sanitize_for_json(activity_log[:200])}

        elif action == 'clear_activity_log':
            save_content("ProspectBuilder_ActivityLog", [])
            response = {'success': True, 'message': 'Activity log cleared'}

        elif action == 'save_work_state':
            config_id = get_form_data('config_id', '')
            state_json = get_form_data('state_data', '{}')
            if config_id:
                state = json.loads(state_json)
                work_states = load_content("ProspectBuilder_WorkStates", {})
                work_states[config_id] = state
                save_content("ProspectBuilder_WorkStates", work_states)
                response = {'success': True}
            else:
                response = {'success': False, 'message': 'No config_id'}

        elif action == 'load_work_state':
            config_id = get_form_data('config_id', '')
            work_states = load_content("ProspectBuilder_WorkStates", {})
            state = work_states.get(config_id, {})
            response = {'success': True, 'state': state}

        elif action == 'save_session':
            session_json = get_form_data('session_data', '{}')
            session = json.loads(session_json)
            sessions = load_sessions()

            session_id = session.get('id', '')
            if not session_id:
                session['id'] = make_id('sess')
                session['createdAt'] = now_str()
                sessions.append(session)
            else:
                for i, s in enumerate(sessions):
                    if s.get('id') == session_id:
                        session['createdAt'] = s.get('createdAt', now_str())
                        sessions[i] = session
                        break
                else:
                    session['createdAt'] = now_str()
                    sessions.append(session)

            session['updatedAt'] = now_str()
            # Keep only last 20 sessions
            if len(sessions) > 20:
                sessions = sessions[-20:]
            save_sessions(sessions)
            response = {'success': True, 'session': sanitize_for_json(session), 'message': 'Session saved'}

        elif action == 'list_sessions':
            sessions = load_sessions()
            response = {'success': True, 'sessions': sanitize_for_json(sessions)}

        elif action == 'load_session':
            session_id = get_form_data('session_id', '')
            sessions = load_sessions()
            session = None
            for s in sessions:
                if s.get('id') == session_id:
                    session = s
                    break
            if session:
                response = {'success': True, 'session': sanitize_for_json(session)}
            else:
                response = {'success': False, 'message': 'Session not found'}

        elif action == 'delete_session':
            session_id = get_form_data('session_id', '')
            sessions = load_sessions()
            sessions = [s for s in sessions if s.get('id') != session_id]
            save_sessions(sessions)
            response = {'success': True, 'message': 'Session deleted'}

        # ==========================================================
        # SEARCH HELPERS
        # ==========================================================
        elif action == 'search_involvements':
            search_term = get_form_data('search_term', '')
            program_id = get_form_data('program_id', '')
            division_id = get_form_data('division_id', '')

            where_clauses = ["o.OrganizationStatusId = 30"]
            if search_term:
                safe_term = search_term.replace("'", "''")
                where_clauses.append("(o.OrganizationName LIKE '%{0}%' OR CAST(o.OrganizationId AS VARCHAR) = '{0}')".format(safe_term))
            if program_id:
                where_clauses.append("d.ProgId = {0}".format(int(program_id)))
            if division_id:
                where_clauses.append("o.DivisionId = {0}".format(int(division_id)))

            sql = """
                SELECT TOP 50 o.OrganizationId, o.OrganizationName,
                    p.Name as ProgramName, d.Name as DivisionName,
                    (SELECT COUNT(*) FROM OrganizationMembers WHERE OrganizationId = o.OrganizationId) as MemberCount
                FROM Organizations o
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE {0}
                ORDER BY o.OrganizationName
            """.format(" AND ".join(where_clauses))
            results = []
            for r in q.QuerySql(sql):
                results.append({
                    'orgId': r.OrganizationId,
                    'name': safe_str(r.OrganizationName),
                    'program': safe_str(r.ProgramName),
                    'division': safe_str(r.DivisionName),
                    'memberCount': r.MemberCount,
                })
            response = {'success': True, 'involvements': sanitize_for_json(results)}

        elif action == 'get_programs':
            sql = """
                SELECT DISTINCT p.Id, p.Name
                FROM Program p
                JOIN Division d ON d.ProgId = p.Id
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE o.OrganizationStatusId = 30
                ORDER BY p.Name
            """
            programs = []
            for r in q.QuerySql(sql):
                programs.append({'id': r.Id, 'name': safe_str(r.Name)})
            response = {'success': True, 'programs': programs}

        elif action == 'get_divisions':
            prog_id = int(get_form_data('program_id', '0'))
            sql = """
                SELECT DISTINCT d.Id, d.Name
                FROM Division d
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE d.ProgId = {0}
                  AND o.OrganizationStatusId = 30
                ORDER BY d.Name
            """.format(prog_id)
            divisions = []
            for r in q.QuerySql(sql):
                divisions.append({'id': r.Id, 'name': safe_str(r.Name)})
            response = {'success': True, 'divisions': divisions}

        elif action == 'get_subgroups':
            org_id = int(get_form_data('org_id', '0'))
            subgroups = []
            if org_id > 0:
                sg_sql = """
                    SELECT DISTINCT sg.Name,
                        (SELECT COUNT(*) FROM SubGroupMembers sgm
                         WHERE sgm.SubGroupId = sg.Id) as MemberCount
                    FROM SubGroup sg
                    WHERE sg.OrganizationId = {0}
                    ORDER BY sg.Name
                """.format(org_id)
                try:
                    for r in q.QuerySql(sg_sql):
                        subgroups.append({'name': safe_str(r.Name), 'count': r.MemberCount or 0})
                except:
                    pass
            response = {'success': True, 'subgroups': subgroups}

        elif action == 'get_member_types':
            sql = """
                SELECT Id, Description FROM lookup.MemberType ORDER BY Description
            """
            mt_list = []
            for r in q.QuerySql(sql):
                mt_list.append({'id': int(r.Id), 'description': safe_str(r.Description)})
            response = {'success': True, 'memberTypes': mt_list}

        elif action == 'search_tags':
            search_term = get_form_data('search_term', '')
            safe_term = search_term.replace("'", "''")
            sql = """
                SELECT TOP 20 t.Name,
                    (SELECT COUNT(*) FROM TagPerson tp WHERE tp.Id = t.Id) as PersonCount
                FROM Tag t
                WHERE t.Name LIKE '%{0}%'
                  AND t.TypeId = 1
                ORDER BY t.Name
            """.format(safe_term)
            tags = []
            for r in q.QuerySql(sql):
                tags.append({'name': safe_str(r.Name), 'count': r.PersonCount})
            response = {'success': True, 'tags': tags}

        elif action == 'search_extra_values':
            search_term = get_form_data('search_term', '')
            safe_term = search_term.replace("'", "''")
            sql = """
                SELECT TOP 30 DISTINCT Field
                FROM PeopleExtra
                WHERE Field LIKE '%{0}%'
                ORDER BY Field
            """.format(safe_term)
            fields = []
            for r in q.QuerySql(sql):
                fields.append(safe_str(r.Field))
            response = {'success': True, 'fields': fields}

        # ============================================================
        # JOURNEY: Individual Engagement Timeline
        # ============================================================
        elif action == 'get_person_journey':
            people_id = int(get_form_data('peopleId', '0'))
            if not people_id:
                response = {'success': False, 'error': 'No person ID provided'}
            else:
                person_sql = """
                SELECT
                    p.PeopleId,
                    p.Name2 AS Name,
                    p.Age,
                    p.EmailAddress,
                    p.CellPhone,
                    ms.Description AS MemberStatus
                FROM People p
                LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                WHERE p.PeopleId = {0}
                """.format(people_id)

                person_info = q.QuerySqlTop1(person_sql)
                if not person_info:
                    response = {'success': False, 'error': 'Person not found'}
                else:
                    can_view_giving = False
                    for role in ["Finance", "FinanceAdmin", "Admin"]:
                        if model.UserIsInRole(role):
                            can_view_giving = True
                            break
                    if can_view_giving:
                        try:
                            q.QuerySqlTop1("SELECT TOP 1 ContributionId FROM Contribution")
                        except:
                            can_view_giving = False

                    giving_section = ""
                    if can_view_giving and SHOW_GIVING_IN_JOURNEY:
                        giving_section = """
                        UNION ALL
                        SELECT TOP 1 c.ContributionDate AS EventDate,
                            'First Contribution' AS EventType,
                            'Started giving' AS Description,
                            'giving' AS Category, 6 AS SortOrder
                        FROM Contribution c
                        WHERE c.PeopleId = {0}
                        AND c.ContributionTypeId != 99 AND c.ContributionAmount > 0
                        ORDER BY c.ContributionDate""".format(people_id)

                    events_sql = """
                    WITH JourneyEvents AS (
                        SELECT p.CreatedDate AS EventDate,
                            'Added to System' AS EventType,
                            'First contact/registration' AS Description,
                            'system' AS Category, 1 AS SortOrder
                        FROM People p WHERE p.PeopleId = {0}
                        UNION ALL
                        SELECT TOP 1 om.CreatedDate AS EventDate,
                            'First Program Joined' AS EventType,
                            o.OrganizationName AS Description,
                            'program' AS Category, 2 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        WHERE om.PeopleId = {0} AND o.OrganizationStatusId = 30
                        ORDER BY om.CreatedDate
                        UNION ALL
                        SELECT TOP 1 a.MeetingDate AS EventDate,
                            'First Attendance' AS EventType,
                            o.OrganizationName AS Description,
                            'attendance' AS Category, 3 AS SortOrder
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                        ORDER BY a.MeetingDate
                        UNION ALL
                        SELECT TOP 1 om.CreatedDate AS EventDate,
                            'Small Group Joined' AS EventType,
                            o.OrganizationName AS Description,
                            'smallgroup' AS Category, 4 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        JOIN Division d ON o.DivisionId = d.Id
                        WHERE om.PeopleId = {0} AND o.OrganizationStatusId = 30
                        AND d.ProgId IN (1128)
                        ORDER BY om.CreatedDate
                        UNION ALL
                        SELECT TOP 1 EventDate, EventType, Description, Category, SortOrder
                        FROM (
                            SELECT MIN(a.MeetingDate) AS EventDate,
                                'Started Serving' AS EventType,
                                o.OrganizationName + ' (Volunteer)' AS Description,
                                'serving' AS Category, 5 AS SortOrder
                            FROM Attend a
                            JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                            WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                            AND a.AttendanceTypeId IN (10, 20)
                            GROUP BY o.OrganizationName
                            UNION ALL
                            SELECT om.CreatedDate AS EventDate,
                                'Started Serving' AS EventType,
                                o.OrganizationName + ' (' + ISNULL(mt.Description, 'Leadership') + ')' AS Description,
                                'serving' AS Category, 5 AS SortOrder
                            FROM OrganizationMembers om
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                            WHERE om.PeopleId = {0} AND om.MemberTypeId > 100
                        ) AS ServingEvents
                        ORDER BY EventDate
                        {1}
                    )
                    SELECT * FROM JourneyEvents
                    WHERE EventDate IS NOT NULL
                    ORDER BY EventDate, SortOrder
                    """.format(people_id, giving_section)

                    events = q.QuerySql(events_sql)

                    recent_sql = """
                    SELECT
                        (SELECT TOP 1 CONVERT(varchar, a.MeetingDate, 101) + ' - ' + o.OrganizationName
                         FROM Attend a JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                         WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                         ORDER BY a.MeetingDate DESC) AS LastAttendance,
                        (SELECT TOP 1 CONVERT(varchar, om.EnrollmentDate, 101) + ' - ' + o.OrganizationName
                         FROM OrganizationMembers om JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                         WHERE om.PeopleId = {0}
                         ORDER BY om.EnrollmentDate DESC) AS LastSignup,
                        (SELECT COUNT(*) FROM Attend a
                         WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                         AND a.MeetingDate >= DATEADD(day, -90, GETDATE())) AS AttendanceCount90Days,
                        (SELECT DATEDIFF(day, MAX(a.MeetingDate), GETDATE())
                         FROM Attend a WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1) AS DaysSinceLastAttendance,
                        (SELECT COUNT(DISTINCT om.OrganizationId)
                         FROM OrganizationMembers om WHERE om.PeopleId = {0}
                         AND om.MemberTypeId > 100 AND om.InactiveDate IS NULL) AS CurrentlyServingCount
                    """.format(people_id)

                    recent_activity = q.QuerySqlTop1(recent_sql)

                    icon_map = {
                        'system': {'icon': 'fa-user-plus', 'color': '#6c757d'},
                        'program': {'icon': 'fa-sitemap', 'color': '#ffc107'},
                        'attendance': {'icon': 'fa-calendar-check-o', 'color': '#007bff'},
                        'smallgroup': {'icon': 'fa-users', 'color': '#28a745'},
                        'serving': {'icon': 'fa-hands-helping', 'color': '#17a2b8'},
                        'giving': {'icon': 'fa-gift', 'color': '#e83e8c'}
                    }

                    events_list = []
                    for event in events:
                        if event.Category == 'giving' and (not SHOW_GIVING_IN_JOURNEY or not can_view_giving):
                            continue
                        icon_info = icon_map.get(event.Category, {'icon': 'fa-circle', 'color': '#6c757d'})
                        date_str = str(event.EventDate).split(' ')[0] if event.EventDate else ''
                        events_list.append({
                            'date': date_str,
                            'event': safe_str(event.EventType),
                            'description': safe_str(event.Description),
                            'type': safe_str(event.Category),
                            'icon': icon_info['icon'],
                            'color': icon_info['color']
                        })

                    # Factor-based engagement scoring (consistent with detail view)
                    j_a90 = int(recent_activity.AttendanceCount90Days) if recent_activity and recent_activity.AttendanceCount90Days else 0
                    j_ds = int(recent_activity.DaysSinceLastAttendance) if recent_activity and recent_activity.DaysSinceLastAttendance is not None else None
                    j_srv = int(recent_activity.CurrentlyServingCount) if recent_activity and recent_activity.CurrentlyServingCount else 0
                    j_grp = len([e for e in events_list if e['type'] in ('program', 'smallgroup', 'serving')])

                    def _j_score_recency(d):
                        if d is None: return 0
                        if d <= 7: return 100
                        if d <= 14: return 85
                        if d <= 30: return 70
                        if d <= 60: return 40
                        if d <= 90: return 20
                        return 5
                    def _j_score_frequency(n):
                        if n >= 10: return 100
                        if n >= 7: return 80
                        if n >= 4: return 60
                        if n >= 2: return 40
                        if n >= 1: return 20
                        return 0
                    def _j_score_groups(n):
                        if n >= 4: return 100
                        if n >= 3: return 80
                        if n >= 2: return 60
                        if n >= 1: return 40
                        return 0
                    def _j_score_serving(n):
                        if n >= 3: return 100
                        if n >= 2: return 80
                        if n >= 1: return 60
                        return 0

                    j_factors = {
                        'attend_recency': {'score': _j_score_recency(j_ds), 'weight': 30},
                        'attend_frequency': {'score': _j_score_frequency(j_a90), 'weight': 30},
                        'group_involvement': {'score': _j_score_groups(j_grp), 'weight': 20},
                        'serving': {'score': _j_score_serving(j_srv), 'weight': 20},
                    }
                    j_total_w = sum(f['weight'] for f in j_factors.values())
                    j_weighted = sum(f['score'] * f['weight'] for f in j_factors.values())
                    score = int(round(j_weighted / j_total_w)) if j_total_w > 0 else 0
                    score = max(0, min(100, score))

                    days_since = j_ds
                    response = {
                        'success': True,
                        'person': {
                            'name': safe_str(person_info.Name),
                            'age': person_info.Age or 0,
                            'email': safe_str(person_info.EmailAddress or ''),
                            'phone': safe_str(person_info.CellPhone or ''),
                            'member_status': safe_str(person_info.MemberStatus or 'Unknown'),
                            'engagement_score': score
                        },
                        'journey': events_list,
                        'insights': {
                            'journey_length': str(len(events_list)) + ' events',
                            'entry_point': events_list[1]['description'] if len(events_list) > 1 else 'Not yet engaged',
                            'total_events': len(events_list)
                        },
                        'recent_activity': {
                            'last_attendance': safe_str(recent_activity.LastAttendance) if recent_activity else None,
                            'last_signup': safe_str(recent_activity.LastSignup) if recent_activity else None,
                            'attendance_90_days': recent_activity.AttendanceCount90Days if recent_activity else 0,
                            'days_since_attendance': days_since,
                            'currently_serving': recent_activity.CurrentlyServingCount if recent_activity else 0,
                            'engagement_status': 'Active' if days_since and days_since <= 30 else 'Inactive' if days_since and days_since > 90 else 'Occasional'
                        }
                    }

        # ============================================================
        # JOURNEY: Family Engagement Timeline
        # ============================================================
        elif action == 'get_family_journey':
            people_id = int(get_form_data('peopleId', '0'))
            if not people_id:
                response = {'success': False, 'error': 'No person ID provided'}
            else:
                can_view_giving = False
                for role in ["Finance", "FinanceAdmin", "Admin"]:
                    if model.UserIsInRole(role):
                        can_view_giving = True
                        break
                if can_view_giving:
                    try:
                        q.QuerySqlTop1("SELECT TOP 1 ContributionId FROM Contribution")
                    except:
                        can_view_giving = False

                family_sql = """
                SELECT f.PeopleId, f.Name2 AS Name, f.Age,
                    f.PositionInFamilyId,
                    fp.Description AS FamilyPosition,
                    ms.Description AS MemberStatus,
                    p.FamilyId
                FROM People p
                JOIN People f ON p.FamilyId = f.FamilyId
                LEFT JOIN lookup.FamilyPosition fp ON f.PositionInFamilyId = fp.Id
                LEFT JOIN lookup.MemberStatus ms ON f.MemberStatusId = ms.Id
                WHERE p.PeopleId = {0}
                AND f.IsDeceased = 0 AND f.ArchivedFlag = 0
                ORDER BY f.PositionInFamilyId, f.Age DESC
                """.format(people_id)

                family_members_data = q.QuerySql(family_sql)
                if not family_members_data:
                    response = {'success': False, 'error': 'Family not found'}
                else:
                    family_members = []
                    family_id = 0
                    icon_map = {
                        'system': {'icon': 'fa-user-plus', 'color': '#6c757d'},
                        'program': {'icon': 'fa-users', 'color': '#17a2b8'},
                        'attendance': {'icon': 'fa-calendar-check-o', 'color': '#007bff'},
                        'smallgroup': {'icon': 'fa-users', 'color': '#28a745'},
                        'serving': {'icon': 'fa-hands-helping', 'color': '#fd7e14'},
                        'giving': {'icon': 'fa-heart', 'color': '#dc3545'}
                    }

                    for member in family_members_data:
                        family_id = member.FamilyId
                        giving_union = ""
                        if can_view_giving and SHOW_GIVING_IN_JOURNEY:
                            giving_union = """
                            UNION ALL
                            SELECT TOP 1 c.ContributionDate AS EventDate,
                                'Started Giving' AS EventType,
                                'First contribution of $' + CAST(c.ContributionAmount AS VARCHAR(20)) AS Description,
                                'giving' AS Category, 5 AS SortOrder
                            FROM Contribution c WHERE c.PeopleId = {0}
                            AND c.ContributionTypeId != 99
                            ORDER BY c.ContributionDate ASC""".format(member.PeopleId)

                        member_sql = """
                        WITH MemberJourney AS (
                            SELECT p.CreatedDate AS EventDate,
                                'Added to System' AS EventType,
                                'Profile created in database' AS Description,
                                'system' AS Category, 1 AS SortOrder
                            FROM People p WHERE p.PeopleId = {0}
                            UNION ALL
                            SELECT TOP 1 om.EnrollmentDate AS EventDate,
                                'Joined Program' AS EventType,
                                ISNULL(pr.Name, 'Unknown Program') AS Description,
                                'program' AS Category, 2 AS SortOrder
                            FROM OrganizationMembers om
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            LEFT JOIN Division d ON o.DivisionId = d.Id
                            LEFT JOIN Program pr ON d.ProgId = pr.Id
                            WHERE om.PeopleId = {0} AND o.OrganizationStatusId = 30
                            AND om.EnrollmentDate IS NOT NULL
                            ORDER BY om.EnrollmentDate ASC
                            UNION ALL
                            SELECT TOP 1 a.MeetingDate AS EventDate,
                                'First Attendance' AS EventType,
                                o.OrganizationName AS Description,
                                'attendance' AS Category, 3 AS SortOrder
                            FROM Attend a
                            JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                            WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                            ORDER BY a.MeetingDate ASC
                            UNION ALL
                            SELECT TOP 1 om.EnrollmentDate AS EventDate,
                                'Joined Small Group' AS EventType,
                                o.OrganizationName AS Description,
                                'smallgroup' AS Category, 4 AS SortOrder
                            FROM OrganizationMembers om
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            LEFT JOIN Division d ON o.DivisionId = d.Id
                            WHERE om.PeopleId = {0} AND d.ProgId = 1128
                            AND om.EnrollmentDate IS NOT NULL
                            ORDER BY om.EnrollmentDate ASC
                            UNION ALL
                            SELECT TOP 1 a.MeetingDate AS EventDate,
                                'Started Serving' AS EventType,
                                o.OrganizationName + ' (' + at.Description + ')' AS Description,
                                'serving' AS Category, 6 AS SortOrder
                            FROM Attend a
                            JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                            LEFT JOIN lookup.AttendType at ON a.AttendanceTypeId = at.Id
                            WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                            AND a.AttendanceTypeId IN (10, 20)
                            ORDER BY a.MeetingDate ASC
                            UNION ALL
                            SELECT TOP 1 om.EnrollmentDate AS EventDate,
                                'Leadership Role' AS EventType,
                                o.OrganizationName + ' (' + ISNULL(mt.Description, 'Leader') + ')' AS Description,
                                'serving' AS Category, 6 AS SortOrder
                            FROM OrganizationMembers om
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                            WHERE om.PeopleId = {0} AND om.MemberTypeId > 100
                            AND om.EnrollmentDate IS NOT NULL
                            ORDER BY om.EnrollmentDate ASC
                            {1}
                        )
                        SELECT TOP 10 EventDate, EventType, Description, Category, SortOrder
                        FROM MemberJourney
                        WHERE EventDate IS NOT NULL
                        ORDER BY EventDate ASC, SortOrder ASC
                        """.format(member.PeopleId, giving_union)

                        member_events = q.QuerySql(member_sql)
                        journey_events = []
                        for event in member_events:
                            if event.Category == 'giving' and (not SHOW_GIVING_IN_JOURNEY or not can_view_giving):
                                continue
                            icon_info = icon_map.get(event.Category, {'icon': 'fa-circle', 'color': '#6c757d'})
                            date_str = str(event.EventDate).split(' ')[0] if event.EventDate else ''
                            journey_events.append({
                                'date': date_str,
                                'event': safe_str(event.EventType),
                                'description': safe_str(event.Description),
                                'type': safe_str(event.Category),
                                'icon': icon_info['icon'],
                                'color': icon_info['color']
                            })

                        m_score = 20
                        if len(journey_events) > 1:
                            m_score += 30

                        family_members.append({
                            'person': {
                                'people_id': member.PeopleId,
                                'name': safe_str(member.Name),
                                'age': member.Age or 0,
                                'position': safe_str(member.FamilyPosition or 'Family Member'),
                                'member_status': safe_str(member.MemberStatus or 'Unknown'),
                                'engagement_score': m_score
                            },
                            'journey': journey_events,
                            'insights': {
                                'journey_length': str(len(journey_events)) + ' events',
                                'entry_point': journey_events[0]['description'] if journey_events else 'Not engaged',
                                'total_events': len(journey_events)
                            }
                        })

                    total_members = len(family_members)
                    engaged_members = sum(1 for m in family_members if m['person']['engagement_score'] >= 40)
                    avg_score = sum(m['person']['engagement_score'] for m in family_members) / total_members if total_members > 0 else 0

                    response = {
                        'success': True,
                        'family_info': {
                            'family_id': family_id,
                            'total_members': total_members,
                            'engaged_members': engaged_members,
                            'avg_engagement': int(avg_score),
                            'timeline_start': '2023-01-01',
                            'timeline_end': str(datetime.datetime.now()).split(' ')[0]
                        },
                        'family_members': family_members
                    }

        # ==========================================================
        # PROSPECT GROUPS
        # ==========================================================
        elif action == 'load_prospect_groups':
            gdata = load_groups_data()
            groups = gdata.get('groups', [])
            assignments = gdata.get('assignments', {})
            efforts = gdata.get('efforts', [])
            changelog = gdata.get('changeLog', [])
            response = {
                'success': True,
                'groups': groups,
                'assignments': assignments,
                'efforts': efforts[-200:],  # Last 200 for UI
                'changeLog': changelog
            }

        elif action == 'save_prospect_group':
            group_json = get_form_data('group_data', '{}')
            group = json.loads(group_json)
            gdata = load_groups_data()
            groups = gdata.get('groups', [])

            group_id = group.get('id', '')
            if not group_id:
                group['id'] = make_id('grp')
                group['createdAt'] = now_str()
                group['createdBy'] = model.UserPeopleId
                groups.append(group)
            else:
                for i, g in enumerate(groups):
                    if g.get('id') == group_id:
                        group['createdAt'] = g.get('createdAt', now_str())
                        group['createdBy'] = g.get('createdBy', model.UserPeopleId)
                        groups[i] = group
                        break
                else:
                    group['createdAt'] = now_str()
                    group['createdBy'] = model.UserPeopleId
                    groups.append(group)

            group['updatedAt'] = now_str()
            gdata['groups'] = groups
            save_groups_data(gdata)
            response = {'success': True, 'group': group, 'message': 'Prospect group saved'}

        elif action == 'delete_prospect_group':
            group_id = get_form_data('group_id', '')
            gdata = load_groups_data()
            gdata['groups'] = [g for g in gdata.get('groups', []) if g.get('id') != group_id]
            # Clean up assignments, efforts, changelog for this group
            if group_id in gdata.get('assignments', {}):
                del gdata['assignments'][group_id]
            gdata['efforts'] = [e for e in gdata.get('efforts', []) if e.get('groupId') != group_id]
            gdata['changeLog'] = [c for c in gdata.get('changeLog', []) if c.get('groupId') != group_id]
            save_groups_data(gdata)
            # Also drop the metrics cache entry for this group
            try:
                _gm = load_group_metrics_all()
                if group_id in _gm:
                    del _gm[group_id]
                    save_group_metrics_all(_gm)
            except:
                pass
            response = {'success': True, 'message': 'Prospect group deleted'}

        elif action == 'load_group_metrics':
            # Returns the cached metrics for one group (read-only, fast).
            group_id = get_form_data('group_id', '')
            metrics = get_group_metrics(group_id)
            response = {
                'success': True,
                'metrics': sanitize_for_json(metrics) if metrics else None,
                'windows': METRIC_WINDOWS
            }

        elif action == 'refresh_group_metrics':
            # On-demand recompute for one group. Synchronous; one group's
            # queries are bounded by its scope so this should be reasonably fast.
            group_id = get_form_data('group_id', '')
            gdata = load_groups_data()
            group = None
            for g in gdata.get('groups', []):
                if g.get('id') == group_id:
                    group = g
                    break
            if not group:
                response = {'success': False, 'message': 'Group not found'}
            else:
                try:
                    metrics = compute_group_metrics(group)
                    set_group_metrics(group_id, metrics)
                    response = {
                        'success': True,
                        'metrics': sanitize_for_json(metrics),
                        'windows': METRIC_WINDOWS
                    }
                except Exception as e:
                    response = {'success': False, 'message': safe_str(e)}

        elif action == 'get_funnel_people':
            # Returns the people in one bucket of the funnel for a specific
            # org+window. Buckets: converted, engaged, no_show, dropped.
            # Each person carries:
            #   - isCurrentProspect: still has an active prospect MemberType
            #     in OrganizationMembers for this org (no InactiveDate yet)
            #   - crossOrgAttends: other orgs in the group scope they have
            #     attended in window (the "engaged elsewhere" highlight)
            group_id = get_form_data('group_id', '')
            org_id = get_form_data('org_id', '0')
            window_days = get_form_data('window_days', '90')
            state = (get_form_data('state', 'converted') or 'converted').lower()
            try:
                org_id = int(org_id)
                window_days = int(window_days)
            except:
                response = {'success': False, 'message': 'Invalid org_id or window'}
                org_id = 0

            if org_id and state in ('converted', 'engaged', 'no_show', 'dropped'):
                gdata = load_groups_data()
                group = None
                for g in gdata.get('groups', []):
                    if g.get('id') == group_id:
                        group = g
                        break
                if not group:
                    response = {'success': False, 'message': 'Group not found'}
                else:
                    prospect_types_csv = _safe_int_csv(group.get('memberTypes', []) or [311])
                    converted_types = (group.get('convertedAttendTypeIds')
                                       or load_settings().get('converted_attend_type_ids')
                                       or [30])
                    converted_types_csv = _safe_int_csv(converted_types)

                    # Build the scope-wide list of orgs for the cross-org check.
                    scope_orgs_sql_part = ''
                    org_filter, needs_os_join = _group_scope_org_filter(group, 'om', 'o', 'os')
                    os_join_for_discovery = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId" if needs_os_join else ""
                    scope_orgs_query = """
                        SELECT DISTINCT o.OrganizationId
                        FROM Organizations o
                        {os_join}
                        WHERE o.OrganizationStatusId = 30
                          {scope_filter}
                    """.format(
                        os_join=os_join_for_discovery,
                        scope_filter=org_filter.replace('om.', 'o.') if 'om.' in org_filter else org_filter
                    )
                    scope_org_ids = []
                    try:
                        for r in q.QuerySql(scope_orgs_query):
                            scope_org_ids.append(int(r.OrganizationId))
                    except:
                        pass
                    if not scope_org_ids:
                        scope_org_ids = [org_id]
                    scope_org_ids_csv = ','.join(str(x) for x in scope_org_ids)

                    # State-specific filter clauses against PiW.
                    # PiW = prospects in window for this org.
                    state_clause = ''
                    if state == 'converted':
                        state_clause = """
                            AND EXISTS (
                                SELECT 1 FROM Attend a WITH (NOLOCK)
                                WHERE a.PeopleId = p.PeopleId
                                  AND a.OrganizationId = {org}
                                  AND a.AttendanceFlag = 1
                                  AND a.AttendanceTypeId IN ({ctypes})
                                  AND a.MeetingDate >= DATEADD(day, -{w}, GETDATE())
                            )
                        """.format(org=int(org_id), ctypes=converted_types_csv, w=int(window_days))
                    elif state == 'engaged':
                        state_clause = """
                            AND EXISTS (
                                SELECT 1 FROM Attend a WITH (NOLOCK)
                                WHERE a.PeopleId = p.PeopleId
                                  AND a.OrganizationId = {org}
                                  AND a.AttendanceFlag = 1
                                  AND a.MeetingDate >= DATEADD(day, -{w}, GETDATE())
                            )
                            AND NOT EXISTS (
                                SELECT 1 FROM Attend a2 WITH (NOLOCK)
                                WHERE a2.PeopleId = p.PeopleId
                                  AND a2.OrganizationId = {org}
                                  AND a2.AttendanceFlag = 1
                                  AND a2.AttendanceTypeId IN ({ctypes})
                                  AND a2.MeetingDate >= DATEADD(day, -{w}, GETDATE())
                            )
                        """.format(org=int(org_id), ctypes=converted_types_csv, w=int(window_days))
                    elif state == 'no_show':
                        state_clause = """
                            AND NOT EXISTS (
                                SELECT 1 FROM Attend a WITH (NOLOCK)
                                WHERE a.PeopleId = p.PeopleId
                                  AND a.OrganizationId = {org}
                                  AND a.AttendanceFlag = 1
                                  AND a.MeetingDate >= DATEADD(day, -{w}, GETDATE())
                            )
                            AND EXISTS (
                                SELECT 1 FROM OrganizationMembers om
                                WHERE om.OrganizationId = {org}
                                  AND om.PeopleId = p.PeopleId
                            )
                        """.format(org=int(org_id), w=int(window_days))
                    elif state == 'dropped':
                        state_clause = """
                            AND NOT EXISTS (
                                SELECT 1 FROM Attend a WITH (NOLOCK)
                                WHERE a.PeopleId = p.PeopleId
                                  AND a.OrganizationId = {org}
                                  AND a.AttendanceFlag = 1
                                  AND a.MeetingDate >= DATEADD(day, -{w}, GETDATE())
                            )
                            AND NOT EXISTS (
                                SELECT 1 FROM OrganizationMembers om
                                WHERE om.OrganizationId = {org}
                                  AND om.PeopleId = p.PeopleId
                            )
                        """.format(org=int(org_id), w=int(window_days))

                    # Conversion date is only relevant for the converted bucket.
                    conv_date_select = ("CONVERT(VARCHAR(10), "
                                        "(SELECT MIN(a2.MeetingDate) FROM Attend a2 WITH (NOLOCK) "
                                        "WHERE a2.PeopleId = p.PeopleId "
                                        "AND a2.OrganizationId = {org} "
                                        "AND a2.AttendanceFlag = 1 "
                                        "AND a2.AttendanceTypeId IN ({ctypes}) "
                                        "AND a2.MeetingDate >= DATEADD(day, -{w}, GETDATE())), "
                                        "120) AS FirstConvertedDate") \
                        .format(org=int(org_id), ctypes=converted_types_csv, w=int(window_days)) \
                        if state == 'converted' else "NULL AS FirstConvertedDate"

                    funnel_sql = """
                        SELECT p.PeopleId,
                               ISNULL(p.Name2, '') AS Name2,
                               p.Age,
                               ISNULL(p.EmailAddress, '') AS EmailAddress,
                               ISNULL(p.CellPhone, '') AS CellPhone,
                               -- Still a current prospect on this org's roster?
                               CASE WHEN EXISTS (
                                   SELECT 1 FROM OrganizationMembers om2
                                   WHERE om2.OrganizationId = {org}
                                     AND om2.PeopleId = p.PeopleId
                                     AND om2.MemberTypeId IN ({ptypes})
                                     AND om2.InactiveDate IS NULL
                               ) THEN 1 ELSE 0 END AS IsCurrentProspect,
                               {conv_date_select}
                        FROM People p
                        WHERE p.PeopleId IN (
                            SELECT DISTINCT et.PeopleId
                            FROM EnrollmentTransaction et
                            WHERE et.OrganizationId = {org}
                              AND et.MemberTypeId IN ({ptypes})
                              AND et.TransactionStatus = 0
                              AND et.EnrollmentDate <= GETDATE()
                              AND (et.InactiveDate IS NULL OR et.InactiveDate >= DATEADD(day, -{w}, GETDATE()))
                        )
                        {state_clause}
                        ORDER BY p.LastName, p.FirstName
                    """.format(
                        org=int(org_id),
                        ptypes=prospect_types_csv,
                        w=int(window_days),
                        conv_date_select=conv_date_select,
                        state_clause=state_clause
                    )

                    try:
                        people = []
                        people_by_pid = {}
                        pids = []
                        for r in q.QuerySql(funnel_sql):
                            try:
                                age_val = r.Age
                            except:
                                age_val = None
                            person = {
                                'peopleId': r.PeopleId,
                                'name': safe_str(r.Name2),
                                'age': int(age_val) if age_val else None,
                                'email': safe_str(r.EmailAddress),
                                'cellPhone': safe_str(r.CellPhone),
                                'isCurrentProspect': bool(int(r.IsCurrentProspect or 0)),
                                'firstConvertedDate': safe_str(r.FirstConvertedDate) if r.FirstConvertedDate else '',
                                'crossOrgAttends': []
                            }
                            people.append(person)
                            people_by_pid[r.PeopleId] = person
                            pids.append(r.PeopleId)

                        # Second query: cross-org attendance within the group's scope.
                        # Lets the UI highlight "this prospect is engaged elsewhere".
                        if pids and len(scope_org_ids) > 1:
                            pids_csv = ','.join(str(p) for p in pids)
                            cross_sql = """
                                SELECT a.PeopleId, a.OrganizationId,
                                       ISNULL(o.OrganizationName, '') AS OrganizationName,
                                       CONVERT(VARCHAR(10), MAX(a.MeetingDate), 120) AS LastAttend
                                FROM Attend a WITH (NOLOCK)
                                JOIN Organizations o ON o.OrganizationId = a.OrganizationId
                                WHERE a.PeopleId IN ({pids})
                                  AND a.OrganizationId IN ({scope})
                                  AND a.OrganizationId <> {org}
                                  AND a.AttendanceFlag = 1
                                  AND a.MeetingDate >= DATEADD(day, -{w}, GETDATE())
                                GROUP BY a.PeopleId, a.OrganizationId, o.OrganizationName
                                ORDER BY a.PeopleId, MAX(a.MeetingDate) DESC
                            """.format(
                                pids=pids_csv,
                                scope=scope_org_ids_csv,
                                org=int(org_id),
                                w=int(window_days)
                            )
                            try:
                                for r in q.QuerySql(cross_sql):
                                    pid = r.PeopleId
                                    if pid in people_by_pid:
                                        # Cap to 5 entries so the UI doesn't blow up.
                                        if len(people_by_pid[pid]['crossOrgAttends']) < 5:
                                            people_by_pid[pid]['crossOrgAttends'].append({
                                                'orgId': int(r.OrganizationId),
                                                'orgName': safe_str(r.OrganizationName),
                                                'lastAttend': safe_str(r.LastAttend)
                                            })
                            except:
                                pass

                        response = {'success': True, 'state': state, 'people': people}
                    except Exception as e:
                        response = {'success': False, 'message': safe_str(e)}

        elif action == 'assign_to_group':
            group_id = get_form_data('group_id', '')
            pids_str = get_form_data('people_ids', '')
            notes = get_form_data('notes', '')
            pids = [int(x) for x in pids_str.split(',') if x.strip()]

            gdata = load_groups_data()
            if group_id not in gdata.get('assignments', {}):
                gdata['assignments'][group_id] = {}

            assigned = 0
            user_id = model.UserPeopleId
            for pid in pids:
                pid_key = str(pid)
                if pid_key not in gdata['assignments'][group_id]:
                    # Capture initial member status for change detection
                    initial_status = ''
                    try:
                        person = model.GetPerson(pid)
                        if person:
                            initial_status = safe_str(getattr(person, 'MemberStatus', ''))
                    except:
                        pass
                    gdata['assignments'][group_id][pid_key] = {
                        'assignedAt': now_str(),
                        'assignedBy': user_id,
                        'status': 'active',
                        'initialMemberStatus': initial_status,
                        'notes': notes,
                        'fromGroup': ''
                    }
                    assigned += 1

            save_groups_data(gdata)
            response = {'success': True, 'assigned': assigned}

        elif action == 'remove_from_group':
            group_id = get_form_data('group_id', '')
            pids_str = get_form_data('people_ids', '')
            pids = [int(x) for x in pids_str.split(',') if x.strip()]

            gdata = load_groups_data()
            removed = 0
            grp_assignments = gdata.get('assignments', {}).get(group_id, {})
            for pid in pids:
                pid_key = str(pid)
                if pid_key in grp_assignments:
                    del grp_assignments[pid_key]
                    removed += 1

            save_groups_data(gdata)
            response = {'success': True, 'removed': removed}

        elif action == 'update_group_assignment':
            group_id = get_form_data('group_id', '')
            people_id = get_form_data('people_id', '')
            new_status = get_form_data('new_status', '')
            notes = get_form_data('notes', '')

            gdata = load_groups_data()
            grp_assignments = gdata.get('assignments', {}).get(group_id, {})
            if people_id in grp_assignments:
                if new_status:
                    grp_assignments[people_id]['status'] = new_status
                if notes:
                    grp_assignments[people_id]['notes'] = notes
                grp_assignments[people_id]['updatedAt'] = now_str()
                save_groups_data(gdata)
                response = {'success': True, 'message': 'Assignment updated'}
            else:
                response = {'success': False, 'message': 'Assignment not found'}

        elif action == 'move_prospect_group':
            from_group = get_form_data('from_group_id', '')
            to_group = get_form_data('to_group_id', '')
            pids_str = get_form_data('people_ids', '')
            notes = get_form_data('notes', '')
            pids = [int(x) for x in pids_str.split(',') if x.strip()]

            gdata = load_groups_data()
            moved = 0
            user_id = model.UserPeopleId
            from_assignments = gdata.get('assignments', {}).get(from_group, {})
            if to_group not in gdata.get('assignments', {}):
                gdata['assignments'][to_group] = {}

            for pid in pids:
                pid_key = str(pid)
                old_assignment = from_assignments.get(pid_key, {})
                # Add to destination
                gdata['assignments'][to_group][pid_key] = {
                    'assignedAt': now_str(),
                    'assignedBy': user_id,
                    'status': 'active',
                    'initialMemberStatus': old_assignment.get('initialMemberStatus', ''),
                    'notes': notes,
                    'fromGroup': from_group
                }
                # Remove from source
                if pid_key in from_assignments:
                    del from_assignments[pid_key]
                moved += 1

            save_groups_data(gdata)
            response = {'success': True, 'moved': moved}

        elif action == 'log_group_effort':
            effort_json = get_form_data('effort_data', '{}')
            effort = json.loads(effort_json)
            effort['id'] = make_id('eff')
            effort['timestamp'] = now_str()
            effort['userId'] = model.UserPeopleId
            try:
                user = model.GetPerson(model.UserPeopleId)
                effort['userName'] = safe_str(user.Name2) if user else 'Unknown'
            except:
                effort['userName'] = 'Unknown'

            gdata = load_groups_data()
            efforts = gdata.get('efforts', [])
            efforts.insert(0, effort)
            gdata['efforts'] = efforts
            save_groups_data(gdata)

            # Also write to central activity log
            try:
                group_name = ''
                for g in gdata.get('groups', []):
                    if g.get('id') == effort.get('groupId'):
                        group_name = g.get('name', '')
                        break
                effort_type = effort.get('effortType', 'contact')
                result = effort.get('result', '')
                desc = effort.get('description', '')
                detail_parts = [effort_type]
                if result:
                    detail_parts.append(result)
                if desc:
                    detail_parts.append(desc)
                log_entry = {
                    'configId': '',
                    'configName': '',
                    'groupId': effort.get('groupId', ''),
                    'groupName': group_name,
                    'peopleId': effort.get('peopleId', 0),
                    'personName': effort.get('personName', ''),
                    'actionType': 'group_effort',
                    'actionDetail': ' / '.join(detail_parts),
                    'timestamp': effort.get('timestamp', now_str()),
                    'userName': effort.get('userName', 'Unknown'),
                    'userId': effort.get('userId', 0),
                    'source': 'group'
                }
                activity_log = load_content("ProspectBuilder_ActivityLog", [])
                activity_log.insert(0, log_entry)
                if len(activity_log) > 500:
                    activity_log = activity_log[:500]
                save_content("ProspectBuilder_ActivityLog", activity_log)
            except:
                pass

            response = {'success': True, 'effort': effort}

        elif action == 'detect_group_changes':
            group_id = get_form_data('group_id', '')
            gdata = load_groups_data()
            assignments = gdata.get('assignments', {}).get(group_id, {})
            changes = []

            if assignments:
                pids = [int(k) for k in assignments.keys()]
                pid_list = ','.join(str(p) for p in pids)

                # Check for new attendance since assignment
                for pid_key, assignment in assignments.items():
                    pid = int(pid_key)
                    assigned_at = assignment.get('assignedAt', '2020-01-01')

                    # Check attendance
                    att_sql = """
                        SELECT TOP 1 m.MeetingDate
                        FROM Attend a WITH (NOLOCK)
                        JOIN Meetings m WITH (NOLOCK) ON a.MeetingId = m.MeetingId
                        WHERE a.PeopleId = {0}
                            AND a.AttendanceFlag = 1
                            AND m.MeetingDate >= '{1}'
                        ORDER BY m.MeetingDate ASC
                    """.format(pid, assigned_at[:10])
                    try:
                        att_result = q.QuerySql(att_sql)
                        for row in att_result:
                            if row.MeetingDate:
                                changes.append({
                                    'id': make_id('cl'),
                                    'groupId': group_id,
                                    'peopleId': pid,
                                    'changeType': 'attended',
                                    'description': 'Attended on ' + safe_str(row.MeetingDate)[:10],
                                    'detectedAt': now_str(),
                                    'acknowledged': False
                                })
                                break
                    except:
                        pass

                    # Check member status change
                    try:
                        person = model.GetPerson(pid)
                        if person:
                            current_status = safe_str(getattr(person, 'MemberStatus', ''))
                            initial = assignment.get('initialMemberStatus', '')
                            if initial and current_status and current_status != initial:
                                changes.append({
                                    'id': make_id('cl'),
                                    'groupId': group_id,
                                    'peopleId': pid,
                                    'changeType': 'status_change',
                                    'description': 'Status changed: ' + initial + ' -> ' + current_status,
                                    'detectedAt': now_str(),
                                    'acknowledged': False
                                })
                    except:
                        pass

                    # Check new org enrollments since assignment
                    enroll_sql = """
                        SELECT TOP 3 o.OrganizationName, om.EnrollmentDate
                        FROM OrganizationMembers om WITH (NOLOCK)
                        JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
                        WHERE om.PeopleId = {0}
                            AND om.EnrollmentDate >= '{1}'
                        ORDER BY om.EnrollmentDate DESC
                    """.format(pid, assigned_at[:10])
                    try:
                        enroll_result = q.QuerySql(enroll_sql)
                        for row in enroll_result:
                            if row.OrganizationName:
                                changes.append({
                                    'id': make_id('cl'),
                                    'groupId': group_id,
                                    'peopleId': pid,
                                    'changeType': 'joined_group',
                                    'description': 'Joined: ' + safe_str(row.OrganizationName),
                                    'detectedAt': now_str(),
                                    'acknowledged': False
                                })
                    except:
                        pass

            # Deduplicate against existing changelog entries
            existing = gdata.get('changeLog', [])
            existing_keys = set()
            for ec in existing:
                existing_keys.add('{0}_{1}_{2}'.format(ec.get('groupId'), ec.get('peopleId'), ec.get('changeType')))

            new_changes = []
            for ch in changes:
                ch_key = '{0}_{1}_{2}'.format(ch.get('groupId'), ch.get('peopleId'), ch.get('changeType'))
                if ch_key not in existing_keys:
                    new_changes.append(ch)
                    existing_keys.add(ch_key)

            if new_changes:
                gdata['changeLog'] = new_changes + existing
                save_groups_data(gdata)

            response = {'success': True, 'newChanges': len(new_changes), 'totalChanges': len(gdata.get('changeLog', []))}

        elif action == 'acknowledge_group_change':
            change_id = get_form_data('change_id', '')
            gdata = load_groups_data()
            for ch in gdata.get('changeLog', []):
                if ch.get('id') == change_id:
                    ch['acknowledged'] = True
                    break
            save_groups_data(gdata)
            response = {'success': True}

        elif action == 'get_group_stats':
            gdata = load_groups_data()
            groups = gdata.get('groups', [])
            efforts = gdata.get('efforts', [])
            changelog = gdata.get('changeLog', [])

            stats = {}
            now = datetime.datetime.now()
            thirty_days_ago = (now - datetime.timedelta(days=30)).strftime('%Y-%m-%d')

            for g in groups:
                gid = g['id']

                # Count prospects from the ministry level query (same logic as get_group_prospect_details)
                level = g.get('level', 'program')
                prog_id = g.get('programId', 0)
                div_id = g.get('divisionId', 0)
                org_id = g.get('orgId', 0)
                member_types = g.get('memberTypes', [])
                min_enroll_days = g.get('minEnrollDays', 0)
                min_stale_days = g.get('minStaleDays', 0)

                where_clause = ''
                join_clause = ''
                if level == 'involvement' and org_id:
                    where_clause = 'om.OrganizationId = {0}'.format(org_id)
                elif level == 'division' and div_id:
                    where_clause = 'o.DivisionId = {0} AND o.OrganizationStatusId = 30'.format(div_id)
                elif prog_id:
                    join_clause = 'JOIN Division d ON o.DivisionId = d.Id'
                    where_clause = 'd.ProgId = {0} AND o.OrganizationStatusId = 30'.format(prog_id)

                if member_types and where_clause:
                    mt_list = ','.join(str(mt) for mt in member_types)
                    where_clause += ' AND om.MemberTypeId IN ({0})'.format(mt_list)

                if min_enroll_days and min_enroll_days > 0 and where_clause:
                    where_clause += ' AND om.EnrollmentDate <= DATEADD(day, -{0}, GETDATE())'.format(int(min_enroll_days))

                stale_filter = ''
                if min_stale_days and min_stale_days > 0:
                    stale_filter = """
                        AND p.PeopleId NOT IN (
                            SELECT tn.AboutPersonId FROM TaskNote tn WITH (NOLOCK)
                            WHERE tn.AboutPersonId = p.PeopleId
                              AND tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())
                        )
                    """.format(int(min_stale_days))

                total_count = 0
                if where_clause:
                    count_sql = """
                        SELECT COUNT(DISTINCT p.PeopleId) as Cnt
                        FROM OrganizationMembers om WITH (NOLOCK)
                        JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
                        {0}
                        JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
                        WHERE {1}
                          AND p.IsDeceased = 0
                          {2}
                    """.format(join_clause, where_clause, stale_filter)
                    try:
                        for row in q.QuerySql(count_sql):
                            total_count = int(row.Cnt) if row.Cnt else 0
                    except:
                        pass

                # Recent efforts (from group effort log)
                recent_efforts = sum(1 for e in efforts if e.get('groupId') == gid and e.get('timestamp', '') >= thirty_days_ago)

                # Unacknowledged changes
                unacked = sum(1 for c in changelog if c.get('groupId') == gid and not c.get('acknowledged'))

                stats[gid] = {
                    'activeCount': total_count,
                    'totalAssigned': total_count,
                    'recentEfforts': recent_efforts,
                    'staleCount': 0,
                    'unackedChanges': unacked
                }

            response = {'success': True, 'stats': stats}

        elif action == 'get_group_prospect_details':
            group_id = get_form_data('group_id', '')
            gdata = load_groups_data()
            assignments = gdata.get('assignments', {}).get(group_id, {})
            efforts = gdata.get('efforts', [])

            # Find the group definition to get its ministry level
            group_def = None
            for g in gdata.get('groups', []):
                if g.get('id') == group_id:
                    group_def = g
                    break

            # Build last-effort map for this group
            prospect_last_effort = {}
            prospect_effort_count = {}
            for e in efforts:
                if e.get('groupId') == group_id:
                    pid = str(e.get('peopleId', ''))
                    if pid:
                        prospect_effort_count[pid] = prospect_effort_count.get(pid, 0) + 1
                        ts = e.get('timestamp', '')
                        if pid not in prospect_last_effort or ts > prospect_last_effort[pid]:
                            prospect_last_effort[pid] = ts

            # Query prospects from the group's ministry level
            level = group_def.get('level', 'program') if group_def else 'program'
            prog_id = group_def.get('programId', 0) if group_def else 0
            div_id = group_def.get('divisionId', 0) if group_def else 0
            org_id = group_def.get('orgId', 0) if group_def else 0

            where_clause = ''
            join_clause = ''
            member_types = group_def.get('memberTypes', []) if group_def else []

            if level == 'involvement' and org_id:
                where_clause = 'om.OrganizationId = {0}'.format(org_id)
            elif level == 'division' and div_id:
                where_clause = 'o.DivisionId = {0} AND o.OrganizationStatusId = 30'.format(div_id)
            elif prog_id:
                join_clause = 'JOIN Division d ON o.DivisionId = d.Id'
                where_clause = 'd.ProgId = {0} AND o.OrganizationStatusId = 30'.format(prog_id)

            # Filter by member types if specified
            if member_types and where_clause:
                mt_list = ','.join(str(mt) for mt in member_types)
                where_clause += ' AND om.MemberTypeId IN ({0})'.format(mt_list)

            # Filter by minimum enrollment days
            min_enroll_days = group_def.get('minEnrollDays', 0) if group_def else 0
            if min_enroll_days and min_enroll_days > 0 and where_clause:
                where_clause += ' AND om.EnrollmentDate <= DATEADD(day, -{0}, GETDATE())'.format(int(min_enroll_days))

            # Filter by minimum days with no contact (stale threshold)
            min_stale_days = group_def.get('minStaleDays', 0) if group_def else 0

            prospects = []
            if where_clause:
                # Prospect member types (220=Member, 230=Prospect, 310=Leader, 500=Visitor)
                # Build stale filter as subquery if needed
                stale_filter = ''
                if min_stale_days and min_stale_days > 0:
                    stale_filter = """
                        AND p.PeopleId NOT IN (
                            SELECT tn.AboutPersonId FROM TaskNote tn WITH (NOLOCK)
                            WHERE tn.AboutPersonId = p.PeopleId
                              AND tn.CreatedDate >= DATEADD(day, -{0}, GETDATE())
                        )
                    """.format(int(min_stale_days))

                sql = """
                    SELECT DISTINCT TOP 500 p.PeopleId, p.Name2, p.EmailAddress, p.CellPhone, p.Age,
                           p.PictureId,
                           ms.Description as MemberStatus,
                           om.MemberTypeId,
                           mt.Description as MemberType,
                           om.EnrollmentDate
                    FROM OrganizationMembers om WITH (NOLOCK)
                    JOIN Organizations o WITH (NOLOCK) ON om.OrganizationId = o.OrganizationId
                    {0}
                    JOIN People p WITH (NOLOCK) ON om.PeopleId = p.PeopleId
                    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                    WHERE {1}
                      AND p.IsDeceased = 0
                      {2}
                    ORDER BY p.Name2
                """.format(join_clause, where_clause, stale_filter)

                seen_pids = set()
                for row in q.QuerySql(sql):
                    pid_key = str(row.PeopleId)
                    if row.PeopleId in seen_pids:
                        continue
                    seen_pids.add(row.PeopleId)

                    assignment = assignments.get(pid_key, {})
                    prospects.append({
                        'PeopleId': row.PeopleId,
                        'Name2': safe_str(row.Name2),
                        'EmailAddress': safe_str(row.EmailAddress),
                        'CellPhone': safe_str(row.CellPhone),
                        'Age': row.Age,
                        'MemberStatus': safe_str(row.MemberStatus),
                        'MemberType': safe_str(row.MemberType),
                        'PictureId': row.PictureId,
                        'EnrollmentDate': safe_str(row.EnrollmentDate)[:10] if row.EnrollmentDate else '',
                        'assignedAt': assignment.get('assignedAt', ''),
                        'status': assignment.get('status', 'new'),
                        'notes': assignment.get('notes', ''),
                        'fromGroup': assignment.get('fromGroup', ''),
                        'effortCount': prospect_effort_count.get(pid_key, 0),
                        'lastEffort': prospect_last_effort.get(pid_key, ''),
                        'isAssigned': pid_key in assignments
                    })

            response = {'success': True, 'prospects': sanitize_for_json(prospects)}

        # ==========================================================
        # PROSPECT SENDER ACTIONS
        # ==========================================================
        elif action == 'load_senders':
            senders = load_senders()
            response = {'success': True, 'senders': sanitize_for_json(senders)}

        elif action == 'save_sender':
            sender_json = get_form_data('sender_data', '{}')
            sender = json.loads(sender_json)
            senders = load_senders()

            sender_id = sender.get('id', '')
            if not sender_id:
                sender['id'] = make_id('snd')
                sender['createdAt'] = now_str()
                senders.append(sender)
            else:
                for i, s in enumerate(senders):
                    if s.get('id') == sender_id:
                        sender['createdAt'] = s.get('createdAt', now_str())
                        sender['last_sent'] = s.get('last_sent', '')
                        sender['last_sent_count'] = s.get('last_sent_count', 0)
                        senders[i] = sender
                        break
                else:
                    sender['createdAt'] = now_str()
                    senders.append(sender)

            sender['updatedAt'] = now_str()
            save_senders(senders)
            response = {'success': True, 'sender': sanitize_for_json(sender), 'message': 'Sender saved'}

        elif action == 'delete_sender':
            sender_id = get_form_data('sender_id', '')
            senders = load_senders()
            senders = [s for s in senders if s.get('id') != sender_id]
            save_senders(senders)
            response = {'success': True, 'message': 'Sender deleted'}

        elif action == 'preview_sender':
            sender_json = get_form_data('sender_data', '{}')
            sender = json.loads(sender_json)
            result = execute_sender(sender, dry_run=True, triggered_by='preview')
            response = {'success': True, 'result': sanitize_for_json(result)}

        elif action == 'run_sender':
            sender_id = get_form_data('sender_id', '')
            senders = load_senders()
            sender = None
            for s in senders:
                if s.get('id') == sender_id:
                    sender = s
                    break
            if not sender:
                response = {'success': False, 'message': 'Sender not found'}
            else:
                result = execute_sender(sender, dry_run=False, triggered_by='manual')
                response = {'success': True, 'result': sanitize_for_json(result)}

        elif action == 'run_sender_oneoff':
            # One-off send: uses sender config but ignores schedule/last_sent
            sender_json = get_form_data('sender_data', '{}')
            sender = json.loads(sender_json)
            # Override lookback to get all prospects regardless of last_sent
            sender['lookback'] = sender.get('oneoff_lookback', 'last_7_days')
            sender['last_sent'] = ''  # Clear to get all in window
            result = execute_sender(sender, dry_run=False, triggered_by='oneoff')
            response = {'success': True, 'result': sanitize_for_json(result)}

        elif action == 'get_sender_log':
            log = load_sender_log()
            sender_id = get_form_data('sender_id', '')
            if sender_id:
                log = [l for l in log if l.get('sender_id') == sender_id]
            response = {'success': True, 'log': sanitize_for_json(log[-50:])}

        elif action == 'get_roles':
            sql = """
                SELECT DISTINCT r.RoleName
                FROM Roles r
                JOIN UserRole ur ON r.RoleId = ur.RoleId
                JOIN Users u ON ur.UserId = u.UserId
                WHERE u.IsApproved = 1
                ORDER BY r.RoleName
            """
            roles = []
            for row in q.QuerySql(sql):
                roles.append(safe_str(row.RoleName))
            response = {'success': True, 'roles': roles}

        elif action == 'get_programs_divisions':
            programs = []
            for row in q.QuerySql("SELECT Id, Name FROM Program ORDER BY Name"):
                programs.append({'id': int(row.Id), 'name': safe_str(row.Name)})
            divisions = []
            for row in q.QuerySql("SELECT Id, Name, ProgId FROM Division ORDER BY Name"):
                divisions.append({'id': int(row.Id), 'name': safe_str(row.Name), 'progId': int(row.ProgId)})
            response = {'success': True, 'programs': programs, 'divisions': divisions}

        elif action == 'drop_from_org':
            pid = int(get_form_data('people_id', '0'))
            oid = int(get_form_data('org_id', '0'))
            if pid > 0 and oid > 0:
                try:
                    person = model.GetPerson(pid)
                    if person:
                        model.RemoveFromOrg(person, oid)
                        response = {'success': True, 'message': 'Removed PeopleId {0} from Org {1}'.format(pid, oid)}
                    else:
                        response = {'success': False, 'message': 'Person not found'}
                except Exception as e:
                    response = {'success': False, 'message': safe_str(e)}
            else:
                response = {'success': False, 'message': 'Invalid people_id or org_id'}

        elif action == 'get_prospect_health':
            sender_json = get_form_data('sender_data', '{}')
            sender = json.loads(sender_json)
            source = sender.get('source', {})
            scope = source.get('scope', '')
            member_types = source.get('member_types', '311')
            recip_member_types = sender.get('recipient_member_types', '')

            mt_filter = ''
            if member_types:
                mt_filter = "AND om.MemberTypeId IN ({0})".format(member_types)

            # Build org filter
            needs_os = scope in ('program', 'division')
            os_join = "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId" if needs_os else ""
            org_filter = ''
            if scope == 'org':
                org_filter = "AND om.OrganizationId = {0}".format(source.get('org_id', 0))
            elif scope == 'program':
                org_filter = "AND os.ProgId = {0}".format(source.get('program_id', 0))
            elif scope == 'division':
                p = source.get('program_id', '')
                d = source.get('division_id', '')
                if p and d:
                    org_filter = "AND os.ProgId = {0} AND os.DivId = {1}".format(p, d)
                elif p:
                    org_filter = "AND os.ProgId = {0}".format(p)

            # Get all prospects (no date filter — show ALL current prospects)
            health_sql = '''
                SELECT p.PeopleId, p.Name, p.EmailAddress, p.CellPhone,
                       p.Age,
                       om.EnrollmentDate, om.OrganizationId,
                       o.OrganizationName,
                       DATEDIFF(day, om.EnrollmentDate, GETDATE()) as DaysInGroup
                FROM People p
                JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
                JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                {0}
                WHERE o.OrganizationStatusId = 30
                    AND p.IsDeceased = 0
                    {1}
                    {2}
                ORDER BY o.OrganizationName, om.EnrollmentDate
            '''.format(os_join, org_filter, mt_filter)

            # Get leader counts per org
            recip_mt_filter = ''
            if recip_member_types:
                recip_mt_filter = "AND om.MemberTypeId IN ({0})".format(recip_member_types)
            leader_sql = '''
                SELECT om.OrganizationId, COUNT(DISTINCT p.PeopleId) as LeaderCount
                FROM People p
                JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
                JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                {0}
                WHERE o.OrganizationStatusId = 30
                    AND p.IsDeceased = 0
                    {1}
                    {2}
                GROUP BY om.OrganizationId
            '''.format(os_join, org_filter, recip_mt_filter)

            # Get contact effort counts (TaskNotes with keywords)
            settings = load_settings()
            contact_methods = settings.get('contact_methods', [])
            keyword_ids = [str(cm.get('keywordId', 0)) for cm in contact_methods if cm.get('keywordId')]

            effort_counts = {}

            # Build health data by org
            orgs = {}
            try:
                for r in q.QuerySql(health_sql):
                    oid = r.OrganizationId
                    if oid not in orgs:
                        orgs[oid] = {'name': safe_str(r.OrganizationName), 'prospects': [], 'leaders': 0}
                    days = int(r.DaysInGroup) if r.DaysInGroup else 0
                    orgs[oid]['prospects'].append({
                        'people_id': r.PeopleId,
                        'name': safe_str(r.Name),
                        'email': safe_str(r.EmailAddress) or '',
                        'cell_phone': safe_str(r.CellPhone) or '',
                        'age': safe_str(r.Age) if r.Age else '',
                        'days': days,
                        'enrolled': safe_str(r.EnrollmentDate)[:10] if r.EnrollmentDate else '',
                        'efforts': 0,
                        'last_effort': '',
                        'priority': 0
                    })
            except Exception as e:
                response = {'success': False, 'message': safe_str(e)}

            # Get effort counts for all prospects found
            all_prospect_pids = []
            for oid in orgs:
                for p in orgs[oid]['prospects']:
                    all_prospect_pids.append(str(p['people_id']))
            if all_prospect_pids:
                # Batch in groups of 500 to avoid query limits
                for batch_start in range(0, len(all_prospect_pids), 500):
                    batch = all_prospect_pids[batch_start:batch_start+500]
                    effort_sql = '''
                        SELECT tn.AboutPersonId as PeopleId, COUNT(*) as EffortCount,
                               MAX(tn.CreatedDate) as LastEffort
                        FROM TaskNote tn WITH (NOLOCK)
                        WHERE tn.AboutPersonId IN ({0})
                            AND tn.CreatedDate >= DATEADD(day, -180, GETDATE())
                        GROUP BY tn.AboutPersonId
                    '''.format(','.join(batch))
                    try:
                        for r in q.QuerySql(effort_sql):
                            effort_counts[r.PeopleId] = {
                                'count': int(r.EffortCount),
                                'last': safe_str(r.LastEffort)[:10] if r.LastEffort else ''
                            }
                    except:
                        pass

            # Apply effort counts to prospects
            for oid in orgs:
                for p in orgs[oid]['prospects']:
                    effort = effort_counts.get(p['people_id'], {'count': 0, 'last': ''})
                    p['efforts'] = effort['count']
                    p['last_effort'] = effort['last']
                    p['priority'] = p['days'] - (effort['count'] * 7)
                    if p['priority'] < 0:
                        p['priority'] = 0

            # Add leader counts
            try:
                for r in q.QuerySql(leader_sql):
                    oid = r.OrganizationId
                    if oid in orgs:
                        orgs[oid]['leaders'] = int(r.LeaderCount)
            except:
                pass

            # Get meeting info per org
            meeting_info = {}
            if orgs:
                org_ids = ','.join(str(oid) for oid in orgs)
                meeting_sql = '''
                    SELECT m.OrganizationId,
                           COUNT(*) as MeetingCount,
                           MAX(m.MeetingDate) as LastMeeting,
                           MIN(m.MeetingDate) as FirstMeeting
                    FROM Meetings m
                    WHERE m.OrganizationId IN ({0})
                        AND m.MeetingDate >= DATEADD(day, -90, GETDATE())
                        AND m.DidNotMeet = 0
                    GROUP BY m.OrganizationId
                '''.format(org_ids)
                try:
                    for r in q.QuerySql(meeting_sql):
                        oid = r.OrganizationId
                        count = int(r.MeetingCount) if r.MeetingCount else 0
                        last = safe_str(r.LastMeeting)[:10] if r.LastMeeting else ''
                        first = safe_str(r.FirstMeeting)[:10] if r.FirstMeeting else ''
                        # Calculate frequency
                        freq = ''
                        if count >= 12:
                            freq = 'Weekly'
                        elif count >= 6:
                            freq = 'Bi-weekly'
                        elif count >= 3:
                            freq = 'Monthly'
                        elif count > 0:
                            freq = 'Occasional'
                        else:
                            freq = 'None'
                        meeting_info[oid] = {
                            'count_90d': count,
                            'last_meeting': last,
                            'frequency': freq
                        }
                except:
                    pass

            # Sort prospects within each org by priority (desc)
            for oid in orgs:
                orgs[oid]['prospects'].sort(key=lambda p: -p['priority'])

            # Convert to list sorted by org name
            org_list = []
            for oid in sorted(orgs, key=lambda k: orgs[k]['name']):
                o = orgs[oid]
                avg_days = sum(p['days'] for p in o['prospects']) / len(o['prospects']) if o['prospects'] else 0
                avg_efforts = sum(p['efforts'] for p in o['prospects']) / len(o['prospects']) if o['prospects'] else 0
                no_contact = sum(1 for p in o['prospects'] if p['efforts'] == 0)
                stale = sum(1 for p in o['prospects'] if p['days'] > 14 and p['efforts'] == 0)
                mi = meeting_info.get(oid, {})
                org_list.append({
                    'org_id': oid,
                    'name': o['name'],
                    'leaders': o['leaders'],
                    'total_prospects': len(o['prospects']),
                    'avg_days': round(avg_days, 1),
                    'avg_efforts': round(avg_efforts, 1),
                    'no_contact': no_contact,
                    'stale': stale,
                    'meetings_90d': mi.get('count_90d', 0),
                    'last_meeting': mi.get('last_meeting', ''),
                    'meeting_freq': mi.get('frequency', ''),
                    'prospects': o['prospects']
                })

            response = {'success': True, 'orgs': sanitize_for_json(org_list), 'total_orgs': len(org_list),
                        'total_prospects': sum(len(o['prospects']) for o in orgs.values())}

        elif action == 'get_email_templates':
            sql = """
                SELECT c.Id, c.Name, c.Title, c.TypeId
                FROM Content c
                WHERE c.TypeId IN (2, 7)
                ORDER BY c.Title
            """
            templates = []
            for row in q.QuerySql(sql):
                templates.append({
                    'id': int(row.Id),
                    'name': safe_str(row.Name),
                    'title': safe_str(row.Title) or safe_str(row.Name),
                    'type_id': int(row.TypeId)
                })
            response = {'success': True, 'templates': templates}

    except Exception as e:
        response = {'success': False, 'message': safe_str(e), 'trace': safe_str(traceback.format_exc())}

    print json.dumps(sanitize_for_json(response))

# ============================================================
# GENERATE UI (GET request)
# ============================================================
else:
    def generate_ui():
        return '''
<script>
if (window.location.pathname.indexOf('/PyScript/') > -1) {
    window.location.href = window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search;
}
</script>
<style>
/* ============================================================ */
/* PROSPECT BUILDER CSS - .pb- prefix                           */
/* ============================================================ */
:root {
    --pb-primary: #2c3e50;
    --pb-accent: #3498db;
    --pb-success: #27ae60;
    --pb-warning: #f39c12;
    --pb-danger: #e74c3c;
    --pb-light: #ecf0f1;
    --pb-border: #ddd;
    --pb-text: #333;
    --pb-muted: #7f8c8d;
    --pb-white: #fff;
    --pb-shadow: 0 2px 8px rgba(0,0,0,0.1);
    --pb-radius: 6px;
}
.page-header { display: none !important; }
.pb-container { max-width: 1400px; margin: 0 auto; padding: 15px; font-family: 'Segoe UI', sans-serif; color: var(--pb-text); }
.pb-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid var(--pb-primary); }
.pb-header h2 { margin: 0; color: var(--pb-primary); font-size: 1.5em; }
.pb-tabs { display: flex; gap: 0; border-bottom: 2px solid var(--pb-border); margin-bottom: 15px; }
.pb-tab { padding: 10px 20px; cursor: pointer; border: 1px solid transparent; border-bottom: none; color: var(--pb-muted); font-weight: 600; transition: all 0.2s; background: none; border-radius: var(--pb-radius) var(--pb-radius) 0 0; }
.pb-tab:hover { color: var(--pb-accent); background: rgba(52,152,219,0.05); }
.pb-tab.active { color: var(--pb-accent); border-color: var(--pb-border); border-bottom: 2px solid var(--pb-white); margin-bottom: -2px; background: var(--pb-white); }
.pb-tab-content { display: none; }
.pb-tab-content.active { display: block; }
/* Dismissible per-tab "What is this?" banners. Shown by default on the
   first visit; localStorage persists the user's dismiss decision per-tab. */
.pb-tab-intro { display:flex; gap:12px; align-items:flex-start;
                background:#eef5fb; border:1px solid #cfe3ff;
                border-radius:8px; padding:12px 14px; margin-bottom:14px;
                font-size:0.88em; line-height:1.5; color:#264960; }
.pb-tab-intro-icon { font-size:1.6em; line-height:1; flex-shrink:0; }
.pb-tab-intro-body { flex:1; min-width:0; }
.pb-tab-intro-link { display:inline-block; margin-top:6px; color:var(--pb-primary);
                     font-weight:600; cursor:pointer; text-decoration:underline; }
.pb-tab-intro-close { background:none; border:0; color:#5a7a92; cursor:pointer;
                      font-size:1.4em; line-height:1; padding:0 4px;
                      flex-shrink:0; }
.pb-tab-intro-close:hover { color:var(--pb-danger); }

/* Cards */
.pb-card { background: var(--pb-white); border: 1px solid var(--pb-border); border-radius: var(--pb-radius); padding: 15px; margin-bottom: 12px; box-shadow: var(--pb-shadow); }
.pb-card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
.pb-card-title { font-weight: 700; font-size: 1.1em; color: var(--pb-primary); }

/* Buttons */
.pb-btn { display: inline-flex; align-items: center; gap: 5px; padding: 7px 14px; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); cursor: pointer; font-size: 0.9em; font-weight: 600; transition: all 0.2s; background: var(--pb-white); color: var(--pb-text); }
.pb-btn:hover { transform: translateY(-1px); box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
.pb-btn-primary { background: var(--pb-accent); color: var(--pb-white); border-color: var(--pb-accent); }
.pb-btn-primary:hover { background: #2980b9; }
.pb-btn-success { background: var(--pb-success); color: var(--pb-white); border-color: var(--pb-success); }
.pb-btn-success:hover { background: #219a52; }
.pb-btn-danger { background: var(--pb-danger); color: var(--pb-white); border-color: var(--pb-danger); }
.pb-btn-danger:hover { background: #c0392b; }
.pb-btn-warning { background: var(--pb-warning); color: var(--pb-white); border-color: var(--pb-warning); }
.pb-btn-sm { padding: 4px 10px; font-size: 0.8em; }
.pb-btn-group { display: flex; gap: 6px; flex-wrap: wrap; }

/* Forms */
.pb-form-group { margin-bottom: 12px; }
.pb-form-group label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 0.9em; color: var(--pb-primary); }
.pb-input, .pb-select { width: 100%; padding: 8px 10px; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); font-size: 0.9em; box-sizing: border-box; }
.pb-input:focus, .pb-select:focus { outline: none; border-color: var(--pb-accent); box-shadow: 0 0 0 2px rgba(52,152,219,0.15); }
.pb-textarea { width: 100%; padding: 8px 10px; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); font-size: 0.9em; min-height: 60px; box-sizing: border-box; resize: vertical; }

/* Modal */
.pb-modal-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 10000; justify-content: center; align-items: flex-start; padding-top: 40px; }
.pb-modal-overlay.active { display: flex; }
.pb-modal { background: var(--pb-white); border-radius: var(--pb-radius); width: 90%; max-width: 800px; max-height: 85vh; overflow-y: auto; box-shadow: 0 10px 40px rgba(0,0,0,0.3); }
.pb-modal-header { display: flex; justify-content: space-between; align-items: center; padding: 15px 20px; border-bottom: 1px solid var(--pb-border); background: var(--pb-light); border-radius: var(--pb-radius) var(--pb-radius) 0 0; }
.pb-modal-header h3 { margin: 0; color: var(--pb-primary); }
.pb-modal-close { background: none; border: none; font-size: 1.5em; cursor: pointer; color: var(--pb-muted); padding: 0 5px; }
.pb-modal-body { padding: 20px; }
.pb-modal-footer { padding: 15px 20px; border-top: 1px solid var(--pb-border); display: flex; justify-content: flex-end; gap: 8px; }

/* Table */
.pb-table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
.pb-table th { background: var(--pb-light); padding: 8px 10px; text-align: left; font-weight: 700; border-bottom: 2px solid var(--pb-border); color: var(--pb-primary); position: sticky; top: 0; }
.pb-table td { padding: 8px 10px; border-bottom: 1px solid var(--pb-border); }
.pb-table tr:hover { background: rgba(52,152,219,0.04); }
.pb-table-wrap { max-height: 600px; overflow-y: auto; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); }

/* Badges */
.pb-badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 0.75em; font-weight: 700; }
.pb-badge-primary { background: rgba(52,152,219,0.15); color: var(--pb-accent); }
.pb-badge-success { background: rgba(39,174,96,0.15); color: var(--pb-success); }
.pb-badge-warning { background: rgba(243,156,18,0.15); color: var(--pb-warning); }
.pb-badge-danger { background: rgba(231,76,60,0.15); color: var(--pb-danger); }
.pb-badge-muted { background: var(--pb-light); color: var(--pb-muted); }
.pb-badge-accent { background: rgba(155,89,182,0.15); color: #9b59b6; }

/* Contact badges */
.pb-contact-badge { display: inline-flex; gap: 3px; }
.pb-contact-code { display: inline-block; padding: 1px 5px; border-radius: 3px; font-size: 0.75em; font-weight: 700; background: var(--pb-light); color: var(--pb-muted); }
.pb-contact-code.has-count { background: rgba(52,152,219,0.15); color: var(--pb-accent); }

/* Flag badges */
.pb-flag-badge { display: inline-block; padding: 2px 6px; border-radius: 3px; font-size: 0.7em; font-weight: 700; margin: 1px; }
.pb-flag-on { background: rgba(39,174,96,0.2); color: var(--pb-success); border: 1px solid rgba(39,174,96,0.3); }
.pb-flag-off { background: var(--pb-light); color: var(--pb-muted); border: 1px solid var(--pb-border); opacity: 0.5; }

/* Prospect cards */
.pb-prospect-card { background: var(--pb-white); border: 1px solid var(--pb-border); border-radius: var(--pb-radius); padding: 8px 12px; margin-bottom: 4px; transition: all 0.2s; }
.pb-prospect-card:hover { box-shadow: var(--pb-shadow); border-color: var(--pb-accent); }

/* Action dropdown */
.pb-action-wrap { position: relative; }
.pb-action-trigger { padding: 4px 10px; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); background: var(--pb-white); cursor: pointer; font-size: 0.8em; font-weight: 600; color: var(--pb-accent); white-space: nowrap; }
.pb-action-trigger:hover { background: rgba(52,152,219,0.08); border-color: var(--pb-accent); }
.pb-action-menu { display: none; position: absolute; right: 0; top: 100%; min-width: 280px; background: var(--pb-white); border: 1px solid var(--pb-border); border-radius: var(--pb-radius); box-shadow: 0 6px 20px rgba(0,0,0,0.15); z-index: 500; max-height: 400px; overflow-y: auto; }
.pb-action-menu.open { display: block; }
/* Variant that opens upward -- used when the trigger sits near the bottom
   of a tall card (e.g. Single view) where dropping down would render
   off-screen. */
.pb-action-menu.pb-action-menu-up { top: auto; bottom: 100%;
    box-shadow: 0 -6px 20px rgba(0,0,0,0.15); }
.pb-action-menu-section { padding: 4px 0; border-bottom: 1px solid var(--pb-light); }
.pb-action-menu-section:last-child { border-bottom: none; }
.pb-action-menu-label { padding: 4px 12px; font-size: 0.7em; font-weight: 700; color: var(--pb-muted); text-transform: uppercase; letter-spacing: 0.5px; }
.pb-action-menu-item { padding: 6px 12px; cursor: pointer; font-size: 0.85em; display: flex; align-items: center; gap: 8px; }
.pb-action-menu-item:hover { background: rgba(52,152,219,0.06); }
.pb-action-menu-item .pb-ami-icon { width: 20px; text-align: center; flex-shrink: 0; }
.pb-action-menu-item .pb-ami-health { font-size: 0.75em; color: var(--pb-muted); margin-left: auto; }
.pb-prospect-card.processed { border-left: 4px solid var(--pb-success); opacity: 0.7; }
.pb-prospect-card.deferred { border-left: 4px solid var(--pb-warning); }
.pb-prospect-card.skipped { border-left: 4px solid var(--pb-muted); opacity: 0.5; }

/* Single view */
.pb-single-nav { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; padding: 10px; background: var(--pb-light); border-radius: var(--pb-radius); }
.pb-single-detail { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
.pb-single-detail.full-width { grid-template-columns: 1fr; }
.pb-detail-section { background: var(--pb-white); border: 1px solid var(--pb-border); border-radius: var(--pb-radius); padding: 12px; }
.pb-detail-section h4 { margin: 0 0 8px 0; color: var(--pb-primary); font-size: 0.95em; border-bottom: 1px solid var(--pb-border); padding-bottom: 5px; }
.pb-detail-row { display: flex; margin-bottom: 4px; font-size: 0.9em; }
.pb-detail-label { font-weight: 600; min-width: 130px; color: var(--pb-muted); }
.pb-detail-value { flex: 1; }

/* Action bar */
.pb-action-bar { display: flex; gap: 8px; align-items: center; padding: 10px 15px; background: var(--pb-light); border-radius: var(--pb-radius); margin-bottom: 15px; flex-wrap: wrap; }

/* Progress */
.pb-progress { height: 6px; background: var(--pb-light); border-radius: 3px; overflow: hidden; margin: 4px 0 8px 0; }
.pb-progress-bar { height: 100%; background: var(--pb-success); transition: width 0.3s; border-radius: 4px; }

/* View toggle */
.pb-view-toggle { display: flex; gap: 0; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); overflow: hidden; }
.pb-view-btn { padding: 6px 12px; cursor: pointer; background: var(--pb-white); border: none; font-size: 0.85em; font-weight: 600; color: var(--pb-muted); transition: all 0.2s; }
.pb-view-btn.active { background: var(--pb-accent); color: var(--pb-white); }

/* Checkbox */
.pb-checkbox { display: inline-flex; align-items: center; gap: 5px; cursor: pointer; font-size: 0.9em; }
.pb-checkbox input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; }

/* Search results dropdown */
.pb-search-results { position: absolute; top: 100%; left: 0; right: 0; background: var(--pb-white); border: 1px solid var(--pb-border); border-top: none; border-radius: 0 0 var(--pb-radius) var(--pb-radius); max-height: 200px; overflow-y: auto; z-index: 100; box-shadow: var(--pb-shadow); }
.pb-search-item { padding: 8px 10px; cursor: pointer; font-size: 0.9em; border-bottom: 1px solid var(--pb-light); }
.pb-search-item:hover { background: rgba(52,152,219,0.05); }
.pb-search-wrap { position: relative; }

/* Loading */
.pb-loading { text-align: center; padding: 30px; color: var(--pb-muted); }
.pb-spin { display: inline-block; width: 20px; height: 20px; border: 3px solid var(--pb-light); border-top-color: var(--pb-accent); border-radius: 50%; animation: pb-spin 0.8s linear infinite; }
@keyframes pb-spin { to { transform: rotate(360deg); } }

/* Inline info-help icon. Use next to any label that needs a tooltip
   definition. Native title="..." attribute drives the tooltip text, so it
   works everywhere without extra JS. Click is a no-op; hover is the
   affordance. */
.pb-info {
    display: inline-flex;
    align-items: center; justify-content: center;
    width: 14px; height: 14px;
    background: rgba(0,0,0,0.10);
    color: var(--pb-muted);
    border-radius: 50%;
    font-size: 10px; font-weight: 800;
    margin-left: 4px;
    cursor: help;
    vertical-align: middle;
    transition: background 0.15s ease, color 0.15s ease;
}
.pb-info:hover { background: var(--pb-accent); color: #fff; }

/* Generic stat-card for grid-aligned metric rows (Group Management + workspace) */
.pb-stat-card {
    border: 1px solid var(--pb-border);
    border-radius: 10px;
    padding: 14px 16px;
    text-align: center;
    box-shadow: var(--pb-shadow);
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    min-width: 0; min-height: 86px;
}
.pb-stat-card-num   { font-size: 1.9em; font-weight: 800; line-height: 1.1; }
.pb-stat-card-label { font-size: 0.78em; font-weight: 700;
                      text-transform: uppercase; letter-spacing: 0.6px;
                      margin-top: 4px; }
@media (max-width: 1100px) {
    .pb-stat-grid-4 { grid-template-columns: repeat(2, minmax(0,1fr)) !important; }
}
@media (max-width: 700px) {
    .pb-stat-grid-4 { grid-template-columns: 1fr !important; }
}

/* Empty state */
.pb-empty { text-align: center; padding: 40px 20px; color: var(--pb-muted); }
.pb-empty-icon { font-size: 2.5em; margin-bottom: 10px; }

/* Dashboard empty-state get-started steps */
.pb-dash-empty-steps {
    display: flex; flex-direction: column; gap: 10px;
    max-width: 580px; margin: 8px auto 0; text-align: left;
}
.pb-dash-empty-step {
    background: var(--pb-white); border: 1px solid var(--pb-border);
    border-radius: var(--pb-radius); padding: 12px 14px;
    display: flex; gap: 12px; align-items: flex-start;
    box-shadow: var(--pb-shadow);
}
.pb-dash-empty-step-num {
    display: inline-flex; align-items: center; justify-content: center;
    width: 26px; height: 26px; flex: 0 0 26px;
    background: var(--pb-accent); color: #fff;
    border-radius: 50%; font-weight: 800; font-size: 0.85em;
}

/* Flex utils */
.pb-flex { display: flex; }
.pb-flex-between { display: flex; justify-content: space-between; align-items: center; }
.pb-flex-center { display: flex; align-items: center; }
.pb-gap-sm { gap: 6px; }
.pb-gap-md { gap: 12px; }
.pb-mt-sm { margin-top: 8px; }
.pb-mt-md { margin-top: 15px; }
.pb-mb-sm { margin-bottom: 8px; }
.pb-text-muted { color: var(--pb-muted); font-size: 0.85em; }
.pb-text-sm { font-size: 0.85em; }
.pb-bold { font-weight: 700; }

/* Photo */
.pb-photo { width: 50px; height: 50px; border-radius: 50%; object-fit: cover; border: 2px solid var(--pb-border); }
.pb-photo-lg { width: 100px; height: 100px; }

/* Chip/tag list */
.pb-chip { display: inline-flex; align-items: center; gap: 3px; padding: 1px 7px; border-radius: 10px; background: var(--pb-light); font-size: 0.75em; margin: 1px; }
.pb-chip-remove { cursor: pointer; color: var(--pb-danger); font-weight: 700; }

/* Filter bar */
.pb-filter-bar { display: flex; gap: 8px; align-items: center; padding: 6px 8px; background: var(--pb-light); border-radius: var(--pb-radius); margin-bottom: 6px; flex-wrap: wrap; }
.pb-filter-bar .pb-input { width: auto; min-width: 200px; }
.pb-filter-bar .pb-select { width: auto; min-width: 120px; }

/* Config list */
.pb-config-item { display: flex; justify-content: space-between; align-items: center; padding: 12px 15px; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); margin-bottom: 8px; background: var(--pb-white); transition: all 0.2s; }
.pb-config-item:hover { border-color: var(--pb-accent); box-shadow: var(--pb-shadow); }

/* Session list */
.pb-session-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 15px; border: 1px solid var(--pb-border); border-radius: var(--pb-radius); margin-bottom: 6px; }

/* Group Management */
.pb-grp-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 12px; }
.pb-grp-card { background: var(--pb-white); border: 1px solid var(--pb-border); border-radius: var(--pb-radius); padding: 15px; transition: all 0.2s; border-left: 4px solid var(--pb-accent); cursor: pointer; position: relative; }
.pb-grp-card:hover { box-shadow: var(--pb-shadow); border-color: var(--pb-accent); }
.pb-grp-card-title { font-weight: 700; font-size: 1.05em; color: var(--pb-primary); margin-bottom: 4px; }
.pb-grp-card-desc { font-size: 0.82em; color: var(--pb-muted); margin-bottom: 10px; min-height: 18px; }
.pb-grp-stats { display: flex; gap: 12px; flex-wrap: wrap; }
.pb-grp-stat { text-align: center; min-width: 55px; }
.pb-grp-stat-num { font-size: 1.3em; font-weight: 700; color: var(--pb-primary); line-height: 1.2; }
.pb-grp-stat-label { font-size: 0.65em; color: var(--pb-muted); text-transform: uppercase; letter-spacing: 0.3px; }
.pb-grp-badge-changes { position: absolute; top: 8px; right: 8px; background: var(--pb-danger); color: #fff; font-size: 0.7em; font-weight: 700; padding: 2px 6px; border-radius: 10px; }
.pb-grp-actions { display: flex; gap: 6px; margin-top: 10px; padding-top: 8px; border-top: 1px solid var(--pb-light); }
.pb-grp-detail-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; padding-bottom: 10px; border-bottom: 2px solid var(--pb-border); }
.pb-grp-detail-tabs { display: flex; gap: 0; border-bottom: 1px solid var(--pb-border); margin-bottom: 12px; }
.pb-grp-detail-tab { padding: 8px 16px; cursor: pointer; font-weight: 600; font-size: 0.9em; color: var(--pb-muted); border-bottom: 2px solid transparent; transition: all 0.2s; background: none; border-top: none; border-left: none; border-right: none; }
.pb-grp-detail-tab:hover { color: var(--pb-accent); }
.pb-grp-detail-tab.active { color: var(--pb-accent); border-bottom-color: var(--pb-accent); }
.pb-grp-prospect-row { transition: background 0.15s; }
.pb-grp-prospect-row:hover { background: rgba(52,152,219,0.04); }
.pb-grp-prospect-row.stale td:first-child { border-left: 3px solid var(--pb-warning); }
.pb-grp-prospect-row.dropped { opacity: 0.5; }
.pb-grp-prospect-row.dropped td:first-child { border-left: 3px solid var(--pb-muted); }
.pb-grp-timeline-item { display: flex; gap: 10px; padding: 8px 0; border-bottom: 1px solid var(--pb-light); font-size: 0.88em; }
.pb-grp-timeline-icon { width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; flex-shrink: 0; font-size: 0.9em; }
.pb-grp-change-item { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border-bottom: 1px solid var(--pb-light); font-size: 0.88em; }
.pb-grp-change-item.unacked { background: rgba(231,76,60,0.05); border-left: 3px solid var(--pb-danger); }
.pb-grp-color-swatches { display: flex; gap: 6px; flex-wrap: wrap; }
.pb-grp-color-swatch { width: 28px; height: 28px; border-radius: 50%; cursor: pointer; border: 2px solid transparent; transition: all 0.15s; }
.pb-grp-color-swatch:hover, .pb-grp-color-swatch.selected { border-color: var(--pb-primary); transform: scale(1.15); }
.pb-grp-no-config { text-align: center; padding: 40px 20px; color: var(--pb-muted); }

/* Dashboard cards + chart panels (Dashboard tab) */

/* Loading state: parent gets .pb-dash-loading while a refresh is in flight.
   Dims the content + animates the pill so users have an obvious "this is
   working" signal during the multi-second chart fetches. */
.pb-dash-content { transition: opacity 0.18s ease, filter 0.18s ease; }
.pb-dash-content.pb-dash-loading { opacity: 0.55; filter: grayscale(0.25); pointer-events: none; }
.pb-dash-loading-pill {
    display: inline-flex; align-items: center; gap: 8px;
    background: var(--pb-accent); color: #fff;
    padding: 5px 12px; border-radius: 999px;
    font-size: 0.78em; font-weight: 700;
    letter-spacing: 0.3px;
    margin-left: 12px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.18);
    animation: pbDashPillPulse 1.5s ease-in-out infinite;
}
.pb-dash-loading-spinner {
    width: 12px; height: 12px;
    border: 2px solid rgba(255,255,255,0.4);
    border-top-color: #fff;
    border-radius: 50%;
    animation: pbDashSpin 0.8s linear infinite;
}
@keyframes pbDashPillPulse { 0%,100% { opacity: 1; } 50% { opacity: 0.7; } }
@keyframes pbDashSpin { to { transform: rotate(360deg); } }

.pb-dash-card { background: var(--pb-white); border: 1px solid var(--pb-border);
                border-radius: 10px; padding: 14px 16px; box-shadow: var(--pb-shadow);
                display: flex; flex-direction: column; min-width: 0; }
.pb-dash-label { font-size: 0.78em; color: var(--pb-muted); text-transform: uppercase;
                 letter-spacing: 0.6px; font-weight: 700; }
.pb-dash-value { font-size: 2em; font-weight: 800; color: var(--pb-primary);
                 line-height: 1.1; margin: 6px 0 4px 0; overflow: hidden;
                 text-overflow: ellipsis; }
.pb-dash-sub   { font-size: 0.78em; color: var(--pb-muted); }
.pb-dash-sub.pb-dash-range-hint { color: var(--pb-accent); font-style: italic; }
.pb-dash-chart-panel { background: var(--pb-white); border: 1px solid var(--pb-border);
                       border-radius: 10px; padding: 14px 16px; box-shadow: var(--pb-shadow); }
.pb-dash-chart-title { font-weight: 700; color: var(--pb-primary); font-size: 0.95em;
                       margin-bottom: 8px; }
@media (max-width: 1100px) {
    #pb-dash-cards { grid-template-columns: repeat(3, minmax(0, 1fr)) !important; }
}
@media (max-width: 700px) {
    #pb-dash-cards { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }
}
</style>

<div class="pb-container">
    <!-- Auto-update banner — hidden by default, populated client-side
         when the version check finds a newer release on DisplayCache. -->
    <div id="appUpdateBanner" style="display:none;background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;align-items:center;gap:10px;"></div>
    <div class="pb-header">
        <div class="pb-flex-center pb-gap-md">
            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAAAXNSR0IArs4c6QAAAERlWElmTU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAAUKADAAQAAAABAAAAUAAAAAAx4ExPAAAmPklEQVR4Ae18B5hV9Z32+7+99za9d6YwA1NogiAoASEWjBqN0Zg1MZ9rNJtszCbrpqy76YmJa7KJyWaTiA1FQIo0pcNUpvd+Z+69c3uv5/tdEvchPs9uAIF9vu/h6HnmznDm/M95z6++7+8McGO7gcANBG4gcAOBGwjcQOAGAjcQuCIE8vMln/1sgwxoEF7i77NLPO66HPa/cjGPP1KRt7lFtjUZCqwJugI5iVBcFI1zAZGAm2QSQXtUJjzSbhWff/75kejPvlaQZ1Jza8zi5Eqtkl8iVIn5AZlidsolOPrd5+1vnDkzbrsuSP03i1x3AN9+cfUjKn7kO23vTZm7z7kQcUchYxykAkAjZSjMEECdJ4r3q0S93qTIVauO1TaypD6fJZBMcOCJGMS5MqSKdBjhy0ZfOxR58B++M3Dyv7m/a/5j/jVf4aIFXn1+/SMaRP/9599uU7Sf8iAVTqBADmSSA+vIgeX0OJOBFLa3x/m1wZilMhwpEIejsjfH46jgCDweB6ZmSLEkUvNeqOw+naGMv2LALt8+MeEPXbTUdfvIu14rfeGhlvKibPn3Xn6hiwUWopBIADOBJyfgUnQRjMBL0IcI7QEOqImn0BJIYGM0Ca2Uw2ycg9DIg8zMB1PwwFPwYe8PIdPqKPvc5vCj1+s+PrzOdQNw6615Xxxss2kx54dSCohoZZXoT+ClP09HgI4Yw4SEh0UWHsTkGzG62hD9vNbAwxHw8cMJhtc64pyAQGdacvUCITytEVTo4p8qLNSqP3xz1+P76wLgHTdX5KnkwjtO7epDvo4hlABkFPMEZHU82oV0FbNRYCtZ2mNiDl9UpaBOBxf6t1gSuDmQxJe1cTxZksTxqST8tgRSoSRkmQJE/IDaHy/7zFasvB6AfXiN6wLgXdtqPmYfcxvUXh9cZFbklRcA5BNA6T0NZBovLg7Y/Bymg+TOf8IPAnJ1HzjMz6QgmY5BluJYnI5PeZJgSQ5KSjo8ipF1+bGPf/jmrsf31wNAfrZZ8vHYoB1mjQwjXopltKqa3DBJ8W6eQn83WdEceHgpxud+EGPcN+cZHAS0gI6TUoJ5V8Lw6HkJPn9EirZBcHNegjeVAN+XgD6bh4CVQ6aYW5Wfn6+5HqBdvAY50rXdNq8vyzWIuSVBtw9Dzhh85KpRAqYnwaDVMGRn81GVK8DDtQJkqeNMyUvg2R8nkQiTB5NZcgSyWMe4L/6DBLlV2RiZlOPQUJx7s9fOMiIurMyMIbOIj5QoVfjIx33NX/8R9l3bO/rLs19zADfeUtCUKUxodgzM46g9BrmGhxIVsDoXKKCvCkECimQC5UoemJB8kyxTRFclowQjps9i+iynxKthPlQaB1GZLaagaIKf1aPt+AT3x30LTDjhwEO6FG/totR9X8f/IwA+/UhG3nItf0t0IbIiFU8p+AnM83m8Xm+UndvpEvbtOmpdSD+rQh227Ng/gLcWgmhulHLyuSgzk1mpKc6lY6GQXFRlZPCSRXrInZ0RDgNeYLcP6KRjlPT5/HwKlVS6lNlS0KjjlIEXoDRzWF0+xVZbInj5O0n8dB+Dqii+/vYmk/ntM/br1p3QI7/8bfszeZ9JeKPfOtcdkvROJhPecMqgIGspoUKijCwsT8ufE8qFh3ePJ2a6w/zH5pwh9f1llDFHg/AEObIksjACbpbi3AyVLgFyaUY7KIkIqAZUkCXyyH156Z9TluFoZwraqW7kyzkYzDyuoK6CVavGUan0w3mUwbfA0Jsn4bYfF9/72gHXKxfdFf+hh0pLbqvmrSzL5jWI+MiJJZhzyonDj7yA1x29vYGLjr3sj5cN4B++UvCgZyH8/Z+8Eej3B5Et5DMjx5jMk+K+4o9GDot4bH21lm1bbmH1TRaGWIpDjx/c+bkE81E5IqDMSeGN8uqfuo8SJZCvoKKaANVTK5JhYNDRQ5DKqGUjbxVRXSiiFo9HhWFSxUeAaptpFyWeqRjOjyXgj6WQI+JQG+Zw0+0CHODU43f9ffSB+27XibY1i1tKzYnbJMn40rAvKu4fi2KAyiAm4GNRsTSpNknO7RvhP/AvL4yOXDZyf/6FywLwqw9m6UvNwiM//A93tzqW3CRgKdV0NG0qbP9owHfrRRdBP1SsNKgkf6wQezPWWVKoNgkQJ0vq8XNchzPFQn6q46j1yCWQcgnATAIyi8Az6anbIPAklHmFFAeFZI18qrQFYtplPAjJlUH9cLp18UzEcKYjiu29HM56eNyyCgnLKZZhSb44UqHjJAFHGIe7wzjYl+g/Pc4ddYV5zXImWEzPJEjPcXRzI0+yboXKfWhEeOtv35rwXHT9l/zxspJIhl6+eMoVm0n6kpZGS0J1ykPRnf7jePjjxSuuXr2aReOiFpFUo3e6A3jBNgNufALV8hBWZTD2SAGPrIwPD/ntDMU8F7n1LMdxIV6KLZB5asE4BbW9khSBmCTro6qFi3KIUAnkpjufp2Mc1KHEOQlUhPxdVTw8LeeYJB6HzRbEieMe4fOjGDw7xe12Rfk76fGcAyYi2Sp9Ey8Re18KTl4oTdW828oOieUhQ3WZYgNd/8Vuf/Ht/I+fLwvAaDipDIX5vdnS1G1a6hiCdHPpHpYxNnfxKkNDk0UmU8aa8tJqQSoxg4BbgihfDBUVgL1ODu/PJSCirJOtinOlJrAqMw9matdkSgG5Fxm0kPoTOjEvHQvp+3ShLaCiUEaAq+mB5VJYSERTCFCvPDsfRfdQAn+Y4TAZVMDLz4EpO7P/iNV2C6K98xcCKyYuXJ5Ap+zibE5yYhSXyBNUzLP6MSt/e1ZmopoOuPYA2ucj4wker14vToXFxIyQV11o/lMp5F0MoNU6PvHQffeLE5yY51xwIqaVIl+mxs3yMBQyISUIPoJEFiyEEmzGEUXnVBzhVIKlqP8QCsi1qfuQkGuLKbbxye0JL8TICiOUZIJUR8ZiPHBJPnUvAih4YirGiVhQ5XJbb21mFaW5aB+emJbIg6G9ey++qv/6fCFsySlb5fKSknEBmwv6qSq/wu2yLHCilQ3KylNWmYSNy/hYYhGn4E8QMwLcT+u/RHvq7rvv5udZLI8W5GXU21wRLKkpx3v2MeREA7BolDBbZMjKU0OiIJToN6ORJIL+2IWvoUgCvmAMgTD9LJpAPEG2QqmYTwYpEvIJVAKMHoCMLFPGT4GfiBGQcfTP+PCbuTiyM9Qw6OVQzojzinXFdP4ztF+0LXiaBODyZHyOOiEKCTE4glHmmppJ9l501GV9vCwAX5uZCS8VGN+0pPhUriXvqFQm+dNhyo5gq6vV6i93e73/8tprr6Ueuvvu10tKy+4rL8lr6WjvgW9uCiU5YkilQhgMUihVEiRFYsqsEhgKKINIydzSbceF1oO6CrGMsjTFVq+bjhFRGUP9rkgElq59hHTJMSoYA1QgBn2IOt3ItTvRs89K2VrE6fUa5g2EVCe7B9IItn+AxrPPPsv74/d/8A162IIKRZLYIAZXkn94MinfOxzjWYELZesHh1/y18sCMH3WcxMOiiurd5SUnPp5kYx7YrU+TplVSC7Ie65ZpzKedvme/u1rr3nufWCbRcBXYGhsCiYWQZZMDo1GDG+ch7eOe9Hr4KAhSkZLYJZkKVBolqA4Qw5dlo7iHt1mGkw1tSrpzyIqMilTOb1xLLj98HsDEESCUElSKDTqINbKscGRYN1uD1TElem1qoRExKgMR9rMc2jP+uazz24zgn9zjpTUF2ohXUmBbSil+OHQzNQY/fsVb5cN4J9WOpr4g8/8zKfh1leoUuS+CXR5BUQSsKdatPKFYEHtGW8gli0VxeByzFxgmn81mSKLioMvZBzFNGZQi8EX043EItjT5ee8wRiTi3jc1jUlWFZpZHqlkJhnHmb9KbzXY+faemyUYcniqO5LUQwVSEWQk0XnZiuRrRNhkULJ+lvPcHKzBTv2Hs3v6BvcW1ZaEcstLMy1ZGZJFQo5XE4XNzs2yl6Z7Iff5fzNUCTSfcXI/fkXLwTUKz9JvuRLRXP/mkeWeNYlhD1CwZ0xrssfdz79zW8bRKR1fP+Xf+ScEQmWGN3sqSUGKpp5qKvPArFW6LF60TPqAJ+IhhKJCC5KEP/cEeEGYjpWV2yAuaAUR98/i7XcGNZkCFGUo0R5cwkMGQakBCLYrG6Mz7pxLq7CC+9FuZDHDs94G1u7bi237vZ7WFZZNfLyLDDSw6Kkx8U5hiClC3EswjpOnxrfvfOtn/3+97//Jd3/FXcjHxHAC9Czx4rEn80Rse/MhXh6B4F4iLiolds+SR2vAM6ZGYinWtmSRRbcxIujuC4Tv2qfx47zCxiZ9aJ6cT1UMhGXa+1mD5RqsX8+gkFnFA/fUovWmTAVyt34QpkSh+0RHCIpILOkACZFCncvL8Jda6oQGB6H1+PDbb8ZhJNg+Mo/PIPFN93CnTnTwVqPHebGh4eY1WpHKExEBllhXo4J69csw/2fvA9avR6vvvzyqee+9a0vTE5O/le8vByDokDz0bdWd7JNIRK9m0nVSlswlbn55nzhUE8/G54L4uElBtakisKdEkIZDOE3749i3hPDtk2rUFiag4ayDO4zaysxRd2UlOPYQY8E24rleOBzG6DLt+DMyUFkkNVOGgsQkevAVxg4TmLA4TNj6DhyDi0WCfvaqx0YiUjx7R+/AIk5D//5wk/YSz97gWs/dZ75J2exJmpHdsSLMWJzB6x+dPSOYc/bb3EBv5s98Njnc0oqKrZM9veds87PT14uGlcFwPSig/74/AlX/JVP3lcvX2+RrmihVmI0yLGa8nxI3TZMhjjU64VoKtRiY7EGSzIY7lyei5sK5EzjGGVrc8RMrVHgwIAHa80ilNUUIqLVYl+rFXeuyMLdFQpsrc/EHTdXsftvqcRddXrkEcGwr2Oa/bIvgqf+8bvcsD3MXnvxBzj4zgkIY3FWxvlxC/Piua+qsDqPQ4x4A3cigUBSwGXkl6CrewBnjx3Bpgc+pTRnZK4fbDu33+X1Oi4HxKsG4J8XTX3hnrqNB3YNtMRCcRQV53IeJkVBxMWoj+fWlZlZQbYKkkwTohIFqW9CJImzlxkMkOlVmLSFMDhkxe0VJhjLcqAwqNE1MsHd2VzI+CIpgsR3CWJhiBBnYqmUGcwa9uhLJ/GJzz7FWaNydvadP6D1ZA80lPWrU17cJo5gC7E3VV+pgU4gQO45K0kAEtSuXYvnnn8OCws+7Nl/EkOdZ1jdhrtUJo08//jRw6/TvVABemnb1QYQH7+5/GMlkVDTgjuMIWpYEwo1qiXE08fDbFKTx1WWZhINr2fqsgKoiwpIGMoAMTpwuKLcz/f04lYznxXlG6EsyiESgWrMVIq90zaFdTcTCPmZUGVYIMnKgshiwEuvHkOnS4qa1XewwfYTOPbuEYg5ahGTQaxAFPd/QoCm53PAq6wBKzPBtMIHaasXBoWalW/aivLaavbir99gfR1tUKmkyKtaWqxKBs71Dw4OXRp8VGVd6oGXelyUL5heusgEvTCFt4972KO3mjnOEUCTmhLCQDd7bkDBGbM0nNqoYVKtliN84Hc4qdiew6YcASuWaqhEoWKNCul073bb0kK4qVD+9kvvcQXZJqbRaDip0QTPghs/2tGKzZ95ijlIUx85fxYckbIqLgI9Zf/CNKM9Rh3NGz7IHnCBR2qVZ7uThexJ2B3d+OFjT7JucT6i1EklUyIc27OTy65sYrmL6h/Azp276QLSjNtf3a46gEO24IkNZZqko3WSX1BUwPGp2+ALhbDZQ2hWybDJKGQpkjajxCzH+U6GsJ/YsADn5gfJ8jIx3jOLxZVUQEeJbVVrwYUDuMWshCcRYMG4A5Ek3TAp73OnO5nNG0PJ4iaub3gME5PzTEoTCyqOhHjywBixOxNdHAucdaNh8RiYI87mt4dgJWrMxgvA6TkFa6wHYhJowuQBXo8X1rF+GLUZy7P1+swZp3P2r6JHB1CZf3W3f/3hWJuDYz0pYlNiPDETCvikPnLU19L8CxW/SrKubLkYZXIBqoRx5CSjiPoiLJASXCgz/ME4RGlRJE3zqHVEbBHxatHBmaQulsJCg0bAlucpGS8QhD4jk+KnicV9DhYPp9AoTKCSWB4VFVDUbDAp3V0u8Swcz2dLFsXP5dzGSxCNBXk4jMzEAnQBK8zRBbA4BWiCwTkzBoFCY6murym+VFSuugUCI9GhYOUr9atKa3//nyMUW25CMp64wEwvuIOQEimAVBwSH1H51IUoDDJI5SKcaJ/C8LgdKnJd56QdRgIx2dkJ37QV/VQwzxIds7IqA21HumBWy0mdc0CqLkRXzzjmZ23UIpN1myPIoNmQkQUCkMCrzifw7xH3eLPYJu2y6alwTPtQg9r/75I9SX6rL44msliassEURwkqFGUhnwdMJOVpjRn6/0UAgZePLby65tNlX92yNKaUagywU/mvJBZlhipdulDwJXwSjOIka0oQGQth3hfDyYU4Ts06IOOlcNYVxIYCLf1eFC/1ORCmMPDiYy1I0kCNoiQTu46OwEasjS8VxejkPLE1RDbIVNi5EMLiRBDVpPSVE0ehr+PBZ5a/pF0+R/Udwy7u7u1bR3d/o6I7nB8fJ/qACm8heYeSYEwTtULqu80mLWaEQmI3Lm27BhYI9E1HSZzjh4mSUgplcs6XFDAlOUkWTRN5qCP4LnUiMWJchORyMiL+AmRdfgrkBWYD6cYxbLdGsds6DUbAqdVG3NlYgoUIBfojs5DLFehkGpz0OeCh4nh8egE1GjEnlMpZUJQJP1FhYqmXZmuIsZHy4WOCiQ+g2MZejfi/p5wmuTQ/RhApSXPWkqhPiiqc5CU5Fj1W1+Zi189JQrzE7ZoAONTW5ppzlZ1RyAWbfb4QhOYs9I6PY0uVFm6FgBpPIf3Ho4QAGEjgURPlLCXyIEG9KlEONA9DLk7UlVouQ45BDimNGXWcn6KIzcegLYy+OS8RETJEic5KUGEslKjpWOIlpWVIOJOICyVUGs0Q708CfVBUR1i8mcYjsGeRmfdeMivNepMTIEUJRET1YYwSE5/WVeotGBmZCk9PjF5SAkmf86rXgemT0saVF6rbZmcDObXNq8pc1nl2pm/sAhFaV2HEoJ3G24jB4RIhbtTmpuxJIxrEPpuJoM0nZa4xQ8rliOMsUychJifBkY4Bl8fPOsas6Ju0EZlKckI8CV48CFF2JUyZBYznHYNEXQSLSgU58YU6qgOV9HvaMsFSnsIUdo1SktfF/3lROLp4oD2JBRLDRhNCnKd9ktpMjVYBZXEjBvqH+iZOHvyRJxIh1eWvb9cKQBw4NeWsXLV5tra88FNhn583ODaNEXfkgi6SZyRZh57d97YVM5c3xLWSrvHIZ7Zgy4N3Yulta7Hozk9AQQVzy0OfRVlDA2u8/9OYoWmsU8fPsBcfbyFClQONB9JQZhIj5GyagkZucZEWI12d7ImnHoePhKxoyAttfRQ/3hsV7jsd3RBJsTt6OqP5KZJIc7MEaJsQYiQpREechDFivvMal0CRvQiJhcEXT509uf+vQ/enI64ZgOnTe2ZmIs1NS+4tKchWj49OYszmxRDJbjlqEaZJ1rx5kQFbG7NZIT/F9p/qx4kRO+kbSUiMWUiSnjnvibAde4/h5V9vZ+LxbvzgsRamN2jwyrFZ3JSrpnPFiA4bQdaiJUyizqI4Ow8uEIeaCIUsXwdNutrxi26SP1M0gESE7AK56ukxhnwVDynqhw95eUxAhby6kCivxs0QcaHJqTMHnpq12S5Z4rymALqDwWBJQYG6vKx0ddojhsen4aV5jgDNp1HkQgGVMFkGqgmLjWgxSlmx28aEM6NsrO0km+hsRaz/HMq947g9T4jNGxcxKdWDrV029FljcNN4sI+EpRCdd3p6FEltCZVMGthnuzHf2soVkeb8XF+QOQm4GGXgGMXc9O4hptvqpfEEKpN6PXEmyTNBvXwzUjxxAt6Rvz10cP/7l2p96eOuKYDpBY6dOXemprpu07LGWsv8nB1jMzZEySJWNdchQbUbi1FBTOUMU1D7RnzdIprAal5SgCbKhjVVucgmSktApIJEp0aESpf2fg/VjtlwULaNUByUKDTgxfyYnZuCHRao9BmcOV+G4/SwzoW1jJOawUk04IRKmuBSgRPIMUeSaIJ6ZtnSWohr17JIXMBpea5vvPG7X/wbXfIltXDpe0tv1xxAWoPrHZq8v3lpQ65MIiNrmaPek0NjYx2qKkswTzNIWTSNYKVW75WT8zQCl8AMDU9SaET/hPtCgU1dGXIKjJgYd2PCI4YhrwAnu4aJ/RagtrIMBUWF7fv373qWnwg0O5JKuSOpZXlLGlCRK6cJrxiTUucjpYyu0RAzlKNBc0s1sm/eyKLKYpaIxzwavusrr//2hR/TtVJhc3nbNQfwl7/65T12Z+CJ4eFRXklhwYVC2k2iUJR04eaWeiiVckQWHHRjSihpZOMUFc52TxReXxhiyqJGcvOVq8opZjHYw3IYihchQm55uq0PFosZpZUlyakFz+fOdXa+HPbYdiU8U+ZEPFhmMZn4SxqWklKnpNpRyWoW16FlRQtqli6DpXARuX48kgjM7Y7aux/f/frLbxBsl2V5H8B8TerAD05OX0V6S87TQZp+6zqxiyTcFMrz85jJH8DcvBMz8x7UlxfDT/2rdWEUleXZeFyvwOAsachF2TBaiILK0SMVDNANi6HJ1xMZKsFI+wgV2GroTHocbOu2e4LBM39ecyDsXfiCXBu8JRyHpjRXg57OUzh65Lg/v6rJVZhFY2Iy3iyLuc+650b27dy5h0Y+cEnlykX39BcfrymAP/3p926RKE0N03M2PHRzDvv9oZOk3AmQaTRinuZ75ygrR8ryoC8qR5AjcV3CwZJNZEMJXZZSS7uBypEIwu4k+JZMIlIlENCIq5voKyG1dS+/c5DLrV2Zcf/9d/3LgQMHvprV2Jjw9A5lepxO2dG3/xPeqW4iKRTghReeyJrfv8d63pZ6a5aeHnE9f4HCR/jmmgKYnV/+RIwvZVKlGiaxDSu0Sew9dhw3L6mDkuLhZFrsodJCIBBDWUqkp8+OuCATCa8VQgr0SHppGj8MoT4bQp2FpvPdGByZJaKGx3WPjUMnDbCpsSG89B/uTz//t6s2ttTlpe79x4nhW5dqcM+GOjiFRuw4NMzJolr/3pMdl0XVXyqm1ywGbtmy5ePG4oYvTVh9fJfLC8fgWWwximH3R9FBrZiQZMl0MrERAWAwm6EngVwklyNEcy98pQ4JJqZMKQBfYwZTWRAiQd6x4Edb9wh++/Z+xryT7JmNufjFO51YX5bClwoFijeOjig/sSo771NmIb/QpERlSzm2rq9ht62tW79q+WLt2a6Zc243vVt2FbdrAaDk9oef+Vo0Fv9+Ve1S6cyMA34aYevo7CCWJIR6tQBjRGuN0h4jgALEzkjJGkUiCbQEosdLyUOlJS5RjnCCj6hQDS9FqZGROfQOTuB3bx/A3HAffn0PiU5UU24nw/q7ZSbsPz4LjUWB5YkIzvXYMUBz0wLfAtRxH9RKoaS8InvFhnVL1qvM2UePH293Xi0MryqAqzfeueQzT37tJa0572Hn3LRo7ZoVFOdcUJBobsnJwuEj72J5hoLmnWnZaAijRMvPub1Ix8g5u+/CSzUCIgLSjMrgKLkzT0RsixMnz/bgyMk2/OrVneCsY/jW2lysbc7ArlPzODRDPKEaXKcjymIkHcxSAZ2xvAq6ZfXcBM3t9E+6mZ20Y6l9EvkWQUY4FOa9vKf9nasF4FWJgQYK949/96dfUpgKviiSqpXemXma6UtSlnUhP9uAd9/vxp0b1qH97Am8PnICjyzOwmqikEwkoJ93z6KzfZrrHxlhx1o7sLS6ilhoIhe0GpohJPaFWrWegQGE3HYsUyZx1zIL1m2kzO304uRUFM2bPo1dXa8iRZRYVp6W+1iFjpWtLAdV4DSdxNH/1ML5Igh5KZ6SxqLw2POuFnjp83xkC7z14/e2PPKVb/zBlFN2/7nuGXEyHsPyuny8d/gwFDozF4kzlmPRonPAik9sux0dxB6PDvZhVbGeXuciEYjPg0FIeYEYlBTpI6OzVoxQgjjf1wWP04GxgR7kJtxYq+Nhc6UeK7eSOkfU1XZ6bWxPspp78ot/i/MzNlausuPJIjUL2mjwiF4ykUWonQ24wAKeNKMDjUJM/KIc2pLifKlWbzt6oueKJhE+DP5HAZB/7+Nf/eKSFetesnlShVaq6xqrc+Ah/i9N3Y8MDtDTTw8CKdi65eU4dm4Yi0qycMvGj3H7Bh043d7O1lbokEttnJQmNNMDbRrqVlWxADJ5EdpjUAYcqKPGf12WBMuqM9By51JoKYYePdCL7w8o8IWvfxNrm8rY3n0HuQd1NuayBtHhT6CXOMOOPuqZu6YQnrdBEPZBEiLucGwCMtssf3fnVNmpXmt6JuaS9d8PA/fB91cEYFX96uKPPfjYLwrL6p4cHrWJ5uYcNCVlhEmngD8Qwvun+uCkMbSQZxrajGIWJtcpyjWhZ9hKw5MJ9k9//zcYS6rw8uEz9MoDx5qKdTASw6klxcxAE1caGqA00vhrNRENjdly1LSUovbuVTR0zse//eZ9/HhAjc9+80fcJzc14/zADPvVb7az6PwYF1SI2R131GHF3etRsm45BkW53EFfHo2GkIhUmscEBiOcXUP4dbtH8/jXvpW17509hwgImnu98u2KAMxd1Pio2lj8xPjEPAGnR0NtESYp2x4/009tkxTN9UWYoOA/3HUSa9evYydbR7CZhPEwjYQOT9iQor6srLScBYUGnLSGGbcwjjVVNI2ll8JEQOYSC11TYUY5EQtl6+qRd1P9hWnUn/zyEE4rV3EyUwHGhofhjkrY4IQdrcd2YVMJn60ja85S0RRrKkRvNyWokzEyRfk6LC2i2UwNyVUGMzwkd7ze6uZ9+R+/Vb98WXP+a6++uofgu2JLvCIA5WO9rROBxEx+UWFj5aJS+cTUPOm+brIQISLhKOJJ0mZ1Wgz3dDMZEZhSdRaN8EYhpu6hjKYOOnonsfuV3+KtN99m9376Ee7Y6Dykk6Msg2RPI03dlzYXo6ClCsZFRTS5YELMtgD7wZP4fbsLt3/+q9j+xkF2+tgR5gzE0D82h/XyHmwj4Bw0oT9lpwFMyvxSlwOdwwvcaz0hbLLMMYFzHowYGweRsfsGY9xdd93NgkxVXVNTKT3wzp53r9QGrwhAV7pHsI+18RL+A7OOsJnmjMuLivNYlLi5WVLW0qNk9aQppovi9w/s5LZuu4P1DtH7mBQT06Mdo0N92LX7INS5TbBOEd+3bRt27H8H6ymDikk0YqR9CCWkmpAuQl0HvKfaMUagH43qkFmzmrXRAFKMqHjbeBv8pGuURMchoeb3PPF9VrWWGxdr2ZlRL351yosiiRMbS3iMI+tPUnUwTJLA3hnSpzOr4fGHWH5+/rKayoKBg+8e7LkSEK8IwA8Wss/N2iZ6T++0ZBo7Z2fdOWqdLmfF8jrMzS3gbOsAqqrLYZ0cZ0GvndNlFTOP10+0lR/nzx6mjkSAxQ0NkKvUjOPJmLGkErs6u7B4sZmV1uZfmI1mqSTGOwfx67c68OaCFLqG9VwwoWBBEtU5IXUtXlL3PJPooYe2tELNbVuex7Y88jG2sjoTDTXZWLmyCNWFRDpkZjABsT48onHPjS3gmMuI2sWNTEweI0PQ2dN9fsepEycGPrivy/n6kQD880LJycHu/qn+0zuyc3KHp+dcmUq1JstCwz/d1HZJ1WacOvQ2q1+yGI0NVXjveCcGOo6BE2ciI8OIFU1VKCRWeNXyZmSU1OKnL/yO3bvYAH6UKC0qtB98qQ85d/0dqjbcR7SJlolJmHe5fLBSjRlP8eG19uHeBh0eNIlZms5XkVTKj/ohIs3YJEoyi4RHBk26MY04MFLxdrTPYhgV7NY1zeTqg/te/OnPHv7jH3535HJAu/jYqwHgB+eLjPS0dgy1v/eGKSNzkKxNl5lpyuGL5MztjcI324mqJSvRf74Nw8NTxMnV0LtwQswSI0OlD7eqoQCrW+rYrqOtUDpGUEYa8vO7eyBtuh8Ny9fTy9oSjJH7ne0ah4BAFNDUvoLeiY3QVMFq9QL9+ZQwdg65cK5jCikH/UWPZIy4G5q/8Xjg6hxClCoFCQ08/fCwFTl1K0ask8PPPPOlJ782MTFKeumVbzSAcm02Hb3FWrJm04oYT3uP1pSz4fzpQ2aTQUlDPH7UNK/HNI33ZtKs88Z1i9E/ZOWGRmfZp+5Zg2DQw+345//Dvtyowjd6pLjl4a/D7QnApJWjZ3AW65aXEfMvxZ73etB64gwCwTC8wxRPaR5QQy5fTsOZY4EkR3wCGiwytpgm//NIYDdS/TgnFQ7/zd6Fn0ST4reGZmYuWfv9nxC6mhb4F+vQ62zR2Ymh4fnx8+/wI65dMkPuaDSekOZVNhmXNjcIy8ryEaTAH6LgfueGeiakVzNdxEJLFFo2QdMCP3vrCD719D9hykkkbK6BGGsjEpRcDFoZ3tx7DpPTjgsTVWLSRKLUVy83hDhzLMYeXaWnqQTG3psMEemVgpH2BJdqP++Nfv877zme7pzzHnb6fLTC1dmuSi/8Vy4lNjLS2wf09jU1Nf3aO9dTOdSFVUqdadWSRbl1xTnmrDZ6m72kNB9Wq4O6FQv2J0P06urdkBoL0ZTJIUMnx+DkArkgD0dPDV74+zINNQXop1dg01pH+8T76WqYEeuPacr8C644TFLB+E0mye4ub+qdJ7rsx+ka00TqVd+umQtfwpVK6+rqzCtX0sxVTl41J1JXvX92uFStFGWdP/a2+rEv/5NcKlfTn9JiGByzYcu6Ory6r4NaOTm6+qfotS8+2tv7oJRLuJmhM97I+Kl4IXkqiScc4/OHaBBpS+uCf/ASruMjHfK/CeCHLzwdTmQSiURXWVurzsirMpXk6I1We0BtNGu1Ir5I2DPh5pNCl5qatMb0JrVrfmrSTe/MLcRioTn7SH+EpmVylDS/OZNko6FQaO7DC9z4/gYCNxC4gcANBG4gcAOBGwjcQOAGAjcQuIHA/x8I/F+WU2CmKbQwyAAAAABJRU5ErkJggg==" width="80" height="80" style="flex-shrink:0;border-radius:10px;">
            <div>
                <h2 style="margin:0;">Prospector</h2>
                <div class="pb-text-muted pb-text-sm" style="font-style:italic;">Finding people for Christ</div>
            </div>
        </div>
        <div class="pb-btn-group">
            <button class="pb-btn pb-btn-sm" onclick="pbSwitchTab('sessions')">Activity Log</button>
        </div>
    </div>

    <!-- TABS -->
    <div class="pb-tabs">
        <div class="pb-tab active" data-tab="dashboard" onclick="pbSwitchTab('dashboard')">Dashboard</div>
        <div class="pb-tab" data-tab="configs" onclick="pbSwitchTab('configs')">Prospect Management</div>
        <div class="pb-tab" data-tab="groups" onclick="pbSwitchTab('groups')">Group Management</div>
        <div class="pb-tab" data-tab="senders" onclick="pbSwitchTab('senders')">Prospect Sender</div>
        <div class="pb-tab" data-tab="sessions" onclick="pbSwitchTab('sessions')">Activity Log</div>
        <div class="pb-tab" data-tab="settings" onclick="pbSwitchTab('settings')">Settings</div>
        <div class="pb-tab" data-tab="help" onclick="pbSwitchTab('help')">Help</div>
    </div>

    <!-- TAB: DASHBOARD (outreach-health KPIs + Chart.js trends) -->
    <div id="pb-tab-dashboard" class="pb-tab-content active">
        <div class="pb-tab-intro" data-intro-id="dashboard" style="display:none;">
            <span class="pb-tab-intro-icon">&#128202;</span>
            <div class="pb-tab-intro-body">
                <strong>Dashboard:</strong> outreach health at a glance. A <em>touch</em> is a <strong>Note</strong> (not a pending Task) tagged with one of your <a class="pb-tab-intro-link" onclick="pbSwitchTab('settings')">Contact Method Keywords</a> (Phone Call, Email, Visit, etc.). A <em>conversion</em> is a person's FIRST converted-type attendance in an involvement during the window, AND they had previously attended that same involvement with a non-converted attendance type (a real prospect&nbsp;&rarr;&nbsp;member transition in the Attend table). Cards show point-in-time totals and rolling windows. The date range below affects rolling cards (touches, new conversions, conversion rate); "Active Prospects" and "Overdue" are always live counts. <strong>Scope:</strong> all metrics are limited to involvements covered by the groups you have configured in <a class="pb-tab-intro-link" onclick="pbSwitchTab('groups')">Group Management</a>. People or involvements outside those groups are invisible here.
                <a class="pb-tab-intro-link" onclick="pbSwitchTab('help')">Open the full Help &raquo;</a>
            </div>
            <button class="pb-tab-intro-close" onclick="pbDismissIntro('dashboard')" title="Hide this tip">&times;</button>
        </div>

        <div id="pb-dash-rangebar" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;
                                          padding:10px 12px;background:var(--pb-light);
                                          border:1px solid var(--pb-border);border-radius:8px;
                                          margin-bottom:14px;">
            <label style="font-size:0.85em;color:var(--pb-muted);font-weight:600;">Start</label>
            <input type="date" id="pb-dash-start" class="pb-input" style="width:auto;">
            <label style="font-size:0.85em;color:var(--pb-muted);font-weight:600;">End</label>
            <input type="date" id="pb-dash-end" class="pb-input" style="width:auto;">
            <button class="pb-btn pb-btn-primary pb-btn-sm" onclick="pbDashRefresh()">Refresh</button>
            <span id="pb-dash-loading-pill" class="pb-dash-loading-pill" style="display:none;">
              <span class="pb-dash-loading-spinner"></span> Refreshing&hellip;
            </span>
            <span id="pb-dash-asof" style="margin-left:auto;font-size:0.8em;color:var(--pb-muted);"></span>
        </div>

        <!-- Empty state for first-time use: hidden by default; shown by
             pbDashRenderKpis when the server reports hasGroups=false. -->
        <div id="pb-dash-empty" class="pb-empty pb-dash-empty" style="display:none;">
            <div class="pb-empty-icon" style="font-size:3em;">&#127919;</div>
            <h3 style="margin:8px 0 4px;color:var(--pb-primary);">Your dashboard is ready. It's just a little lonely.</h3>
            <p style="max-width:560px;margin:6px auto 18px;line-height:1.5;">
                Once you build a <b>Group</b>, this is where you'll see your outreach health:
                who's overdue for contact, how many people are converting in each involvement,
                and your weekly touch pace. Two minutes to set up, lifetime to use.
            </p>
            <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap;margin-bottom:18px;">
                <button class="pb-btn pb-btn-primary" onclick="pbSwitchTab('groups')">Build your first group</button>
                <button class="pb-btn pb-btn-secondary" onclick="pbSwitchTab('help')">Read the quickstart</button>
            </div>
            <div class="pb-dash-empty-steps">
                <div class="pb-dash-empty-step"><span class="pb-dash-empty-step-num">1</span><span><b>Create a Group.</b> Pick the involvements (or a saved query) that hold your prospects.</span></div>
                <div class="pb-dash-empty-step"><span class="pb-dash-empty-step-num">2</span><span><b>Set Contact Methods.</b> Tell PB which TaskNote keywords count as a real touch (Phone, Email, Visit, etc.).</span></div>
                <div class="pb-dash-empty-step"><span class="pb-dash-empty-step-num">3</span><span><b>Come back here.</b> Cards and trends fill in automatically. No batch job. No setup.</span></div>
            </div>
        </div>

        <div id="pb-dash-content" class="pb-dash-content">
        <div id="pb-dash-cards"
             style="display:grid;grid-template-columns:repeat(6, minmax(0, 1fr));gap:12px;margin-bottom:18px;">
            <div class="pb-dash-card" data-key="activeProspects"><div class="pb-dash-label">Active Prospects <span class="pb-info" title="Distinct people currently in prospect status across the involvements in your Group Management groups. The sub-text shows the per-involvement membership count (a person who is a prospect in 2 involvements counts twice in memberships). The conversion-rate denominator uses the per-involvement count, not the distinct-people count.">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">Distinct people in prospect status</div></div>
            <div class="pb-dash-card" data-key="touches7d"><div class="pb-dash-label">Touches Sent (7d) <span class="pb-info" title="Count of Notes (not pending Tasks) tagged with one of your Contact Method Keywords (Phone Call, Email, Visit, etc.) in the last 7 days. Configure the keywords in Settings.">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">Contact-method touches, 7 days</div></div>
            <div class="pb-dash-card" data-key="touches30d"><div class="pb-dash-label">Touches Sent (30d) <span class="pb-info" title="Same definition as Touches Sent (7d), but over the last 30 days. Tasks (pending intentions) are not counted, only Notes (records of completed contact).">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">Contact-method touches, 30 days</div></div>
            <div class="pb-dash-card" data-key="newConversions30d"><div class="pb-dash-label">New Conversions (30d) <span class="pb-info" title="Distinct people whose FIRST converted-type attendance in an involvement happened in the last 30 days, AND who had previously attended that same involvement with a non-converted attendance type. The transition from prospect to member, detected via the Attend table.">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">Distinct people converted</div></div>
            <div class="pb-dash-card" data-key="overdueCount"><div class="pb-dash-label">Overdue <span class="pb-info" title="Active prospects who have no contact-method Note in the last 14 days. Tasks do not reset the clock; only Notes (records of completed contact) do.">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">No touch in <span id="pb-dash-overdue-days">14</span>+ days</div></div>
            <div class="pb-dash-card" data-key="conversionRate90d"><div class="pb-dash-label">Conversion Rate (90d) <span class="pb-info" title="Per-involvement conversion rate. Numerator: (Person, Involvement) pairs where the conversion happened in the last 90 days. Denominator: (Person, Involvement) prospect pairs whose membership overlapped the last 90 days. Each involvement is its own conversion scenario.">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">Per involvement (Person &times; Involvement pair)</div></div>
            <div class="pb-dash-card" data-key="avgOrgsBeforeConv90d"><div class="pb-dash-label">Avg Involvements Before Conversion (90d) <span class="pb-info" title="Of the people who converted in the last 90 days, the average number of distinct involvements they were a prospect in. A behavioral signal of how much pipeline exploration happens before landing as a member.">?</span></div><div class="pb-dash-value">&mdash;</div><div class="pb-dash-sub">Involvements tried as prospect</div></div>
        </div>

        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;">
            <div class="pb-dash-chart-panel">
                <div class="pb-dash-chart-title">Weekly Touches Sent: Notes vs Tasks (last 13 weeks)</div>
                <div style="position:relative;height:260px;"><canvas id="pb-dash-chart-touches"></canvas></div>
            </div>
            <div class="pb-dash-chart-panel">
                <div class="pb-dash-chart-title">Weekly Conversions (last 13 weeks)</div>
                <div style="position:relative;height:260px;"><canvas id="pb-dash-chart-conversions"></canvas></div>
            </div>
        </div>

        <div class="pb-dash-chart-panel">
            <div class="pb-dash-chart-title" style="display:flex;align-items:center;justify-content:space-between;">
                <span>Conversion Rate by Group: 90-Day Rolling (last 6 months) <span class="pb-info" title="Each point is the per-involvement conversion rate over the 90-day window ending at that month-end (current month uses today). Same math as the Conversion Rate card, so the latest point should match the card value.">?</span></span>
                <button id="pb-dash-groups-toggle-details" type="button" onclick="pbDashToggleGroupDetails()" style="font-size:12px;font-weight:500;padding:3px 10px;border:1px solid #ccc;border-radius:4px;background:#fff;cursor:pointer;">Show numbers</button>
            </div>
            <div style="position:relative;height:340px;"><canvas id="pb-dash-chart-groups"></canvas></div>
            <div id="pb-dash-groups-details" style="display:none;margin-top:12px;font-size:12px;overflow-x:auto;"></div>
        </div>
        </div><!-- /#pb-dash-content -->
    </div>

    <!-- TAB: CONFIGURATIONS (combined with workspace) -->
    <div id="pb-tab-configs" class="pb-tab-content">
        <div class="pb-tab-intro" data-intro-id="configs" style="display:none;">
            <span class="pb-tab-intro-icon">&#128203;</span>
            <div class="pb-tab-intro-body">
                <strong>Prospect Management:</strong> build named configurations that pull prospects out of an involvement, tag, or saved query. Each configuration becomes a workspace your team works through one person at a time (List / Single / Batch view). Configurations are reusable: build once, run as often as you want.
                <a class="pb-tab-intro-link" onclick="pbSwitchTab('help')">Open the full Help &raquo;</a>
            </div>
            <button class="pb-tab-intro-close" onclick="pbDismissIntro('configs')" title="Hide this tip">&times;</button>
        </div>
        <!-- Config overview (card list) -->
        <div id="pb-configs-overview">
            <div class="pb-flex-between pb-mb-sm">
                <div class="pb-text-muted">Create and manage prospect configurations</div>
                <button class="pb-btn pb-btn-primary" onclick="pbShowConfigModal()">+ New Configuration</button>
            </div>
            <div id="pb-config-list">
                <div class="pb-loading"><span class="pb-spin"></span> Loading configurations...</div>
            </div>
        </div>

        <!-- Workspace (shows when config is opened) -->
        <div id="pb-configs-workspace" style="display:none;">
        <div id="pb-workspace-empty" class="pb-empty" style="display:none;">
            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHgAAAB4CAYAAAA5ZDbSAAAAAXNSR0IArs4c6QAAAERlWElmTU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAAeKADAAQAAAABAAAAeAAAAAAI4lXuAAAtaklEQVR4Ae19B3hVVbr2u/Y+Pf2kh0ASaggthI5SBBWkiNgQ7F3HsVzLWOaKOtYZddQZG46FGcWKBRWRohSlt9AJoSWk9+ScnH72+t91EMfrdZ57n/8fQvyfs2Cfss8ua3/v+vq3VoBoi1IgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFIgSoEoBaIUiFLgpFBAnJSrdsBFJSCsFkufcBgFUsocw5AmaMKjC+2YrssD3boFjhw8CP/PuzJ+PExb1ll6BcLID/E8XsZmQAQsQpZbdOy5OxAoeRgwfn5e9HsHUaBnT1h13XYtYFkLWL3cpBBWqWtWqXFT33/Yvwua9QWTyXZGfDycZrO5SNMsj/G3rdx8keM0m9RMNgm+R77DwutZ1pt121UXAXoHPdJJvc2vioNtNluezy/fghTjshwShYkGesQYiDcBycRV0wBXSKDMp6G4WcO+FqA9GKFfNV9TrCaY850Co7oIDEjVkB6no2uGHcKk4ViLgfWVIXy6y49DtUEIIb+MiQlc43aj/qQicJIv/qsBOAFIahWWlTZdG3RJtzB62cJo8Ui4AkAMAe6fIqEehmIWMRbAbBEICE3ubtNEiUuga4zEiGQDqWYJEw8MBA14yccDBzqR1tUBb2sQ5dVe2FJj8NrOIJ5c5gZk+NsuMjitAiBn/zrbr0YMBTXLHVJql96ZH0J3cxi7yVeeMBCitiQmSLJRcRI4yloQO0hu6xp0sbpeR61P4GgrUO8VKIg14A/KyDFBnq9zRKSn8mQhsGVnM2rK3Jg5wA6H04LVpcE8r24+asjQtl8nvACFWudv6emIMQxxbf8kieFJYextJDDsuUZQNIKaoNTuTxp3o51ov1um44p8h3xkfJL8/dQ0ubrZgoNuAUrkyEAQfG9rC4AGGjleQ3YXB8E3sHl7E24uNMvuaWYEw8Zv+lEw/OTyv6qPvwqAWxqs48liPSdlGGig5AyEFbDiR5HsoIhWZq96GBtlklUBSMyVeJrkahXnVNWJ6U5DJNkE3NTRVv6gjtO5ebwh+PzqbCAj1Q4Tr6XEt7/JJ2YPIWcDg0pMthHqw6+x8XE6fwtKMSeOPFRE7t1dc5x7Va8V39r4BGYCKvjFIwU+KDOhgiLZQw6meqY9xh+8EjIopd2mixcPmfA+OTuPOvmy3BCsHBpuTwgOh46kBLO0203Cw+9llR7MyE/Gn5ZrWtiQs3mp77j96lqnBzg2Fmm0ZM8uTDLgIFiNNHeUWOb/SEsg8JSuEZC/rdFkRdCEhwZoIlUzkIGQzAoReyItyZHvTbTLmqAdVSEpnvjOK1fWtYtZXcIIKFM7zarEtEhyWqEAbmwJ4Iw4yOHdLWLtAd+UFCCugUb6rw1hCqlO3sKW88JSXHZZbhjmsEQ5Saw41kSUaSMhxPc9bTqWVutYS6NqVHxYzI33ITcQQkrIELpicz6laPUhscUtunjcIl/zY4/PJBYdDKGO3N1Gbk8lwEkODWbKhRpa08EQ/aokizDHW7Bkjy8xbDKtDhvhQ52cWv+te52egwMGLlRcOpA+b0md0rUCh+nnHvTooGGMJBpePTOB04ZKdKEv69kl0EZQa4MCrRwBBXZJzicH2+kfK91N4JwBgtrmR3zvZDR3i8VfS9vx+DwXujrCOK2Lhl7U09k8r67Wi6n5TtxHhe0NGOfxdkv/GwU7+Y5ODbDD4cikuBw7NFmCUhOLas046qe1Gy9xXqGBcfSHR3U3kE43Sdm5a+s0zNhokkN3m4VyXH0BgfX9g+hpIuLUsZdu0XHUFUYsHeEGn8Qbt7rkjNGGCGhOeShprFi3o0UuXlEhFm6qRmyrCzO6B/HISA1n9rNi0VbP2el0sWtpoHdyTP9L906osv+y82R/cTqd8T6Xr1fYQBepSxvdlEZNM1UlBdzlPyWgrltmhw3t3R6xyueVcoBTiqlZYdnNZoisOIlYO5DTleKa4OrKlSXQu5tFJACtDKwLXzRhVc8Q8mNpZCWbUbRexx2XBGT/LClM5O4BDI4ojo4o9LMupkxOpcwPoHrTTvnxumYx/4tqaO1SZjmt4vPNbqnr4TNCodDqE/TJyoKjoSE2LxwO5obDhhO6MExCq7ZYtH0ej0dFz05561AOVqHGQEDe1tTUPoNUzeHTaxFFGiFDyF8L6zHuXwfdWBTnsK5ub/dP0ekOpTEceXF2UGTqBqhaRRNdpVgzrZ54Ast3MzdNeTRW+jS55Fa+N1B+6xy+kRHMXYK+rvqWkyjFkB78rEKYalMAq42WHOKJuKEhs7VS/LZrHS6/ir70tqD40wo/8xnqatoZKSnY1txsmcAkx4yqKpxOpyqH17VouomumUSIdkIoFG5gh5aZde3PwaCHse9T1yLP3xG3t+iWGfRfXxFCZE7oKnBGNkVtHLmPdFNeaA0F384Gic01Bg40GcqPreDuxP7JIvbPQ4OyuSUs6LJC4aSs6D6Ul4p74zKASj8jVTz/SJsmDzdDVLQKHGvSsH0/xfbAIHrRJZIMWI/ZxJGQFEJBtkRXJ5DL0GXvNKCX00Dm6LFS71vEK/P/upXAvmK0UJwkUjo0pGhyxBMWHK6W3CM5EkTPbgyCnN43BuO4FWRbkWzXGFUzcKguiM+3ufHB2lZ4fOF2s1ncGwz6X+oIGv/SPToEYGZyhgWD2sqMGBHz8Cg9EkZcVWmgvI3kYq9SYwT6Mfh/WlcdvZM0WeOB+PJQGJ8dCKG0IYwUxo+HEYzh3PIcBhlOoImhrO1eDVsaCG6jQJBxZQWOjVdkPAPJTBDV0lhKoBltphVNFxmN1Ml5NJ5UKLOFn10cMFTFsHJft65xGDE0R04Y202cllSN3IaNqOAQW75XwycHNbl0D10vWntTi2Ix+/REDMm2QGess7TCi60HPdjLBEUL79ct04rJQ+JhtWjyxlcqBX+TjIZeHAgHFv4SACd7X0cATKa1fsoYw4xnxmpYckTiS26/3CTykjRM62XCJQMs6JlmwY4Wgfd2+LB4rx91LWEmC8KUrBpaaCUzn4DeFN+5BD3VIuHklkawMqmz0xyMPRPExpAWEdXkYTB5hD4Jx6NWtM7hpVXdRO4vbxfY3QS5s8EQe1warIye5CTThWI2yh0yY0xBHK6ZkIjJ/RwRUDdsa8bSLS1YfzSA3Y0SZGm2E6Tks1GtXDPRiTunJcspj5eJ8obAgVQZGFLPy/3yc5+8vSd6ddLuEBMTk97eHtw7Okt3JlNPfnFYEeBf3Y6//YC9ihMXZZlw6fAYnNvPRh9Vx7bKAD7Z5cOKEj9Ka0NIMxmYlBZGdwKsODYjXiCO+jeie3kPM7fjMevj91OXVsmJE01j4ETFpW3c3ARb+dKvlmgwqNSH9YnB+aMI6qBYxNl1rNvTho9WNWD9bjeOtYXhiwRGf3iQnz9P5BkMzBqbhIwkM15YVMcQqD4hFPKtPHHvjnr/edf+/fe1WAqYtyvunayZq+igqgyO/wcQ//vNfvJD5KN6CURIeVrPGFw4JA5n5duQHq9Tx4bktyVe8X2JD60tQfSJCWO4M4RkGmKR0CWfLBLx4rtKPqj248Pysuqz6kddWEdp0IpK3Y60rFiMHxiHAV0p6GkEbKHo/ej7RizZ0ogmNzse6Qn1+I8XUlf95cYxw/vS1iiMwYrtbg404wpa22//8tEnb+//oqv/bze32+1dfD5jj5QiIYNZnwT6pCXtfHwVI/5FSqn9bCSwMm/Pvfx6+L1ebPr+OzTXVEV+6pFlwdl9bTizpxn9EiViZFi2MiLlY/bAZjXDyghWiBZZOBBEmCwraZkJom0i8maLCRaHGRqTxu4YmwxZLMLKQIYgouUNQXy3v10sK26L6NWAL4S45HQUjhiJopGj5Ievvyyqy4+y2/+D88Gup/BZGRhjXJwSgb3WNTklHPYviTxAB76cdID5LCYhbOuJ11AlSvMZLfqq4V9xwQlwAwRExwXX3IKZV97ESg1d1leV4YHr5ogpGW4m+Wlc1Rqo9xiwUA53TdTQ26mhDzkvNzcemel2OONMsFNJR7JLylKnleULSbT7w6ijiK0h15fV+nGwyhcxlMrrg8rqRQz1dN9UEwak63h7pwev/n2+nDjtfOwsb8GBnZvFH267Bq4WFUNTfVWbIqHiVzVofyAn7zUkMcQIqcCmFr5qaLbbtP70jY+PUB7dUe1/GIr/lm6E+NyfkxRD22hlppsNZFNfqoxPhC4RTH8glFRiMIzc3gW44NrfysGjxsHrYZae7Od2uyCC7bh/uAU94jXcuDqE93YHmNqTOFQf5hbCkhIV3lDEPzFu+a6IfuJrxBeO3JDHnHjnx0g7ftCUPhb56gSroJ8rv9rnElXHKmHzt2L5e/OwbMkSShM/Ep2pSGJQJMZhh8Y+19c3oLa2FkZYmfIa9boZhXEG1jVztKgm5DenAtzIrSMdOMkvSkx7/XKrJpE+2RlEPF2XxU3mSP3UcVGtOmAgJS0NUy+5VA4cOR4JzjTiaoqU1+RmOcXyZUvx9/+8Cd/PjmcljZSPbg0JG0Uus3too4G0uSaEitYQ6smdMyi6bxlkRouPZTkcMyECqxFlE+mdxujXmhqJx9cFEEf2zkjQWJ/Fd8aqdYoZhp/l42c40C3NhgnzG8QxI1O2M6csGciYft5MOW3mTHjNCdBMVjgsFhTmJCGWkqKG6mP5shWYP/8dsWtXMboyxdVoWFTViY++/pmhkHftSSbzL17+xNj+xR//nTs1s/VGIyRedVIHTyHIXiqmlS1mNNHdOc5hqqrCAltsggwHXWLYmLPl3X94Sozrm8aQpAO3zn0G6165DytmOWFQr1qZtbfamfuxW/C7Za34nnHq3KxYOb0PQWoKiCuzw9TFoUi1hnK8I+UB1MNmnvdhlQ631YpNh9vl4UYdUwoScOdoDTGJRNlBUz/ejtV7WjHr8R0I2dIx64bfYPj4c6TX7UPJnj3YWVwsGmtr4Pf7qfOtyMxMx+DCQZg5baIs6NkFn368EPf9/iFRUXaYAsS2RErflH8nLTvrtahKLfPplMguVqu8PEOTV2Rqsm+MRZpV2argBou8/aJe8i/X57Oc1S4feWWBdPl9NH+kvPaeh+WEbE26b02WrrndZPjt4VJ+OkY+ez0Rpbx99W9vGt9v2mbcd0lPQy4ZL42Pxsjwu6Nk6O0RMvj3YXwfzu8jpfHhaVIunSA/mzsoct7f3pwfOe+2C3ne12dJuWKyfPc/i6TZpMnzZl5gfLttv/HM+yuN06dfK+NSulMUsJxTcNNp3VlUMJuZECRE+m6OTZfTZ11vbNt9UFbX1Mpzpk5X93Dpuvn6UwXKD0qiQ25PmMLLhdDz20Ja3wpmhVLIzQNjw9S6AtUsUZ8zKQuvXpIhY3ceEGuP0d/dsB29Bg2X/Xt0EytWr5PuPSvF7Hw79OQ4mDOTOB7M+MNHVThY4WH6zyWKtxeLVRt2iBvP7kK/k0YPLWZBq1pn4IJqkWOI3Ek5Lbg99lEldh1ywcXz9u3dI75euQW3Ts8WS7e34sKHt+C6G3+Df/z9Tcxf8AmeuO8/xKEd6xEUdljiu8Aalw5bTAo0eyrmXHstrrrhRmT2GIzqhnZsXfmV+Mfbb9FKT8Qt98+VpaUHrYdL9pyj6/YSKUN7OoTSP7lJRxhZP7kdmJvxz9F069NNIXHrEurhXrSsm2l8abR0b5+RjZa9lSKBcfvfD7OIOcuP4ObLLhaxC96RdiuNcUrzSOw6Eq1Q2oWb4hFq2DWrv+W7RB/qxMiB6mflalE8G450aAndpazdJCL7eCFdFVHTzl218pvIeV3S4kVJXVBe9cftmDJtJl55+S/ivgceEn968nGmGjNhziiMhEHtIR/sQTfsjDubaOFNLEjDFVdP490Y+Xrgavn8P1bgT0+/IJ74/Z2orKkWdz32rKytqjLt3b7xZas1fpvf33aQN+yw1pEcfOKhaCKFl1Bs7aX9078hoKW280Mi3Zrfz86RzqAXick2Uc+SmXeYLMgvLMLrr7yK3du3iVRLAJf1ZfS/3Q+RwIL1BAfqmoNy2VYGo1mLodqF4zMwY0zqD7Wz3GFQNOSdCxTeL6ThhVbP5I7FyqyUxMI1zB38cN6kEWly28FW7KuyE/Sv8M6C98V9v7sbcb0nITarEHpLJeJDHiQFPUjxe3CZrUk+kNaEb5duRGuXbug7YICIcehi4qj+IjYtF8t3tqH467dQNGQIbrj+Gix87wOH3+9Jl9Lo0Jj0qQBY4UBGCu1NSox7JxAIV/Ch83VdS7rh3K6MJ5MtZRCrDvnxaanAgy+8JkePG4dFH7wreiQJzO5Nsas4uNENEWeXA/o5RVVLSDa5AmL84BQ8fUMvisUWOMMBmOKtx40rb4sQqQNZxlEmQjXbcKAqiLGDnCz30VFR78Xg3gm4eUZXPPi3vXjiySdE3/w+mDH9PIGU/kjtNwV6TQli2+uQHPIiO+BFQaAd1w1oF73yDZHY1i7efGuVqNI0OWjkEEp/Wta9UuTr61ycVaGLkm/fFU88+hDaXC6xbu13fcxmx1eGEeywXPEpA1iB7PNxIooMbqb15QgEjDPHDEgUBfnUrRTX89ezCN3aB5PPvxjjh+Tj008+F/1jWnBBH3Iws0Aqbyjq2oQlIw5njMgQdoaNWl0evLDwiFi8uhY35dD95iF0bxjhYNbg0BdCHNsCvbEdFzx/EK8vr6SADmFsoRMPXdkLyzfUYeMRC15+7WXxh0efFN+vWYu0EVeS4w/DUb0PmRTLOQE3htIXn0ipkD+ZQZbRzETR5dJLglj97XqxoriU0jwda3aUiS+3VYjY3iNQs+lrOONN4uabb5Dz5r2le/0Mj0njK/X8HdGUIjrljbMBVyll+cIn5VSnmvSYHPis2IWZF15Ad8iKnl1SxQXnnYOP93joDvEQhps0Jh887P2z75Rg8K2b8P5mJ9KLbsDZF98ld9RD/GWTC+FDNRA76arsPApRTNW3qwxvLqnAhsMenHPZA7LvWfdiXVUvnH7bVjzNgTHi9NOxtbxNfvThx7BmD2WygmFNcm8qAx35vjaM87sxlWJ+jAjDNqgQuPhmONPT5QTGyy/SvHAsXYJHLrkN9//n67h2Qm98et9kxBRNx5tvvY02ODCgaCgHZnByenqkyOiU070jO6Bpmu39iJt0QS/5mxm5sltevty875BcsGqXDIV9sqGx1sjtkS9Z1C5n9LHJqwdYZEGaSQ4bebpc9PkXBjM1EXdKvTz73PNSN1vlsAxdPjwuXr55rlM+PTFBTsxlYpkj6Z577/vxWPXhHwsWRPbf/tDT8s/vf83ZiiaZPOxa2X34DbIosaecZXXKL+LMsrE/YyxXm6Scf4aUbWt4Jouty96S8u406R0NWZxqknPjMuWo7HHyD899KAOGlP1ufEOaHUnyzSVrjavv+L26T9BiiRnYUcTtaCv6Xz2XYRimW4UIOl/4uGwiM0jaSy/9FTEJifCWtzLGJbB3zz7haaeIPH08KtRMhK5d8OCD58sLZ06jS6RS/D7qdV/Egr7zjtsxpKhQ3vfIH8Xjm7ch6GqmLHUgr1ch3vrHbbjq8tnsh5/HU4YLG7plZ0X6ldE1R65ds5F5B1rr1jiYmsoQF/IjW4YwhPlh5wAedg7LUM69HA3L6mXb+geQft0MxNw3XdjmvYn890MYeKgV+311+ObLlajXEqXfmiJCmgU15UdEbq/evIAqB5Pd+WFn5KYn+aWzAMzHdNcXFGDanj2m7xIS04bPmDFd7qxqEkEaVI0un4yLi0F7u0dccf1Ncvasi4SNuWDlOLEchjzJGDAhjrhH/ODyuuDI6i3ueuplWbJzMx668TJx0dXXyvOv+S1SHCZRz4rJVGWFq6QzW/Fe5bloqG/2iNLSowoDhPw+hL1tNAcEGN1EU5NA0h4Ja4Ef5V9/Kj+9YJUIhl1wfrZIzlk4SNobhfA1s5okyFCslyAf3IdFi9OFXjASkgVjHJwyIzuHATUTmDakmd8xrRMBzNxsXSotqPrUwkH9GUxwor60jHkZphfLG0TXLj3lbx+YK2+/+SaxbvNOnMEMT5IzWQ7KcaJbikP4GZZscvnl0TqXqGjycDqKD0dKdokFLz1Hggaw6L0FIiElA4NHj0N1czwBdiAj0Y7YGAfKGz0Rau/ce4zlllbe0kBrQxmnxcQiaItHc9iLI+0eZB8zYDvkwZ7yZWJ1mHFw4YCxv0yM+aISvUtYk93I+i3GxdsMVp34XPC7GmG0UHqEg6ipa4MzgyEdobJLmtngvo5onQrg+vp6NbLTcnO7oY4TjTavWYHEpGTGjwsZ4NDx4F0301qOxYtPPYLFH8xH4cgxGDR0OJMUmRAsr/RxOkJrSwuqyw9jz5b1qCgtRr9cG8ZM7c6pKD58+dpD+GpBF/QaMARpXXIQ70yOBDz2bl1PoSwFzxX2WAZKVBCEAQ0jPgNefwsa/C5U0shqbQ8hgczeNTcIt9mCBsbRMzhx2cnCeVlKLmfiq4ZSwUUZ76c/72uuRvDYfpbmerBu+1GhZi6qASsM5Zx3TOtUADPrFPZ6vUzdati9YxvefPoR3P/c68cTBhRuMcz93nPz5XLw0GH4ZOHH2Lx6hdj8/Rr4WBBAScq6ZV3EspZWJQLSzV58+tvuctxpWZSxDIL4Q9i3pxGz/7xPrl1eJxKTU+ClTvcHmHuOCHge4mqQPfoNRU5ON1HWVAUjuxDt3mR47K1oJtBtxCXcxOBLXwP98gyspZ9+YVFIptAlDzBLSQYGPW6oTGiQaUSD/nLgWDGTKFYkJGfA39bAZwlAt9qr/1kufHKB7lQAx8fHN3q9/obi4l1xpaUPE4B2xMSqqnWpdDG5zBBW3RDnjs6XI/rfI8pqb5AVtY2MQ7fxEINZJ6vM65oJV20F7r/7bjHnpY2i6KMjSGGJTxNnNOxoNCFv2Fnib//xgExMJcG9HhKcOr65GXdfebEwh5thdzgwZvwEHH3rDdg42+2s8ZPFwS8Wo8XfDDd94TprGM8utWJrFQMuzE69s94igt0COJcrDjQT3DYFMIeMGjYhJpVDVZvRvf9ApCQnoeGwsquk3wSdNaMd0zoVwEyatwvdunVn8c48lfjXmYMVrIfV6CjT0CJfAKWVrUiKtYr0pBjWZqWI4b1SCe3xRsHKYzgOemVj9eplctk3q7B2wzY0NbdhcGaKuP20EXLc6aN5XU5BMFh8SzCUobW+tEH2GTRUlu+nSD99mugxcJTskvW58B7dIK772714m6K/5d0q+EUjXqgMYkm9lflgigyef4xVm88fZD9TPbCxqlPNP/ZRxIdMNgQ8TZwx047f/+42ufVgk1j42TZ21HSgZ89uR/Yw7dgRrVMBrB6YQazXGYO6UCUCwqyQCLXWCHv3vgxFemR1U7to8QRQy5mCCmDFtWTASKpXVxV2EXT5Iv3CyuTE9ClTMX3KNLX/h8ZphOQt0O0hNjyRg4GbhZmn4ROnYfva5cIUaMCx1ixcd9OteGTu/Zg2+y7k9BqF3G79cbi+Dks5WNwsdD5+Lm1vSTAZMvvIZcMkE20ATqvwWuIYBqW1XFuMW353J664aJrYMfdJVB07zIFh+4LgshMd005pqPKXHtEwwoc1k8nQhF7OQEJhHMtizp42HTXNbrgZ5SOO4nBNG3LSYuGwWmgRewi+n1ytfGE25dqqIKX6KMlOkbhmiN/UpgTj8cMUQOogxcHltHBrucJW8cYNqD2yB4Vjz0ZYi8HAXln45rO3hGFJRGp2gQz62sQ69iNAS5g9IbAaU4hqE2iBCU5HHI7ZklFjjkV97R6cM/lMvDbvFdRyYNx+yy1MTbrdFrPtZhbfUZN3TOt0AKvHlkZotd2esDEYDF519OgRxyXnT2WOnTqstT1SNMcFzEQDubhnZgJq6bvur2hB7y6JgjPx1ekKYPWuPvy4qW/cuJcjgP9UNklxvTpmP8V+WXWz8DGcWLzqc4paVnzoaaJn/kAMG9gdq7+cL5pcrMBPypZHPH4R0K0I0moPRt4tCLJ8J2Cy023LwCGfD0frSnDpnAvw1Asv4kAt68j+43a5dfN6ViBZXwoF29+LdK6DXjojwBbd7Lg6GPA+T67rFgoFtIqyY+Km6y5HRZOXa2pwcrbHRa6xoM3DkhmzCfsqmmlI2WHlHJXSqlaRlqhmBP+zqYDViW8KXPV/x5FGFqXHyABLbLcfbsSx6gZ4QyYRl5SK9Us+IO6a3Lhhg3CwSGDOrPNRV10i1mxaJfzeAEIW+r/UsSEFsGZiSNLg/jautVWJ2IwYPPP0k/Kmux8Q73+7C688/RhWfLmQl7Pti4uxXMUyH++JvnTEe6cC2GyOGcRIw3tGyH8rzPZM5HGui68dR0p3Ma3rw7WXX4RDx2rw+duviaJRY0RdqxcNbZFJSahu8ogeGfEoJnAOVnokxDBgERHCioyMlxBVohYR4BWsvFAA989NFspoO1jdJo6WVQsPwUvJykHZno2oLNkodALoC4VEye7t6N9/AK686nIksxqhvalcBF2sovQ2cF09F+cs6SgaNkBecPU1mH39Heg3fCy+XrMerz7+gNixdjlvbWtgCdBFXq/7UEeA+tN78ME7R2NJy/lcIuE15oKT0WM8RNHlkGamDpfeTweTUQTqz7vuulPeRvdn+JDRuPrGG+SoqZeIusZWokcO8vtlt8xUpHLuSml1KyYXdSNXR0COyO3jTymEyxvE4k1HYTbpcmJhtliyrQKVx8rxzarNiE/rJjd+/poo27WOh5uQ2ud0lgyNlE/dMxvzX35OfLN6Pa65c65M79pVyJBfBhiWNFOC0PJneJSzIo9U4/CBUi6IWCbWfv0ZXK2MYmnWGpZ3zWIyZM2poHSnAJjgzmF89k1WWlitk+6V4dwzET68W8imSgaKX2WigAH+dKbnShbi7EmTZcmBI6KuuhzP/PVlmTt0gqisa8TKzxbI8dMvQWJ8DFewCwqlXycMyubkbU5djMwt4EQzlw/f7qxCc3sACXYTy4DMosHtx+t/nMvppXko27UWVSWbMGlEFm44L1tuLWkRH6xuZEqwn3zysQfx2GNPYs3GYtFzyFlISu/Gqk67pFEoXM2NaKmrhKu+DG11ZQyCMerBJQeEybzSrIlbAwF3x/hEvzCCTj3ApphJjOx/Kmw0q656GdI5EK51S2D4qKoEPdzNL0DvPhRG94sgj20E9rLiRU3RJYexTgv3PvgQxp5/BZ544G5lYctHn34W5TUtQiUo6P6Iou4p6JmVSMPHha2ldZzfbcL+bevQs3+R5HgS7770Ryz95COYWSQXdNXRt7Zi161pSC3MZtDUieaABX/+uhUvvrEF8YmpsvzokciAEazcYEg5ogQMrgqgXC+lBsjOkrJ9C9fxeikcbH+XO4O/QPcO23VqdTDX4KA5+zkXA01OuvIlaBlD0brqM4b4GKolARHm+7HvmMEZCxmbR7eHUwdZ/MaYH2BP5kx6iVXLPkNt2RHZJbsLPn77dcGFcXDNnJnCHzZEXUs7Khs9srSyWRyubVMch10bVqJ4/SoMGDIMf3vqQbFq8cdwsMry+nxOZWljqIRlPJL3bz9Uj5wkK2I5A3zC2B4YeeYoLPpio2ht5eRlRxqkNZH9oasVaFGgkltN33MZird0zfygEWp/RBrB7UTxRAymwwD9+Y1OKcCaob8gDf/4uEm3wzLwYjR9+zELHjjg1aIb9C8RageqvoelP0V2fPdIVIuzGjhLvy+L08lhSmy7a3Bw13qxY2sx+ZfTPNd9xzLY/bjk3LNkr7yuopYgezkQkhPjROnW7+TzD93DaU8mufj9t8W+HZtoeNnwwkQrjrQYss7kEO8+Mlhm9MkQe4M2uWxjjehCoypF57JMPdIwbdZkfLWkGE1hLgvQ9xIgbwy7Sc5tPsTrmLfIsPc3jDWXdwZgTwB9ygA2m+0jwqHA85bcQj3+wmfQsnYpQm00Spg14tQFgkxuZfwX1RthHTgNRkw24gaMhCU9m9xiQdhgEE7N8ksuYJ5xN8vpfYhluYcvRBG8fxfefndhxOCaPG5YREd+/NaL4sWn/ygC9FPra6qEm/FrpSfHcCmJ6/rp8pZvuBbHBDtm9LeIHt2dGD0wWQzul4rSugD1ugV2JgqcPXKQntdTLlywSMCZz77Gwj5sJjRvDcL1pX1ZrO+hD68stE7TThnARPA5mr8DEi99DoG2EDwljNNauUyOl/Pg/czLsJYZHi6MVbsVjhGz4CiYwMoNejntNLxqNyO4a6FEwz5h6nM2J30FYarfjbkjLHJUGpMROSau+ewWby9aIt7/8HP54YJ3xJZ1q1CYonP5Bg0ZcWZWPKoZhwYKOS+pyS/FmooQ4m1CusqaYa5rEmkpMbDH25CX4eBkMq4SwIlVaKhCAuPLbyw9JPySfU3IYwlvGxImXSn9e1cIw906wmJxLGL+uaGzIHxqYtHWuD6G3z3V2mcMTNnD0fL1+1RjnH1AKynczopSlskgWYloFbIlEPQ5/Wuek75tn4lg1X7u42+6g1JR4yotBdIoOk20zq/G23vXi4+m2nkNP0Y6JdbWW7Cm+oiwxwlMKrBiaLYJLpMZ/+CsxDco0WF3YvGRJqzivKa/3jFADi1IENUNfnxQ0oTE90vlnPN6iQxWbaoJ5Uo3o4bVfDsPCfq0HIgchFy5NMTAS4ALfiRe+DAa510Xz+jbH3j0Rdw6RTslHKwJ7UaKsslxU+6hqKVbc2AH1S6nmFA8h2t30CKlbeLsQ/1awQKPnfDvW4nAgbXCsGUB/c4HhlwL9JgIlHM6yYFvhWPAJGkfMgNH1nwiytu8uHRYvIxNdKB3PMRZmcDknhYMKnAir0cCFmz34ul1nGCRP5NrKk2ArNiGK8Yl4tFZmSKbi6z07R6PM4pSuUh4kqjlIuEZicwccSzQomJV3SEuxCbw4jovrXyuqdh8mBxcibDPIxPOvBbBil2su97fx2SKW/mDLj7lIJ8KDtZYAjlFi0mFJW8U2jjVUuldZTuxYJK1c4zDm+MVRY8D7UgFuk+gQTOWIjGHx9IAYzJdsKLCfuXfEVz6CBqfmyQSZr+ApKtelZ/Pu1Tcttov3rirL2cq2mlwBxgU44x+nvbIG6Xy4W9ahGPy3QgkjpChNX8hdD5khU3wfbeHYphr/ndNYclGBnqk29BDVRCp/AQZNrC7HGbWcm1psiPIxbrUD4O1A7Ku2RCV1RtEs1EJx9gr4N+7Rg+H/HfygO+4nfLW8Rxss+UgFHzE0nOkxVZ0Mdy7N5JhybHKPaE4lpUb+JnjThkxcbRWB1La5ZzGSYknQCfwqjGZrhIQCWdcAz0mHu6v/wRhjxe2EXPk1hVLxCZWN7IgHs4u8fTEDNw27yCe+bJKxE67D1bOVvB8dg/ra5QnY8b2mjDe3B3EB1zJx9PgQgGNOyt7I1mvxZIPBHYdRfAgVQejZHcsa0fRwHQOSE58kAHxaKGU1NFi++6dFNcs7UhIh9FwJNdiif24M+jiDgeYyxOezgnAV9mHz4LIHEzjamsEXKVrJYvVUEkjlKE/pA4kp5J14gnyiRZJCJHplAuluJ7xaUnwYvpNhLVwOoJlWwh6orAWXYh9q1dg8XeHkeG0Ye4bB/DO6ibEXfQUzGl5cH3yAKPSQXBmN/JYMDIkLoi+nAGqKjT+uj2MNccCmGh3w1HXBB+BDTa0IY6F9gvLDby4wYev5xagqL8TT35Wg4HJUpzbhetOc4AWlxwmX7M8yOemOWhUUw2dci7ucICFsMygOjuL0QwYFLdBNw0p5RIp8Fj9CMXBSgxnDKaRpUQhOTaOMekIuARWvauIkdoIdIhlN2EmaKyJ6bD2mcBZD6nQ0/Nh6Xc2qugufbRoC0qaY5F49cv0apzwHfgedlrl3p3L0ENvwwUEp28icEWRhYutcRlCrnq6+BhwgEv/z+xGmLigZqzdhK0twOxFHrpUJlycpyGNS+TNW1oTWddjajcgl+t1HXRz9Z96hikjfWSoTYbms6OntHU4wNJk5UKDon+48VhOcO9yLqlxhECybNWsqo8JWhXDkcpiZVQrwqnkUhZHHXehIiAfB/YHIvJYBbKPKcR2GrU6jWtaVhxBWkImjPJNMNpqkHTTu9StQxh14tTPoZfA/cVjSKrZipl5uqx0SxHPGW82luDsr+CgY1F9SqyGxeXA+X3MXG5Rw/uHDMz50oPRGRqeG88JbS1ueBjRemVTO5q5EIyNyzX8ZZ+OHfxTPqqMh+NutwbtYRYsHDyl6B6n6CnpAqtg7dNZbXMTxfV4si7XBSTANnIqU3BgrhUDruI6ROlEU4lj6uTMnnRr6Hsqff1jU2Cf+KIygRpVNZcRjE8kyH60PHsGYhklc4z9bcTqFRY7OXgd2uZdjJuGm2WRI4DMNIcoGpkm5y1rxNZtTWIkbboE/u2G21YZuHqgGYe5ttd6+sg3D9Dx6GhqZqqGUpeBD/YH8ewGDgh1+x/MAqaQixmHnudMcCxoaKAy7wTtR/Kcsr4wB6xJeT6DDlMJ3gAOf4tySSJiWlnQMRlcWpZBj8RcriA6iDo5nf6vg+pZOQCq+zxWuVXHgadlxLorxoY19zEYn9wA5y2fwpQ1KKLfAw21aFv9HuSm51gowHOp8ntlWnDHOclyVm8d975VLgpTNCyqNuHLg/yDW2qNTHbhvB6a7Oo047taiNUVYeyqDUf+NI+SIvxfz+0bWhALWIqzgh2iyOk87dQD/E9acCpfTF+CPYYLl3F2l6RMDXclMgpJNnIygUWMk/OCaXipZWbj6OQ6krkQC5WolaKZnC9ZOqMkgGjYB3zzGBJv+QhaUg7CHjfaitfCaK0Dts6DI9zKP8xhRNaqrPbrOJfrY95bCDy3zcAiag2uzhQxwLhmS+TvLjVHYONONaKEOKppxvdcJ2AJE1KrOsva0KpzP2+dCeCf9S01FmZPD90Is8pDDOZEvQLGFvMIONENJxyXi4rg6hGU8UUkFFef2JTwVCW3MUnkdv4WotVMHRxpypgjx8fS1s22c0U8Xuag24RpPXWsLGfpK5f8V6wZkST800q8RRnzyzs56DZyovrGhISYfZ1FBB9/oH/92okB/sVO22FXeUJkcUnKbHpR5HBkccukbFYF0gp4sjNUKQdXiWY2IpKnVdfiHqGpulladBSjmtZI/5tKH32tjDPzT/ccl/ICdRwP84UUB4jxQa73fKTA662mM3dK87rqCf5v2q8N4P/NM9LH4jK1cFFWM3/40xYLmWKz+ch9ZGG18n9qrK67+FddDMYs+Zd4dBxgec3rgUCA8j3aohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohSIUiBKgSgFohT4/5QC/wf5CjH7W4oL8wAAAABJRU5ErkJggg==" width="80" height="80" style="opacity:0.4;border-radius:12px;">
            <p style="margin-top:10px;">Ready to find people for Christ!<br>Select a configuration to start prospecting.</p>
            <button class="pb-btn pb-btn-primary" onclick="pbSwitchTab('configs')">Go to Configurations</button>
        </div>
        <div id="pb-workspace-active" style="display:none;">
            <!-- Config selector + view toggle -->
            <div class="pb-flex-between pb-mb-sm">
                <div class="pb-flex-center pb-gap-md">
                    <button class="pb-btn pb-btn-sm" onclick="pbBackToConfigs()" style="margin-right:4px;">&larr; Back</button>
                    <strong id="pb-active-config-name" style="color:var(--pb-primary);font-size:1.1em;"></strong>
                    <span id="pb-prospect-count" class="pb-badge pb-badge-primary"></span>
                </div>
                <div class="pb-flex-center pb-gap-md">
                    <div class="pb-view-toggle">
                        <button class="pb-view-btn active" data-view="list" onclick="pbSetView('list')">List</button>
                        <button class="pb-view-btn" data-view="single" onclick="pbSetView('single')">Single</button>
                        <button class="pb-view-btn" data-view="batch" onclick="pbSetView('batch')">Batch</button>
                    </div>
                    <button class="pb-btn pb-btn-sm" onclick="pbRefreshProspects()">Refresh</button>
                    <button class="pb-btn pb-btn-sm" id="pb-dest-panel-btn" onclick="pbToggleDestPanel()" style="background:#e6f7ff;border-color:var(--pb-accent);color:var(--pb-accent);" title="Toggle Destination Health panel"><i class="fa fa-heartbeat"></i> Dest.</button>
                </div>
            </div>

            <!-- Workspace stats: 4 cards + progress bar -->
            <div id="pb-ws-stats-grid" class="pb-stat-grid-4" style="display:grid;grid-template-columns:repeat(4, minmax(0,1fr));gap:12px;margin:8px 0;">
                <div class="pb-stat-card" data-key="pending" style="background:#eef5fb;border-color:#cfe3ff;">
                    <div class="pb-stat-card-num" id="pb-ws-stat-pending" style="color:#1f4e79;">0</div>
                    <div class="pb-stat-card-label" style="color:#1f4e79;">Pending <span class="pb-info" title="People in this configuration who have not been processed, deferred, or skipped yet. Your remaining workload.">?</span></div>
                </div>
                <div class="pb-stat-card" data-key="processed" style="background:#dff6dd;border-color:#a8e2a3;">
                    <div class="pb-stat-card-num" id="pb-ws-stat-processed" style="color:#107c10;">0</div>
                    <div class="pb-stat-card-label" style="color:#107c10;">Processed <span class="pb-info" title="People marked as worked through. Action taken, note logged, or otherwise handled. Hidden from Pending view by default.">?</span></div>
                </div>
                <div class="pb-stat-card" data-key="deferred" style="background:#fff4ce;border-color:#f0e3a0;">
                    <div class="pb-stat-card-num" id="pb-ws-stat-deferred" style="color:#7a6d2e;">0</div>
                    <div class="pb-stat-card-label" style="color:#7a6d2e;">Deferred <span class="pb-info" title="People you intentionally pushed off until later. They come back into the Pending pool the next time you reset deferrals.">?</span></div>
                </div>
                <div class="pb-stat-card" data-key="skipped" style="background:#f3f3f3;border-color:#d4d4d4;">
                    <div class="pb-stat-card-num" id="pb-ws-stat-skipped" style="color:#605e5c;">0</div>
                    <div class="pb-stat-card-label" style="color:#605e5c;">Skipped <span class="pb-info" title="People you decided not to act on this round. Skipped does not mean dismissed, just temporarily passed over.">?</span></div>
                </div>
            </div>
            <div class="pb-flex-between pb-text-sm" style="margin-top:6px;">
                <span id="pb-progress-text" class="pb-text-muted">0 of 0</span>
                <span id="pb-progress-pct" style="font-weight:700;color:var(--pb-accent);">0% worked</span>
            </div>
            <div class="pb-progress"><div id="pb-progress-bar" class="pb-progress-bar" style="width:0%"></div></div>

            <!-- Filter bar -->
            <div class="pb-filter-bar">
                <input type="text" class="pb-input" id="pb-filter-search" placeholder="Search by name..." oninput="pbApplyFilters()" style="min-width:200px;">
                <select class="pb-select" id="pb-filter-status" onchange="pbApplyFilters()">
                    <option value="all">All Status</option>
                    <option value="pending">Pending</option>
                    <option value="processed">Processed</option>
                    <option value="deferred">Deferred</option>
                    <option value="skipped">Skipped</option>
                </select>
                <select class="pb-select" id="pb-filter-flag" onchange="pbApplyFilters()">
                    <option value="all">All Flags</option>
                </select>
                <select class="pb-select" id="pb-filter-sort" onchange="pbApplyFilters()">
                    <option value="name">Sort: Name</option>
                    <option value="age">Sort: Age</option>
                    <option value="enrolled">Sort: Enrollment</option>
                    <option value="contacts">Sort: Contact Count</option>
                    <option value="score">Sort: Score</option>
                </select>
            </div>

            <div style="display:flex;gap:12px;align-items:flex-start;">
            <!-- Main views column -->
            <div style="flex:1;min-width:0;">

            <!-- LIST VIEW -->
            <div id="pb-view-list">
                <div id="pb-list-content"></div>
            </div>

            <!-- SINGLE VIEW -->
            <div id="pb-view-single" style="display:none;">
                <div class="pb-single-nav">
                    <button class="pb-btn pb-btn-sm" onclick="pbSinglePrev()">&larr; Previous</button>
                    <span id="pb-single-index">1 of 1</span>
                    <button class="pb-btn pb-btn-sm" onclick="pbSingleNext()">Next &rarr;</button>
                </div>
                <div id="pb-single-content"></div>
            </div>

            <!-- BATCH VIEW -->
            <div id="pb-view-batch" style="display:none;">
                <div class="pb-action-bar">
                    <label class="pb-checkbox"><input type="checkbox" id="pb-batch-select-all" onchange="pbToggleSelectAll(this.checked)"> Select All Visible</label>
                    <span id="pb-batch-selected-count" class="pb-text-muted">0 selected</span>
                    <div style="flex:1;"></div>
                    <select class="pb-select" id="pb-batch-action" style="width:auto;min-width:180px;">
                        <option value="">-- Choose Action --</option>
                    </select>
                    <button class="pb-btn pb-btn-success pb-btn-sm" onclick="pbExecuteBatchAction()">Apply Action</button>
                </div>
                <div id="pb-batch-content"></div>
            </div>

            </div><!-- end main views column -->

            <!-- Destination Health side panel -->
            <div id="pb-dest-panel" style="display:none;width:320px;flex-shrink:0;position:sticky;top:10px;max-height:calc(100vh - 80px);overflow-y:auto;background:var(--pb-white);border:1px solid var(--pb-border);border-radius:var(--pb-radius);box-shadow:var(--pb-shadow);">
                <div style="padding:10px 12px;border-bottom:1px solid var(--pb-border);background:var(--pb-light);border-radius:var(--pb-radius) var(--pb-radius) 0 0;display:flex;justify-content:space-between;align-items:center;">
                    <strong style="color:var(--pb-primary);font-size:0.95em;"><i class="fa fa-heartbeat"></i> Destinations</strong>
                    <button class="pb-btn pb-btn-sm" onclick="pbToggleDestPanel()" style="padding:1px 6px;font-size:0.8em;">&times;</button>
                </div>
                <div id="pb-dest-panel-content" style="padding:10px;font-size:0.82em;">
                    <div class="pb-text-muted" style="text-align:center;padding:20px;">Click a configuration to load destination health.</div>
                </div>
            </div>

            </div><!-- end flex container -->
        </div>
        </div><!-- end pb-configs-workspace -->
    </div>

    <!-- TAB: SESSIONS -->
    <div id="pb-tab-sessions" class="pb-tab-content">
        <div class="pb-tab-intro" data-intro-id="sessions" style="display:none;">
            <span class="pb-tab-intro-icon">&#128221;</span>
            <div class="pb-tab-intro-body">
                <strong>Activity Log:</strong> a read-only audit trail of everything your team has done in Prospect Builder. Shows who processed, deferred, or skipped each prospect, what contact efforts were logged, and which groups people were assigned to. Filter by source (Prospect Management vs Group Efforts), configuration, or group to drill in.
                <a class="pb-tab-intro-link" onclick="pbSwitchTab('help')">Open the full Help &raquo;</a>
            </div>
            <button class="pb-tab-intro-close" onclick="pbDismissIntro('sessions')" title="Hide this tip">&times;</button>
        </div>
        <div class="pb-flex-between pb-mb-sm">
            <div class="pb-text-muted">Activity history across all prospect management and group outreach</div>
            <div class="pb-flex-center pb-gap-sm" style="flex-wrap:wrap;">
                <select class="pb-select" id="pb-log-source-filter" onchange="pbLoadActivityLog()" style="width:auto;min-width:120px;">
                    <option value="">All Sources</option>
                    <option value="workspace">Prospect Management</option>
                    <option value="group">Group Efforts</option>
                </select>
                <select class="pb-select" id="pb-log-config-filter" onchange="pbLoadActivityLog()" style="width:auto;min-width:150px;">
                    <option value="">All Configurations</option>
                </select>
                <select class="pb-select" id="pb-log-group-filter" onchange="pbLoadActivityLog()" style="width:auto;min-width:150px;">
                    <option value="">All Groups</option>
                </select>
                <button class="pb-btn pb-btn-sm pb-btn-danger" onclick="pbClearActivityLog()">Clear Log</button>
            </div>
        </div>
        <div id="pb-activity-log">
            <div class="pb-loading"><span class="pb-spin"></span> Loading activity...</div>
        </div>
    </div>

    <!-- TAB: PROSPECT GROUPS (By Involvement view + group configuration) -->
    <div id="pb-tab-groups" class="pb-tab-content">
        <div class="pb-tab-intro" data-intro-id="groups" style="display:none;">
            <span class="pb-tab-intro-icon">&#128101;</span>
            <div class="pb-tab-intro-body">
                <strong>Group Management:</strong> health dashboards for ministry-level prospect groups. Pick an involvement (or a saved Group Config) and see who's been contacted recently, who's gone dark, and what efforts have been logged. The card view is for triage; click into any group for the full prospect-by-prospect detail.
                <a class="pb-tab-intro-link" onclick="pbSwitchTab('help')">Open the full Help &raquo;</a>
            </div>
            <button class="pb-tab-intro-close" onclick="pbDismissIntro('groups')" title="Hide this tip">&times;</button>
        </div>

        <div class="pb-flex-between pb-mb-sm">
            <span style="font-weight:700;font-size:1.05em;color:var(--pb-primary);">&#128202; Prospects by Involvement</span>
        </div>

        <!-- Empty state: shown by pbGrpUpdateInlineButtons when no saved groups exist. -->
        <div id="pb-grp-empty-state" class="pb-empty" style="display:none;">
            <div class="pb-empty-icon" style="font-size:3em;">&#129313;</div>
            <h3 style="margin:8px 0 4px;color:var(--pb-primary);">No groups yet. Let's give your prospects a home.</h3>
            <p style="max-width:520px;margin:6px auto 16px;line-height:1.5;">
                A <b>Group</b> bundles related involvements (or a saved query) so the dashboard,
                health view, and senders can speak the same language. Most churches start with one:
                "Adult Prospects" or "First-time Guests."
            </p>
            <button class="pb-btn pb-btn-primary" onclick="pbGrpShowModal(null)">+ Create your first group</button>
        </div>

        <!-- Group card grid: primary entry point. Populated by pbGrpRenderCards
             when pbGrpGroups has 1+. Click a card to load that group's health.
             Mirrors the Prospect Management config-grid pattern for consistency. -->
        <div id="pb-grp-card-grid-wrap" style="display:none;margin-bottom:16px;">
            <div class="pb-flex-between" style="margin-bottom:8px;">
                <span style="font-weight:600;color:var(--pb-primary);">Saved Groups</span>
                <div style="display:flex;gap:8px;">
                    <button class="pb-btn pb-btn-sm pb-btn-primary" onclick="pbGrpShowModal(null)">+ New Group</button>
                    <button class="pb-btn pb-btn-sm pb-btn-secondary" onclick="pbGrpToggleAdvanced()" id="pb-grp-advanced-btn">Manual source &raquo;</button>
                </div>
            </div>
            <div id="pb-grp-card-grid" class="pb-grp-grid"></div>
        </div>

        <!-- Source/Group/Manual selectors. Hidden by default when a card grid
             is shown; revealed via the "Manual source" toggle in the grid header.
             Still the only path for "Manual (Program/Division)" ad-hoc browsing. -->
        <div id="pb-grp-source-area" style="display:none;">
        <div style="display:flex;gap:12px;align-items:end;margin-bottom:16px;flex-wrap:wrap;">
            <div style="min-width:200px;">
                <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Source</label>
                <select id="pb-grp-health-source" class="pb-input" style="height:36px;" onchange="pbGrpHealthSourceChanged()">
                    <option value="group">From Group Config</option>
                    <option value="manual">Manual (Program/Division)</option>
                </select>
            </div>

            <!-- Group selection + inline CRUD (inline so buttons align with the row baseline) -->
            <div id="pb-grp-health-group-wrap" style="display:flex;gap:6px;align-items:end;">
                <div style="min-width:220px;">
                    <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Group</label>
                    <select id="pb-grp-health-group" class="pb-input" style="height:36px;width:100%;" onchange="pbGrpUpdateInlineButtons()">
                        <option value="">-- Select Group --</option>
                    </select>
                </div>
                <button class="pb-btn pb-btn-sm pb-btn-primary" style="height:36px;" onclick="pbGrpShowModal()" title="Create a new group">+ New</button>
                <button id="pb-grp-inline-edit" class="pb-btn pb-btn-sm" style="height:36px;" onclick="pbGrpInlineEdit()" title="Edit selected group" disabled>Edit</button>
            </div>

            <!-- Manual selection (hidden by default) -->
            <div id="pb-grp-health-manual-wrap" style="display:none;">
                <div style="display:flex;gap:12px;align-items:end;">
                    <div style="min-width:180px;">
                        <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Program</label>
                        <select id="pb-grp-health-program" class="pb-input" style="height:36px;" onchange="pbGrpHealthProgramChanged()">
                            <option value="">-- Select Program --</option>
                        </select>
                    </div>
                    <div style="min-width:180px;">
                        <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Division</label>
                        <select id="pb-grp-health-division" class="pb-input" style="height:36px;">
                            <option value="">-- All Divisions --</option>
                        </select>
                    </div>
                    <div style="min-width:120px;">
                        <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Prospect Types</label>
                        <input type="text" id="pb-grp-health-mt" class="pb-input" value="311" placeholder="311,230">
                    </div>
                    <div style="min-width:120px;">
                        <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Leader Types</label>
                        <input type="text" id="pb-grp-health-leader-mt" class="pb-input" value="140,310,320" placeholder="140,310">
                    </div>
                </div>
            </div>

            <button class="pb-btn pb-btn-primary" onclick="pbGrpLoadHealth()">Load Health</button>
        </div>
        </div><!-- /#pb-grp-source-area -->
        <div id="pb-grp-health-summary" style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;"></div>
        <div id="pb-grp-health-content"></div>
    </div>

    <!-- TAB: PROSPECT SENDER -->
    <div id="pb-tab-senders" class="pb-tab-content">
        <div class="pb-tab-intro" data-intro-id="senders" style="display:none;">
            <span class="pb-tab-intro-icon">&#128231;</span>
            <div class="pb-tab-intro-body">
                <strong>Prospect Sender:</strong> scheduled emails that fire from TouchPoint's <code>ScheduledTasks</code> and notify leaders or staff when new prospects show up in their groups. Each sender has its own source query, frequency, recipients, and message body. You can write the body with merge fields like <code>{Count}</code> and <code>{ProspectsLink}</code>. Click "Edit" on a sender to see the full list.
                <a class="pb-tab-intro-link" onclick="pbSwitchTab('help')">Open the full Help &raquo;</a>
            </div>
            <button class="pb-tab-intro-close" onclick="pbDismissIntro('senders')" title="Hide this tip">&times;</button>
        </div>
        <div class="pb-flex-between pb-mb-sm">
            <div class="pb-flex-center pb-gap-md">
                <span style="font-weight:700;font-size:1.05em;color:var(--pb-primary);">Prospect Sender</span>
                <span id="pb-snd-count" class="pb-badge pb-badge-primary"></span>
            </div>
            <button class="pb-btn pb-btn-primary pb-btn-sm" onclick="pbSndShowEditor()">+ New Sender</button>
        </div>
        <div class="pb-text-muted pb-mb-sm" style="font-size:0.85em;">Automatically email staff/roles when new prospects are added. Runs via ScheduledTasks.</div>

        <!-- Sender List -->
        <div id="pb-snd-list"></div>
        <div id="pb-snd-empty" class="pb-empty" style="display:none;">
            <div class="pb-empty-icon" style="font-size:2.8em;">&#128235;</div>
            <h3 style="margin:4px 0;color:var(--pb-primary);">No senders yet. Your inbox is suspiciously quiet.</h3>
            <p style="max-width:520px;margin:6px auto 16px;line-height:1.5;">
                A <b>sender</b> automatically emails leaders the prospects they should follow up
                with, on whatever cadence you want (weekly, daily, monthly). Pick a group, pick
                recipients, pick a cadence. Done.
            </p>
            <button class="pb-btn pb-btn-primary" onclick="pbSndNewSender()">+ Create your first sender</button>
        </div>

        <!-- Sender Editor (hidden by default) -->
        <div id="pb-snd-editor" style="display:none;margin-top:16px;">
            <div class="pb-card" style="padding:20px;">
                <div class="pb-flex-between pb-mb-sm">
                    <span style="font-weight:700;font-size:1.05em;color:var(--pb-primary);" id="pb-snd-editor-title">New Sender</span>
                    <button class="pb-btn pb-btn-sm" onclick="pbSndHideEditor()">Cancel</button>
                </div>

                <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;">
                    <div>
                        <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Sender Name *</label>
                        <input type="text" id="pb-snd-name" class="pb-input" placeholder="e.g. New CG Prospects - Adult Ministry">
                    </div>
                    <div>
                        <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Enabled</label>
                        <select id="pb-snd-enabled" class="pb-input" style="height:36px;">
                            <option value="true">Yes - Active</option>
                            <option value="false">No - Paused</option>
                        </select>
                    </div>
                </div>

                <!-- Source Configuration -->
                <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-bottom:12px;">
                    <label class="pb-text-sm" style="font-weight:700;display:block;margin-bottom:8px;">Prospect Source (which involvements to monitor for new prospects)</label>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
                        <div>
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Scope</label>
                            <select id="pb-snd-source-scope" class="pb-input" style="height:36px;" onchange="pbSndScopeChanged()">
                                <option value="all">All Involvements</option>
                                <option value="program">By Program</option>
                                <option value="division">By Program &amp; Division</option>
                                <option value="org">Specific Organization</option>
                            </select>
                        </div>
                        <div id="pb-snd-program-wrap" style="display:none;">
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Program</label>
                            <select id="pb-snd-program" class="pb-input" style="height:36px;" onchange="pbSndProgramChanged()">
                                <option value="">Loading...</option>
                            </select>
                        </div>
                    </div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:8px;">
                        <div id="pb-snd-division-wrap" style="display:none;">
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Division</label>
                            <select id="pb-snd-division" class="pb-input" style="height:36px;">
                                <option value="">-- All Divisions --</option>
                            </select>
                        </div>
                        <div id="pb-snd-org-wrap" style="display:none;">
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Organization ID</label>
                            <input type="text" id="pb-snd-source-value" class="pb-input" placeholder="Enter org ID">
                        </div>
                    </div>
                    <div style="margin-top:8px;">
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Prospect Member Types (which member types count as "new prospects")</label>
                        <input type="text" id="pb-snd-member-types" class="pb-input" placeholder="e.g. 230,311 (leave blank for all)" value="230,311">
                        <span class="pb-text-sm pb-text-muted">Common: 230=InActive/Prospect, 311=Prospect, 310=New Guest, 301=Visitor</span>
                    </div>
                </div>

                <!-- Schedule -->
                <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-bottom:12px;">
                    <label class="pb-text-sm" style="font-weight:700;display:block;margin-bottom:8px;">Schedule</label>
                    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;">
                        <div>
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Frequency</label>
                            <select id="pb-snd-frequency" class="pb-input" style="height:36px;" onchange="pbSndFreqChanged()">
                                <option value="daily">Daily</option>
                                <option value="weekly">Weekly</option>
                                <option value="monthly">Monthly</option>
                            </select>
                        </div>
                        <div>
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Target Hour (24h)</label>
                            <input type="number" id="pb-snd-target-hour" class="pb-input" value="7" min="0" max="23">
                        </div>
                        <div id="pb-snd-day-wrap" style="display:none;">
                            <label class="pb-text-sm" id="pb-snd-day-label" style="display:block;margin-bottom:4px;">Day</label>
                            <select id="pb-snd-target-day" class="pb-input" style="height:36px;">
                                <option value="0">Monday</option>
                                <option value="1">Tuesday</option>
                                <option value="2">Wednesday</option>
                                <option value="3">Thursday</option>
                                <option value="4">Friday</option>
                                <option value="5">Saturday</option>
                                <option value="6">Sunday</option>
                            </select>
                            <input type="number" id="pb-snd-target-dom" class="pb-input" value="1" min="1" max="28" style="display:none;">
                        </div>
                    </div>
                    <div style="margin-top:8px;">
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Lookback Period</label>
                        <select id="pb-snd-lookback" class="pb-input" style="height:36px;">
                            <option value="since_last">Since Last Send</option>
                            <option value="yesterday">Yesterday Only</option>
                            <option value="last_7_days">Last 7 Days</option>
                            <option value="last_30_days">Last 30 Days</option>
                        </select>
                    </div>
                </div>

                <!-- Recipients -->
                <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-bottom:12px;">
                    <label class="pb-text-sm" style="font-weight:700;display:block;margin-bottom:8px;">Email Recipients</label>
                    <div>
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Send To Mode</label>
                        <select id="pb-snd-send-to-mode" class="pb-input" style="height:36px;" onchange="pbSndSendToChanged()">
                            <option value="involvement_members">Involvement Members (by Member Type)</option>
                            <option value="roles">By Role</option>
                            <option value="specific_people">Specific People</option>
                        </select>
                    </div>
                    <div id="pb-snd-inv-members-wrap" style="margin-top:8px;">
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Recipient Member Types (select which members receive the email)</label>
                        <div id="pb-snd-recipient-mt-list" style="max-height:180px;overflow-y:auto;border:1px solid var(--pb-border);border-radius:var(--pb-radius);padding:8px;background:var(--pb-white);">Loading member types...</div>
                        <input type="hidden" id="pb-snd-recipient-mt" value="">
                        <span class="pb-text-sm pb-text-muted">Leave all unchecked to send to everyone in the organization</span>
                    </div>
                    <div id="pb-snd-roles-wrap" style="display:none;margin-top:8px;">
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Roles (select one or more)</label>
                        <div id="pb-snd-roles-list" style="max-height:150px;overflow-y:auto;border:1px solid var(--pb-border);border-radius:var(--pb-radius);padding:8px;background:var(--pb-white);"></div>
                    </div>
                    <div id="pb-snd-people-wrap" style="display:none;margin-top:8px;">
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">People IDs (comma-separated)</label>
                        <input type="text" id="pb-snd-specific-people" class="pb-input" placeholder="e.g. 3134,21480">
                    </div>
                </div>

                <!-- Email Settings -->
                <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-bottom:12px;">
                    <label class="pb-text-sm" style="font-weight:700;display:block;margin-bottom:8px;">Email Settings</label>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
                        <div>
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">From Email</label>
                            <input type="text" id="pb-snd-from-email" class="pb-input" placeholder="e.g. admin@yourchurch.org">
                        </div>
                        <div>
                            <label class="pb-text-sm" style="display:block;margin-bottom:4px;">From Name</label>
                            <input type="text" id="pb-snd-from-name" class="pb-input" placeholder="e.g. Adult Ministry">
                        </div>
                    </div>
                    <div style="margin-top:8px;">
                        <label class="pb-text-sm" style="display:block;margin-bottom:4px;">Subject (leave blank for auto-generated)</label>
                        <input type="text" id="pb-snd-subject" class="pb-input" placeholder="Auto: {Sender Name} - X New Prospect(s)">
                    </div>
                </div>

                <!-- Email Fields -->
                <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-bottom:12px;">
                    <label class="pb-text-sm" style="font-weight:700;display:block;margin-bottom:8px;">Prospect Fields to Include in Email</label>
                    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px 16px;">
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="email" checked> Email</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="cell_phone" checked> Cell Phone</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="home_phone"> Home Phone</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="address" checked> Address</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="age" checked> Age</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="member_status"> Member Status</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="enrollment_date" checked> Enrollment Date</label>
                        <label style="font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-field-cb" value="person_link" checked> Link to Profile</label>
                    </div>
                </div>

                <!-- Message Body -->
                <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-bottom:12px;">
                    <label class="pb-text-sm" style="font-weight:700;display:block;margin-bottom:8px;">Message Body (override default)</label>
                    <p class="pb-text-sm pb-text-muted" style="margin:0 0 8px;">Leave blank to use the central default from Settings &rarr; Sender Defaults. Anything you enter here overrides it for this sender only.</p>
                    <details class="pb-merge-help" style="margin:0 0 8px;font-size:0.85em;background:var(--pb-light);border:1px solid var(--pb-border);border-radius:6px;padding:8px 12px;">
                        <summary style="cursor:pointer;font-weight:600;color:var(--pb-primary);">Merge fields you can use in the message body</summary>
                        <table style="width:100%;margin-top:8px;border-collapse:collapse;">
                            <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{Count}</code></td><td style="padding:3px 0;">Number of new prospects in the batch (e.g. 3)</td></tr>
                            <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{SenderName}</code></td><td style="padding:3px 0;">This sender configuration's name</td></tr>
                            <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{OrgName}</code></td><td style="padding:3px 0;">The involvement name (when emailing leaders by group)</td></tr>
                            <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{Date}</code></td><td style="padding:3px 0;">Today's date (YYYY-MM-DD)</td></tr>
                            <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{ProspectsLink}</code></td><td style="padding:3px 0;">Ready-made link to the leader's involvement in TouchPoint (<em>Open &lt;involvement name&gt; &raquo;</em>). Only renders when the sender is involvement-based.</td></tr>
                            <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{ProspectsURL}</code></td><td style="padding:3px 0;">Just the URL of the involvement page (no text). Use for custom wording: <code>&lt;a href="{ProspectsURL}"&gt;view the roster&lt;/a&gt;</code></td></tr>
                        </table>
                        <p style="margin:6px 0 0;font-size:0.85em;color:var(--pb-muted);font-style:italic;">Both link tokens silently disappear when the sender pulls from a tag or saved query (no single involvement to link to).</p>
                    </details>
                    <textarea id="pb-snd-message-body" class="pb-input" rows="5" style="font-size:0.9em;line-height:1.5;" placeholder="Leave blank to use the default from Settings"></textarea>
                </div>

                <!-- Actions -->
                <div class="pb-flex-center pb-gap-sm" style="margin-top:16px;">
                    <button class="pb-btn pb-btn-primary" onclick="pbSndSave()">Save Sender</button>
                    <button class="pb-btn" onclick="pbSndPreview()" style="background:#17a2b8;color:#fff;">Preview</button>
                    <button class="pb-btn" onclick="pbSndRunNow()" style="background:var(--pb-warning);color:#fff;">Send Now</button>
                </div>
                <div id="pb-snd-msg" class="pb-text-sm" style="margin-top:8px;"></div>

                <!-- Preview Results -->
                <div id="pb-snd-preview" style="display:none;margin-top:16px;">
                    <div class="pb-card" style="padding:14px;">
                        <h4 style="margin:0 0 8px;color:var(--pb-primary);">Preview Results</h4>
                        <div id="pb-snd-preview-content"></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Send Log -->
        <div style="margin-top:20px;">
            <div class="pb-flex-between pb-mb-sm">
                <span style="font-weight:600;color:var(--pb-primary);">Send History</span>
                <button class="pb-btn pb-btn-sm" onclick="pbSndLoadLog()">Refresh</button>
            </div>
            <div id="pb-snd-log" class="pb-text-muted" style="font-size:0.85em;">Loading...</div>
        </div>

        <!-- ScheduledTasks Setup -->
        <div class="pb-card" style="background:var(--pb-light);padding:14px;margin-top:20px;">
            <h4 style="margin:0 0 8px;color:var(--pb-primary);">ScheduledTasks Setup
                <span id="pb-sched-status" class="pb-text-muted" style="font-size:0.85em;font-weight:400;margin-left:8px;"></span>
            </h4>
            <p class="pb-text-sm pb-text-muted" style="margin:0 0 8px;">One-click install lets this script add itself to TouchPoint's <code>ScheduledTasks</code> content. Other entries in <code>ScheduledTasks</code> are preserved.</p>
            <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px;">
                <button id="pb-sched-install-btn" class="pb-btn pb-btn-sm pb-btn-primary" onclick="pbInstallSched()">&#10004; Install</button>
                <button id="pb-sched-uninstall-btn" class="pb-btn pb-btn-sm" onclick="pbUninstallSched()" style="display:none;">&#10005; Remove</button>
                <button class="pb-btn pb-btn-sm" onclick="pbCheckSchedInstall()">Refresh status</button>
            </div>
            <div id="pb-sched-result" style="display:none;font-size:0.85em;padding:8px 10px;border-radius:var(--pb-radius);margin-bottom:8px;"></div>
            <details style="margin-top:6px;">
                <summary style="cursor:pointer;font-size:0.85em;color:var(--pb-muted);">Manual setup (edit ScheduledTasks yourself)</summary>
                <p class="pb-text-sm pb-text-muted" style="margin:8px 0 4px;">Add this block to your <code>ScheduledTasks</code> Python special content:</p>
                <pre style="background:var(--pb-white);padding:10px;border-radius:var(--pb-radius);font-size:0.85em;border:1px solid var(--pb-border);overflow-x:auto;">Data.run_batch = "true"
model.CallScript("TPxi_ProspectBuilder")</pre>
            </details>
        </div>
    </div>

    <!-- TAB: SETTINGS -->
    <div id="pb-tab-settings" class="pb-tab-content">
        <div class="pb-tab-intro" data-intro-id="settings" style="display:none;">
            <span class="pb-tab-intro-icon">&#9881;</span>
            <div class="pb-tab-intro-body">
                <strong>Settings:</strong> install-wide configuration. Set up contact-method codes (shared with Program Pulse), the default sender message body, the priority scorecard weights, and the scheduler install/enable controls. Most settings are admin-only.
                <a class="pb-tab-intro-link" onclick="pbSwitchTab('help')">Open the full Help &raquo;</a>
            </div>
            <button class="pb-tab-intro-close" onclick="pbDismissIntro('settings')" title="Hide this tip">&times;</button>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">Contact Method Keywords</div>
            <p class="pb-text-muted">Define keyword codes for contact effort tracking. These are shared with Program Pulse. Each method maps a short code (P, E, V, etc.) to a TaskNote keyword from TouchPoint.</p>
            <table style="width:100%;border-collapse:collapse;font-size:0.9em;" id="pb-methods-table">
                <thead>
                    <tr style="border-bottom:2px solid var(--pb-border);">
                        <th style="padding:6px 8px;text-align:left;width:70px;">Code</th>
                        <th style="padding:6px 8px;text-align:left;width:140px;">Label</th>
                        <th style="padding:6px 8px;text-align:left;">Keyword</th>
                        <th style="padding:6px 8px;width:70px;"></th>
                    </tr>
                </thead>
                <tbody id="pb-method-rows"></tbody>
            </table>
            <button class="pb-btn pb-btn-sm pb-mt-sm" onclick="pbAddContactMethod()">+ Add Contact Method</button>
            <div class="pb-mt-md" style="padding:10px;background:var(--pb-light);border-radius:var(--pb-radius);border:1px solid var(--pb-border);">
                <label style="font-weight:600;font-size:0.9em;color:var(--pb-primary);display:block;margin-bottom:4px;">"Other" Contact Weight</label>
                <p class="pb-text-muted" style="font-size:0.8em;margin:0 0 6px 0;">The "Other" category catches unmatched TaskNote activity (automated emails, system notes, etc.). Set a weight to reduce its impact on effort totals. A weight of 1.0 counts fully, 0.5 counts at half, 0 ignores "Other" entirely.</p>
                <select id="pb-other-weight" class="pb-select" style="width:120px;" onchange="pbSettings.other_weight=parseFloat(this.value)">
                    <option value="1">1.0 - Full weight</option>
                    <option value="0.75">0.75</option>
                    <option value="0.5">0.5 - Half weight</option>
                    <option value="0.25">0.25</option>
                    <option value="0">0 - Ignore</option>
                </select>
            </div>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">Sender Defaults</div>
            <p class="pb-text-muted">Default message body used by prospect sender emails (the gray callout above the prospect list). Individual senders can override this in their own editor.</p>
            <details class="pb-merge-help" style="margin:0 0 10px;font-size:0.85em;background:var(--pb-light);border:1px solid var(--pb-border);border-radius:6px;padding:8px 12px;">
                <summary style="cursor:pointer;font-weight:600;color:var(--pb-primary);">Merge fields you can use in the message body</summary>
                <table style="width:100%;margin-top:8px;border-collapse:collapse;">
                    <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{Count}</code></td><td style="padding:3px 0;">Number of new prospects in the batch (e.g. 3)</td></tr>
                    <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{SenderName}</code></td><td style="padding:3px 0;">This sender configuration's name</td></tr>
                    <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{OrgName}</code></td><td style="padding:3px 0;">The involvement name (when emailing leaders by group)</td></tr>
                    <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{Date}</code></td><td style="padding:3px 0;">Today's date (YYYY-MM-DD)</td></tr>
                    <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{ProspectsLink}</code></td><td style="padding:3px 0;">Ready-made link to the leader's involvement in TouchPoint (<em>Open &lt;involvement name&gt; &raquo;</em>). Only renders when the sender is involvement-based.</td></tr>
                    <tr><td style="padding:3px 8px 3px 0;white-space:nowrap;"><code>{ProspectsURL}</code></td><td style="padding:3px 0;">Just the URL of the involvement page (no text). Use for custom wording: <code>&lt;a href="{ProspectsURL}"&gt;view the roster&lt;/a&gt;</code></td></tr>
                </table>
                <p style="margin:6px 0 0;font-size:0.85em;color:var(--pb-muted);font-style:italic;">Both link tokens silently disappear when the sender pulls from a tag or saved query (no single involvement to link to).</p>
            </details>
            <label class="pb-text-sm" style="font-weight:600;display:block;margin-bottom:4px;">Default Message Body</label>
            <textarea id="pb-default-message-body" class="pb-input" rows="5" style="font-size:0.9em;line-height:1.5;width:100%;" placeholder="e.g. We just wanted you to be aware that the following prospect(s) have been added to your group."></textarea>
            <p class="pb-text-muted" style="font-size:0.8em;margin-top:6px;">Saved senders with a blank Message Body in their own editor will use this default. Save Settings to apply.</p>
        </div>
        <div class="pb-card">
            <div class="pb-card-title"><i class="fa fa-bullseye"></i> Priority Scorecard: <em>what should I do?</em></div>
            <p class="pb-text-muted">Compute a composite priority score (0-100) to help staff identify which prospects need the most attention. This is separate from the Engagement Breakdown (which shows current activity levels). The scorecard combines outreach signals, status, and engagement data to prioritize your next actions.</p>
            <div class="pb-mb-sm">
                <label style="font-weight:600;font-size:0.9em;cursor:pointer;">
                    <input type="checkbox" id="pb-scorecard-enabled" onchange="pbSettings.scorecard=pbGetScorecard();pbSettings.scorecard.enabled=this.checked"> Enable Prospect Scoring
                </label>
            </div>
            <table style="width:100%;border-collapse:collapse;font-size:0.9em;">
                <thead>
                    <tr style="border-bottom:2px solid var(--pb-border);">
                        <th style="padding:6px 8px;width:40px;text-align:center;">On</th>
                        <th style="padding:6px 8px;text-align:left;">Factor</th>
                        <th style="padding:6px 8px;width:70px;text-align:center;">Weight</th>
                        <th style="padding:6px 8px;text-align:left;">Description</th>
                    </tr>
                </thead>
                <tbody id="pb-scorecard-rows"></tbody>
            </table>
            <p class="pb-text-muted" style="font-size:0.8em;margin-top:6px;">Weights are relative. They are normalized to total 100% when computing scores. A score of <span style="color:var(--pb-success);font-weight:700;">70+</span> = high priority, <span style="color:var(--pb-warning);font-weight:700;">40 to 69</span> = moderate, <span style="color:var(--pb-danger);font-weight:700;">&lt;40</span> = low priority.</p>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">Scheduled Tasks
                <span id="pb-sched-status-settings" class="pb-text-muted" style="font-size:0.85em;font-weight:400;margin-left:8px;"></span>
            </div>
            <p class="pb-text-muted">Automatic prospect sends fire from TouchPoint's <code>ScheduledTasks</code> content. Toggle off below to pause sends without removing the wrapper. Safe when other scripts share <code>ScheduledTasks</code>.</p>
            <div style="display:flex;align-items:center;gap:10px;padding:10px 12px;margin:8px 0 12px;background:var(--pb-light);border:1px solid var(--pb-border);border-radius:var(--pb-radius);">
                <label class="pb-text-sm" style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:600;margin:0;">
                    <input type="checkbox" id="pb-sched-enabled" onchange="pbSettings.sched_enabled=this.checked">
                    Enable scheduled prospect sends
                </label>
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px;">
                <button id="pb-sched-install-btn2" class="pb-btn pb-btn-sm pb-btn-primary" onclick="pbInstallSched()">&#10004; Install in ScheduledTasks</button>
                <button id="pb-sched-uninstall-btn2" class="pb-btn pb-btn-sm" onclick="pbUninstallSched()" style="display:none;">&#10005; Remove</button>
                <button class="pb-btn pb-btn-sm" onclick="pbCheckSchedInstall()">Refresh status</button>
            </div>
            <div id="pb-sched-result-settings" style="display:none;font-size:0.85em;padding:8px 10px;border-radius:var(--pb-radius);margin-bottom:8px;"></div>
            <p class="pb-text-muted" style="font-size:0.8em;margin-top:6px;">"Install" appends a managed block to <code>ScheduledTasks</code>. "Remove" pulls just that block back out, leaving any other scripts in <code>ScheduledTasks</code> intact.</p>
        </div>
        <div class="pb-mt-md" style="padding:12px;">
            <button class="pb-btn pb-btn-primary" onclick="pbSaveSettings()">Save Settings</button>
            <span id="pb-settings-status" class="pb-text-muted" style="margin-left:10px;"></span>
        </div>
    </div>

    <!-- TAB: HELP -->
    <div id="pb-tab-help" class="pb-tab-content">
        <div class="pb-card">
            <div class="pb-card-title">Getting Started</div>
            <ol style="line-height:1.8;">
                <li><strong>Create a Configuration</strong> - Go to the Configurations tab and click "+ New Configuration". Name it, choose a source (involvement, tag, or saved query), and set which member types to include.</li>
                <li><strong>Configure Display Fields</strong> - Choose which person/family/medical/extra value fields to show in the workspace.</li>
                <li><strong>Set Cross-Query Flags</strong> - Enable flags like "Children Attending Program" or "Spouse in Org" to see family relationship indicators.</li>
                <li><strong>Define Target Actions</strong> - Set up where to place prospects (tag, involvement, or subgroup) when processed.</li>
                <li><strong>Work Prospects</strong> - Open the workspace, choose your config, and work through prospects in List, Single, or Batch view.</li>
                <li><strong>Save Sessions</strong> - Your progress (processed/deferred/skipped) saves automatically. Use Sessions tab to resume later.</li>
            </ol>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">Contact Effort Badges</div>
            <p>Contact badges show communication history from TaskNotes. Configure keyword codes in Settings (e.g., P=Phone, E=Email, V=Visit). Badges display as <span class="pb-contact-code has-count">P(3)</span> <span class="pb-contact-code has-count">E(1)</span> <span class="pb-contact-code">V(0)</span></p>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">Group Management</div>
            <p>Create ministry-level groups to manage and track prospects:</p>
            <ul style="line-height:1.8;">
                <li><strong>Create Groups</strong> - Define a group by program, division, or involvement level with member type filters</li>
                <li><strong>Set Thresholds</strong> - Filter by minimum enrollment days and minimum days without contact</li>
                <li><strong>Target Actions</strong> - Define where prospects can go (tags, involvements) - shown in the action dropdown</li>
                <li><strong>Log Efforts</strong> - Track outreach attempts (calls, emails, visits) and their results</li>
                <li><strong>Contact History</strong> - See existing TouchPoint TaskNote contact efforts alongside group efforts</li>
            </ul>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">Cross-Query Flags</div>
            <p>Flags evaluate family relationships across the database:</p>
            <ul style="line-height:1.8;">
                <li><strong>Children Attending</strong> - Prospect has children enrolled in a specified program</li>
                <li><strong>Parents Not Attending</strong> - Prospect is a parent whose children attend but they don't</li>
                <li><strong>Spouse in Org</strong> - Prospect's spouse is a member of a specified involvement</li>
                <li><strong>Has Extra Value</strong> - Prospect has a specific extra value field set</li>
            </ul>
        </div>
        <div class="pb-card">
            <div class="pb-card-title"><i class="fa fa-bullseye"></i> Priority Scorecard vs <i class="fa fa-bar-chart"></i> Engagement Breakdown</div>
            <p>The person detail view shows two complementary indicators:</p>
            <ul style="line-height:1.8;">
                <li><strong><i class="fa fa-bar-chart"></i> Engagement Breakdown</strong> (<em>what's happening?</em>): shows current activity levels across 4 factors (attendance recency, attendance frequency, group involvement, serving). This is always visible and answers "how engaged is this person right now?"</li>
                <li><strong><i class="fa fa-bullseye"></i> Priority Scorecard</strong> (<em>what should I do?</em>): a configurable composite score that helps prioritize outreach. Combines engagement data with contact history, member status, family connections, and more. Enable in Settings to activate.</li>
            </ul>
            <p><strong>How the Priority Scorecard works:</strong></p>
            <ul style="line-height:1.8;">
                <li><strong>Enable in Settings</strong> - Turn on scoring and configure which factors matter to your ministry</li>
                <li><strong>Each factor scores 0-100</strong> - The system evaluates each prospect against each enabled factor independently</li>
                <li><strong>Weights are relative</strong> - If you set Contact Efforts to 20 and Attendance Recency to 10, contacts count twice as much. Weights are normalized to 100% automatically</li>
                <li><strong>Score colors</strong> - <span style="color:var(--pb-success);font-weight:700;">70+</span> = high priority (strong signals, act now), <span style="color:var(--pb-warning);font-weight:700;">40-69</span> = moderate (some engagement, nurture), <span style="color:var(--pb-danger);font-weight:700;">&lt;40</span> = low priority (little activity)</li>
                <li><strong>Hover for breakdown</strong> - Hover over any score badge to see how each factor contributed</li>
            </ul>
            <p><strong>Scoring Factors:</strong></p>
            <table style="width:100%;border-collapse:collapse;font-size:0.85em;margin-top:6px;">
                <thead><tr style="border-bottom:2px solid var(--pb-border);"><th style="padding:4px 8px;text-align:left;">Factor</th><th style="padding:4px 8px;text-align:left;">What it measures</th><th style="padding:4px 8px;text-align:left;">High score (100)</th><th style="padding:4px 8px;text-align:left;">Low score (0-20)</th></tr></thead>
                <tbody>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Contact Efforts</td><td style="padding:4px 8px;">Intentional outreach (calls, emails, visits)</td><td style="padding:4px 8px;">10+ weighted contacts</td><td style="padding:4px 8px;">0-1 contacts</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Attendance Recency</td><td style="padding:4px 8px;">How recently they last attended anything</td><td style="padding:4px 8px;">Within 7 days</td><td style="padding:4px 8px;">90+ days ago or never</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Attendance Frequency</td><td style="padding:4px 8px;">How often they attended in last 90 days</td><td style="padding:4px 8px;">13+ times (weekly)</td><td style="padding:4px 8px;">0-1 times</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Involvements</td><td style="padding:4px 8px;">Number of active organization memberships</td><td style="padding:4px 8px;">4+ involvements</td><td style="padding:4px 8px;">0 involvements</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Serving Roles</td><td style="padding:4px 8px;">Active leadership or volunteer positions</td><td style="padding:4px 8px;">3+ serving roles</td><td style="padding:4px 8px;">Not serving</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Enrollment Recency</td><td style="padding:4px 8px;">How recently they joined any organization</td><td style="padding:4px 8px;">Within 30 days</td><td style="padding:4px 8px;">365+ days ago</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Family Engaged</td><td style="padding:4px 8px;">Another family member is active in an org</td><td style="padding:4px 8px;">Yes (family connected)</td><td style="padding:4px 8px;">No family engagement</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">Member Status</td><td style="padding:4px 8px;">Where they are in the membership journey</td><td style="padding:4px 8px;">Prospect / Just Added</td><td style="padding:4px 8px;">Already a Member</td></tr>
                <tr style="border-bottom:1px solid var(--pb-border);"><td style="padding:4px 8px;font-weight:600;">TaskNote Activity</td><td style="padding:4px 8px;">Recent notes/tasks about this person</td><td style="padding:4px 8px;">5+ recent notes</td><td style="padding:4px 8px;">0 notes</td></tr>
                </tbody>
            </table>
            <p class="pb-text-muted" style="margin-top:8px;font-size:0.85em;"><strong>Tip:</strong> Member Status scores <em>Prospects</em> higher than <em>Members</em> because prospects need more attention. If you're using this for re-engagement of existing members, consider disabling or lowering the Member Status weight.</p>
        </div>
        <div class="pb-card">
            <div class="pb-card-title">"Other" Contact Weight</div>
            <p>When TouchPoint TaskNotes are analyzed for contact efforts, notes that don't match any configured keyword (Phone, Email, Visit, etc.) are counted as "Other". These often include:</p>
            <ul style="line-height:1.8;">
                <li>Automated system emails and notifications</li>
                <li>Administrative notes and data updates</li>
                <li>Bulk communication sends</li>
                <li>Registration confirmations</li>
            </ul>
            <p>While it's useful to know these interactions happened, they typically don't represent intentional personal outreach. The <strong>Other Weight</strong> setting in the Settings tab lets you reduce their impact:</p>
            <ul style="line-height:1.8;">
                <li><strong>1.0 (Full weight)</strong> - Count "Other" contacts the same as intentional outreach</li>
                <li><strong>0.5 (Half weight)</strong> - A good middle ground: acknowledges the contact but values intentional effort more</li>
                <li><strong>0 (Ignore)</strong> - Only count contacts that match a configured keyword</li>
            </ul>
            <p class="pb-text-muted" style="font-size:0.85em;">The "Other" badge (<span class="pb-contact-code has-count" style="opacity:0.5;">O(4)</span>) will appear dimmed when weight is below 1.0, visually indicating reduced impact. The weighted total affects both sorting and the Prospect Scorecard's Contact Efforts factor.</p>
        </div>
    </div>
</div>

<!-- CONFIG MODAL -->
<div class="pb-modal-overlay" id="pb-config-modal">
    <div class="pb-modal">
        <div class="pb-modal-header">
            <h3 id="pb-config-modal-title">New Configuration</h3>
            <button class="pb-modal-close" onclick="pbCloseConfigModal()">&times;</button>
        </div>
        <div class="pb-modal-body">
            <input type="hidden" id="pb-cfg-id">

            <!-- Name -->
            <div class="pb-form-group">
                <label>Configuration Name</label>
                <input type="text" class="pb-input" id="pb-cfg-name" placeholder="e.g., Sunday Visitors Q1">
            </div>
            <div class="pb-form-group" style="max-width:200px;">
                <label>Max Prospects</label>
                <input type="number" class="pb-input" id="pb-cfg-max-prospects" placeholder="2000" min="100" step="100">
                <small class="pb-text-muted">Default: 2000. Set lower for faster load.</small>
            </div>

            <!-- Source -->
            <div class="pb-form-group">
                <label>Prospect Source</label>
                <select class="pb-select" id="pb-cfg-source-type" onchange="pbToggleSourceFields()">
                    <option value="involvement">Involvement (Organization)</option>
                    <option value="tag">Tag</option>
                    <option value="query">Saved Query Tag</option>
                </select>
            </div>
            <div id="pb-cfg-source-inv" class="pb-form-group">
                <label>Search Involvement</label>
                <div class="pb-search-wrap">
                    <input type="text" class="pb-input" id="pb-cfg-inv-search" placeholder="Type to search..." oninput="pbSearchInvolvements(this.value)">
                    <div class="pb-search-results" id="pb-cfg-inv-results" style="display:none;"></div>
                </div>
                <input type="hidden" id="pb-cfg-org-id">
                <div id="pb-cfg-inv-selected" class="pb-mt-sm pb-text-sm"></div>
            </div>
            <div id="pb-cfg-source-tag" class="pb-form-group" style="display:none;">
                <label>Tag Name</label>
                <div class="pb-search-wrap">
                    <input type="text" class="pb-input" id="pb-cfg-tag-search" placeholder="Type to search tags..." oninput="pbSearchTags(this.value)">
                    <div class="pb-search-results" id="pb-cfg-tag-results" style="display:none;"></div>
                </div>
                <input type="hidden" id="pb-cfg-tag-name">
                <div id="pb-cfg-tag-selected" class="pb-mt-sm pb-text-sm"></div>
            </div>
            <div id="pb-cfg-source-query" class="pb-form-group" style="display:none;">
                <label>Saved Query Code</label>
                <input type="text" class="pb-input" id="pb-cfg-query-id" placeholder="e.g., Test for RSVP (exact query name from Search Builder)">
            </div>

            <!-- Member Types -->
            <div class="pb-form-group">
                <label>Member Types to Include</label>
                <div id="pb-cfg-mt-list" class="pb-flex pb-gap-md" style="flex-wrap:wrap;">
                    <span class="pb-text-muted">Loading member types...</span>
                </div>
            </div>

            <!-- No Contact Threshold -->
            <div class="pb-form-group">
                <label>No Contact Threshold (days)</label>
                <div class="pb-flex-center pb-gap-sm">
                    <input type="number" class="pb-input" id="pb-cfg-no-contact-days" value="90" min="1" max="730" style="width:100px;">
                    <span class="pb-text-muted pb-text-sm">Family members not attending within this many days are flagged</span>
                </div>
            </div>

            <!-- Display Fields -->
            <div class="pb-form-group">
                <label>Display Fields</label>
                <div id="pb-cfg-fields-list" style="max-height:200px;overflow-y:auto;border:1px solid var(--pb-border);border-radius:var(--pb-radius);padding:8px;"></div>
                <div class="pb-mt-sm pb-flex pb-gap-sm">
                    <select class="pb-select" id="pb-cfg-add-field-type" style="width:auto;" onchange="pbPopulateFieldOptions()">
                        <option value="person">Person</option>
                        <option value="family">Family</option>
                        <option value="involvement">Involvements</option>
                        <option value="medical">Medical</option>
                        <option value="extravalue">Extra Value</option>
                        <option value="regquestion">Reg Question</option>
                    </select>
                    <select class="pb-select" id="pb-cfg-add-field-source" style="width:auto;min-width:150px;"></select>
                    <button class="pb-btn pb-btn-sm" onclick="pbAddDisplayField()">Add Field</button>
                </div>
            </div>

            <!-- Cross-Query Flags -->
            <div class="pb-form-group">
                <label>Cross-Query Flags</label>
                <div id="pb-cfg-flags-list"></div>
                <div class="pb-mt-sm pb-flex pb-gap-sm">
                    <select class="pb-select" id="pb-cfg-add-flag-type" style="width:auto;">
                        <option value="children_attending">Children Attending Program</option>
                        <option value="parents_not_attending">Parents Not Attending Program</option>
                        <option value="spouse_in_org">Spouse in Organization</option>
                        <option value="has_extra_value">Has Extra Value</option>
                    </select>
                    <button class="pb-btn pb-btn-sm" onclick="pbAddCrossFlag()">Add Flag</button>
                </div>
            </div>

            <!-- Target Actions -->
            <div class="pb-form-group">
                <label>Target Actions (where to place processed prospects)</label>
                <div id="pb-cfg-actions-list"></div>
                <div class="pb-mt-sm">
                    <div class="pb-flex pb-gap-sm pb-mb-sm">
                        <select class="pb-select" id="pb-cfg-add-action-type" style="width:auto;" onchange="pbToggleActionFields()">
                            <option value="tag">Add to Tag</option>
                            <option value="involvement">Add to Involvement(s)</option>
                        </select>
                    </div>
                    <!-- Tag name input -->
                    <div id="pb-action-tag-fields" style="display:none;">
                        <div class="pb-flex pb-gap-sm">
                            <input type="text" class="pb-input" id="pb-action-tag-name" placeholder="Enter tag name" style="flex:1;">
                            <button class="pb-btn pb-btn-sm pb-btn-success" onclick="pbAddTagAction()">Add Tag</button>
                        </div>
                    </div>
                    <!-- Involvement multi-select (click to add) -->
                    <div id="pb-action-inv-fields" style="display:none;">
                        <p class="pb-text-muted pb-text-sm" style="margin:0 0 6px 0;">Select member type, filter by program/division, then click involvements to add.</p>
                        <div class="pb-flex pb-gap-sm pb-mb-sm" style="flex-wrap:wrap;">
                            <!-- Hardcoded ids removed (220/230/etc were default-seed ids,
                                 wrong on churches that renamed lookup.MemberType -- e.g. FBCH
                                 has Prospect=311, not 230). pbUpdateActionMemberTypeDropdown()
                                 rebuilds this from lookup.MemberType on every modal open. -->
                            <select class="pb-select" id="pb-action-member-type" style="width:auto;min-width:150px;">
                                <option value="" disabled selected>Loading...</option>
                            </select>
                            <select class="pb-select" id="pb-action-prog-filter" style="width:auto;min-width:140px;" onchange="pbActionProgChanged()">
                                <option value="">All Programs</option>
                            </select>
                            <select class="pb-select" id="pb-action-div-filter" style="width:auto;min-width:140px;" onchange="pbActionSearchInv()">
                                <option value="">All Divisions</option>
                            </select>
                        </div>
                        <input type="text" class="pb-input pb-mb-sm" id="pb-action-inv-search" placeholder="Search involvements..." oninput="pbActionSearchInv()">
                        <div id="pb-action-inv-results" style="max-height:220px;overflow-y:auto;border:1px solid var(--pb-border);border-radius:var(--pb-radius);"></div>
                    </div>
                </div>
            </div>
        </div>
        <div class="pb-modal-footer" style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
            <!-- Left-side delete is visually separated from Save/Cancel so
                 a misclick is less likely. Only shows when editing an
                 existing configuration (pbShowConfigModal toggles display). -->
            <button id="pb-config-modal-delete" class="pb-btn pb-btn-sm pb-btn-danger" onclick="pbConfigModalDelete()" style="display:none;">Delete configuration...</button>
            <div style="display:flex;gap:8px;">
                <button class="pb-btn" onclick="pbCloseConfigModal()">Cancel</button>
                <button class="pb-btn pb-btn-primary" onclick="pbSaveConfig()">Save Configuration</button>
            </div>
        </div>
    </div>
</div>

<!-- CONTACT LOG MODAL -->
<div class="pb-modal-overlay" id="pb-contact-modal">
    <div class="pb-modal" style="max-width:500px;">
        <div class="pb-modal-header">
            <h3>Log Contact</h3>
            <button class="pb-modal-close" onclick="pbCloseContactModal()">&times;</button>
        </div>
        <div class="pb-modal-body">
            <input type="hidden" id="pb-contact-pid">
            <div class="pb-form-group">
                <label>Contact Method</label>
                <select class="pb-select" id="pb-contact-keyword"></select>
            </div>
            <div class="pb-form-group">
                <label>Notes</label>
                <textarea class="pb-textarea" id="pb-contact-notes" placeholder="Describe the contact..."></textarea>
            </div>
        </div>
        <div class="pb-modal-footer">
            <button class="pb-btn" onclick="pbCloseContactModal()">Cancel</button>
            <button class="pb-btn pb-btn-success" onclick="pbSubmitContact()">Log Contact</button>
        </div>
    </div>
</div>

<!-- GROUP CREATE/EDIT MODAL -->
<div class="pb-modal-overlay" id="pb-grp-modal">
    <div class="pb-modal" style="max-width:550px;">
        <div class="pb-modal-header">
            <h3 id="pb-grp-modal-title">New Group</h3>
            <button class="pb-modal-close" onclick="pbGrpCloseModal()">&times;</button>
        </div>
        <div class="pb-modal-body">
            <input type="hidden" id="pb-grp-edit-id">
            <div class="pb-form-group">
                <label>Group Name</label>
                <input type="text" class="pb-input" id="pb-grp-edit-name" placeholder="e.g., Youth Outreach">
            </div>
            <div class="pb-form-group">
                <label>Description</label>
                <textarea class="pb-textarea" id="pb-grp-edit-desc" placeholder="What is this group focused on?"></textarea>
            </div>

            <!-- Ministry Level -->
            <div class="pb-form-group">
                <label>Ministry Level</label>
                <select class="pb-select" id="pb-grp-edit-level" onchange="pbGrpLevelChanged()">
                    <option value="program">Program</option>
                    <option value="division">Program / Division</option>
                    <option value="involvement">Involvement</option>
                </select>
            </div>
            <div class="pb-form-group">
                <label>Program</label>
                <select class="pb-select" id="pb-grp-edit-program" onchange="pbGrpProgramChanged()">
                    <option value="">Loading programs...</option>
                </select>
            </div>
            <div class="pb-form-group" id="pb-grp-div-row" style="display:none;">
                <label>Division</label>
                <select class="pb-select" id="pb-grp-edit-division">
                    <option value="">-- Select Division --</option>
                </select>
            </div>
            <div class="pb-form-group" id="pb-grp-inv-row" style="display:none;">
                <label>Involvement</label>
                <div class="pb-search-wrap">
                    <input type="text" class="pb-input" id="pb-grp-edit-inv-search" placeholder="Type to search involvements..." oninput="pbGrpSearchInvolvements(this.value)">
                    <div class="pb-search-results" id="pb-grp-inv-results" style="display:none;"></div>
                </div>
                <input type="hidden" id="pb-grp-edit-org-id">
                <div id="pb-grp-inv-selected" class="pb-mt-sm pb-text-sm"></div>
            </div>

            <!-- Thresholds -->
            <div style="display:flex;gap:12px;">
                <div class="pb-form-group" style="flex:1;">
                    <label>Min Days Enrolled</label>
                    <input type="number" class="pb-input" id="pb-grp-edit-min-days" value="0" min="0" max="730">
                    <small class="pb-text-muted">Only show people enrolled this many days+</small>
                </div>
                <div class="pb-form-group" style="flex:1;">
                    <label>Min Days No Contact</label>
                    <input type="number" class="pb-input" id="pb-grp-edit-stale-days" value="0" min="0" max="730">
                    <small class="pb-text-muted">Only show with no contact in this many days (0=all)</small>
                </div>
            </div>

            <!-- Member Types -->
            <div class="pb-form-group">
                <label>Member Types to Include</label>
                <div id="pb-grp-mt-list" class="pb-flex pb-gap-md" style="flex-wrap:wrap;">
                    <span class="pb-text-muted">Loading...</span>
                </div>
                <small class="pb-text-muted">Select which organization member types to show. Leave unchecked for all.</small>
            </div>

            <!-- Target Actions -->
            <div class="pb-form-group">
                <label>Target Actions (where to place prospects)</label>
                <div id="pb-grp-actions-list"></div>
                <div class="pb-mt-sm">
                    <div class="pb-flex pb-gap-sm pb-mb-sm">
                        <select class="pb-select" id="pb-grp-add-action-type" style="width:auto;" onchange="pbGrpToggleActionFields()">
                            <option value="tag">Add to Tag</option>
                            <option value="involvement">Add to Involvement</option>
                        </select>
                    </div>
                    <div id="pb-grp-action-tag-fields" style="display:none;">
                        <div class="pb-flex pb-gap-sm">
                            <input type="text" class="pb-input" id="pb-grp-action-tag-name" placeholder="Enter tag name" style="flex:1;">
                            <button class="pb-btn pb-btn-sm pb-btn-success" onclick="pbGrpAddTagAction()">Add</button>
                        </div>
                    </div>
                    <div id="pb-grp-action-inv-fields" style="display:none;">
                        <div class="pb-flex pb-gap-sm pb-mb-sm" style="flex-wrap:wrap;">
                            <!-- Same hardcoded-id bug as the main action dropdown.
                                 Rebuilt from lookup.MemberType via pbUpdateActionMemberTypeDropdown. -->
                            <select class="pb-select" id="pb-grp-action-member-type" style="width:auto;min-width:130px;">
                                <option value="" disabled selected>Loading...</option>
                            </select>
                        </div>
                        <input type="text" class="pb-input pb-mb-sm" id="pb-grp-action-inv-search" placeholder="Search involvements..." oninput="pbGrpActionSearchInv(this.value)">
                        <div id="pb-grp-action-inv-results" style="max-height:150px;overflow-y:auto;border:1px solid var(--pb-border);border-radius:var(--pb-radius);"></div>
                    </div>
                </div>
            </div>

            <div class="pb-form-group">
                <label>Color</label>
                <div class="pb-grp-color-swatches" id="pb-grp-color-swatches">
                    <div class="pb-grp-color-swatch selected" style="background:#3498db;" data-color="#3498db" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#27ae60;" data-color="#27ae60" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#e74c3c;" data-color="#e74c3c" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#f39c12;" data-color="#f39c12" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#9b59b6;" data-color="#9b59b6" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#1abc9c;" data-color="#1abc9c" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#e67e22;" data-color="#e67e22" onclick="pbGrpSelectColor(this)"></div>
                    <div class="pb-grp-color-swatch" style="background:#2c3e50;" data-color="#2c3e50" onclick="pbGrpSelectColor(this)"></div>
                </div>
            </div>
        </div>
        <div class="pb-modal-footer" style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
            <!-- Left-side delete kept visually separated from the save/cancel
                 cluster so it's harder to misclick. Only rendered when
                 editing an existing group (pbGrpShowModal toggles display). -->
            <button id="pb-grp-modal-delete" class="pb-btn pb-btn-sm pb-btn-danger" onclick="pbGrpModalDelete()" style="display:none;">Delete group...</button>
            <div style="display:flex;gap:8px;">
                <button class="pb-btn" onclick="pbGrpCloseModal()">Cancel</button>
                <button class="pb-btn pb-btn-primary" onclick="pbGrpSaveGroup()">Save Group</button>
            </div>
        </div>
    </div>
</div>

<!-- ASSIGN TO GROUP MODAL -->
<div class="pb-modal-overlay" id="pb-grp-assign-modal">
    <div class="pb-modal" style="max-width:550px;">
        <div class="pb-modal-header">
            <h3>Assign to Group</h3>
            <button class="pb-modal-close" onclick="pbGrpCloseAssignModal()">&times;</button>
        </div>
        <div class="pb-modal-body">
            <input type="hidden" id="pb-grp-assign-pids">
            <div id="pb-grp-assign-names" class="pb-text-muted pb-mb-sm" style="font-size:0.9em;"></div>
            <div class="pb-form-group">
                <label>Select Group</label>
                <select class="pb-select" id="pb-grp-assign-target"></select>
            </div>
            <div class="pb-form-group">
                <label>Notes (optional)</label>
                <textarea class="pb-textarea" id="pb-grp-assign-notes" placeholder="Why are they being assigned?"></textarea>
            </div>
        </div>
        <div class="pb-modal-footer">
            <button class="pb-btn" onclick="pbGrpCloseAssignModal()">Cancel</button>
            <button class="pb-btn pb-btn-success" onclick="pbGrpSubmitAssign()">Assign</button>
        </div>
    </div>
</div>

<!-- LOG EFFORT MODAL -->
<div class="pb-modal-overlay" id="pb-grp-effort-modal">
    <div class="pb-modal" style="max-width:500px;">
        <div class="pb-modal-header">
            <h3>Log Effort</h3>
            <button class="pb-modal-close" onclick="pbGrpCloseEffortModal()">&times;</button>
        </div>
        <div class="pb-modal-body">
            <input type="hidden" id="pb-grp-effort-gid">
            <input type="hidden" id="pb-grp-effort-pid">
            <div id="pb-grp-effort-person" class="pb-mb-sm" style="font-weight:600;color:var(--pb-primary);"></div>
            <div class="pb-form-group">
                <label>Effort Type</label>
                <select class="pb-select" id="pb-grp-effort-type">
                    <option value="phone">Phone Call</option>
                    <option value="email">Email</option>
                    <option value="text">Text Message</option>
                    <option value="visit">Home Visit</option>
                    <option value="event_invite">Event Invite</option>
                    <option value="meeting">In-Person Meeting</option>
                    <option value="other">Other</option>
                </select>
            </div>
            <div class="pb-form-group">
                <label>Result</label>
                <select class="pb-select" id="pb-grp-effort-result">
                    <option value="connected">Connected</option>
                    <option value="no_answer">No Answer</option>
                    <option value="left_message">Left Message</option>
                    <option value="interested">Interested</option>
                    <option value="declined">Declined</option>
                    <option value="follow_up">Needs Follow-Up</option>
                    <option value="other">Other</option>
                </select>
            </div>
            <div class="pb-form-group">
                <label>Description</label>
                <textarea class="pb-textarea" id="pb-grp-effort-desc" placeholder="Describe the interaction..."></textarea>
            </div>
        </div>
        <div class="pb-modal-footer">
            <button class="pb-btn" onclick="pbGrpCloseEffortModal()">Cancel</button>
            <button class="pb-btn pb-btn-success" onclick="pbGrpSubmitEffort()">Log Effort</button>
        </div>
    </div>
</div>

<!-- MOVE PROSPECT MODAL -->
<div class="pb-modal-overlay" id="pb-grp-move-modal">
    <div class="pb-modal" style="max-width:500px;">
        <div class="pb-modal-header">
            <h3>Move Prospect</h3>
            <button class="pb-modal-close" onclick="pbGrpCloseMoveModal()">&times;</button>
        </div>
        <div class="pb-modal-body">
            <input type="hidden" id="pb-grp-move-from">
            <input type="hidden" id="pb-grp-move-pids">
            <div id="pb-grp-move-names" class="pb-mb-sm" style="font-size:0.9em;font-weight:600;color:var(--pb-primary);"></div>
            <div id="pb-grp-move-from-label" class="pb-text-muted pb-mb-sm" style="font-size:0.85em;"></div>
            <div class="pb-form-group">
                <label>Move To Group</label>
                <select class="pb-select" id="pb-grp-move-target"></select>
            </div>
            <div class="pb-form-group">
                <label>Reason / Notes</label>
                <textarea class="pb-textarea" id="pb-grp-move-notes" placeholder="Why are they being moved?"></textarea>
            </div>
        </div>
        <div class="pb-modal-footer">
            <button class="pb-btn" onclick="pbGrpCloseMoveModal()">Cancel</button>
            <button class="pb-btn pb-btn-warning" onclick="pbGrpSubmitMove()">Move</button>
        </div>
    </div>
</div>

<script>
// ============================================================
// PROSPECT BUILDER JAVASCRIPT
// ============================================================
var pbScriptPath = (function() {
    var p = window.location.pathname;
    var parts = p.split('/');
    for (var i = 0; i < parts.length; i++) {
        if (parts[i] === 'PyScriptForm' || parts[i] === 'PyScript') {
            if (i + 1 < parts.length) {
                return '/PyScriptForm/' + parts[i + 1].split('?')[0];
            }
        }
    }
    return p;
})();

// ============================================================
// AUTO-UPDATE (see TPxi/AutoUpdate/README.md)
// ============================================================
// SCRIPT_NAME is detected from the URL the browser actually loaded —
// resilient to renames and works even if model.URL parsing on the
// server returned something unexpected. Forwarded on every pbAjax call.
var SCRIPT_NAME = (function() {
    try {
        var m = (window.location.pathname || '').match(/\\/PyScript(?:Form)?\\/([^\\/?#]+)/);
        if (m && m[1]) return m[1];
    } catch(e) {}
    return ''' + json.dumps(DC_SCRIPT_ID) + ''';
})();
var APP_VERSION = ''' + json.dumps(APP_VERSION) + ''';
var DC_SCRIPT_ID = ''' + json.dumps(DC_SCRIPT_ID) + ''';
var DC_API_BASE = ''' + json.dumps(DC_API_BASE) + ''';
var APP_UPDATE_AVAILABLE = false;
var APP_LATEST_VERSION = '';

function checkForAppUpdate() {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', DC_API_BASE + '/script-versions', true);
        xhr.timeout = 5000;
        xhr.onreadystatechange = function() {
            if (xhr.readyState !== 4 || xhr.status !== 200) return;
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
    var b = document.getElementById('appUpdateBanner');
    if (!b || !APP_UPDATE_AVAILABLE) return;
    var h = '';
    h += '<div style="font-size:18px">&#128640;</div>';
    h += '<div style="flex:1;font-size:12px;color:#0078d4">';
    h += '<strong>Prospect Builder update available</strong>';
    h += ' &mdash; you have <code>v' + APP_VERSION + '</code>, latest is <code>v' + APP_LATEST_VERSION + '</code>. All saved configurations, groups, sessions, and senders are preserved.';
    h += '</div>';
    h += '<button id="appUpdateBtn" onclick="applyAppUpdate()" style="white-space:nowrap;padding:6px 14px;background:#0078d4;color:#fff;border:0;border-radius:4px;cursor:pointer;">Update Now</button>';
    b.innerHTML = h;
    b.style.display = 'flex';
}

window.applyAppUpdate = function() {
    if (!confirm('Update Prospect Builder from v' + APP_VERSION + ' to v' + APP_LATEST_VERSION + '?\\n\\nYour saved data is stored in separate content slots and will be preserved.')) return;
    var btn = document.getElementById('appUpdateBtn');
    if (btn) { btn.disabled = true; btn.innerHTML = 'Updating...'; }
    pbAjax({action: 'apply_update'}, function(d) {
        if (d && d.success) {
            alert((d.message || 'Updated!') + ' Reloading...');
            window.location.reload(true);
        } else {
            alert('Update failed: ' + ((d && d.message) || 'Unknown error'));
            if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
        }
    });
};

// ============================================================
// STATE
// ============================================================
var pbConfigs = [];
var pbProspects = [];
var pbFilteredProspects = [];
var pbCurrentConfig = null;
var pbCurrentIndex = 0;
var pbProcessedMap = {};
var pbDeferredSet = {};
var pbSkippedSet = {};
var pbCrossFlags = {};
var pbContactMap = {};
var pbFamilyData = {};
var pbExtraValues = {};
var pbRegData = {};
var pbInvData = {};
var pbViewMode = 'list';
var pbCurrentSession = null;
var pbBatchSelected = {};
var pbSettings = {};
var pbDetailCache = {};  // pid -> {family, involvements, extraValues, regData}
var pbScoreMap = {};     // pid -> {score, breakdown}
var pbExpandedPid = null;

// Group Management state
var pbGrpGroups = [];
var pbGrpAssignments = {};
var pbGrpEfforts = [];
var pbGrpChanges = [];
var pbGrpStats = {};
var pbGrpCurrentGroup = null;
var pbGrpDetailTab = 'prospects';
var pbGrpDetailProspects = [];

// ============================================================
// AJAX
// ============================================================
function pbAjax(data, callback) {
    // Always forward the URL-detected install name so the server-side
    // get_pb_script_name() can route apply_update / ScheduledTasks to
    // the actual installed slot (not the hardcoded default).
    data = data || {};
    if (typeof SCRIPT_NAME !== 'undefined' && SCRIPT_NAME && !data.script_name) {
        data.script_name = SCRIPT_NAME;
    }
    var xhr = new XMLHttpRequest();
    xhr.open('POST', pbScriptPath, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status >= 200 && xhr.status < 300) {
            var resp = xhr.responseText;
            try {
                var text = resp;
                var jsonStart = text.indexOf('{');
                var jsonEnd = text.lastIndexOf('}');
                if (jsonStart >= 0 && jsonEnd > jsonStart) {
                    text = text.substring(jsonStart, jsonEnd + 1);
                }
                var d = JSON.parse(text);
                callback(d);
            } catch(e) {
                try {
                    var arrStart = resp.indexOf('[');
                    var arrEnd = resp.lastIndexOf(']');
                    if (arrStart >= 0 && arrEnd > arrStart) {
                        callback(JSON.parse(resp.substring(arrStart, arrEnd + 1)));
                        return;
                    }
                } catch(e2) {}
                console.error('PB parse error:', e, resp.substring(0, 200));
                callback({success: false, message: 'Failed to parse response'});
            }
        } else {
            console.error('PB AJAX error:', xhr.status);
            callback({success: false, message: 'Request failed: ' + xhr.status});
        }
    };
    // Encode data object to form-urlencoded string
    var parts = [];
    for (var key in data) {
        if (data.hasOwnProperty(key)) {
            parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(data[key]));
        }
    }
    xhr.send(parts.join('&'));
}

// ============================================================
// TAB MANAGEMENT
// ============================================================
// ---- Per-tab intro banners ----------------------------------------------
// Show by default; remember dismiss per-tab in localStorage so a user only
// has to close each banner once. Keyed under a single stable namespace so
// future tabs slot in without colliding.
var PB_INTRO_STORAGE_KEY = 'pb_intro_dismissed_v1';

function pbGetDismissedIntros() {
    try {
        var raw = localStorage.getItem(PB_INTRO_STORAGE_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch(e) { return {}; }
}

function pbDismissIntro(introId) {
    var d = pbGetDismissedIntros();
    d[introId] = true;
    try { localStorage.setItem(PB_INTRO_STORAGE_KEY, JSON.stringify(d)); } catch(e) {}
    var el = document.querySelector('.pb-tab-intro[data-intro-id="' + introId + '"]');
    if (el) el.style.display = 'none';
}

function pbApplyIntroVisibility() {
    var dismissed = pbGetDismissedIntros();
    var intros = document.querySelectorAll('.pb-tab-intro');
    for (var i = 0; i < intros.length; i++) {
        var id = intros[i].getAttribute('data-intro-id');
        intros[i].style.display = dismissed[id] ? 'none' : '';
    }
}

function pbSwitchTab(tab) {
    document.querySelectorAll('.pb-tab').forEach(function(t) { t.classList.remove('active'); });
    document.querySelectorAll('.pb-tab-content').forEach(function(t) { t.classList.remove('active'); });
    document.querySelector('.pb-tab[data-tab="' + tab + '"]').classList.add('active');
    document.getElementById('pb-tab-' + tab).classList.add('active');
    // Re-apply intro visibility -- new intros may have been freshly rendered
    // (or the active tab just changed and only its intros are now visible).
    pbApplyIntroVisibility();

    if (tab === 'configs') {
        // If no config is active, show overview; otherwise keep workspace visible
        if (!pbCurrentConfig || document.getElementById('pb-configs-workspace').style.display === 'none') {
            document.getElementById('pb-configs-overview').style.display = '';
            document.getElementById('pb-configs-workspace').style.display = 'none';
            pbLoadConfigs();
        }
    }
    if (tab === 'sessions') {
        // Populate config filter dropdown
        var sel = document.getElementById('pb-log-config-filter');
        sel.innerHTML = '<option value="">All Configurations</option>';
        for (var i = 0; i < pbConfigs.length; i++) {
            sel.innerHTML += '<option value="' + pbConfigs[i].id + '">' + pbConfigs[i].name + '</option>';
        }
        // Populate group filter dropdown
        var gsel = document.getElementById('pb-log-group-filter');
        gsel.innerHTML = '<option value="">All Groups</option>';
        for (var gi = 0; gi < pbGrpGroups.length; gi++) {
            gsel.innerHTML += '<option value="' + pbGrpGroups[gi].id + '">' + pbGrpGroups[gi].name + '</option>';
        }
        pbLoadActivityLog();
    }
    if (tab === 'groups') pbGrpInitTab();
    if (tab === 'senders') pbSndInitTab();
    if (tab === 'settings') pbLoadSettings();
    if (tab === 'dashboard') pbDashInitTab();
}

// ============================================================
// DASHBOARD TAB
// ============================================================
var pbDashChartTouches = null;
var pbDashChartConversions = null;
var pbDashChartGroups = null;
var pbDashLoading = false;
var pbDashLastPayload = null;
var pbDashInitialized = false;
var pbDashChartReadyTimer = null;
// Stashed Chart 3 payload for the "Show numbers" toggle. Set in the
// chart-init block so the toggle can build the breakdown table without
// re-fetching anything.
var pbDashGroupConversion = null;

function pbDashInitTab() {
    // Wire date inputs to defaults (last 30 days) on first open. Subsequent
    // opens preserve whatever the user typed.
    var startInput = document.getElementById('pb-dash-start');
    var endInput = document.getElementById('pb-dash-end');
    if (!startInput || !endInput) return;
    if (!pbDashInitialized) {
        var today = new Date();
        var start = new Date(today.getTime() - (29 * 24 * 60 * 60 * 1000));
        endInput.value = pbDashFmtDate(today);
        startInput.value = pbDashFmtDate(start);
        pbDashInitialized = true;
    }
    pbDashEnsureChartJs(function() { pbDashRefresh(); });
}

function pbDashFmtDate(d) {
    var y = d.getFullYear();
    var m = ('0' + (d.getMonth() + 1)).slice(-2);
    var dd = ('0' + d.getDate()).slice(-2);
    return y + '-' + m + '-' + dd;
}

function pbDashEnsureChartJs(callback) {
    if (window.Chart) { callback(); return; }
    // Already injected, just waiting?
    if (document.getElementById('pb-chartjs-script')) {
        var tries = 0;
        var poll = setInterval(function() {
            tries++;
            if (window.Chart) { clearInterval(poll); callback(); }
            else if (tries > 50) { clearInterval(poll); /* give up, leave cards filled */ }
        }, 100);
        return;
    }
    var s = document.createElement('script');
    s.id = 'pb-chartjs-script';
    s.src = 'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js';
    s.onload = function() { callback(); };
    s.onerror = function() { /* charts will silently fail; cards still work */ };
    document.head.appendChild(s);
}

function pbDashRefresh() {
    if (pbDashLoading) return;
    pbDashLoading = true;
    var start = document.getElementById('pb-dash-start').value || '';
    var end = document.getElementById('pb-dash-end').value || '';
    var asOf = document.getElementById('pb-dash-asof');
    var pill = document.getElementById('pb-dash-loading-pill');
    var content = document.getElementById('pb-dash-content');
    if (pill) pill.style.display = 'inline-flex';
    if (content) content.classList.add('pb-dash-loading');
    if (asOf) asOf.textContent = '';

    // Fire KPIs and Charts in PARALLEL so cards land in ~3s instead of
    // waiting on the (heavier) chart aggregations. Chart 3 reads pre-computed
    // group metrics from cache when available; first load without warm cache
    // can still take a moment.
    var kpisDone = false, chartsDone = false;
    function maybeDone() {
        if (kpisDone && chartsDone) {
            pbDashLoading = false;
            if (pill) pill.style.display = 'none';
            if (content) content.classList.remove('pb-dash-loading');
            if (asOf) {
                var now = new Date();
                asOf.textContent = 'As of ' + now.toISOString().substring(0,10) + ' ' +
                                   now.toTimeString().substring(0,5);
            }
        }
    }

    // ---- KPIs (cards) ----
    pbAjax({action: 'load_dashboard_data', start: start, end: end, parts: 'kpis'}, function(d) {
        kpisDone = true;
        if (!d || !d.success) {
            if (asOf) asOf.textContent = 'Error loading KPIs';
            maybeDone();
            return;
        }
        pbDashLastPayload = d;
        pbDashRenderKpis(d, start, end);
        // Pill keeps spinning until charts arrive -- no per-step text needed.
        maybeDone();
    });

    // ---- Charts (parallel) ----
    pbAjax({action: 'load_dashboard_data', start: start, end: end, parts: 'charts'}, function(d) {
        chartsDone = true;
        if (!d || !d.success) {
            if (asOf && kpisDone) asOf.textContent = 'KPIs loaded, charts errored';
            maybeDone();
            return;
        }
        pbDashEnsureChartJs(function() { pbDashRenderCharts(d); });
        maybeDone();
    });
}

function pbDashFmtPct(v) {
    if (v === null || v === undefined || isNaN(v)) return '-';
    return (Math.round(v * 10) / 10).toFixed(1) + '%';
}
function pbDashFmtNum(v) {
    if (v === null || v === undefined || isNaN(v)) return '-';
    return Number(v).toLocaleString();
}

function pbDashRenderKpis(d, start, end) {
    var k = d.kpis || {};
    var customRange = pbDashIsCustomRange(start, end);

    // Empty state: when no groups exist or the org scope resolves to nothing,
    // show the "get started" panel instead of a wall of zeros.
    var emptyEl = document.getElementById('pb-dash-empty');
    var contentEl = document.getElementById('pb-dash-content');
    var hasGroups = (k.hasGroups === true);
    var hasScope  = (k.hasOrgScope !== false);  // default to true if missing
    if (emptyEl && contentEl) {
        if (!hasGroups || !hasScope) {
            emptyEl.style.display = '';
            contentEl.style.display = 'none';
            // Also clear the loading pill since there's nothing else to show.
            var pill = document.getElementById('pb-dash-loading-pill');
            if (pill) pill.style.display = 'none';
            return;
        } else {
            emptyEl.style.display = 'none';
            contentEl.style.display = '';
        }
    }

    // Override the "Overdue days" sub-text from server-provided value.
    var od = document.getElementById('pb-dash-overdue-days');
    if (od && d.overdueDays) od.textContent = d.overdueDays;

    function setCard(key, value, sub, isRangeSensitive) {
        var card = document.querySelector('.pb-dash-card[data-key="' + key + '"]');
        if (!card) return;
        var vEl = card.querySelector('.pb-dash-value');
        var sEl = card.querySelector('.pb-dash-sub');
        if (vEl) vEl.textContent = value;
        if (sEl) {
            // Only swap the sub-line for cards driven by the user's custom range.
            if (isRangeSensitive && customRange) {
                sEl.textContent = '(custom range)';
                sEl.classList.add('pb-dash-range-hint');
            } else if (sub) {
                sEl.textContent = sub;
                sEl.classList.remove('pb-dash-range-hint');
            }
        }
    }

    // Card 1: Active Prospects (point-in-time, NOT range-sensitive).
    // Surface BOTH unit counts so users can reconcile vs. Group Management's
    // per-membership 'Prospects' total (which counts each P+O once).
    var apSub = 'Distinct people in prospect status';
    if (k.activeProspectMemberships !== undefined && k.activeProspectMemberships !== null) {
        apSub = pbDashFmtNum(k.activeProspects) + ' people · ' +
                pbDashFmtNum(k.activeProspectMemberships) + ' memberships';
    }
    setCard('activeProspects', pbDashFmtNum(k.activeProspects), apSub, false);

    // Card 2: Touches Sent. Use range value when custom, else 7d.
    if (customRange) {
        setCard('touches7d', pbDashFmtNum(k.touchesRange), null, true);
    } else {
        setCard('touches7d', pbDashFmtNum(k.touches7d), 'Contact-method touches, 7 days', true);
    }

    // Card 3: Touches Sent (30d) - swap to range when custom.
    if (customRange) {
        setCard('touches30d', pbDashFmtNum(k.touchesRange), null, true);
    } else {
        setCard('touches30d', pbDashFmtNum(k.touches30d), 'Contact-method touches, 30 days', true);
    }

    // Card 4: New Conversions
    if (customRange) {
        setCard('newConversions30d', pbDashFmtNum(k.newConversionsRange), null, true);
    } else {
        setCard('newConversions30d', pbDashFmtNum(k.newConversions30d), 'Distinct people converted', true);
    }

    // Card 5: Overdue (point-in-time, NOT range-sensitive)
    setCard('overdueCount', pbDashFmtNum(k.overdueCount),
            'No touch in ' + (d.overdueDays || 14) + '+ days', false);

    // Card 6: Conversion Rate (per involvement; pair-based)
    if (customRange) {
        setCard('conversionRate90d', pbDashFmtPct(k.conversionRateRange), null, true);
    } else {
        setCard('conversionRate90d', pbDashFmtPct(k.conversionRate90d), 'Per involvement (Person × Involvement)', true);
    }

    // Card 7: Avg Involvements Tried Before Conversion (90d, point-in-time)
    if (k.avgOrgsBeforeConv90d !== undefined && k.avgOrgsBeforeConv90d !== null) {
        var avgVal = Number(k.avgOrgsBeforeConv90d).toFixed(2);
        var sample = k.avgOrgsBeforeConvSample90d || 0;
        var sub = sample > 0
            ? ('Across ' + sample + ' converted ' + (sample === 1 ? 'person' : 'people'))
            : 'No conversions yet in window';
        setCard('avgOrgsBeforeConv90d', avgVal, sub, false);
    } else {
        setCard('avgOrgsBeforeConv90d', '-', 'Involvements tried as prospect', false);
    }

    var asOf = document.getElementById('pb-dash-asof');
    if (asOf) {
        var when = d.asOf ? d.asOf.substring(0, 16).replace('T', ' ') : '';
        asOf.textContent = when ? ('As of ' + when) : '';
    }
}

function pbDashIsCustomRange(startStr, endStr) {
    if (!startStr || !endStr) return false;
    var today = new Date();
    var defaultStart = new Date(today.getTime() - (29 * 24 * 60 * 60 * 1000));
    return (startStr !== pbDashFmtDate(defaultStart)) || (endStr !== pbDashFmtDate(today));
}

function pbDashRenderCharts(d) {
    if (!window.Chart) return;
    var charts = d.charts || {};

    // ---- Chart 1: Weekly Touches, STACKED Notes vs Tasks ----
    // Notes (IsNote=1) = completed touches -- the "real" outreach signal.
    // Tasks (IsNote=0) = pending intentions. Stacking shows BOTH so staff
    // can see follow-through ratios at a glance.
    var t = charts.dailyTouches || {labels: [], notes: [], tasks: [], data: []};
    var canvas1 = document.getElementById('pb-dash-chart-touches');
    if (canvas1) {
        if (pbDashChartTouches) pbDashChartTouches.destroy();
        var stackOpts = pbDashChartOptions('Touches');
        stackOpts.scales = stackOpts.scales || {};
        stackOpts.scales.x = Object.assign({}, stackOpts.scales.x || {}, {stacked: true});
        stackOpts.scales.y = Object.assign({}, stackOpts.scales.y || {}, {stacked: true, beginAtZero: true});
        stackOpts.plugins = Object.assign({}, stackOpts.plugins || {}, {
            legend: {display: true, position: 'top'},
            tooltip: {mode: 'index', intersect: false}
        });
        pbDashChartTouches = new Chart(canvas1.getContext('2d'), {
            type: 'bar',
            data: {
                labels: t.labels,
                datasets: [
                    {
                        label: 'Notes',
                        data: t.notes || t.data || [],
                        backgroundColor: 'rgba(54, 162, 235, 0.7)',
                        borderColor: 'rgba(54, 162, 235, 1)',
                        borderWidth: 1,
                        stack: 'touches'
                    },
                    {
                        label: 'Tasks',
                        data: t.tasks || [],
                        backgroundColor: 'rgba(150, 150, 150, 0.55)',
                        borderColor: 'rgba(110, 110, 110, 1)',
                        borderWidth: 1,
                        stack: 'touches'
                    }
                ]
            },
            options: stackOpts
        });
    }

    // ---- Chart 2: Daily Conversions (line, green) ----
    var c = charts.dailyConversions || {labels: [], data: []};
    var canvas2 = document.getElementById('pb-dash-chart-conversions');
    if (canvas2) {
        if (pbDashChartConversions) pbDashChartConversions.destroy();
        pbDashChartConversions = new Chart(canvas2.getContext('2d'), {
            type: 'line',
            data: {
                labels: c.labels,
                datasets: [{
                    label: 'Conversions',
                    data: c.data,
                    backgroundColor: 'rgba(75, 192, 75, 0.6)',
                    borderColor: 'rgb(75, 192, 75)',
                    borderWidth: 2,
                    fill: true,
                    tension: 0.25
                }]
            },
            options: pbDashChartOptions('Conversions')
        });
    }

    // ---- Chart 3: Conversion Rate by Group -- 90-day rolling ----
    // Multi-line, X = last 6 month-ends (YYYY-MM label), Y = per-involvement
    // conversion % over the 90-day window ending at that month-end. Same
    // math as Card 6, so the latest point matches the card value. One
    // line per group so trend over time is visible.
    var g = charts.groupConversion || {months: [], windows: [], series: []};
    pbDashGroupConversion = g;
    // If the details panel is currently open, repaint it with the fresh data
    // (otherwise it would still show last refresh's numbers).
    var detailsPanelEl = document.getElementById('pb-dash-groups-details');
    if (detailsPanelEl && detailsPanelEl.style.display !== 'none') {
        pbDashRenderGroupDetails();
    }
    var canvas3 = document.getElementById('pb-dash-chart-groups');
    if (canvas3) {
        if (pbDashChartGroups) pbDashChartGroups.destroy();
        // Stable color palette
        var palette = [
            'rgba(155, 89, 182, 1)',   // purple
            'rgba(52, 152, 219, 1)',   // blue
            'rgba(46, 204, 113, 1)',   // green
            'rgba(230, 126, 34, 1)',   // orange
            'rgba(231, 76, 60, 1)',    // red
            'rgba(241, 196, 15, 1)',   // yellow
            'rgba(26, 188, 156, 1)',   // teal
            'rgba(149, 165, 166, 1)'   // gray
        ];
        var series = g.series || [];
        var datasets = series.map(function(s, idx) {
            var color = palette[idx % palette.length];
            return {
                label: s.name || ('Group ' + (idx + 1)),
                data: s.data || [],
                borderColor: color,
                backgroundColor: color.replace(', 1)', ', 0.18)'),
                borderWidth: 2,
                tension: 0.25,
                fill: false,
                pointRadius: 3,
                pointHoverRadius: 5
            };
        });
        pbDashChartGroups = new Chart(canvas3.getContext('2d'), {
            type: 'line',
            data: { labels: g.months || [], datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { display: true, position: 'top' },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                return ctx.dataset.label + ': ' +
                                    (ctx.parsed.y === null ? '-' : ctx.parsed.y.toFixed(1) + '%');
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { callback: function(v) { return v + '%'; } }
                    }
                }
            }
        });
    }
}

function pbDashChartOptions(yLabel) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
            y: { beginAtZero: true, title: { display: true, text: yLabel } },
            x: { ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 8 } }
        }
    };
}

// ---- Chart 3 details panel: "Show numbers" toggle ----
// Surfaces the raw P-pair and C-pair counts that produced each chart point.
// Lets the user distinguish: did the rate jump because of a real conversion
// surge (numerator up), config drift (window membership changed), or
// denominator collapse (active prospect pool shrunk)?
function pbDashToggleGroupDetails() {
    var panel = document.getElementById('pb-dash-groups-details');
    var btn = document.getElementById('pb-dash-groups-toggle-details');
    if (!panel) return;
    if (panel.style.display === 'none') {
        pbDashRenderGroupDetails();
        panel.style.display = 'block';
        if (btn) btn.textContent = 'Hide numbers';
    } else {
        panel.style.display = 'none';
        if (btn) btn.textContent = 'Show numbers';
    }
}

function pbDashRenderGroupDetails() {
    var panel = document.getElementById('pb-dash-groups-details');
    if (!panel) return;
    var gc = pbDashGroupConversion;
    if (!gc || !gc.series || gc.series.length === 0) {
        panel.innerHTML = '<div style="color:#888;padding:8px 0;">No group data loaded yet.</div>';
        return;
    }
    var months = gc.months || [];
    var windows = gc.windows || [];
    var html = '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
    html += '<thead><tr style="background:#f5f5f5;border-bottom:2px solid #ddd;">';
    html += '<th style="padding:6px 8px;text-align:left;">Group</th>';
    html += '<th style="padding:6px 8px;text-align:left;">Metric</th>';
    for (var i = 0; i < months.length; i++) {
        var title = windows[i] ? ('Window: ' + windows[i]) : '';
        html += '<th style="padding:6px 8px;text-align:right;" title="' + title + '">' + months[i] + '</th>';
    }
    html += '</tr></thead><tbody>';
    for (var s = 0; s < gc.series.length; s++) {
        var ser = gc.series[s];
        var name = ser.name || ('Group ' + (s + 1));
        var rates = ser.data || [];
        var pps = ser.prospectPairs || [];
        var cps = ser.conversionPairs || [];
        // Three rows per group: rate, conversions, prospect pool size.
        html += '<tr style="border-top:1px solid #eee;background:#fafbff;">';
        html += '<td rowspan="3" style="padding:6px 8px;font-weight:600;vertical-align:top;">' + pbEscapeHtml(name) + '</td>';
        html += '<td style="padding:6px 8px;">Rate</td>';
        for (var r = 0; r < months.length; r++) {
            var rv = (rates[r] === undefined || rates[r] === null) ? '-' : (rates[r].toFixed(1) + '%');
            html += '<td style="padding:6px 8px;text-align:right;font-weight:600;">' + rv + '</td>';
        }
        html += '</tr>';
        html += '<tr>';
        html += '<td style="padding:6px 8px;color:#444;" title="Counted once per (Person, Involvement) conversion landing inside the 90-day window. Subset of the prospect involvements row below.">Conversions (in window)</td>';
        for (var c = 0; c < months.length; c++) {
            var cv = (cps[c] === undefined || cps[c] === null) ? '-' : cps[c];
            html += '<td style="padding:6px 8px;text-align:right;color:#444;">' + cv + '</td>';
        }
        html += '</tr>';
        html += '<tr style="border-bottom:1px solid #ddd;">';
        html += '<td style="padding:6px 8px;color:#444;" title="Every (Person, Involvement) prospect membership that was active at any point during the 90-day window. Includes prospects who later converted or went inactive within the window. Jane being a prospect in two involvements counts as 2. Always larger than the live snapshot below.">Prospect involvements (in window)</td>';
        for (var p = 0; p < months.length; p++) {
            var pv = (pps[p] === undefined || pps[p] === null) ? '-' : pps[p];
            html += '<td style="padding:6px 8px;text-align:right;color:#444;">' + pv + '</td>';
        }
        html += '</tr>';
    }
    html += '</tbody></table>';
    // Live snapshot row — lets users see today's actual prospect pool size
    // beside the 90-day in-window pool sizes, so the "why is the chart
    // denominator so much bigger than the Active Prospects card?" question
    // answers itself (the chart includes prospects who churned in/out
    // during the window).
    var kpi = (pbDashLastPayload && pbDashLastPayload.kpis) || {};
    var liveDistinct = (kpi.activeProspects !== undefined) ? kpi.activeProspects : null;
    var liveMemberships = (kpi.activeProspectMemberships !== undefined) ? kpi.activeProspectMemberships : null;
    if (liveDistinct !== null || liveMemberships !== null) {
        html += '<div style="margin-top:10px;padding:8px 10px;background:#f7f9fc;border:1px solid #e2e8f0;border-radius:4px;font-size:12px;color:#333;">';
        html += '<strong>Today (live snapshot, all groups combined):</strong> ';
        if (liveDistinct !== null) html += pbDashFmtNum(liveDistinct) + ' distinct prospects';
        if (liveDistinct !== null && liveMemberships !== null) html += ' &middot; ';
        if (liveMemberships !== null) html += pbDashFmtNum(liveMemberships) + ' prospect involvements';
        html += '. The 90-day numbers above are larger because they include prospects who converted or went inactive within the window.';
        html += '</div>';
    }
    html += '<div style="margin-top:8px;color:#666;font-size:11px;">';
    html += 'Each column is a 90-day rolling window ending at the month label. Hover a column header to see the exact date range. ';
    html += 'Rate = conversions &divide; prospect involvements, both counted <em>per involvement</em> (Jane being a prospect in 2 involvements counts as 2). ';
    html += 'A rising rate driven by <strong>rising conversions</strong> is real outreach impact; a rising rate driven by <strong>falling prospect involvements</strong> is denominator collapse (people aged out, were marked inactive, or memberType changed).';
    html += '</div>';
    panel.innerHTML = html;
}

// Minimal HTML-escape for group names that might contain user-entered chars.
function pbEscapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ============================================================
// PROSPECT SENDER
// ============================================================
var pbSndSenders = [];
var pbSndRoles = [];
var pbSndMemberTypes = [];
var pbSndMemberTypesLoading = false;
var pbSndEditingId = null;

function pbSndInitTab() {
    pbSndLoadSenders();
    pbSndLoadRoles();
    pbSndLoadMemberTypes();
    pbSndLoadLog();
    // Pre-warm program/division names so the Source column renders friendly
    // labels instead of "Program #1115" on the first paint.
    if (pbSndPrograms.length === 0) pbSndLoadProgramsDivisions();
}

function pbSndSourceLabel(src) {
    // Decode whatever source shape this sender has into a human-readable
    // string for the Senders list. Senders can be created with either the
    // newer `scope` model ('org', 'program', 'division') OR the legacy
    // `source_type` model ('involvement', 'tag', 'saved_search').
    if (!src) return '(none)';
    var scope = src.scope || src.source_type || '';
    if (!scope) return '(none)';
    if (scope === 'org' || scope === 'involvement') {
        return 'Involvement: ' + (src.org_name || src.orgName || '#' + (src.org_id || src.orgId || '?'));
    }
    if (scope === 'program' || scope === 'division') {
        var progName = '#' + (src.program_id || '?');
        for (var i = 0; i < pbSndPrograms.length; i++) {
            if (String(pbSndPrograms[i].id) === String(src.program_id)) {
                progName = pbSndPrograms[i].name; break;
            }
        }
        if (scope === 'division') {
            var divName = '#' + (src.division_id || '?');
            for (var j = 0; j < pbSndDivisions.length; j++) {
                if (String(pbSndDivisions[j].id) === String(src.division_id)) {
                    divName = pbSndDivisions[j].name; break;
                }
            }
            return 'Program/Division: ' + progName + ' / ' + divName;
        }
        return 'Program: ' + progName;
    }
    if (scope === 'tag')          return 'Tag: ' + (src.tag_name || src.tagName || '#' + (src.tag_id || ''));
    if (scope === 'saved_search') return 'Search: ' + (src.search_name || src.searchName || '#' + (src.search_id || ''));
    return scope; // last-ditch fallback: at least name the scope key
}

function pbSndLoadMemberTypes() {
    if (pbSndMemberTypes.length > 0 || pbSndMemberTypesLoading) return;
    pbSndMemberTypesLoading = true;
    pbAjax({action: 'get_member_types'}, function(d) {
        pbSndMemberTypesLoading = false;
        if (d.success) pbSndMemberTypes = d.member_types || [];
    });
}

function pbSndLoadSenders() {
    pbAjax({action: 'load_senders'}, function(d) {
        if (d.success) {
            pbSndSenders = d.senders || [];
            pbSndRenderList();
        }
    });
}

function pbSndLoadRoles() {
    pbAjax({action: 'get_roles'}, function(d) {
        if (d.success) pbSndRoles = d.roles || [];
    });
}

function pbSndRenderList() {
    var el = document.getElementById('pb-snd-list');
    var empty = document.getElementById('pb-snd-empty');
    if (pbSndSenders.length === 0) {
        el.innerHTML = '';
        empty.style.display = 'block';
        document.getElementById('pb-snd-count').textContent = '';
        return;
    }
    empty.style.display = 'none';
    document.getElementById('pb-snd-count').textContent = pbSndSenders.length;

    var html = '<div class="pb-table-wrap"><table class="pb-table"><thead><tr>';
    html += '<th>Name</th><th>Source</th><th>Schedule</th><th>Recipients</th><th>Last Sent</th><th>Status</th><th>Actions</th>';
    html += '</tr></thead><tbody>';
    for (var i = 0; i < pbSndSenders.length; i++) {
        var s = pbSndSenders[i];
        var src = s.source || {};
        var srcDesc = pbSndSourceLabel(src);
        var sched = (s.frequency || 'daily') + ' @ ' + (s.target_hour || 7) + ':00';
        var recip = s.send_to_mode === 'roles' ? (s.roles || []).join(', ') : (s.specific_people || []).length + ' people';
        var lastSent = s.last_sent ? new Date(s.last_sent).toLocaleString() + ' (' + (s.last_sent_count || 0) + ')' : 'Never';
        var status = s.enabled !== false ? '<span style="color:var(--pb-success);font-weight:600;">Active</span>' : '<span style="color:var(--pb-muted);">Paused</span>';

        html += '<tr>';
        html += '<td style="font-weight:600;">' + pbEsc(s.name || '') + '</td>';
        html += '<td class="pb-text-sm">' + pbEsc(srcDesc) + '</td>';
        html += '<td class="pb-text-sm">' + pbEsc(sched) + '</td>';
        html += '<td class="pb-text-sm">' + pbEsc(recip) + '</td>';
        html += '<td class="pb-text-sm">' + lastSent + '</td>';
        html += '<td>' + status + '</td>';
        html += '<td><button class="pb-btn pb-btn-sm" onclick="pbSndEdit(' + "'" + s.id + "'" + ')">Edit</button> ';
        html += '<button class="pb-btn pb-btn-sm" onclick="pbSndRunNowById(' + "'" + s.id + "'" + ')" style="background:var(--pb-warning);color:#fff;">Send</button> ';
        html += '<button class="pb-btn pb-btn-sm" onclick="pbSndDelete(' + "'" + s.id + "'" + ')" style="color:var(--pb-danger);">Del</button></td>';
        html += '</tr>';
    }
    html += '</tbody></table></div>';
    el.innerHTML = html;
}

function pbSndShowEditor(sender) {
    pbSndEditingId = sender ? sender.id : null;
    document.getElementById('pb-snd-editor-title').textContent = sender ? 'Edit Sender' : 'New Sender';
    document.getElementById('pb-snd-name').value = sender ? sender.name || '' : '';
    document.getElementById('pb-snd-enabled').value = sender ? String(sender.enabled !== false) : 'true';
    var src = sender ? sender.source || {} : {};
    document.getElementById('pb-snd-source-scope').value = src.scope || src.source_type || 'program';
    document.getElementById('pb-snd-member-types').value = src.member_types || '311';
    document.getElementById('pb-snd-source-value').value = src.org_id || '';
    // Show/hide scope fields and load programs with saved selection
    var _scope = src.scope || src.source_type || 'program';
    document.getElementById('pb-snd-program-wrap').style.display = (_scope === 'program' || _scope === 'division') ? '' : 'none';
    document.getElementById('pb-snd-division-wrap').style.display = _scope === 'division' ? '' : 'none';
    document.getElementById('pb-snd-org-wrap').style.display = _scope === 'org' ? '' : 'none';
    if (_scope === 'program' || _scope === 'division') {
        pbSndLoadProgramsDivisions(src.program_id || '', src.division_id || '');
    }
    document.getElementById('pb-snd-frequency').value = sender ? sender.frequency || 'daily' : 'daily';
    document.getElementById('pb-snd-target-hour').value = sender ? sender.target_hour || 7 : 7;
    document.getElementById('pb-snd-target-day').value = sender ? sender.target_day || 0 : 0;
    document.getElementById('pb-snd-target-dom').value = sender ? sender.target_day_of_month || 1 : 1;
    document.getElementById('pb-snd-lookback').value = sender ? sender.lookback || 'since_last' : 'since_last';
    document.getElementById('pb-snd-send-to-mode').value = sender ? sender.send_to_mode || 'involvement_members' : 'involvement_members';
    var selectedMts = sender ? (sender.recipient_member_types || '').split(',').filter(function(s){return s.trim()}) : ['140','310','320'];
    document.getElementById('pb-snd-recipient-mt').value = selectedMts.join(',');
    // Fetch member types and render directly
    var mtEl = document.getElementById('pb-snd-recipient-mt-list');
    mtEl.innerHTML = '<span class="pb-text-muted">Loading...</span>';
    $.ajax({
        url: window.location.pathname,
        type: 'POST',
        data: {action: 'get_member_types'},
        success: function(resp) {
            var d;
            try { d = typeof resp === 'string' ? JSON.parse(resp) : resp; } catch(e) { d = {}; }
            var mtData = d.member_types || d.memberTypes || [];
            if (!d.success || mtData.length === 0) {
                mtEl.innerHTML = '<span class="pb-text-muted">No member types found (check console)</span>';
                console.log('get_member_types response:', d);
                return;
            }
        var leaderIds = [140, 310, 320, 710, 1510, 1470, 1280];
        var sorted = mtData.slice().sort(function(a, b) {
            var aL = leaderIds.indexOf(a.id) > -1 ? 0 : 1;
            var bL = leaderIds.indexOf(b.id) > -1 ? 0 : 1;
            if (aL !== bL) return aL - bL;
            return (a.description || '').localeCompare(b.description || '');
        });
        var html = '';
        for (var mi = 0; mi < sorted.length; mi++) {
            var mt = sorted[mi];
            var chk = selectedMts.indexOf(String(mt.id)) > -1 ? ' checked' : '';
            var star = leaderIds.indexOf(mt.id) > -1 ? ' <span style="color:var(--pb-accent);font-size:0.8em;">&#9733;</span>' : '';
            html += '<label style="display:block;padding:3px 0;font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-mt-cb" value="' + mt.id + '"' + chk + ' onchange="pbSndUpdateMtHidden()"> ' + (mt.description||'') + ' (' + mt.id + ')' + star + '</label>';
        }
        mtEl.innerHTML = html;
        }
    });
    document.getElementById('pb-snd-specific-people').value = sender ? (sender.specific_people || []).join(',') : '';
    document.getElementById('pb-snd-from-email').value = sender ? sender.from_email || '' : '';
    document.getElementById('pb-snd-from-name').value = sender ? sender.from_name || '' : '';
    document.getElementById('pb-snd-subject').value = sender ? sender.subject || '' : '';
    // Blank for new and existing senders. Empty body falls back to the central default
    // in Settings > Sender Defaults at send time.
    document.getElementById('pb-snd-message-body').value = sender ? (sender.message_body || '') : '';
    // Restore email field checkboxes
    var fields = sender ? sender.email_fields || {} : {};
    var defaults = {email:true,cell_phone:true,home_phone:false,address:true,age:true,member_status:false,enrollment_date:true,person_link:true};
    document.querySelectorAll('.pb-snd-field-cb').forEach(function(cb) {
        cb.checked = fields.hasOwnProperty(cb.value) ? fields[cb.value] : (defaults[cb.value] || false);
    });
    pbSndFreqChanged();
    pbSndSendToChanged();
    pbSndRenderRoles(sender ? sender.roles || [] : []);
    document.getElementById('pb-snd-editor').style.display = 'block';
    document.getElementById('pb-snd-preview').style.display = 'none';
    document.getElementById('pb-snd-msg').innerHTML = '';
}

function pbSndHideEditor() {
    document.getElementById('pb-snd-editor').style.display = 'none';
    pbSndEditingId = null;
}

function pbSndEdit(id) {
    for (var i = 0; i < pbSndSenders.length; i++) {
        if (pbSndSenders[i].id === id) { pbSndShowEditor(pbSndSenders[i]); return; }
    }
}

var pbSndPrograms = [];
var pbSndDivisions = [];

function pbSndScopeChanged() {
    var scope = document.getElementById('pb-snd-source-scope').value;
    document.getElementById('pb-snd-program-wrap').style.display = (scope === 'program' || scope === 'division') ? '' : 'none';
    document.getElementById('pb-snd-division-wrap').style.display = scope === 'division' ? '' : 'none';
    document.getElementById('pb-snd-org-wrap').style.display = scope === 'org' ? '' : 'none';
    if (scope === 'program' || scope === 'division') {
        pbSndLoadProgramsDivisions();
    }
}

function pbSndLoadProgramsDivisions(selectedProgId, selectedDivId) {
    // Element may not exist yet when this is called from pbSndInitTab as a
    // pre-warm (the editor isn't open). All dropdown-population code is
    // guarded so the pre-warm only updates the cache.
    var pSel = document.getElementById('pb-snd-program');
    if (pbSndPrograms.length > 0) {
        if (pSel) {
            pSel.innerHTML = '<option value="">-- Select Program --</option>';
            for (var i = 0; i < pbSndPrograms.length; i++) {
                var sel = String(pbSndPrograms[i].id) === String(selectedProgId) ? ' selected' : '';
                pSel.innerHTML += '<option value="' + pbSndPrograms[i].id + '"' + sel + '>' + pbEsc(pbSndPrograms[i].name) + '</option>';
            }
            if (selectedDivId) { pbSndProgramChanged(); document.getElementById('pb-snd-division').value = selectedDivId; }
        }
        return;
    }
    if (pSel) pSel.innerHTML = '<option value="">Loading programs...</option>';
    $.ajax({
        url: window.location.pathname,
        type: 'POST',
        data: {action: 'get_programs_divisions'},
        success: function(resp) {
            var d; try { d = typeof resp === 'string' ? JSON.parse(resp) : resp; } catch(e) { return; }
            if (!d.success) return;
            pbSndPrograms = d.programs || [];
            pbSndDivisions = d.divisions || [];
            var pSel2 = document.getElementById('pb-snd-program');
            if (pSel2) {
                pSel2.innerHTML = '<option value="">-- Select Program --</option>';
                for (var i = 0; i < pbSndPrograms.length; i++) {
                    var sel = String(pbSndPrograms[i].id) === String(selectedProgId) ? ' selected' : '';
                    pSel2.innerHTML += '<option value="' + pbSndPrograms[i].id + '"' + sel + '>' + pbEsc(pbSndPrograms[i].name) + '</option>';
                }
                if (selectedDivId) { pbSndProgramChanged(); document.getElementById('pb-snd-division').value = selectedDivId; }
            }
            // Pre-warm path: re-render the senders list so the Source column
            // can now show friendly program/division names instead of IDs.
            if (typeof pbSndRenderList === 'function' && pbSndSenders && pbSndSenders.length) {
                pbSndRenderList();
            }
        }
    });
}

function pbSndProgramChanged() {
    var progId = document.getElementById('pb-snd-program').value;
    var dSel = document.getElementById('pb-snd-division');
    dSel.innerHTML = '<option value="">-- All Divisions --</option>';
    for (var i = 0; i < pbSndDivisions.length; i++) {
        if (String(pbSndDivisions[i].progId) === progId) {
            dSel.innerHTML += '<option value="' + pbSndDivisions[i].id + '">' + pbEsc(pbSndDivisions[i].name) + '</option>';
        }
    }
}

function pbSndFreqChanged() {
    var f = document.getElementById('pb-snd-frequency').value;
    var wrap = document.getElementById('pb-snd-day-wrap');
    var sel = document.getElementById('pb-snd-target-day');
    var dom = document.getElementById('pb-snd-target-dom');
    if (f === 'weekly') { wrap.style.display = ''; sel.style.display = ''; dom.style.display = 'none'; document.getElementById('pb-snd-day-label').textContent = 'Day of Week'; }
    else if (f === 'monthly') { wrap.style.display = ''; sel.style.display = 'none'; dom.style.display = ''; document.getElementById('pb-snd-day-label').textContent = 'Day of Month'; }
    else { wrap.style.display = 'none'; }
}

function pbSndSendToChanged() {
    var m = document.getElementById('pb-snd-send-to-mode').value;
    document.getElementById('pb-snd-inv-members-wrap').style.display = m === 'involvement_members' ? '' : 'none';
    document.getElementById('pb-snd-roles-wrap').style.display = m === 'roles' ? '' : 'none';
    document.getElementById('pb-snd-people-wrap').style.display = m === 'specific_people' ? '' : 'none';
}

function pbSndRenderMemberTypes(selectedIds) {
    var el = document.getElementById('pb-snd-recipient-mt-list');
    if (!el) return;
    if (!window._pbSndMtData || window._pbSndMtData.length === 0) {
        el.innerHTML = '<span class="pb-text-muted">Loading member types...</span>';
        if (!window._pbSndMtLoading) {
            window._pbSndMtLoading = true;
            pbAjax({action: 'get_member_types'}, function(d) {
                window._pbSndMtLoading = false;
                if (d.success && d.member_types) {
                    window._pbSndMtData = d.member_types;
                    pbSndRenderMemberTypes(selectedIds);
                } else {
                    el.innerHTML = '<span class="pb-text-muted">No member types found</span>';
                }
            });
        }
        return;
    }
    var pbSndMemberTypes = window._pbSndMtData;
    // Common leader types shown first
    var leaderIds = [140, 310, 320, 710];
    var sorted = pbSndMemberTypes.slice().sort(function(a, b) {
        var aIsLeader = leaderIds.indexOf(a.id) > -1 ? 0 : 1;
        var bIsLeader = leaderIds.indexOf(b.id) > -1 ? 0 : 1;
        if (aIsLeader !== bIsLeader) return aIsLeader - bIsLeader;
        return (a.description || '').localeCompare(b.description || '');
    });
    var html = '';
    for (var i = 0; i < sorted.length; i++) {
        var mt = sorted[i];
        var checked = selectedIds.indexOf(String(mt.id)) > -1 ? ' checked' : '';
        var badge = leaderIds.indexOf(mt.id) > -1 ? ' <span style="color:var(--pb-accent);font-size:0.8em;">&#9733;</span>' : '';
        html += '<label style="display:block;padding:3px 0;font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-mt-cb" value="' + mt.id + '"' + checked + ' onchange="pbSndUpdateMtHidden()"> ' + pbEsc(mt.description) + ' (' + mt.id + ')' + badge + '</label>';
    }
    el.innerHTML = html || '<span class="pb-text-muted">No member types found</span>';
}

function pbSndUpdateMtHidden() {
    var vals = [];
    document.querySelectorAll('.pb-snd-mt-cb:checked').forEach(function(cb) { vals.push(cb.value); });
    document.getElementById('pb-snd-recipient-mt').value = vals.join(',');
}

function pbSndRenderRoles(selectedRoles) {
    var html = '';
    for (var i = 0; i < pbSndRoles.length; i++) {
        var r = pbSndRoles[i];
        var checked = selectedRoles.indexOf(r) > -1 ? ' checked' : '';
        html += '<label style="display:block;padding:3px 0;font-size:0.9em;cursor:pointer;"><input type="checkbox" class="pb-snd-role-cb" value="' + pbEsc(r) + '"' + checked + '> ' + pbEsc(r) + '</label>';
    }
    document.getElementById('pb-snd-roles-list').innerHTML = html || '<span class="pb-text-muted">Loading roles...</span>';
}

function pbSndBuildSender() {
    var scope = document.getElementById('pb-snd-source-scope').value;
    var source = { scope: scope };
    if (scope === 'org') source.org_id = document.getElementById('pb-snd-source-value').value.trim();
    if (scope === 'program' || scope === 'division') source.program_id = document.getElementById('pb-snd-program').value;
    if (scope === 'division') source.division_id = document.getElementById('pb-snd-division').value;
    source.member_types = document.getElementById('pb-snd-member-types').value.trim();

    var roles = [];
    document.querySelectorAll('.pb-snd-role-cb:checked').forEach(function(cb) { roles.push(cb.value); });

    var specificPeople = document.getElementById('pb-snd-specific-people').value.trim();

    return {
        id: pbSndEditingId || '',
        name: document.getElementById('pb-snd-name').value.trim(),
        enabled: document.getElementById('pb-snd-enabled').value === 'true',
        source: source,
        frequency: document.getElementById('pb-snd-frequency').value,
        target_hour: parseInt(document.getElementById('pb-snd-target-hour').value) || 7,
        target_day: parseInt(document.getElementById('pb-snd-target-day').value) || 0,
        target_day_of_month: parseInt(document.getElementById('pb-snd-target-dom').value) || 1,
        lookback: document.getElementById('pb-snd-lookback').value,
        send_to_mode: document.getElementById('pb-snd-send-to-mode').value,
        recipient_member_types: document.getElementById('pb-snd-recipient-mt').value.trim(),
        roles: roles,
        specific_people: specificPeople ? specificPeople.split(',').map(function(s){return s.trim()}) : [],
        from_email: document.getElementById('pb-snd-from-email').value.trim(),
        from_name: document.getElementById('pb-snd-from-name').value.trim(),
        subject: document.getElementById('pb-snd-subject').value.trim(),
        message_body: document.getElementById('pb-snd-message-body').value,
        email_fields: (function(){ var f={}; document.querySelectorAll('.pb-snd-field-cb').forEach(function(cb){f[cb.value]=cb.checked}); return f; })()
    };
}

function pbSndSave() {
    var sender = pbSndBuildSender();
    if (!sender.name) { document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-danger)">Name is required</span>'; return; }
    pbAjax({action: 'save_sender', sender_data: JSON.stringify(sender)}, function(d) {
        if (d.success) {
            document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-success)">Saved!</span>';
            pbSndEditingId = d.sender ? d.sender.id : pbSndEditingId;
            pbSndLoadSenders();
        } else {
            document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-danger)">' + (d.message || 'Error') + '</span>';
        }
    });
}

function pbSndPreview() {
    var sender = pbSndBuildSender();
    // Validate required fields based on scope
    var src = sender.source || {};
    if (src.scope === 'program' && !src.program_id) {
        document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-danger)">Please select a Program first</span>';
        return;
    }
    if (src.scope === 'org' && !src.org_id) {
        document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-danger)">Please enter an Organization ID</span>';
        return;
    }
    document.getElementById('pb-snd-msg').innerHTML = 'Running preview...';
    pbAjax({action: 'preview_sender', sender_data: JSON.stringify(sender)}, function(d) {
        document.getElementById('pb-snd-msg').innerHTML = '';
        if (d.success && d.result) {
            var r = d.result;
            var html = '<div><strong>Prospects found:</strong> ' + r.prospects_found + '</div>';
            html += '<div><strong>Recipients:</strong> ' + r.recipients + '</div>';
            if (r.org_details && r.org_details.length > 0) {
                html += '<div style="margin-top:12px;"><strong>Per-Group Breakdown:</strong></div>';
                for (var oi = 0; oi < r.org_details.length; oi++) {
                    var od = r.org_details[oi];
                    html += '<div style="background:var(--pb-white);border:1px solid var(--pb-border);border-radius:4px;padding:8px 12px;margin:4px 0;">';
                    html += '<strong>' + pbEsc(od.org_name) + '</strong>: ' + od.prospects + ' prospect(s) &rarr; ' + od.leaders + ' leader(s)';
                    if (od.leader_names && od.leader_names.length > 0) html += '<br><span class="pb-text-sm pb-text-muted">Leaders: ' + od.leader_names.map(function(n){return pbEsc(n)}).join(', ') + '</span>';
                    html += '</div>';
                }
                // Show sample email preview for first org
                if (r.sample_email_html) {
                    html += '<div style="margin-top:16px;"><strong>Sample Email Preview</strong> <span class="pb-text-sm pb-text-muted">(what the first group\\\'s leaders would receive)</span></div>';
                    html += '<div style="border:2px solid var(--pb-border);border-radius:8px;margin-top:8px;overflow:hidden;">';
                    html += '<div style="background:#f5f5f5;padding:10px 14px;border-bottom:1px solid var(--pb-border);">';
                    html += '<div style="font-size:0.85em;color:#666;"><strong>Subject:</strong> ' + pbEsc(r.sample_email_subject || '') + '</div>';
                    html += '<div style="font-size:0.85em;color:#666;"><strong>From:</strong> ' + pbEsc(r.sample_email_from || '') + '</div>';
                    html += '<div style="font-size:0.85em;color:#666;"><strong>To:</strong> ' + pbEsc(r.sample_email_to || '') + '</div>';
                    html += '</div>';
                    html += '<div style="padding:16px;background:#fff;">' + r.sample_email_html + '</div>';
                    html += '</div>';
                }
            } else if (r.prospect_names && r.prospect_names.length > 0) {
                html += '<div style="margin-top:8px;"><strong>Prospects:</strong><ul style="margin:4px 0;">';
                for (var i = 0; i < r.prospect_names.length; i++) html += '<li>' + pbEsc(r.prospect_names[i]) + '</li>';
                html += '</ul></div>';
            }
            if (r.errors && r.errors.length > 0) {
                html += '<div style="color:var(--pb-danger);margin-top:8px;">' + r.errors.join('<br>') + '</div>';
            }
            if (r.debug_sql) {
                html += '<details style="margin-top:12px;"><summary style="cursor:pointer;font-size:0.85em;color:var(--pb-muted);">Debug: SQL Query</summary><pre style="background:var(--pb-light);padding:8px;border-radius:4px;font-size:0.8em;overflow-x:auto;margin-top:4px;">' + pbEsc(r.debug_sql) + '</pre></details>';
            }
            if (r.debug_source) {
                html += '<details><summary style="cursor:pointer;font-size:0.85em;color:var(--pb-muted);">Debug: Source Config</summary><pre style="background:var(--pb-light);padding:8px;border-radius:4px;font-size:0.8em;overflow-x:auto;margin-top:4px;">' + pbEsc(JSON.stringify(r.debug_source, null, 2)) + '</pre></details>';
            }
            document.getElementById('pb-snd-preview-content').innerHTML = html;
            document.getElementById('pb-snd-preview').style.display = 'block';
        } else {
            document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-danger)">' + (d.message || 'Preview failed') + '</span>';
        }
    });
}

function pbSndRunNow() {
    if (!confirm('Send emails now? This cannot be undone.')) return;
    var sender = pbSndBuildSender();
    document.getElementById('pb-snd-msg').innerHTML = 'Sending...';
    pbAjax({action: 'run_sender_oneoff', sender_data: JSON.stringify(sender)}, function(d) {
        if (d.success && d.result) {
            document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-success)">' + (d.result.message || 'Sent!') + '</span>';
            pbSndLoadLog();
        } else {
            document.getElementById('pb-snd-msg').innerHTML = '<span style="color:var(--pb-danger)">' + (d.message || 'Send failed') + '</span>';
        }
    });
}

function pbSndRunNowById(id) {
    if (!confirm('Send emails now for this sender?')) return;
    pbAjax({action: 'run_sender', sender_id: id}, function(d) {
        if (d.success) {
            alert(d.result ? d.result.message || 'Sent!' : 'Sent!');
            pbSndLoadSenders();
            pbSndLoadLog();
        } else {
            alert('Error: ' + (d.message || 'Unknown'));
        }
    });
}

function pbSndDelete(id) {
    if (!confirm('Delete this sender?')) return;
    pbAjax({action: 'delete_sender', sender_id: id}, function(d) {
        if (d.success) pbSndLoadSenders();
    });
}

// ============================================================
// GROUP MANAGEMENT - SUB-TABS & PROSPECT HEALTH
// ============================================================
var pbGrpHealthPrograms = [];
var pbGrpHealthDivisions = [];

// Legacy sub-tab switcher kept as a no-op so any stale onclick references won't error.
// The Group Management tab now shows everything in one panel.
function pbGrpSwitchSubTab(tab) { /* no-op */ }

function pbGrpHealthSourceChanged() {
    var src = document.getElementById('pb-grp-health-source').value;
    document.getElementById('pb-grp-health-group-wrap').style.display = src === 'group' ? '' : 'none';
    document.getElementById('pb-grp-health-manual-wrap').style.display = src === 'manual' ? '' : 'none';
    // Repopulate when switching INTO 'group' -- handles the race where
    // pbGrpInitTab fires its populate() before pbGrpLoadData's AJAX returns.
    if (src === 'group') pbGrpHealthPopulateGroups();
}

function pbGrpHealthPopulateGroups() {
    var sel = document.getElementById('pb-grp-health-group');
    sel.innerHTML = '<option value="">-- Select Group --</option>';
    for (var i = 0; i < pbGrpGroups.length; i++) {
        var g = pbGrpGroups[i];
        sel.innerHTML += '<option value="' + g.id + '">' + pbEsc(g.name) + '</option>';
    }
}

function pbGrpHealthLoadPrograms() {
    $.ajax({
        url: window.location.pathname, type: 'POST',
        data: {action: 'get_programs_divisions'},
        success: function(resp) {
            var d; try { d = typeof resp === 'string' ? JSON.parse(resp) : resp; } catch(e) { return; }
            if (!d.success) return;
            pbGrpHealthPrograms = d.programs || [];
            pbGrpHealthDivisions = d.divisions || [];
            var sel = document.getElementById('pb-grp-health-program');
            sel.innerHTML = '<option value="">-- Select Program --</option>';
            for (var i = 0; i < pbGrpHealthPrograms.length; i++) {
                sel.innerHTML += '<option value="' + pbGrpHealthPrograms[i].id + '">' + pbEsc(pbGrpHealthPrograms[i].name) + '</option>';
            }
        }
    });
}

function pbGrpHealthProgramChanged() {
    var progId = document.getElementById('pb-grp-health-program').value;
    var dSel = document.getElementById('pb-grp-health-division');
    dSel.innerHTML = '<option value="">-- All Divisions --</option>';
    for (var i = 0; i < pbGrpHealthDivisions.length; i++) {
        if (String(pbGrpHealthDivisions[i].progId) === progId) {
            dSel.innerHTML += '<option value="' + pbGrpHealthDivisions[i].id + '">' + pbEsc(pbGrpHealthDivisions[i].name) + '</option>';
        }
    }
}

function pbGrpLoadHealth() {
    var srcMode = document.getElementById('pb-grp-health-source').value;
    var senderData;

    if (srcMode === 'group') {
        var groupId = document.getElementById('pb-grp-health-group').value;
        if (!groupId) { alert('Please select a Group'); return; }
        // Find the group config and convert to sender format
        var grp = null;
        for (var i = 0; i < pbGrpGroups.length; i++) {
            if (pbGrpGroups[i].id === groupId) { grp = pbGrpGroups[i]; break; }
        }
        if (!grp) { alert('Group not found'); return; }

        var scope = 'program';
        if (grp.level === 'involvement') scope = 'org';
        else if (grp.level === 'division') scope = 'division';

        var mt = (grp.memberTypes || []).join(',') || '311';

        senderData = {
            source: {
                scope: scope,
                program_id: grp.programId || '',
                division_id: grp.divisionId || '',
                org_id: grp.orgId || '',
                member_types: mt
            },
            recipient_member_types: '140,310,320',
            _group_id: groupId
        };
    } else {
        var progId = document.getElementById('pb-grp-health-program').value;
        if (!progId) { alert('Please select a Program'); return; }
        var divId = document.getElementById('pb-grp-health-division').value;
        var mt = document.getElementById('pb-grp-health-mt').value.trim();
        var leaderMt = document.getElementById('pb-grp-health-leader-mt').value.trim();

        senderData = {
            source: {
                scope: divId ? 'division' : 'program',
                program_id: progId,
                division_id: divId,
                member_types: mt || '311'
            },
            recipient_member_types: leaderMt || '140,310,320'
        };
    }

    // Track the group id (if any) so we can wire group-metrics affordances.
    pbHealthCurrentGroupId = (srcMode === 'group') ? (senderData._group_id || '') : '';
    var content = document.getElementById('pb-grp-health-content');
    var summary = document.getElementById('pb-grp-health-summary');
    content.innerHTML = '<div class="pb-loading" style="text-align:center;padding:30px;"><span class="pb-spin"></span> Loading prospect health...</div>';
    summary.innerHTML = '';

    pbAjax({action: 'get_prospect_health', sender_data: JSON.stringify(senderData)}, function(d) {
        if (!d.success) { content.innerHTML = '<span style="color:var(--pb-danger)">' + (d.message || 'Error') + '</span>'; return; }
        // Gather all prospect PIDs and fetch contact efforts
        var allPids = [];
        for (var oi = 0; oi < d.orgs.length; oi++) {
            for (var pi = 0; pi < d.orgs[oi].prospects.length; pi++) {
                allPids.push(d.orgs[oi].prospects[pi].people_id);
            }
        }
        function afterRender() {
            // Once the dashboard is in the DOM, layer in cached group metrics
            // when viewing a saved group. Manual mode skips this since metrics
            // are keyed by group id.
            if (pbHealthCurrentGroupId) pbHealthFetchAndApplyMetrics();
        }
        if (allPids.length > 0) {
            pbAjax({action: 'load_contact_efforts', people_ids: allPids.join(',')}, function(d2) {
                var cmap = (d2.success && d2.contactMap) ? d2.contactMap : {};
                pbRenderHealthDashboard(d, summary, content, cmap);
                afterRender();
            });
        } else {
            pbRenderHealthDashboard(d, summary, content, {});
            afterRender();
        }
    });
}

// ---- Group metrics (cached attendance/conversion stats per involvement) ----
var pbHealthCurrentGroupId = '';
var pbHealthMetrics = null;       // {computedAt, byOrg: {orgId: {orgName, windows: {90:..,180:..,365:..}}}}
var pbHealthWindow = '90';        // active window key
var pbHealthWindowsAvail = [90, 180, 365];

function pbHealthFetchAndApplyMetrics() {
    if (!pbHealthCurrentGroupId) return;
    pbAjax({action: 'load_group_metrics', group_id: pbHealthCurrentGroupId}, function(d) {
        if (!d.success) return;
        pbHealthMetrics = d.metrics || null;
        if (Array.isArray(d.windows) && d.windows.length) pbHealthWindowsAvail = d.windows;
        pbHealthRenderMetricsUI();
    });
}

function pbHealthRenderMetricsUI() {
    // Inject (or replace) the metrics toolbar at the top of the health summary
    // and update each org card's metrics bar.
    var summary = document.getElementById('pb-grp-health-summary');
    if (!summary) return;
    var existingTb = document.getElementById('pb-health-metrics-toolbar');
    if (existingTb) existingTb.parentNode.removeChild(existingTb);

    if (!pbHealthCurrentGroupId) {
        pbHealthClearMetricBars();
        return;
    }

    var computedAt = pbHealthMetrics && pbHealthMetrics.computedAt ? pbHealthMetrics.computedAt : '';
    var computedLabel = computedAt ? computedAt.substring(0, 16).replace('T', ' ') : 'never';

    var tb = document.createElement('div');
    tb.id = 'pb-health-metrics-toolbar';
    tb.style.cssText = 'width:100%;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;background:#f8f9fa;border:1px solid var(--pb-border);border-radius:6px;padding:8px 12px;margin-bottom:8px;';

    var winOpts = '';
    for (var wi = 0; wi < pbHealthWindowsAvail.length; wi++) {
        var w = pbHealthWindowsAvail[wi];
        var sel = String(w) === String(pbHealthWindow) ? ' selected' : '';
        winOpts += '<option value="' + w + '"' + sel + '>' + w + 'd</option>';
    }

    var left = '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">' +
        '<span style="font-weight:600;font-size:0.9em;color:var(--pb-primary);">Conversion Metrics</span>' +
        '<label class="pb-text-sm" style="font-weight:600;">Window:</label>' +
        '<select class="pb-input" style="height:30px;padding:2px 8px;font-size:0.85em;" onchange="pbHealthChangeWindow(this.value)">' + winOpts + '</select>' +
        '<span class="pb-text-muted" style="font-size:0.8em;">Last computed: ' + computedLabel + '</span>' +
        '</div>';
    var right = '<button class="pb-btn pb-btn-sm pb-btn-primary" onclick="pbHealthRefreshMetrics()" id="pb-health-refresh-btn">Refresh Now</button>';

    tb.innerHTML = left + right;
    // Insert as the FIRST child of the summary so it sits above the count tiles.
    summary.insertBefore(tb, summary.firstChild);

    pbHealthApplyMetricsToCards();
}

function pbHealthChangeWindow(w) {
    pbHealthWindow = String(w);
    pbHealthApplyMetricsToCards();
}

function pbHealthRefreshMetrics() {
    if (!pbHealthCurrentGroupId) return;
    var btn = document.getElementById('pb-health-refresh-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Refreshing...'; }
    pbAjax({action: 'refresh_group_metrics', group_id: pbHealthCurrentGroupId}, function(d) {
        if (btn) { btn.disabled = false; btn.textContent = 'Refresh Now'; }
        if (!d.success) { alert('Refresh failed: ' + (d.message || 'unknown')); return; }
        pbHealthMetrics = d.metrics || null;
        if (Array.isArray(d.windows) && d.windows.length) pbHealthWindowsAvail = d.windows;
        pbHealthRenderMetricsUI();
    });
}

function pbHealthClearMetricBars() {
    var bars = document.querySelectorAll('.pb-health-metrics-bar');
    for (var i = 0; i < bars.length; i++) {
        bars[i].style.display = 'none';
        bars[i].innerHTML = '';
    }
}

// Funnel state metadata: colors + labels + plural names for the people-list header.
var pbHealthStates = {
    converted: { label: 'Converted', color: '#107c10', plural: 'converted', emoji: '\u{1F7E2}' },
    engaged:   { label: 'Engaged',   color: '#f0ad4e', plural: 'engaged (still prospect)', emoji: '\u{1F7E1}' },
    no_show:   { label: 'No-show',   color: '#d13438', plural: 'no-show (on roster, never attended)', emoji: '\u{1F534}' },
    dropped:   { label: 'Dropped',   color: '#797673', plural: 'dropped (no attend, off roster)', emoji: '⚫' }
};

function pbHealthApplyMetricsToCards() {
    pbHealthClearMetricBars();
    if (!pbHealthMetrics || !pbHealthMetrics.byOrg) return;
    var byOrg = pbHealthMetrics.byOrg;
    var w = String(pbHealthWindow);

    var bars = document.querySelectorAll('.pb-health-metrics-bar');
    for (var i = 0; i < bars.length; i++) {
        var bar = bars[i];
        var oid = bar.getAttribute('data-orgid');
        var orgMetrics = byOrg[String(oid)];
        if (!orgMetrics || !orgMetrics.windows || !orgMetrics.windows[w]) continue;
        var m = orgMetrics.windows[w];
        var rate = m.conversionRate || 0;
        var rateColor = rate >= 30 ? '#107c10' : rate >= 10 ? '#f0ad4e' : '#d13438';

        function segHtml(state, count) {
            var meta = pbHealthStates[state];
            var clickable = count > 0;
            var style = 'color:' + meta.color + ';font-weight:700;'
                      + (clickable ? 'cursor:pointer;text-decoration:underline dotted;' : '');
            var attrs = clickable
                ? ' data-orgid="' + oid + '" data-window="' + w + '" data-state="' + state + '" onclick="pbHealthToggleFunnel(this)" title="Click to see who"'
                : '';
            return '<span style="' + style + '"' + attrs + '>'
                 + count + ' <span class="pb-text-muted" style="font-weight:normal;color:#999;">' + meta.label.toLowerCase() + '</span>'
                 + '</span>';
        }

        bar.style.display = '';
        bar.style.cssText = 'display:flex;flex-wrap:wrap;gap:14px;padding:8px 16px;background:#eef5fb;border-bottom:1px solid #d6e4ee;font-size:0.85em;align-items:center;';
        bar.innerHTML =
            '<span><strong>' + (m.prospects || 0) + '</strong> <span class="pb-text-muted">prospects</span></span>' +
            '<span style="color:#999;">|</span>' +
            segHtml('converted', m.converted || 0) +
            segHtml('engaged',   m.engaged   || 0) +
            segHtml('no_show',   m.noShow    || 0) +
            segHtml('dropped',   m.dropped   || 0) +
            '<span style="color:#999;">|</span>' +
            '<span style="color:' + rateColor + ';font-weight:700;">' + rate + '% <span class="pb-text-muted" style="font-weight:normal;color:#999;">conv</span></span>' +
            '<span class="pb-text-muted" style="font-size:0.8em;margin-left:auto;">' + w + 'd window</span>';

        // Ensure a funnel-detail panel exists right after this bar.
        var sib = bar.nextElementSibling;
        if (!sib || !sib.classList.contains('pb-health-funnel-detail')) {
            var det = document.createElement('div');
            det.className = 'pb-health-funnel-detail';
            det.style.display = 'none';
            bar.parentNode.insertBefore(det, bar.nextSibling);
        } else {
            sib.style.display = 'none';
            sib.innerHTML = '';
            sib.removeAttribute('data-loaded-key');
        }
    }
}

function pbHealthToggleFunnel(el) {
    var oid = el.getAttribute('data-orgid');
    var w = el.getAttribute('data-window');
    var state = el.getAttribute('data-state');
    // Find the metrics bar this segment belongs to.
    var bar = el.closest ? el.closest('.pb-health-metrics-bar') : null;
    if (!bar) {
        var node = el;
        while (node && !(node.classList && node.classList.contains('pb-health-metrics-bar'))) node = node.parentNode;
        bar = node;
    }
    if (!bar) return;
    var panel = bar.nextElementSibling;
    if (!panel || !panel.classList.contains('pb-health-funnel-detail')) return;

    var loadKey = state + '|' + w;
    // Collapse if already open for the same state+window.
    if (panel.style.display !== 'none' && panel.getAttribute('data-loaded-key') === loadKey) {
        panel.style.display = 'none';
        return;
    }

    var meta = pbHealthStates[state] || {label: state, color: '#666'};
    panel.style.cssText = 'background:#fafafa;border-bottom:1px solid #e0e0e0;padding:10px 16px;border-left:3px solid ' + meta.color + ';';
    panel.innerHTML = '<div class="pb-loading" style="text-align:center;padding:8px;"><span class="pb-spin"></span> Loading ' + meta.label.toLowerCase() + ' people...</div>';

    pbAjax({
        action: 'get_funnel_people',
        group_id: pbHealthCurrentGroupId,
        org_id: oid,
        window_days: w,
        state: state
    }, function(d) {
        if (!d.success) {
            panel.innerHTML = '<span style="color:var(--pb-danger);">' + (d.message || 'Failed to load') + '</span>';
            return;
        }
        var people = d.people || [];
        if (!people.length) {
            panel.innerHTML = '<span class="pb-text-muted">No people in this bucket for the ' + w + 'd window.</span>';
            panel.setAttribute('data-loaded-key', loadKey);
            return;
        }
        var html = '<div style="font-weight:600;margin-bottom:6px;font-size:0.9em;color:' + meta.color + ';">'
                 + people.length + ' ' + (meta.plural || meta.label.toLowerCase()) + ' in ' + w + 'd'
                 + '</div>';
        html += '<table style="width:100%;border-collapse:collapse;font-size:0.85em;">';
        html += '<thead><tr style="background:#eef5fb;">'
              + '<th style="padding:5px 8px;text-align:left;">Name</th>'
              + '<th style="padding:5px 8px;text-align:center;">Age</th>'
              + '<th style="padding:5px 8px;text-align:left;">Contact</th>'
              + (state === 'converted'
                    ? '<th style="padding:5px 8px;text-align:left;">First Member Attend</th>'
                    : '<th style="padding:5px 8px;text-align:left;">Engagement Elsewhere</th>')
              + '<th style="padding:5px 8px;"></th>'
              + '</tr></thead><tbody>';
        for (var i = 0; i < people.length; i++) {
            var p = people[i];
            var contact = '';
            if (p.cellPhone) contact += '<a href="tel:' + p.cellPhone + '">' + pbEsc(p.cellPhone) + '</a>';
            if (p.cellPhone && p.email) contact += '<br>';
            if (p.email) contact += '<a href="mailto:' + p.email + '" style="font-size:0.85em;">' + pbEsc(p.email) + '</a>';
            if (!contact) contact = '<span class="pb-text-muted">-</span>';

            // Current-prospect tag — shown when applicable so we can spot
            // people still flagged as prospects even though they may be in
            // the engaged/no-show/dropped bucket.
            var cpTag = p.isCurrentProspect
                ? ' <span style="background:#fff4ce;color:#7a5c00;padding:1px 6px;border-radius:8px;font-size:0.7em;font-weight:600;" title="Still has an active prospect MemberType in this org">PROSPECT</span>'
                : '';

            var lastCol;
            if (state === 'converted') {
                lastCol = p.firstConvertedDate || '-';
            } else {
                var xo = p.crossOrgAttends || [];
                if (xo.length) {
                    var parts = [];
                    for (var ci = 0; ci < xo.length; ci++) {
                        parts.push('<span style="background:#e1f5fe;color:#01579b;padding:2px 6px;border-radius:8px;font-size:0.85em;margin-right:3px;" title="Last attend ' + xo[ci].lastAttend + '">'
                                 + pbEsc(xo[ci].orgName) + '</span>');
                    }
                    lastCol = parts.join('');
                } else {
                    lastCol = '<span class="pb-text-muted">None in scope</span>';
                }
            }

            html += '<tr style="border-bottom:1px solid #f0f0f0;">'
                  + '<td style="padding:5px 8px;font-weight:600;">' + pbEsc(p.name || '') + cpTag + '</td>'
                  + '<td style="padding:5px 8px;text-align:center;">' + (p.age != null ? p.age : '-') + '</td>'
                  + '<td style="padding:5px 8px;">' + contact + '</td>'
                  + '<td style="padding:5px 8px;">' + lastCol + '</td>'
                  + '<td style="padding:5px 8px;text-align:right;"><a href="/Person2/' + p.peopleId + '" target="_blank" title="Open profile" style="text-decoration:none;">&#128279;</a></td>'
                  + '</tr>';
        }
        html += '</tbody></table>';
        panel.innerHTML = html;
        panel.setAttribute('data-loaded-key', loadKey);
    });
}

// Cache of {peopleId -> {memberStatus, peopleId}} drawn from the health
// payload, so pbHealthComputeOrgScores can build the member_data dict the
// compute_scores action expects without needing another DB round-trip.
var pbHealthMemberCache = {};
// Track which orgs we've already scored so re-expanding doesn't re-fire.
var pbHealthScoresLoadedOrgs = {};

function pbRenderHealthDashboard(d, summaryEl, contentEl, contactMap) {
    contactMap = contactMap || {};
    // Reset state -- a fresh health load means previous scores are stale.
    pbHealthMemberCache = {};
    pbHealthScoresLoadedOrgs = {};
    // Mirror contact map onto pbGrpContactMap so pbComputeScores can read
    // weightedTotal/total per person via its existing accessor pattern.
    pbGrpContactMap = contactMap;
    for (var ci = 0; ci < (d.orgs || []).length; ci++) {
        var _g = d.orgs[ci];
        for (var cj = 0; cj < (_g.prospects || []).length; cj++) {
            var _p = _g.prospects[cj];
            pbHealthMemberCache[String(_p.people_id)] = {
                memberStatus: _p.member_status || '',
                peopleId: _p.people_id
            };
        }
    }
    var totalP = d.total_prospects || 0;
    var totalO = d.total_orgs || 0;
    var totalStale = 0; var totalNoContact = 0;
    for (var oi = 0; oi < d.orgs.length; oi++) { totalStale += d.orgs[oi].stale || 0; totalNoContact += d.orgs[oi].no_contact || 0; }
    // Wrap in a grid so the four metric cards always line up as 4 columns
    // on wide screens (1100+), 2 columns on tablet, 1 on phone. The parent
    // pb-grp-health-summary container is display:flex which let the cards
    // collapse to full-width bands -- a grid keeps them aligned cleanly.
    summaryEl.innerHTML =
        '<div class="pb-stat-grid-4" style="display:grid;grid-template-columns:repeat(4, minmax(0,1fr));gap:12px;width:100%;">' +
            '<div class="pb-stat-card" style="background:#dff6dd;border-color:#a8e2a3;">'
                + '<div class="pb-stat-card-num" style="color:#107c10;">' + totalO + '</div>'
                + '<div class="pb-stat-card-label" style="color:#107c10;">Involvements '
                + '<span class="pb-info" title="Number of involvements in scope for the selected group or manual filter. Each involvement is counted once.">?</span></div>'
            + '</div>' +
            '<div class="pb-stat-card" style="background:#deecf9;border-color:#a8d4f1;">'
                + '<div class="pb-stat-card-num" style="color:#0078d4;">' + totalP + '</div>'
                + '<div class="pb-stat-card-label" style="color:#0078d4;">Prospects '
                + '<span class="pb-info" title="Total (Person, Involvement) prospect memberships across the involvements in scope. A person who is a prospect in 2 involvements counts twice.">?</span></div>'
            + '</div>' +
            '<div class="pb-stat-card" style="background:#fff4ce;border-color:#f0e3a0;">'
                + '<div class="pb-stat-card-num" style="color:#7a6d2e;">' + totalNoContact + '</div>'
                + '<div class="pb-stat-card-label" style="color:#7a6d2e;">No Contact '
                + '<span class="pb-info" title="Prospect memberships whose person has not been logged with any TaskNote contact within the configured No Contact Threshold (default 30 days). Set in Settings.">?</span></div>'
            + '</div>' +
            '<div class="pb-stat-card" style="background:#fde7e9;border-color:#f3b6bb;">'
                + '<div class="pb-stat-card-num" style="color:#d13438;">' + totalStale + '</div>'
                + '<div class="pb-stat-card-label" style="color:#d13438;">Stale (14d+) '
                + '<span class="pb-info" title="Prospect memberships whose person has been enrolled for more than 14 days AND who has either never attended this involvement, or has not attended in 14+ days. They are quietly going cold.">?</span></div>'
            + '</div>' +
        '</div>';

    var html = '';
    for (var gi = 0; gi < d.orgs.length; gi++) {
        var og = d.orgs[gi];
        var healthColor = og.stale > 0 ? '#d13438' : og.no_contact > 0 ? '#f0ad4e' : '#107c10';
        html += '<div class="pb-card pb-health-org-card" data-orgid="' + og.org_id + '" style="margin-bottom:12px;border-left:4px solid ' + healthColor + ';padding:0;overflow:hidden;">';
        // Placeholder for cached group metrics (Prospects / Attended / Converted / Conv%).
        // Filled in by pbHealthApplyMetrics after pbGrpLoadHealth fetches the metrics cache.
        html += '<div class="pb-health-metrics-bar" data-orgid="' + og.org_id + '" style="display:none;"></div>';
        html += '<div style="padding:12px 16px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;" onclick="pbSndToggleHealthOrg(this)">';
        html += '<div><strong style="font-size:1em;">' + pbEsc(og.name) + '</strong>';
        html += ' <span class="pb-text-sm pb-text-muted">(' + og.total_prospects + ' prospects, ' + og.leaders + ' leaders)</span>';
        // Meeting info
        if (og.meeting_freq) {
            var meetColor = og.last_meeting ? '#107c10' : '#d13438';
            html += ' <span style="font-size:0.8em;color:' + meetColor + ';margin-left:8px;">&#128197; ' + og.meeting_freq;
            if (og.last_meeting) html += ', last ' + og.last_meeting;
            if (og.meetings_90d) html += ' (' + og.meetings_90d + ' in 90d)';
            html += '</span>';
        }
        html += '</div>';
        html += '<div style="display:flex;gap:8px;">';
        if (og.stale > 0) html += '<span style="background:#fde7e9;color:#d13438;padding:2px 8px;border-radius:10px;font-size:0.8em;font-weight:600;">' + og.stale + ' stale</span>';
        if (og.no_contact > 0) html += '<span style="background:#fff4ce;color:#797673;padding:2px 8px;border-radius:10px;font-size:0.8em;font-weight:600;">' + og.no_contact + ' no contact</span>';
        html += '<span style="color:#999;font-size:0.85em;">avg ' + og.avg_days + 'd</span>';
        html += '</div></div>';
        // Always start collapsed; user clicks the header to expand. Prior logic
        // auto-expanded groups with <=5 prospects which was inconsistent.
        html += '<div style="display:none;padding:0 16px 12px;">';
        html += '<table style="width:100%;border-collapse:collapse;font-size:0.85em;">';
        html += '<tr style="background:#f8f9fa;"><th style="padding:6px 8px;text-align:left;">Name</th><th style="padding:6px 8px;text-align:left;">Contact</th><th style="padding:6px 8px;text-align:center;">Days</th><th style="padding:6px 8px;text-align:left;">Contact History</th><th style="padding:6px 8px;text-align:center;">Efforts</th><th style="padding:6px 8px;text-align:center;">Priority</th><th style="padding:6px 8px;text-align:center;">Actions</th></tr>';
        for (var pi = 0; pi < og.prospects.length; pi++) {
            var pr = og.prospects[pi];
            var prColor = pr.priority > 21 ? '#d13438' : pr.priority > 7 ? '#f0ad4e' : '#107c10';
            var prBg = pr.efforts === 0 && pr.days > 7 ? 'background:#fff8f0;' : '';
            html += '<tr style="' + prBg + 'cursor:pointer;" onclick="pbHealthToggleDetail(this,' + pr.people_id + ')">';
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;"><span style="font-weight:600;color:#2c3e50;">' + pbEsc(pr.name) + '</span>';
            if (pr.age) html += ' <span style="color:#999;font-size:0.85em;">(' + pr.age + ')</span>';
            // Score badge placeholder -- populated by pbHealthComputeOrgScores
            // after the user expands this org card (lazy, only if scorecard
            // is enabled). Empty span keeps the DOM stable for later inject.
            html += ' <span class="pb-health-score-badge" id="pb-health-score-' + pr.people_id + '"></span>';
            html += '</td>';
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;">';
            if (pr.cell_phone) html += '<a href="tel:' + pr.cell_phone + '" onclick="event.stopPropagation()">' + pbEsc(pr.cell_phone) + '</a>';
            if (pr.cell_phone && pr.email) html += '<br>';
            if (pr.email) html += '<a href="mailto:' + pr.email + '" onclick="event.stopPropagation()" style="font-size:0.85em;">' + pbEsc(pr.email) + '</a>';
            if (!pr.cell_phone && !pr.email) html += '<span style="color:#ccc;">-</span>';
            html += '</td>';
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;text-align:center;font-weight:600;">' + pr.days + '</td>';
            // Contact History badges
            var hcm = contactMap[String(pr.people_id)];
            var chHtml = '';
            if (hcm) {
                var hMethods = pbSettings.contact_methods || [];
                chHtml = '<span class="pb-contact-badge">';
                for (var hmi = 0; hmi < hMethods.length; hmi++) {
                    var hcode = hMethods[hmi].code;
                    var hcnt = (hcm.methods && hcm.methods[hcode]) || 0;
                    chHtml += '<span class="pb-contact-code' + (hcnt > 0 ? ' has-count' : '') + '">' + hcode + '(' + hcnt + ')</span>';
                }
                var hOther = (hcm.methods && hcm.methods['O']) || 0;
                if (hOther > 0) chHtml += '<span class="pb-contact-code has-count">O(' + hOther + ')</span>';
                chHtml += '</span>';
                if (hcm.lastDate) chHtml += '<div class="pb-text-muted" style="font-size:0.75em;">Last: ' + hcm.lastDate.substring(0, 10) + '</div>';
            } else {
                chHtml = '<span class="pb-text-muted" style="font-size:0.85em;">None</span>';
            }
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;">' + chHtml + '</td>';
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;text-align:center;">' + (pr.efforts > 0 ? pr.efforts + '<span style="color:#999;font-size:0.8em;"> (' + pr.last_effort + ')</span>' : '<span style="color:#d13438;">0</span>') + '</td>';
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;text-align:center;"><span style="color:' + prColor + ';font-weight:700;">' + pr.priority + '</span></td>';
            html += '<td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;text-align:center;white-space:nowrap;">';
            // Profile link stops propagation so opening the profile doesn't
            // also toggle the row's expand state.
            html += '<a href="/Person2/' + pr.people_id + '#ministry" target="_blank" class="pb-btn pb-btn-sm" onclick="event.stopPropagation()" style="padding:2px 6px;font-size:0.8em;text-decoration:none;" title="Open TouchPoint profile in a new tab">&#128100;</a> ';
            // Chevron that visually signals row expand -- rotates when open.
            // The click stays on the row (no stopPropagation) so tapping the
            // chevron toggles the expand the same as clicking elsewhere.
            html += '<span class="pb-health-chev" id="pb-health-chev-' + pr.people_id + '" title="Click anywhere on this row to expand"'
                  + ' style="display:inline-block;color:#888;font-weight:700;width:16px;text-align:center;transition:transform 0.18s ease;">&#10095;</span>';
            html += '</td>';
            html += '</tr>';
            // Expandable detail row (hidden) — store org_id for drop action
            html += '<tr class="pb-health-detail-row" id="pb-health-detail-' + pr.people_id + '" data-orgid="' + og.org_id + '" style="display:none;"><td colspan="7" style="padding:0;border-bottom:2px solid #e0e0e0;"><div style="padding:12px 16px;background:#fafafa;" id="pb-health-detail-content-' + pr.people_id + '"><span class="pb-text-muted">Loading...</span></div></td></tr>';
        }
        html += '</table></div></div>';
    }
    if (!d.orgs.length) html = '<div style="text-align:center;padding:30px;color:var(--pb-muted);">No groups with prospects found.</div>';
    contentEl.innerHTML = html;
}

function pbHealthToggleDetail(rowEl, peopleId) {
    var detailRow = document.getElementById('pb-health-detail-' + peopleId);
    if (!detailRow) return;
    var chev = document.getElementById('pb-health-chev-' + peopleId);
    if (detailRow.style.display !== 'none') {
        detailRow.style.display = 'none';
        if (chev) { chev.style.transform = ''; chev.style.color = '#888'; }
        return;
    }
    detailRow.style.display = '';
    if (chev) { chev.style.transform = 'rotate(90deg)'; chev.style.color = 'var(--pb-accent)'; }
    var contentEl = document.getElementById('pb-health-detail-content-' + peopleId);
    if (contentEl.dataset.loaded === 'true') return;

    contentEl.innerHTML = '<span class="pb-text-muted"><span class="pb-spin"></span> Loading person details...</span>';
    pbAjax({action: 'load_person_detail', people_id: peopleId, org_id: 0, ev_fields: ''}, function(d) {
        contentEl.dataset.loaded = 'true';
        if (!d.success) { contentEl.innerHTML = '<span style="color:var(--pb-danger)">Error loading details</span>'; return; }

        // Data is nested under d.detail
        var det = d.detail || d;
        var p = det.profile || {};
        var eng = det.engagement || {};
        var milestones = det.milestones || [];
        var famObj = det.family || {};
        var famMembers = famObj.FamilyDetailList || famObj.members || [];
        if (!Array.isArray(famMembers)) famMembers = [];
        var invs = det.involvements || [];
        var journey = det.journeyEvents || [];

        var html = '<div style="display:grid;grid-template-columns:auto 1fr 1fr 1fr;gap:16px;font-size:0.9em;">';

        // Column 0: Photo (click to enlarge -- reuses pbShowPhotoModal)
        html += '<div>';
        if (p.mediumPhoto || p.largePhoto) {
            var thumb = p.mediumPhoto || p.largePhoto;
            var big   = p.largePhoto  || p.mediumPhoto;
            html += '<img src="' + thumb + '" data-large="' + big + '"'
                  + ' onclick="event.stopPropagation();pbShowPhotoModal(this.dataset.large || this.src)"'
                  + ' style="width:70px;height:70px;border-radius:50%;object-fit:cover;cursor:zoom-in;"'
                  + ' title="Click to enlarge"'
                  + ' onerror="this.style.display=&quot;none&quot;">';
        }
        html += '</div>';

        // Column 1: Profile + Engagement
        html += '<div>';
        if (p.address && p.address.length > 3) html += '<div style="margin-bottom:3px;">&#127968; ' + pbEsc(p.address) + '</div>';
        if (p.homePhone) html += '<div>&#128222; Home: <a href="tel:' + p.homePhone + '">' + pbEsc(p.homePhone) + '</a></div>';
        if (p.workPhone) html += '<div>&#128188; Work: <a href="tel:' + p.workPhone + '">' + pbEsc(p.workPhone) + '</a></div>';
        if (p.employer) html += '<div>&#127970; ' + pbEsc(p.employer) + '</div>';
        // Engagement breakdown
        if (eng.score !== undefined) {
            var engColor = eng.score >= 60 ? '#107c10' : eng.score >= 30 ? '#f0ad4e' : '#d13438';
            html += '<div style="margin-top:8px;padding:8px 10px;background:#f8f9fa;border-radius:6px;border:1px solid #e8e8e8;">';
            html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">';
            html += '<strong>Engagement</strong>';
            html += '<span style="background:' + engColor + ';color:#fff;padding:2px 10px;border-radius:12px;font-weight:700;font-size:0.9em;">' + eng.score + '</span>';
            html += '</div>';
            if (eng.level) html += '<div style="font-size:0.85em;color:#666;margin-bottom:4px;">' + pbEsc(eng.level) + '</div>';
            // Factor bars
            var factors = eng.factors || {};
            var factorLabels = {attendRecency:'Attend Recency',attendFrequency:'Attend Frequency',groupInvolvement:'Groups',serving:'Serving'};
            for (var fk in factorLabels) {
                var fv = factors[fk];
                if (fv !== undefined && fv !== null) {
                    var fVal = typeof fv === 'object' ? (fv.score || fv.value || 0) : fv;
                    var barColor = fVal >= 60 ? '#107c10' : fVal >= 30 ? '#f0ad4e' : '#d13438';
                    html += '<div style="display:flex;align-items:center;gap:6px;margin:2px 0;font-size:0.8em;">';
                    html += '<span style="width:100px;color:#666;">' + factorLabels[fk] + '</span>';
                    html += '<div style="flex:1;background:#e8e8e8;border-radius:3px;height:8px;"><div style="width:' + Math.min(fVal, 100) + '%;background:' + barColor + ';border-radius:3px;height:8px;"></div></div>';
                    html += '<span style="width:25px;text-align:right;font-weight:600;color:' + barColor + ';">' + fVal + '</span>';
                    html += '</div>';
                }
            }
            // Extra stats
            var statParts = [];
            if (eng.groupCount !== undefined) statParts.push(eng.groupCount + ' groups');
            if (eng.servingCount !== undefined) statParts.push(eng.servingCount + ' serving');
            if (statParts.length) html += '<div style="font-size:0.8em;color:#999;margin-top:4px;">' + statParts.join(' &middot; ') + '</div>';
            html += '</div>';
        }
        html += '</div>';

        // Column 2: Milestones + Journey
        html += '<div>';
        if (milestones.length > 0) {
            html += '<strong>Milestones:</strong>';
            for (var mi = 0; mi < milestones.length; mi++) {
                var ms = milestones[mi];
                html += '<div style="padding:1px 0;font-size:0.9em;">' + (ms.icon || '') + ' ' + pbEsc(ms.label) + ': <strong>' + pbEsc(ms.date) + '</strong></div>';
            }
        }
        if (journey.length > 0) {
            html += '<strong style="margin-top:6px;display:block;">Recent Journey:</strong>';
            for (var ji = 0; ji < Math.min(journey.length, 4); ji++) {
                var je = journey[ji];
                var jeType = je.type || '';
                var jeColor = jeType === 'join' ? '#107c10' : jeType === 'attend' ? '#0078d4' : '#666';
                html += '<div style="padding:1px 0;font-size:0.85em;color:' + jeColor + ';">' + pbEsc(je.label || je.description || '') + '</div>';
            }
        }
        html += '</div>';

        // Column 3: Family + Involvements
        html += '<div>';
        if (famMembers.length > 0) {
            html += '<strong>Family (' + famMembers.length + '):</strong>';
            for (var fi = 0; fi < famMembers.length; fi++) {
                var fm = famMembers[fi];
                var fmName = fm.Name2 || fm.Name || fm.name || '';
                var fmAge = fm.Age || fm.age || '';
                var fmStatus = fm.MemberStatus || '';
                html += '<div style="padding:1px 0;font-size:0.9em;">&#128100; ' + pbEsc(fmName);
                if (fmAge) html += ' (' + fmAge + ')';
                if (fmStatus) html += ' <span style="font-size:0.8em;color:#999;">' + pbEsc(fmStatus) + '</span>';
                html += '</div>';
            }
        }
        if (invs.length > 0) {
            html += '<strong style="margin-top:6px;display:block;">Involvements (' + invs.length + '):</strong>';
            for (var ii = 0; ii < Math.min(invs.length, 5); ii++) {
                var inv = invs[ii];
                var invName = inv.name || inv.Name || '';
                var invAttend = inv.attended || 0;
                var invLast = inv.lastAttend || '';
                html += '<div style="padding:1px 0;font-size:0.85em;">&#128205; ' + pbEsc(invName);
                if (invLast) html += ' <span style="color:#999;">last ' + pbEsc(invLast) + '</span>';
                else if (invAttend === 0) html += ' <span style="color:#d13438;font-size:0.85em;">no meetings</span>';
                html += '</div>';
            }
            if (invs.length > 5) html += '<div class="pb-text-muted" style="font-size:0.85em;">+' + (invs.length - 5) + ' more</div>';
        }
        // Previous involvements (dropped). Often the most useful storyline
        // line -- a prospect who was active in 3 groups and dropped out of
        // all 3 reads very different from one who never joined anything.
        var pastInvs = det.pastInvolvements || [];
        if (pastInvs.length > 0) {
            html += '<strong style="margin-top:6px;display:block;color:#777;">Previous involvements (' + pastInvs.length + '):</strong>';
            for (var pii = 0; pii < Math.min(pastInvs.length, 5); pii++) {
                var pinv = pastInvs[pii];
                var pName = pinv.name || '';
                var dropped = pinv.droppedDate || '';
                var enrolled = pinv.enrollDate || '';
                html += '<div style="padding:1px 0;font-size:0.85em;color:#777;">&#10006; ' + pbEsc(pName);
                if (enrolled || dropped) {
                    html += ' <span style="color:#999;">';
                    if (enrolled) html += pbEsc(enrolled);
                    if (enrolled && dropped) html += ' &rarr; ';
                    if (dropped) html += pbEsc(dropped);
                    html += '</span>';
                }
                html += '</div>';
            }
            if (pastInvs.length > 5) html += '<div class="pb-text-muted" style="font-size:0.85em;">+' + (pastInvs.length - 5) + ' more</div>';
        }
        html += '</div>';

        html += '</div>';

        // === Priority Scorecard (lazy-shown when enabled in Settings) ===
        // Use whatever score we computed at org-expand time. If it's missing
        // -- the user didn't expand the org first -- show nothing rather
        // than firing another ajax inline.
        var scData = pbScoreMap[String(peopleId)] || pbScoreMap[peopleId];
        if (scData && pbSettings && pbSettings.scorecard && pbSettings.scorecard.enabled) {
            var _scColor = pbGetScoreColor(scData.score);
            html += '<div style="padding:8px 12px;background:#f0f4ff;border-radius:6px;border-left:3px solid '
                  + _scColor + ';margin-top:8px;font-size:0.85em;">';
            html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">';
            html += '<span style="font-weight:700;color:var(--pb-primary);">&#127919; Priority Scorecard '
                  + '<span style="font-weight:400;color:#94a3b8;font-size:0.9em;">(what should I do?)</span></span>';
            html += '<span style="background:' + _scColor + ';color:#fff;padding:2px 10px;border-radius:10px;'
                  + 'font-weight:700;">' + scData.score + ' / 100</span>';
            html += '</div>';
            html += pbBuildScoreDetail(scData);
            html += '</div>';
        }

        // Action bar
        html += '<div style="margin-top:10px;padding-top:10px;border-top:1px solid #e0e0e0;">';

        // Row 1: Contact methods
        html += '<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:6px;">';
        html += '<span style="font-weight:600;font-size:0.85em;color:#666;width:55px;">Contact:</span>';
        var methods = (pbSettings && pbSettings.contact_methods) ? pbSettings.contact_methods : [];
        for (var cmi = 0; cmi < methods.length; cmi++) {
            var cm = methods[cmi];
            var cmColors = {P:'#0078d4',E:'#107c10',T:'#6b69d6',V:'#d83b01',M:'#8764b8'};
            var cmColor = cmColors[cm.code] || '#666';
            html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthQuickLog(' + peopleId + ',&quot;' + (cm.keyword || '').replace(/"/g,'') + '&quot;)" style="padding:3px 10px;font-size:0.8em;background:' + cmColor + ';color:#fff;" title="Log ' + pbEsc(cm.label) + '">' + pbEsc(cm.code) + ' ' + pbEsc(cm.label) + '</button>';
        }
        html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthLogContact(' + peopleId + ')" style="padding:3px 10px;font-size:0.8em;">&#128221; Notes</button>';
        html += '</div>';

        // Row 2: Action — search any active involvement, then Add To / Move / Drop
        html += '<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">';
        html += '<span style="font-weight:600;font-size:0.85em;color:#666;width:55px;">Action:</span>';

        // Target: type-ahead search any active involvement by name.
        // Stores the selected orgId/orgName in hidden fields read by pbHealthDoAction.
        html += '<div style="position:relative;display:inline-block;">';
        html += '<input type="text" id="pb-health-action-input-' + peopleId + '" class="pb-input"'
              + ' placeholder="Type involvement name..." autocomplete="off"'
              + ' style="height:30px;font-size:0.85em;width:240px;padding:2px 6px;"'
              + ' oninput="pbHealthActionSearch(' + peopleId + ', this.value)"'
              + ' onfocus="pbHealthActionSearch(' + peopleId + ', this.value)" />';
        html += '<input type="hidden" id="pb-health-action-orgid-' + peopleId + '" value="" />';
        html += '<input type="hidden" id="pb-health-action-orgname-' + peopleId + '" value="" />';
        html += '<div id="pb-health-action-results-' + peopleId + '"'
              + ' style="display:none;position:absolute;top:32px;left:0;background:#fff;border:1px solid #ccc;'
              + 'max-height:240px;overflow-y:auto;width:340px;z-index:1000;'
              + 'box-shadow:0 2px 8px rgba(0,0,0,0.15);border-radius:4px;"></div>';
        html += '</div>';
        html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthDoAction(' + peopleId + ',&quot;add&quot;)" style="padding:3px 10px;font-size:0.8em;background:#107c10;color:#fff;" title="Add to selected involvement">&#10133; Add To</button>';
        html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthDoAction(' + peopleId + ',&quot;move&quot;)" style="padding:3px 10px;font-size:0.8em;background:#0078d4;color:#fff;" title="Move (drop current + add to selected)">&#10132; Move</button>';
        html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthDropFromOrg(' + peopleId + ')" style="padding:3px 10px;font-size:0.8em;background:#d13438;color:#fff;" title="Remove from this involvement">&#10060; Drop</button>';
        html += '<a href="/Person2/' + peopleId + '" target="_blank" class="pb-btn pb-btn-sm" style="padding:3px 10px;font-size:0.8em;text-decoration:none;">&#128100;</a>';
        html += '<a href="/Person2/' + peopleId + '#ministry" target="_blank" class="pb-btn pb-btn-sm" style="padding:3px 10px;font-size:0.8em;text-decoration:none;">&#128205;</a>';
        html += '</div>';

        html += '</div>';

        contentEl.innerHTML = html;
    });
}

// ---- Type-ahead involvement search for the Action row ----
var pbHealthActionSearchTimers = {};       // peopleId -> debounce timer
var pbHealthActionDocClickBound = false;   // bind outside-click handler once

function pbHealthActionSearch(peopleId, term) {
    // Debounce keystrokes so we don't flood the server while typing.
    if (pbHealthActionSearchTimers[peopleId]) clearTimeout(pbHealthActionSearchTimers[peopleId]);
    pbHealthActionSearchTimers[peopleId] = setTimeout(function() {
        pbHealthActionDoSearch(peopleId, term);
    }, 220);
    pbHealthActionBindOutsideClick();
}

function pbHealthActionDoSearch(peopleId, term) {
    var results = document.getElementById('pb-health-action-results-' + peopleId);
    if (!results) return;
    term = (term || '').trim();
    if (term.length < 2) {
        results.style.display = 'none';
        results.innerHTML = '';
        return;
    }
    results.style.display = '';
    results.innerHTML = '<div style="padding:8px 10px;color:#666;font-size:0.85em;">Searching...</div>';

    pbAjax({action: 'search_involvements', search_term: term}, function(d) {
        if (!d.success) {
            results.innerHTML = '<div style="padding:8px 10px;color:#d13438;font-size:0.85em;">Search failed</div>';
            return;
        }
        var invs = d.involvements || [];
        if (!invs.length) {
            results.innerHTML = '<div style="padding:8px 10px;color:#666;font-size:0.85em;">No matches</div>';
            return;
        }
        var html = '';
        for (var i = 0; i < invs.length && i < 30; i++) {
            var v = invs[i];
            var safeName = (v.name || '').replace(/"/g, '&quot;');
            var breadcrumb = '';
            if (v.program) breadcrumb += pbEsc(v.program);
            if (v.division) breadcrumb += (breadcrumb ? ' / ' : '') + pbEsc(v.division);
            html += '<div class="pb-health-action-result" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0;font-size:0.85em;"'
                  + ' onmouseover="this.style.background=\\'#eef5fb\\'" onmouseout="this.style.background=\\'\\'"'
                  + ' onclick="pbHealthActionPick(' + peopleId + ',' + v.orgId + ',&quot;' + safeName + '&quot;)">'
                  + '<div style="font-weight:600;color:#2c3e50;">' + pbEsc(v.name) + '</div>'
                  + (breadcrumb ? '<div style="font-size:0.8em;color:#999;">' + breadcrumb + '</div>' : '')
                  + '</div>';
        }
        results.innerHTML = html;
    });
}

function pbHealthActionPick(peopleId, orgId, orgName) {
    var input = document.getElementById('pb-health-action-input-' + peopleId);
    var oidEl = document.getElementById('pb-health-action-orgid-' + peopleId);
    var onEl = document.getElementById('pb-health-action-orgname-' + peopleId);
    var results = document.getElementById('pb-health-action-results-' + peopleId);
    if (input) input.value = orgName;
    if (oidEl) oidEl.value = String(orgId);
    if (onEl) onEl.value = orgName;
    if (results) { results.style.display = 'none'; results.innerHTML = ''; }
}

function pbHealthActionBindOutsideClick() {
    if (pbHealthActionDocClickBound) return;
    pbHealthActionDocClickBound = true;
    document.addEventListener('click', function(e) {
        // Close any open results dropdown that wasn't part of the click target.
        var openLists = document.querySelectorAll('[id^="pb-health-action-results-"]');
        for (var i = 0; i < openLists.length; i++) {
            var el = openLists[i];
            if (el.style.display === 'none') continue;
            // If click is inside this results element or its sibling input, keep open.
            var parent = el.parentNode;
            if (parent && (parent.contains(e.target))) continue;
            el.style.display = 'none';
        }
    });
}

function pbHealthDoAction(peopleId, mode) {
    var oidEl = document.getElementById('pb-health-action-orgid-' + peopleId);
    var onEl = document.getElementById('pb-health-action-orgname-' + peopleId);
    var orgId = oidEl ? parseInt(oidEl.value, 10) : 0;
    var orgName = onEl ? onEl.value : '';
    var input = document.getElementById('pb-health-action-input-' + peopleId);
    // If the user typed something then changed it without selecting, clear stale selection
    if (input && orgName && input.value !== orgName) {
        alert('The typed text does not match your selected involvement. Pick one from the dropdown.');
        return;
    }
    if (!orgId) { alert('Type an involvement name and choose one from the dropdown first.'); return; }

    var target = {pb_type: 'involvement', orgId: orgId, orgName: orgName};
    var label = orgName || ('Org #' + orgId);

    if (mode === 'move') {
        if (!confirm('MOVE to ' + label + '? This will remove from current and add to ' + label)) return;
        var detailRow = document.getElementById('pb-health-detail-' + peopleId);
        var curOrgId = detailRow ? detailRow.dataset.orgid : '';
        if (!curOrgId) { alert('Could not determine current organization to move from.'); return; }
        pbAjax({action: 'drop_from_org', people_id: peopleId, org_id: curOrgId}, function(d1) {
            if (d1.success) {
                pbAjax({action: 'process_action', target_data: JSON.stringify(target), people_ids: String(peopleId)}, function(d2) {
                    if (d2.success) {
                        alert('Moved to ' + label);
                        var row = detailRow.previousElementSibling;
                        if (row) row.style.display = 'none';
                        detailRow.style.display = 'none';
                    } else { alert('Dropped but failed to add: ' + (d2.message || '')); }
                });
            } else { alert('Failed to drop: ' + (d1.message || '')); }
        });
    } else {
        if (!confirm('Add to ' + label + '? Person stays in current involvement too.')) return;
        pbAjax({action: 'process_action', target_data: JSON.stringify(target), people_ids: String(peopleId)}, function(d) {
            if (d.success) { alert('Added to ' + label); }
            else { alert('Error: ' + (d.message || 'Unknown')); }
        });
    }
}

function pbHealthDropFromOrg(peopleId) {
    // Find the org_id from the detail row
    var detailRow = document.getElementById('pb-health-detail-' + peopleId);
    var orgId = detailRow ? detailRow.dataset.orgid : '';
    if (!orgId) { alert('Could not determine organization'); return; }
    if (!confirm('Remove this person from the involvement? This cannot be undone.')) return;

    pbAjax({action: 'drop_from_org', people_id: peopleId, org_id: orgId}, function(d) {
        if (d.success) {
            alert('Removed from involvement');
            // Hide the prospect row
            var row = detailRow.previousElementSibling;
            if (row) row.style.display = 'none';
            detailRow.style.display = 'none';
        } else {
            alert('Error: ' + (d.message || 'Unknown'));
        }
    });
}

function pbHealthQuickLog(peopleId, keyword) {
    var notes = prompt('Quick note (optional):') || keyword + ' contact logged';
    if (notes === null) return;
    pbAjax({action: 'log_contact', people_id: peopleId, keyword: keyword, note_text: notes}, function(d) {
        if (d.success) {
            // Clear cached detail so it reloads
            var cel = document.getElementById('pb-health-detail-content-' + peopleId);
            if (cel) cel.dataset.loaded = '';
            alert('Contact logged!');
        } else {
            alert('Error: ' + (d.message || 'Unknown'));
        }
    });
}

function pbHealthLogContact(peopleId) {
    // Ensure settings are loaded (needed for contact_methods)
    if (!pbSettings || !pbSettings.contact_methods) {
        pbAjax({action: 'load_settings'}, function(d) {
            if (d.success) {
                pbSettings = d.settings || {};
                pbHealthLogContact(peopleId); // Retry
            }
        });
        return;
    }
    // Use the existing contact modal
    if (typeof pbShowContactModal === 'function') {
        pbShowContactModal(peopleId);
        return;
    }
}

function pbSndToggleHealthOrg(el) {
    var next = el.nextElementSibling;
    if (!next) return;
    var opening = (next.style.display === 'none');
    next.style.display = opening ? 'block' : 'none';
    // Lazy-compute scorecard scores when an org is first opened. No-op when
    // scorecard is disabled in Settings -- the row badges stay empty.
    if (opening) {
        var card = el.closest ? el.closest('.pb-health-org-card') : null;
        if (card) pbHealthComputeOrgScores(card);
    }
}

function pbHealthComputeOrgScores(cardEl) {
    if (!cardEl) return;
    if (!pbSettings || !pbSettings.scorecard || !pbSettings.scorecard.enabled) return;
    var orgId = cardEl.getAttribute('data-orgid');
    if (pbHealthScoresLoadedOrgs[orgId]) return;  // already computed
    var badges = cardEl.querySelectorAll('.pb-health-score-badge');
    if (!badges.length) return;
    var pids = [];
    for (var b = 0; b < badges.length; b++) {
        var bid = badges[b].id.replace('pb-health-score-', '');
        if (bid) pids.push(bid);
        badges[b].innerHTML = '<span style="color:#bbb;font-size:0.8em;">scoring...</span>';
    }
    if (!pids.length) return;
    pbHealthScoresLoadedOrgs[orgId] = true;

    // Build member_data the way pbComputeScores does -- from caches we
    // already have in scope (pbHealthMemberCache + pbGrpContactMap).
    var memberData = {};
    for (var pi = 0; pi < pids.length; pi++) {
        var pk = pids[pi];
        var cm = pbGrpContactMap[pk] || {};
        var meta = pbHealthMemberCache[pk] || {};
        memberData[pk] = {
            memberStatus:  meta.memberStatus || '',
            weightedTotal: cm.weightedTotal || cm.total || 0,
            taskNoteTotal: cm.total || 0
        };
    }
    pbAjax({
        action: 'compute_scores',
        people_ids: pids.join(','),
        scorecard_config: JSON.stringify(pbSettings.scorecard),
        member_data: JSON.stringify(memberData)
    }, function(d) {
        if (!d || !d.success) {
            // Clear the "scoring..." placeholders so failure doesn't sit there.
            for (var bb = 0; bb < badges.length; bb++) badges[bb].innerHTML = '';
            return;
        }
        var scores = d.scores || {};
        for (var k in scores) {
            if (!scores.hasOwnProperty(k)) continue;
            pbScoreMap[k] = scores[k];
            var badgeEl = document.getElementById('pb-health-score-' + k);
            if (badgeEl) {
                var sc = scores[k];
                var col = pbGetScoreColor(sc.score);
                badgeEl.innerHTML = '<span title="Priority Scorecard -- '
                    + sc.score + ' / 100" style="display:inline-block;background:'
                    + col + ';color:#fff;padding:1px 7px;border-radius:10px;'
                    + 'font-size:0.75em;font-weight:700;margin-left:6px;'
                    + 'vertical-align:middle;">' + sc.score + '</span>';
            }
        }
        // For any pid that didn't come back, clear its placeholder.
        for (var bb2 = 0; bb2 < badges.length; bb2++) {
            if (!badges[bb2].innerHTML || badges[bb2].innerHTML.indexOf('scoring') !== -1) {
                badges[bb2].innerHTML = '';
            }
        }
    });
}

function pbSndLogTypeBadge(entry) {
    // Show what kind of run this log entry came from so a preview isn't
    // mistaken for an actual email send. dry_run=true wins regardless of
    // trigger, since that's the row where Emails=0 by definition.
    var label, bg, fg;
    if (entry && entry.dry_run) {
        label = 'PREVIEW';
        bg = '#f0f0f0'; fg = '#666';
    } else {
        switch ((entry && entry.triggered_by) || 'manual') {
            case 'batch':   label = 'SCHEDULED'; bg = '#cfe3ff'; fg = '#1c4f8c'; break;
            case 'oneoff':  label = 'ONE-OFF';   bg = '#ffe0b2'; fg = '#7a4a00'; break;
            case 'preview': label = 'PREVIEW';   bg = '#f0f0f0'; fg = '#666'; break;
            default:        label = 'MANUAL';    bg = '#d4edda'; fg = '#155724'; break;
        }
    }
    return '<span style="display:inline-block;padding:1px 7px;border-radius:10px;font-size:0.7em;font-weight:700;letter-spacing:0.3px;background:'
         + bg + ';color:' + fg + ';">' + label + '</span>';
}

function pbSndLoadLog() {
    pbAjax({action: 'get_sender_log'}, function(d) {
        var el = document.getElementById('pb-snd-log');
        if (!d.success || !d.log || d.log.length === 0) {
            el.innerHTML = '<span class="pb-text-muted">No send history yet.</span>';
            return;
        }
        var html = '<table class="pb-table"><thead><tr><th>Time</th><th>Type</th><th>Sender</th><th>Prospects</th><th>Recipients</th><th>Emails</th><th>Errors</th></tr></thead><tbody>';
        for (var i = d.log.length - 1; i >= 0; i--) {
            var l = d.log[i];
            html += '<tr' + (l.dry_run ? ' style="opacity:0.78;"' : '') + '>';
            html += '<td>' + pbEsc(l.timestamp || '') + '</td>';
            html += '<td>' + pbSndLogTypeBadge(l) + '</td>';
            html += '<td>' + pbEsc(l.sender_name || '') + '</td>';
            html += '<td>' + (l.prospects || 0) + '</td><td>' + (l.recipients || 0) + '</td>';
            html += '<td>' + (l.emails_sent || 0) + '</td>';
            html += '<td>' + (l.errors || 0) + '</td></tr>';
        }
        html += '</tbody></table>';
        el.innerHTML = html;
    });
}

function pbEsc(s) {
    if (typeof pbEsc._el === 'undefined') pbEsc._el = document.createElement('div');
    pbEsc._el.textContent = s || '';
    return pbEsc._el.innerHTML;
}

// ============================================================
// CONFIGURATIONS
// ============================================================
function pbLoadConfigs() {
    pbAjax({action: 'load_configs'}, function(d) {
        if (d.success) {
            pbConfigs = d.configs || [];
            pbRenderConfigList();
        }
    });
}

function pbRenderConfigList() {
    var el = document.getElementById('pb-config-list');
    if (!pbConfigs.length) {
        el.innerHTML = '<div class="pb-empty"><div class="pb-empty-icon" style="font-size:2.5em;">&#9935;</div><p>No configurations yet. Time to stake your claim!<br>Click "+ New Configuration" to get started.</p></div>';
        return;
    }
    // Per-card stats (Prospects / Contacted-30d / No Contact) were removed:
    // each one fired a separate get_config_stats ajax, scaling the page open
    // cost linearly with the number of configs. Drilling into a config
    // already shows the same numbers (and more), so the card-level summary
    // wasn't pulling its weight. Stats can be re-added behind a lazy
    // "expand" affordance later if needed.
    var html = '<div class="pb-grp-grid">';
    for (var i = 0; i < pbConfigs.length; i++) {
        var c = pbConfigs[i];
        var srcLabel = c.source ? (c.source.pb_type === 'involvement' ? (c.source.orgName || 'Org #' + (c.source.orgId || '')) : c.source.pb_type === 'tag' ? 'Tag: ' + (c.source.tagName || '') : 'Query: ' + (c.source.queryId || '')) : 'No source';

        html += '<div class="pb-grp-card" style="border-left-color:var(--pb-accent);cursor:pointer;" data-idx="' + i + '" onclick="pbLaunchConfig(parseInt(this.dataset.idx))">';
        html += '<div class="pb-grp-card-title">' + (c.name || 'Untitled') + '</div>';
        html += '<div class="pb-text-muted" style="font-size:0.75em;margin-bottom:12px;">' + srcLabel + '</div>';
        html += '<div class="pb-grp-actions">';
        html += '<button class="pb-btn pb-btn-sm pb-btn-primary" data-idx="' + i + '" onclick="event.stopPropagation();pbLaunchConfig(parseInt(this.dataset.idx))">Open</button>';
        html += '<button class="pb-btn pb-btn-sm" data-idx="' + i + '" onclick="event.stopPropagation();pbEditConfig(parseInt(this.dataset.idx))">Edit</button>';
        html += '</div>';
        html += '</div>';
    }
    html += '</div>';
    el.innerHTML = html;
}


function pbShowConfigModal(config) {
    var modal = document.getElementById('pb-config-modal');
    document.getElementById('pb-config-modal-title').textContent = config ? 'Edit Configuration' : 'New Configuration';
    document.getElementById('pb-cfg-id').value = config ? (config.id || '') : '';
    document.getElementById('pb-cfg-name').value = config ? (config.name || '') : '';
    document.getElementById('pb-cfg-max-prospects').value = config ? (config.maxProspects || '') : '';
    // Reveal the Delete button only when editing an existing config so a
    // new-configuration flow can't accidentally delete anything.
    var delBtn = document.getElementById('pb-config-modal-delete');
    if (delBtn) delBtn.style.display = config ? 'inline-block' : 'none';

    // Source
    var srcType = config && config.source ? config.source.pb_type : 'involvement';
    document.getElementById('pb-cfg-source-type').value = srcType;
    document.getElementById('pb-cfg-org-id').value = config && config.source ? (config.source.orgId || '') : '';
    document.getElementById('pb-cfg-inv-selected').textContent = config && config.source && config.source.orgName ? config.source.orgName + ' (#' + config.source.orgId + ')' : '';
    document.getElementById('pb-cfg-tag-name').value = config && config.source ? (config.source.tagName || '') : '';
    document.getElementById('pb-cfg-tag-selected').textContent = config && config.source ? (config.source.tagName || '') : '';
    document.getElementById('pb-cfg-query-id').value = config && config.source ? (config.source.queryId || '') : '';
    pbToggleSourceFields();

    // Member types - load from database
    var mts = config ? (config.memberTypes || []) : [];
    pbLoadMemberTypes(mts);
    document.getElementById('pb-cfg-no-contact-days').value = config ? (config.noContactDays || 90) : 90;

    // Display fields
    pbConfigFields = config ? (config.displayFields || []).slice() : [
        {fieldType: 'person', sourceField: 'Name2', label: 'Full Name', visible: true},
        {fieldType: 'person', sourceField: 'EmailAddress', label: 'Email', visible: true},
        {fieldType: 'person', sourceField: 'CellPhone', label: 'Cell Phone', visible: true},
        {fieldType: 'person', sourceField: 'Age', label: 'Age', visible: true},
        {fieldType: 'person', sourceField: 'MemberStatus', label: 'Member Status', visible: true},
    ];
    pbRenderConfigFields();
    pbPopulateFieldOptions();

    // Cross-query flags
    pbConfigFlags = config ? (config.crossFlags || []).slice() : [];
    pbRenderConfigFlags();

    // Target actions
    pbConfigActions = config ? (config.targetActions || []).slice() : [];
    pbRenderConfigActions();
    // Reset action fields - show tag fields since tag is default
    document.getElementById('pb-cfg-add-action-type').value = 'tag';
    document.getElementById('pb-action-tag-fields').style.display = '';
    document.getElementById('pb-action-inv-fields').style.display = 'none';
    document.getElementById('pb-action-inv-results').innerHTML = '';

    modal.classList.add('active');
}

var pbMemberTypesFromDB = [];
function pbLoadMemberTypes(selectedIds) {
    if (pbMemberTypesFromDB.length) {
        pbRenderMemberTypeCheckboxes(selectedIds);
        // CRITICAL: the modal HTML re-inserts hardcoded <option value="230">
        // every time it opens. Without this call on the cached path the
        // user picks "Prospect" from a dropdown that says value="230" --
        // a number that means "Prospect" on default seeds but "InActive"
        // on churches that customized lookup.MemberType. The wrong id then
        // gets saved into the action JSON and downstream SetMemberType
        // lands the person under the wrong member type. Bug discovered
        // 2026-06-15 at FBCH where Prospect is id=311, not 230.
        pbUpdateActionMemberTypeDropdown();
        return;
    }
    pbAjax({action: 'get_member_types'}, function(d) {
        if (d.success) {
            pbMemberTypesFromDB = d.memberTypes || [];
            pbRenderMemberTypeCheckboxes(selectedIds);
            // Also update the target action member type dropdown
            pbUpdateActionMemberTypeDropdown();
        }
    });
}

function pbRenderMemberTypeCheckboxes(selectedIds) {
    var el = document.getElementById('pb-cfg-mt-list');
    var html = '<label class="pb-checkbox"><input type="checkbox" value="0" class="pb-cfg-mt" onchange="pbToggleAllMemberTypes(this.checked)"> <strong>All Types</strong></label>';
    for (var i = 0; i < pbMemberTypesFromDB.length; i++) {
        var mt = pbMemberTypesFromDB[i];
        var checked = selectedIds.indexOf(mt.id) >= 0 ? ' checked' : '';
        html += '<label class="pb-checkbox"><input type="checkbox" value="' + mt.id + '" class="pb-cfg-mt"' + checked + '> ' + mt.description + ' (' + mt.id + ')</label>';
    }
    el.innerHTML = html;
    // Check "All" if nothing specific is selected
    if (!selectedIds.length) {
        var allCb = el.querySelector('input[value="0"]');
        if (allCb) allCb.checked = true;
    }
}

function pbToggleAllMemberTypes(checked) {
    document.querySelectorAll('.pb-cfg-mt').forEach(function(cb) {
        if (cb.value !== '0') cb.checked = false;
    });
}

function pbUpdateActionMemberTypeDropdown() {
    var sel = document.getElementById('pb-action-member-type');
    if (!sel) return;
    sel.innerHTML = '';
    pbMemberTypeLabels = {};
    for (var i = 0; i < pbMemberTypesFromDB.length; i++) {
        var mt = pbMemberTypesFromDB[i];
        pbMemberTypeLabels[String(mt.id)] = mt.description;
        var opt = document.createElement('option');
        opt.value = mt.id;
        opt.textContent = mt.description;
        if (mt.description === 'Member') opt.selected = true;
        sel.appendChild(opt);
    }
}

function pbCloseConfigModal() {
    document.getElementById('pb-config-modal').classList.remove('active');
}

var pbConfigFields = [];
var pbConfigFlags = [];
var pbConfigActions = [];

function pbRenderConfigFields() {
    var el = document.getElementById('pb-cfg-fields-list');
    if (!pbConfigFields.length) {
        el.innerHTML = '<div class="pb-text-muted" style="padding:8px;">No fields added yet.</div>';
        return;
    }
    var html = '';
    for (var i = 0; i < pbConfigFields.length; i++) {
        var f = pbConfigFields[i];
        html += '<div class="pb-flex-between" style="padding:3px 0;border-bottom:1px solid #eee;">';
        html += '<span><span class="pb-badge pb-badge-muted">' + f.fieldType + '</span> ' + f.label + '</span>';
        html += '<button class="pb-btn pb-btn-danger pb-btn-sm" onclick="pbConfigFields.splice(' + i + ',1);pbRenderConfigFields();" style="padding:2px 6px;">&times;</button>';
        html += '</div>';
    }
    el.innerHTML = html;
}

function pbPopulateFieldOptions() {
    var type = document.getElementById('pb-cfg-add-field-type').value;
    var sel = document.getElementById('pb-cfg-add-field-source');
    sel.innerHTML = '';
    var catalog = {
        person: [
            {s:'Name2',l:'Full Name'},{s:'FirstName',l:'First Name'},{s:'LastName',l:'Last Name'},
            {s:'NickName',l:'Nickname'},{s:'PreferredName',l:'Preferred Name'},{s:'EmailAddress',l:'Email'},
            {s:'CellPhone',l:'Cell Phone'},{s:'HomePhone',l:'Home Phone'},{s:'Age',l:'Age'},
            {s:'BDate',l:'Date of Birth'},{s:'Gender',l:'Gender'},{s:'MaritalStatus',l:'Marital Status'},
            {s:'MemberStatus',l:'Member Status'},{s:'MemberType',l:'Member Type'},
            {s:'JoinDate',l:'Join Date'},{s:'FullAddress',l:'Full Address'},{s:'CampusName',l:'Campus'},
            {s:'EnrollmentDate',l:'Enrollment Date'}
        ],
        family: [
            {s:'FamilyDetail',l:'Family Summary (rich)'},{s:'SpouseName',l:'Spouse Name'},{s:'SpouseDetail',l:'Spouse Detail'},
            {s:'Parents',l:'Parent(s)'},{s:'ParentPhones',l:'Parent Phone(s)'},
            {s:'ParentEmails',l:'Parent Email(s)'},{s:'Children',l:'Children'},{s:'FamilyMembers',l:'All Family Members'}
        ],
        involvement: [
            {s:'CurrentInvolvements',l:'Current Involvements'}
        ],
        medical: [
            {s:'emcontact',l:'Emergency Contact'},{s:'emphone',l:'Emergency Phone'},
            {s:'doctor',l:'Doctor'},{s:'docphone',l:'Doctor Phone'},
            {s:'insurance',l:'Insurance'},{s:'policy',l:'Policy #'},
            {s:'MedAllergy',l:'Allergies'},{s:'CustodyIssue',l:'Custody Issue'}
        ]
    };
    if (catalog[type]) {
        for (var i = 0; i < catalog[type].length; i++) {
            var opt = document.createElement('option');
            opt.value = catalog[type][i].s;
            opt.textContent = catalog[type][i].l;
            sel.appendChild(opt);
        }
    } else if (type === 'extravalue') {
        var opt = document.createElement('option');
        opt.value = '__custom__';
        opt.textContent = '(Type field name below)';
        sel.appendChild(opt);
    } else if (type === 'regquestion') {
        var opt = document.createElement('option');
        opt.value = '__custom__';
        opt.textContent = '(Type question text below)';
        sel.appendChild(opt);
    }
}

function pbAddDisplayField() {
    var type = document.getElementById('pb-cfg-add-field-type').value;
    var sel = document.getElementById('pb-cfg-add-field-source');
    var sourceField = sel.value;
    var label = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].textContent : sourceField;

    if (sourceField === '__custom__') {
        sourceField = prompt('Enter the field name:');
        if (!sourceField) return;
        label = sourceField;
    }
    pbConfigFields.push({fieldType: type, sourceField: sourceField, label: label, visible: true});
    pbRenderConfigFields();
}

function pbRenderConfigFlags() {
    var el = document.getElementById('pb-cfg-flags-list');
    if (!pbConfigFlags.length) {
        el.innerHTML = '<div class="pb-text-muted">No cross-query flags configured.</div>';
        return;
    }
    var html = '';
    for (var i = 0; i < pbConfigFlags.length; i++) {
        var f = pbConfigFlags[i];
        html += '<div class="pb-flex-between" style="padding:5px 0;border-bottom:1px solid #eee;">';
        html += '<div>';
        html += '<span class="pb-badge pb-badge-primary">' + f.pb_type + '</span> ';
        html += '<strong>' + (f.label || f.pb_type) + '</strong>';
        if (f.programId) html += ' <span class="pb-text-muted">(Program: ' + f.programId + ')</span>';
        if (f.orgId) html += ' <span class="pb-text-muted">(Org: ' + f.orgId + ')</span>';
        if (f.evField) html += ' <span class="pb-text-muted">(EV: ' + f.evField + ')</span>';
        html += '</div>';
        html += '<button class="pb-btn pb-btn-danger pb-btn-sm" onclick="pbConfigFlags.splice(' + i + ',1);pbRenderConfigFlags();" style="padding:2px 6px;">&times;</button>';
        html += '</div>';
    }
    el.innerHTML = html;
}

function pbAddCrossFlag() {
    var type = document.getElementById('pb-cfg-add-flag-type').value;
    var flag = {pb_type: type, id: type + '_' + Date.now(), enabled: true};

    if (type === 'children_attending' || type === 'parents_not_attending') {
        var pid = prompt('Enter Program ID:');
        if (!pid) return;
        flag.programId = parseInt(pid);
        flag.label = (type === 'children_attending' ? 'Children in Prog ' : 'Parents Not in Prog ') + pid;
    } else if (type === 'spouse_in_org') {
        var oid = prompt('Enter Organization ID:');
        if (!oid) return;
        flag.orgId = parseInt(oid);
        flag.label = 'Spouse in Org ' + oid;
    } else if (type === 'has_extra_value') {
        var ev = prompt('Enter Extra Value field name:');
        if (!ev) return;
        flag.evField = ev;
        flag.label = 'Has EV: ' + ev;
    }

    pbConfigFlags.push(flag);
    pbRenderConfigFlags();
}

function pbRenderConfigActions() {
    var el = document.getElementById('pb-cfg-actions-list');
    if (!pbConfigActions.length) {
        el.innerHTML = '<div class="pb-text-muted">No target actions configured.</div>';
        return;
    }
    var html = '';
    for (var i = 0; i < pbConfigActions.length; i++) {
        var a = pbConfigActions[i];
        html += '<div class="pb-flex-between" style="padding:5px 0;border-bottom:1px solid #eee;">';
        html += '<div>';
        html += '<span class="pb-badge pb-badge-success">' + a.pb_type + '</span> ';
        html += '<strong>' + (a.label || a.pb_type) + '</strong>';
        html += '</div>';
        html += '<button class="pb-btn pb-btn-danger pb-btn-sm" onclick="pbConfigActions.splice(' + i + ',1);pbRenderConfigActions();" style="padding:2px 6px;">&times;</button>';
        html += '</div>';
    }
    el.innerHTML = html;
}

function pbToggleActionFields() {
    var type = document.getElementById('pb-cfg-add-action-type').value;
    document.getElementById('pb-action-tag-fields').style.display = type === 'tag' ? '' : 'none';
    document.getElementById('pb-action-inv-fields').style.display = type === 'involvement' ? '' : 'none';
    if (type === 'involvement') pbLoadActionPrograms();
}

function pbLoadActionPrograms() {
    pbAjax({action: 'get_programs'}, function(d) {
        if (!d.success) return;
        var sel = document.getElementById('pb-action-prog-filter');
        sel.innerHTML = '<option value="">All Programs</option>';
        for (var i = 0; i < (d.programs || []).length; i++) {
            var p = d.programs[i];
            var opt = document.createElement('option');
            opt.value = p.id;
            opt.textContent = p.name;
            sel.appendChild(opt);
        }
    });
}

function pbActionProgChanged() {
    var progId = document.getElementById('pb-action-prog-filter').value;
    var divSel = document.getElementById('pb-action-div-filter');
    divSel.innerHTML = '<option value="">All Divisions</option>';
    if (progId) {
        pbAjax({action: 'get_divisions', program_id: progId}, function(d) {
            if (!d.success) return;
            for (var i = 0; i < (d.divisions || []).length; i++) {
                var div = d.divisions[i];
                var opt = document.createElement('option');
                opt.value = div.id;
                opt.textContent = div.name;
                divSel.appendChild(opt);
            }
        });
    }
    pbActionSearchInv();
}

var pbActionSearchTimer = null;
function pbActionSearchInv() {
    clearTimeout(pbActionSearchTimer);
    var term = document.getElementById('pb-action-inv-search').value;
    var progId = document.getElementById('pb-action-prog-filter').value;
    var divId = document.getElementById('pb-action-div-filter').value;
    if (!term && !progId && !divId) {
        document.getElementById('pb-action-inv-results').innerHTML = '<div style="padding:15px;text-align:center;color:var(--pb-muted);">Select a program or type to search</div>';
        return;
    }
    pbActionSearchTimer = setTimeout(function() {
        var params = {action: 'search_involvements', search_term: term || ''};
        if (progId) params.program_id = progId;
        if (divId) params.division_id = divId;
        pbAjax(params, function(d) {
            if (!d.success) return;
            var el = document.getElementById('pb-action-inv-results');
            var html = '';
            // Build set of already-added orgIds for visual feedback
            var addedOrgs = {};
            for (var a = 0; a < pbConfigActions.length; a++) {
                if (pbConfigActions[a].pb_type === 'involvement') addedOrgs[pbConfigActions[a].orgId] = true;
            }
            for (var i = 0; i < (d.involvements || []).length; i++) {
                var inv = d.involvements[i];
                var already = addedOrgs[inv.orgId];
                html += '<div class="pb-search-item" data-orgid="' + inv.orgId + '" data-name="' + (inv.name || '').replace(/"/g, '&quot;') + '" onclick="pbActionClickInv(parseInt(this.dataset.orgid), this.dataset.name, this)" style="' + (already ? 'background:rgba(39,174,96,0.1);' : '') + 'cursor:pointer;">';
                html += (already ? '<span class="pb-badge pb-badge-success" style="margin-right:5px;">added</span>' : '<span class="pb-badge pb-badge-primary" style="margin-right:5px;">+ click to add</span>');
                html += '<strong>' + inv.name + '</strong> (#' + inv.orgId + ')';
                html += '<br><span class="pb-text-muted" style="margin-left:5px;">' + (inv.program || '') + ' / ' + (inv.division || '') + ' - ' + inv.memberCount + ' members</span>';
                html += '</div>';
            }
            el.innerHTML = html || '<div style="padding:15px;text-align:center;color:var(--pb-muted);">No results</div>';
        });
    }, 300);
}

var pbMemberTypeLabels = {};  // Populated dynamically from DB

function pbActionClickInv(orgId, name, el) {
    // Check if already added
    for (var i = 0; i < pbConfigActions.length; i++) {
        if (pbConfigActions[i].pb_type === 'involvement' && pbConfigActions[i].orgId === orgId) {
            return; // Already added, ignore
        }
    }
    // Get selected member type
    var mtSel = document.getElementById('pb-action-member-type');
    var mtId = mtSel ? parseInt(mtSel.value) : 220;
    var mtLabel = (mtSel && mtSel.selectedIndex >= 0) ? mtSel.options[mtSel.selectedIndex].text : (pbMemberTypeLabels[String(mtId)] || 'Member');
    // Add immediately as a target action
    pbConfigActions.push({
        pb_type: 'involvement',
        orgId: orgId,
        orgName: name,
        memberTypeId: mtId,
        label: name + ' (#' + orgId + ') as ' + mtLabel,
        subgroups: []
    });
    pbRenderConfigActions();
    // Update the clicked item visual
    if (el) {
        el.style.background = 'rgba(39,174,96,0.1)';
        var badge = el.querySelector('.pb-badge');
        if (badge) { badge.className = 'pb-badge pb-badge-success'; badge.style.marginRight = '5px'; badge.textContent = 'added as ' + mtLabel; }
    }
}

function pbAddTagAction() {
    var name = document.getElementById('pb-action-tag-name').value.trim();
    if (!name) { alert('Please enter a tag name.'); return; }
    pbConfigActions.push({pb_type: 'tag', tagName: name, label: 'Tag: ' + name});
    pbRenderConfigActions();
    document.getElementById('pb-action-tag-name').value = '';
}

function pbSaveConfig() {
    var config = {
        id: document.getElementById('pb-cfg-id').value || '',
        name: document.getElementById('pb-cfg-name').value,
        source: pbBuildSourceFromModal(),
        memberTypes: [],
        displayFields: pbConfigFields,
        crossFlags: pbConfigFlags,
        targetActions: pbConfigActions,
        noContactDays: parseInt(document.getElementById('pb-cfg-no-contact-days').value) || 90,
        maxProspects: parseInt(document.getElementById('pb-cfg-max-prospects').value) || 0
    };

    var allChecked = false;
    document.querySelectorAll('.pb-cfg-mt:checked').forEach(function(cb) {
        var v = parseInt(cb.value);
        if (v === 0) allChecked = true;
        else config.memberTypes.push(v);
    });
    if (allChecked) config.memberTypes = [];  // Empty = all types (no filter)

    if (!config.name) { alert('Please enter a configuration name.'); return; }

    pbAjax({action: 'save_config', config_data: JSON.stringify(config)}, function(d) {
        if (d.success) {
            pbCloseConfigModal();
            pbLoadConfigs();
        } else {
            alert('Error: ' + (d.message || 'Save failed'));
        }
    });
}

function pbBuildSourceFromModal() {
    var type = document.getElementById('pb-cfg-source-type').value;
    if (type === 'involvement') {
        return {pb_type: 'involvement', orgId: parseInt(document.getElementById('pb-cfg-org-id').value) || 0, orgName: document.getElementById('pb-cfg-inv-selected').textContent};
    } else if (type === 'tag') {
        return {pb_type: 'tag', tagName: document.getElementById('pb-cfg-tag-name').value};
    } else {
        return {pb_type: 'query', queryId: document.getElementById('pb-cfg-query-id').value};
    }
}

function pbEditConfig(idx) { pbShowConfigModal(pbConfigs[idx]); }

function pbDeleteConfig(id, onSuccess) {
    if (!confirm('Delete this configuration?')) return;
    pbAjax({action: 'delete_config', config_id: id}, function(d) {
        if (d.success) {
            pbLoadConfigs();
            if (typeof onSuccess === 'function') onSuccess();
        }
    });
}

// Delete from inside the Edit Configuration modal. Routes through the same
// pbDeleteConfig so the server-side action is identical to the card delete.
function pbConfigModalDelete() {
    var id = document.getElementById('pb-cfg-id').value;
    if (!id) return;
    pbDeleteConfig(id, function() {
        pbCloseConfigModal();
    });
}

function pbToggleSourceFields() {
    var type = document.getElementById('pb-cfg-source-type').value;
    document.getElementById('pb-cfg-source-inv').style.display = type === 'involvement' ? '' : 'none';
    document.getElementById('pb-cfg-source-tag').style.display = type === 'tag' ? '' : 'none';
    document.getElementById('pb-cfg-source-query').style.display = type === 'query' ? '' : 'none';
}

// Source search helpers
var pbSearchTimer = null;
function pbSearchInvolvements(term) {
    clearTimeout(pbSearchTimer);
    if (term.length < 2) { document.getElementById('pb-cfg-inv-results').style.display = 'none'; return; }
    pbSearchTimer = setTimeout(function() {
        pbAjax({action: 'search_involvements', search_term: term}, function(d) {
            if (!d.success) return;
            var el = document.getElementById('pb-cfg-inv-results');
            var html = '';
            for (var i = 0; i < (d.involvements || []).length; i++) {
                var inv = d.involvements[i];
                html += '<div class="pb-search-item" data-orgid="' + inv.orgId + '" data-name="' + inv.name.replace(/"/g, '&quot;') + '" onclick="pbSelectInvolvement(parseInt(this.dataset.orgid), this.dataset.name)">';
                html += '<strong>' + inv.name + '</strong> (#' + inv.orgId + ')';
                html += '<br><span class="pb-text-muted">' + (inv.program || '') + ' / ' + (inv.division || '') + ' - ' + inv.memberCount + ' members</span>';
                html += '</div>';
            }
            el.innerHTML = html || '<div class="pb-search-item pb-text-muted">No results</div>';
            el.style.display = '';
        });
    }, 300);
}

function pbSelectInvolvement(orgId, name) {
    document.getElementById('pb-cfg-org-id').value = orgId;
    document.getElementById('pb-cfg-inv-selected').textContent = name + ' (#' + orgId + ')';
    document.getElementById('pb-cfg-inv-results').style.display = 'none';
    document.getElementById('pb-cfg-inv-search').value = '';
}

function pbSearchTags(term) {
    clearTimeout(pbSearchTimer);
    if (term.length < 2) { document.getElementById('pb-cfg-tag-results').style.display = 'none'; return; }
    pbSearchTimer = setTimeout(function() {
        pbAjax({action: 'search_tags', search_term: term}, function(d) {
            if (!d.success) return;
            var el = document.getElementById('pb-cfg-tag-results');
            var html = '';
            for (var i = 0; i < (d.tags || []).length; i++) {
                var t = d.tags[i];
                html += '<div class="pb-search-item" data-name="' + t.name.replace(/"/g, '&quot;') + '" onclick="pbSelectTag(this.dataset.name)">';
                html += t.name + ' <span class="pb-text-muted">(' + t.count + ' people)</span>';
                html += '</div>';
            }
            el.innerHTML = html || '<div class="pb-search-item pb-text-muted">No results</div>';
            el.style.display = '';
        });
    }, 300);
}

function pbSelectTag(name) {
    document.getElementById('pb-cfg-tag-name').value = name;
    document.getElementById('pb-cfg-tag-selected').textContent = name;
    document.getElementById('pb-cfg-tag-results').style.display = 'none';
    document.getElementById('pb-cfg-tag-search').value = '';
}

// ============================================================
// WORKSPACE - Launch config and load prospects
// ============================================================
function pbLaunchConfig(idx) {
    pbCurrentConfig = pbConfigs[idx];
    pbProcessedMap = {};
    pbDeferredSet = {};
    pbSkippedSet = {};
    pbCrossFlags = {};
    pbContactMap = {};
    pbFamilyData = {};
    pbExtraValues = {};
    pbRegData = {};
    pbInvData = {};
    pbDetailCache = {};
    pbExpandedPid = null;
    pbCurrentIndex = 0;
    pbBatchSelected = {};

    document.getElementById('pb-active-config-name').textContent = pbCurrentConfig.name;
    document.getElementById('pb-workspace-empty').style.display = 'none';
    document.getElementById('pb-workspace-active').style.display = '';

    // Populate batch action dropdown
    pbPopulateBatchActions();
    // Populate flag filter dropdown
    pbPopulateFlagFilter();

    // Show workspace, hide config overview (within same tab)
    document.getElementById('pb-configs-overview').style.display = 'none';
    document.getElementById('pb-configs-workspace').style.display = '';
    document.getElementById('pb-workspace-active').style.display = '';
    // Make sure we're on the configs tab
    pbSwitchTab('configs');
    pbDestHealthLoaded = false;
    pbDestHealth = {};
    // Load all prospect groups (for action menu integration)
    pbGrpLoadAll();
    // Load saved work state, then load prospects
    pbLoadWorkState(pbCurrentConfig.id, function() {
        pbLoadProspectsData();
    });
}

function pbBackToConfigs() {
    document.getElementById('pb-configs-workspace').style.display = 'none';
    document.getElementById('pb-configs-overview').style.display = '';
    pbLoadConfigs();
}

function pbPopulateBatchActions() {
    var sel = document.getElementById('pb-batch-action');
    sel.innerHTML = '<option value="">-- Choose Action --</option>';
    sel.innerHTML += '<option value="__processed">Mark as Processed</option>';
    sel.innerHTML += '<option value="__deferred">Mark as Deferred</option>';
    sel.innerHTML += '<option value="__skipped">Mark as Skipped</option>';
    var actions = pbCurrentConfig.targetActions || [];
    for (var i = 0; i < actions.length; i++) {
        sel.innerHTML += '<option value="action_' + i + '">' + actions[i].label + '</option>';
    }
    // Add prospect group options
    for (var gi = 0; gi < pbGrpGroups.length; gi++) {
        sel.innerHTML += '<option value="grp_' + pbGrpGroups[gi].id + '">Group: ' + pbGrpGroups[gi].name + '</option>';
    }
}

function pbPopulateFlagFilter() {
    var sel = document.getElementById('pb-filter-flag');
    sel.innerHTML = '<option value="all">All Flags</option>';
    var flags = pbCurrentConfig.crossFlags || [];
    for (var i = 0; i < flags.length; i++) {
        sel.innerHTML += '<option value="' + flags[i].id + '_on">' + flags[i].label + ': Yes</option>';
        sel.innerHTML += '<option value="' + flags[i].id + '_off">' + flags[i].label + ': No</option>';
    }
}

var pbDestHealth = {};
var pbDestHealthLoaded = false;

function pbGetDestProspectTypes() {
    // Collect the member type IDs from target actions - these are the types
    // people are being ASSIGNED AS in destinations (e.g., Prospect 311)
    var actions = pbCurrentConfig.targetActions || [];
    var typeSet = {};
    for (var i = 0; i < actions.length; i++) {
        if (actions[i].pb_type === 'involvement' && actions[i].memberTypeId) {
            typeSet[actions[i].memberTypeId] = true;
        }
    }
    var ids = Object.keys(typeSet);
    return ids.length ? ids.join(',') : '';
}

function pbLoadDestinationHealth() {
    if (pbDestHealthLoaded) return;
    var actions = pbCurrentConfig.targetActions || [];
    var orgIds = [];
    for (var i = 0; i < actions.length; i++) {
        if (actions[i].pb_type === 'involvement' && actions[i].orgId) {
            orgIds.push(actions[i].orgId);
        }
    }
    if (!orgIds.length) return;
    var ncd = pbCurrentConfig.noContactDays || 90;
    var prospectTypes = pbGetDestProspectTypes();
    pbAjax({action: 'get_destination_health', org_ids: orgIds.join(','), no_contact_days: ncd, prospect_types: prospectTypes}, function(d) {
        if (d.success) {
            pbDestHealth = d.health || {};
            pbDestHealthLoaded = true;
            // Update health hints in any open menu without re-rendering
            pbUpdateHealthHints();
        }
    });
}

function pbRenderDestHealth() {
    var actions = pbCurrentConfig.targetActions || [];
    var invActions = actions.filter(function(a) { return a.pb_type === 'involvement'; });
    if (!invActions.length) return;

    // Summary line
    var totalProspects = 0, totalStale = 0, totalMembers = 0;
    for (var i = 0; i < invActions.length; i++) {
        var h = pbDestHealth[String(invActions[i].orgId)] || {};
        totalProspects += h.prospectCount || 0;
        totalStale += h.staleProspects || 0;
        totalMembers += h.totalMembers || 0;
    }
    document.getElementById('pb-health-summary').innerHTML =
        invActions.length + ' destination' + (invActions.length > 1 ? 's' : '') +
        ' &middot; ' + totalMembers + ' total members' +
        ' &middot; ' + totalProspects + ' prospects' +
        (totalStale > 0 ? ' &middot; <span style="color:var(--pb-danger);">' + totalStale + ' stale</span>' : '');

    // Detail cards
    var html = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:8px;margin-top:6px;">';
    for (var i = 0; i < invActions.length; i++) {
        var a = invActions[i];
        var h = pbDestHealth[String(a.orgId)] || {};
        var staleRatio = (h.prospectCount > 0) ? Math.round(h.staleProspects / h.prospectCount * 100) : 0;
        var healthColor = staleRatio > 60 ? 'var(--pb-danger)' : staleRatio > 30 ? 'var(--pb-warning)' : 'var(--pb-success)';
        var ageLabel = '';
        if (h.orgAgeDays) {
            if (h.orgAgeDays < 90) ageLabel = 'New group (' + h.orgAgeDays + 'd)';
            else if (h.orgAgeDays < 365) ageLabel = Math.round(h.orgAgeDays / 30) + ' months old';
            else ageLabel = Math.round(h.orgAgeDays / 365 * 10) / 10 + ' years';
        }
        html += '<div class="pb-card" style="padding:10px;border-left:4px solid ' + healthColor + ';">';
        html += '<div class="pb-flex-between">';
        html += '<div class="pb-bold" style="font-size:0.95em;">' + (h.name || a.label) + '</div>';
        var statusLabel = h.totalMembers === 0 ? 'Empty' : staleRatio <= 30 ? 'Healthy' : staleRatio <= 60 ? 'Moderate' : 'Overloaded';
        if (h.totalMembers === 0) healthColor = 'var(--pb-muted)';
        html += '<span class="pb-badge" style="background:' + healthColor + ';color:#fff;">' + statusLabel + '</span>';
        html += '</div>';
        // Leader
        if (h.leader) html += '<div class="pb-text-muted pb-text-sm">Leader: ' + h.leader + (h.leaderEmail ? ' (' + h.leaderEmail + ')' : '') + '</div>';
        // Stats grid
        var kpiStyle = 'padding:4px;border-radius:4px;cursor:pointer;transition:transform 0.1s;';
        var oid = a.orgId;
        html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px;margin-top:6px;text-align:center;font-size:0.8em;">';
        html += '<div style="' + kpiStyle + 'background:var(--pb-light);" onclick="pbDrillDown(' + oid + ',&quot;total&quot;,&quot;All Members&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;">' + (h.totalMembers || 0) + '</div><div class="pb-text-muted">Total</div></div>';
        html += '<div style="' + kpiStyle + 'background:rgba(52,152,219,0.1);" onclick="pbDrillDown(' + oid + ',&quot;prospects&quot;,&quot;Prospects&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;color:var(--pb-accent);">' + (h.prospectCount || 0) + '</div><div class="pb-text-muted">Prospects</div></div>';
        html += '<div style="' + kpiStyle + 'background:' + (h.staleProspects > 0 ? 'rgba(231,76,60,0.1)' : 'var(--pb-light)') + ';" onclick="pbDrillDown(' + oid + ',&quot;stale&quot;,&quot;Stale Prospects&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;' + (h.staleProspects > 0 ? 'color:var(--pb-danger);' : '') + '">' + (h.staleProspects || 0) + '</div><div class="pb-text-muted">Stale</div></div>';
        html += '</div>';
        // Row 2: Graduated, Dropped, Meetings, Avg Attend
        html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:4px;margin-top:4px;text-align:center;font-size:0.8em;">';
        html += '<div style="' + kpiStyle + 'background:rgba(39,174,96,0.1);" onclick="pbDrillDown(' + oid + ',&quot;graduated&quot;,&quot;Graduated (Prospect &rarr; Member)&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;color:var(--pb-success);">' + (h.graduated || 0) + '</div><div class="pb-text-muted">Graduated</div></div>';
        html += '<div style="' + kpiStyle + 'background:' + (h.dropped > 0 ? 'rgba(231,76,60,0.1)' : 'var(--pb-light)') + ';" onclick="pbDrillDown(' + oid + ',&quot;dropped&quot;,&quot;Dropped Prospects&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;' + (h.dropped > 0 ? 'color:var(--pb-danger);' : '') + '">' + (h.dropped || 0) + '</div><div class="pb-text-muted">Dropped</div></div>';
        html += '<div style="padding:4px;background:var(--pb-light);border-radius:4px;"><div class="pb-bold" style="font-size:1.3em;">' + (h.meetings90d || 0) + '</div><div class="pb-text-muted">Mtgs (90d)</div></div>';
        html += '<div style="padding:4px;background:var(--pb-light);border-radius:4px;"><div class="pb-bold" style="font-size:1.3em;">' + (h.avgAttendance || 0) + '</div><div class="pb-text-muted">Avg Attend</div></div>';
        html += '</div>';
        if (h.lastMeeting) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:2px;">Last meeting: ' + h.lastMeeting + '</div>';
        if (h.headCountWarning) html += '<div style="margin-top:3px;padding:3px 6px;background:rgba(243,156,18,0.15);border-radius:4px;font-size:0.75em;color:var(--pb-warning);font-weight:600;">&#9888; ' + h.headCountMeetings + ' of ' + h.meetings90d + ' meetings used headcount only (no named attendance)</div>';
        // Member type breakdown
        if (h.typeCounts) {
            html += '<div style="margin-top:4px;display:flex;flex-wrap:wrap;gap:4px;">';
            for (var mt in h.typeCounts) {
                html += '<span class="pb-chip">' + mt + ': ' + h.typeCounts[mt] + '</span>';
            }
            html += '</div>';
        }
        if (ageLabel) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:4px;">' + ageLabel + (h.orgCreated ? ' (created ' + h.orgCreated + ')' : '') + '</div>';
        html += '</div>';
    }
    html += '</div>';
    document.getElementById('pb-health-detail').innerHTML = html;
}

function pbToggleHealthPanel() {
    var detail = document.getElementById('pb-health-detail');
    var arrow = document.getElementById('pb-health-arrow');
    if (detail.style.display === 'none') {
        detail.style.display = '';
        arrow.innerHTML = '&#9660;';
    } else {
        detail.style.display = 'none';
        arrow.innerHTML = '&#9654;';
    }
}

function pbShowDestDetail(orgId) {
    // Close the action menu
    if (pbOpenMenuPid) {
        var menu = document.getElementById('pb-amenu-' + pbOpenMenuPid);
        if (menu) menu.classList.remove('open');
        pbOpenMenuPid = null;
    }
    // Show modal with destination health detail
    var overlay = document.createElement('div');
    overlay.className = 'pb-modal-overlay active';
    overlay.innerHTML = '<div class="pb-modal" style="max-width:600px;">'
        + '<div class="pb-modal-header"><h3>Destination Health</h3><button class="pb-modal-close" onclick="this.closest(&quot;.pb-modal-overlay&quot;).remove()">&times;</button></div>'
        + '<div class="pb-modal-body" id="pb-dest-detail-body"><div class="pb-loading"><span class="pb-spin"></span> Loading...</div></div></div>';
    document.body.appendChild(overlay);

    // Load health if not loaded yet
    var ncd = pbCurrentConfig.noContactDays || 90;
    var prospectTypes = pbGetDestProspectTypes();
    pbAjax({action: 'get_destination_health', org_ids: String(orgId), no_contact_days: ncd, prospect_types: prospectTypes}, function(d) {
        var body = document.getElementById('pb-dest-detail-body');
        if (!body) return;
        if (!d.success) { body.innerHTML = '<div class="pb-text-muted">Failed to load.</div>'; return; }
        var h = (d.health || {})[String(orgId)] || {};
        if (!h.name) { body.innerHTML = '<div class="pb-text-muted">No data found.</div>'; return; }

        var staleRatio = (h.prospectCount > 0) ? Math.round(h.staleProspects / h.prospectCount * 100) : 0;
        var healthColor = h.totalMembers === 0 ? 'var(--pb-muted)' : staleRatio > 60 ? 'var(--pb-danger)' : staleRatio > 30 ? 'var(--pb-warning)' : 'var(--pb-success)';
        var statusLabel = h.totalMembers === 0 ? 'Empty' : staleRatio <= 30 ? 'Healthy' : staleRatio <= 60 ? 'Moderate' : 'Overloaded';
        var kpiStyle = 'padding:6px;border-radius:4px;cursor:pointer;transition:transform 0.1s;text-align:center;';
        var oid = orgId;

        var html = '<div style="border-left:4px solid ' + healthColor + ';padding-left:12px;">';
        html += '<div class="pb-flex-between"><div class="pb-bold" style="font-size:1.1em;">' + h.name + '</div>';
        html += '<span class="pb-badge" style="background:' + healthColor + ';color:#fff;">' + statusLabel + '</span></div>';
        if (h.leader) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:2px;">Leader: ' + h.leader + (h.leaderEmail ? ' (' + h.leaderEmail + ')' : '') + '</div>';

        // KPI grid - clickable
        html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-top:10px;font-size:0.85em;">';
        html += '<div style="' + kpiStyle + 'background:var(--pb-light);" onclick="pbDrillDown(' + oid + ',&quot;total&quot;,&quot;All Members&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.5em;">' + (h.totalMembers || 0) + '</div><div class="pb-text-muted">Total</div></div>';
        html += '<div style="' + kpiStyle + 'background:rgba(52,152,219,0.1);" onclick="pbDrillDown(' + oid + ',&quot;prospects&quot;,&quot;Prospects&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.5em;color:var(--pb-accent);">' + (h.prospectCount || 0) + '</div><div class="pb-text-muted">Prospects</div></div>';
        html += '<div style="' + kpiStyle + 'background:' + (h.staleProspects > 0 ? 'rgba(231,76,60,0.1)' : 'var(--pb-light)') + ';" onclick="pbDrillDown(' + oid + ',&quot;stale&quot;,&quot;Stale Prospects&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.5em;' + (h.staleProspects > 0 ? 'color:var(--pb-danger);' : '') + '">' + (h.staleProspects || 0) + '</div><div class="pb-text-muted">Stale</div></div>';
        html += '</div>';
        // Row 2
        html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;margin-top:6px;font-size:0.85em;text-align:center;">';
        html += '<div style="' + kpiStyle + 'background:rgba(39,174,96,0.1);" onclick="pbDrillDown(' + oid + ',&quot;graduated&quot;,&quot;Graduated&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;color:var(--pb-success);">' + (h.graduated || 0) + '</div><div class="pb-text-muted">Graduated</div></div>';
        html += '<div style="' + kpiStyle + 'background:' + (h.dropped > 0 ? 'rgba(231,76,60,0.1)' : 'var(--pb-light)') + ';" onclick="pbDrillDown(' + oid + ',&quot;dropped&quot;,&quot;Dropped&quot;)" onmouseover="this.style.transform=&quot;scale(1.05)&quot;" onmouseout="this.style.transform=&quot;scale(1)&quot;"><div class="pb-bold" style="font-size:1.3em;' + (h.dropped > 0 ? 'color:var(--pb-danger);' : '') + '">' + (h.dropped || 0) + '</div><div class="pb-text-muted">Dropped</div></div>';
        html += '<div style="padding:6px;background:var(--pb-light);border-radius:4px;"><div class="pb-bold" style="font-size:1.3em;">' + (h.meetings90d || 0) + '</div><div class="pb-text-muted">Mtgs (90d)</div></div>';
        html += '<div style="padding:6px;background:var(--pb-light);border-radius:4px;"><div class="pb-bold" style="font-size:1.3em;">' + (h.avgAttendance || 0) + '</div><div class="pb-text-muted">Avg Attend</div></div>';
        html += '</div>';
        if (h.lastMeeting) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:4px;">Last meeting: ' + h.lastMeeting + '</div>';
        if (h.headCountWarning) html += '<div style="margin-top:4px;padding:3px 6px;background:rgba(243,156,18,0.15);border-radius:4px;font-size:0.75em;color:var(--pb-warning);font-weight:600;">&#9888; ' + h.headCountMeetings + ' of ' + h.meetings90d + ' meetings used headcount only</div>';
        // Member type breakdown
        if (h.typeCounts) {
            html += '<div style="margin-top:6px;display:flex;flex-wrap:wrap;gap:4px;">';
            for (var mt in h.typeCounts) {
                html += '<span class="pb-chip">' + mt + ': ' + h.typeCounts[mt] + '</span>';
            }
            html += '</div>';
        }
        var ageLabel = '';
        if (h.orgAgeDays) {
            if (h.orgAgeDays < 90) ageLabel = 'New group (' + h.orgAgeDays + 'd)';
            else if (h.orgAgeDays < 365) ageLabel = Math.round(h.orgAgeDays / 30) + ' months old';
            else ageLabel = Math.round(h.orgAgeDays / 365 * 10) / 10 + ' years';
        }
        if (ageLabel) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:4px;">' + ageLabel + (h.orgCreated ? ' (created ' + h.orgCreated + ')' : '') + '</div>';
        html += '</div>';
        body.innerHTML = html;
    });
}

function pbDrillDown(orgId, drillType, title) {
    // Show a modal with the people list
    var overlay = document.createElement('div');
    overlay.className = 'pb-modal-overlay active';
    overlay.innerHTML = '<div class="pb-modal" style="max-width:700px;">'
        + '<div class="pb-modal-header"><h3>' + title + '</h3><button class="pb-modal-close" onclick="this.closest(&quot;.pb-modal-overlay&quot;).remove()">&times;</button></div>'
        + '<div class="pb-modal-body" id="pb-drill-body"><div class="pb-loading"><span class="pb-spin"></span> Loading...</div></div></div>';
    document.body.appendChild(overlay);

    var prospectTypes = pbGetDestProspectTypes();
    var ncd = pbCurrentConfig.noContactDays || 90;
    pbAjax({action: 'get_destination_people', org_id: orgId, drill_type: drillType, prospect_types: prospectTypes, no_contact_days: ncd}, function(d) {
        var body = document.getElementById('pb-drill-body');
        if (!body) return;
        if (!d.success || !d.people || !d.people.length) {
            body.innerHTML = '<div class="pb-empty pb-text-muted">No people found.</div>';
            return;
        }
        // Summary stats
        var totalEnrolled = 0, minDays = 9999, maxDays = 0, withAttend = 0;
        for (var s = 0; s < d.people.length; s++) {
            var sp = d.people[s];
            if (sp.daysEnrolled != null) { totalEnrolled++; if (sp.daysEnrolled < minDays) minDays = sp.daysEnrolled; if (sp.daysEnrolled > maxDays) maxDays = sp.daysEnrolled; }
            if (sp.attendCount > 0) withAttend++;
        }
        var html = '<div style="padding:6px 10px;background:var(--pb-light);border-radius:var(--pb-radius);margin-bottom:8px;font-size:0.82em;display:flex;gap:12px;flex-wrap:wrap;">';
        html += '<span><strong>' + d.people.length + '</strong> people</span>';
        if (totalEnrolled > 0) html += '<span>Enrolled: <strong>' + minDays + 'd</strong> to <strong>' + maxDays + 'd</strong></span>';
        html += '<span>With attendance: <strong>' + withAttend + '</strong> (' + Math.round(withAttend/d.people.length*100) + '%)</span>';
        html += '</div>';
        html += '<div style="font-size:0.85em;max-height:55vh;overflow-y:auto;">';
        for (var i = 0; i < d.people.length; i++) {
            var p = d.people[i];
            html += '<div style="display:flex;gap:8px;align-items:flex-start;padding:6px 0;border-bottom:1px solid #f0f0f0;">';
            if (p.photoUrl) {
                html += '<img src="' + p.photoUrl + '" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0;" onerror="this.style.display=&quot;none&quot;">';
            } else {
                html += '<div style="width:36px;height:36px;border-radius:50%;background:var(--pb-light);display:flex;align-items:center;justify-content:center;font-size:14px;color:var(--pb-muted);flex-shrink:0;">&#128100;</div>';
            }
            html += '<div style="flex:1;min-width:0;">';
            // Row 1: Name + member type + days enrolled
            html += '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">';
            html += '<a href="/Person2/' + p.pid + '" target="_blank" style="color:var(--pb-primary);font-weight:600;">' + p.name + '</a>';
            html += '<span class="pb-badge pb-badge-muted" style="font-size:0.7em;">' + (p.memberType || '') + '</span>';
            if (p.daysEnrolled != null) {
                var enrollColor = p.daysEnrolled > 180 ? 'var(--pb-danger)' : p.daysEnrolled > 90 ? 'var(--pb-warning)' : 'var(--pb-accent)';
                html += '<span class="pb-badge" style="font-size:0.65em;background:rgba(52,152,219,0.15);color:' + enrollColor + ';">added ' + p.daysEnrolled + ' days ago' + (p.enrollDate ? ' (' + p.enrollDate + ')' : '') + '</span>';
            }
            if (p.age) html += '<span class="pb-text-muted" style="font-size:0.8em;">Age ' + p.age + '</span>';
            html += '</div>';
            // Row 2: Contact + attendance
            html += '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:1px;font-size:0.82em;">';
            if (p.phone) html += '<span class="pb-text-muted">&#9742; ' + pbFmtPhone(p.phone) + '</span>';
            if (p.email) html += '<span class="pb-text-muted">&#9993; ' + p.email + '</span>';
            // Attendance info
            if (p.attendCount > 0) {
                var attendColor = p.daysSince != null && p.daysSince > (pbCurrentConfig.noContactDays || 90) ? 'var(--pb-danger)' : 'var(--pb-success)';
                html += '<span style="color:' + attendColor + ';font-weight:600;">' + p.attendCount + ' attended</span>';
                if (p.lastAttend) html += '<span class="pb-text-muted">last ' + p.lastAttend + (p.daysSince != null ? ' (' + p.daysSince + 'd)' : '') + '</span>';
            } else {
                html += '<span style="color:var(--pb-danger);font-weight:600;">0 attended</span>';
            }
            if (p.droppedDate) html += '<span style="color:var(--pb-danger);font-weight:600;">dropped ' + p.droppedDate + '</span>';
            html += '</div>';
            html += '</div></div>';
        }
        html += '</div>';
        if (d.people.length >= 200) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:6px;">Showing first 200</div>';
        body.innerHTML = html;
    });
}

function pbUpdateHealthHints() {
    // Health data loaded - update hint text in any existing menus
    var actions = pbCurrentConfig ? (pbCurrentConfig.targetActions || []) : [];
    var invActions = actions.filter(function(a) { return a.pb_type === 'involvement'; });
    // Update all .pb-ami-health spans that exist in the DOM
    document.querySelectorAll('.pb-ami-health').forEach(function(el) {
        var title = el.getAttribute('title');
        if (title !== 'prospects/stale/total') return;
        // Find which org this belongs to by checking sibling onclick
        var item = el.closest('.pb-action-menu-item');
        if (!item) return;
        var plusIcon = item.querySelector('.pb-ami-icon');
        if (!plusIcon) return;
        var oc = plusIcon.getAttribute('onclick') || '';
        for (var i = 0; i < invActions.length; i++) {
            if (oc.indexOf(',' + actions.indexOf(invActions[i]) + ')') >= 0) {
                var hd = pbDestHealth[String(invActions[i].orgId)];
                if (hd) {
                    var txt = hd.prospectCount + 'p';
                    if (hd.staleProspects > 0) txt += ' / ' + hd.staleProspects + 's';
                    txt += ' / ' + hd.totalMembers + 't';
                    el.textContent = txt;
                }
                break;
            }
        }
    });
}

function pbRefreshProspects() { pbLoadProspectsData(); }

function pbLoadProspectsData() {
    document.getElementById('pb-list-content').innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading prospects...</div>';

    // Phase 1: Load core prospect data
    pbAjax({action: 'load_prospects', config_data: JSON.stringify(pbCurrentConfig)}, function(d) {
        if (!d.success) {
            document.getElementById('pb-list-content').innerHTML = '<div class="pb-empty">Error: ' + (d.message || 'Failed to load') + '</div>';
            return;
        }
        pbProspects = d.prospects || [];

        // Note: work state PIDs are not cleaned here to avoid race conditions
        // Stale PIDs in processedMap/deferredSet/skippedSet are harmless -
        // they simply won't match any prospect in the current list

        document.getElementById('pb-prospect-count').textContent = pbProspects.length + ' prospects';
        pbApplyFilters();

        if (!pbProspects.length) return;
        var pids = pbProspects.map(function(p) { return p.PeopleId; }).join(',');

        // Load lightweight bulk data only: contact efforts + cross-query flags
        // Detail data (family, involvements, EVs, reg) loaded on-demand per person

        // Contact efforts (small query, needed for badges in list)
        pbAjax({action: 'load_contact_efforts', people_ids: pids}, function(d4) {
            if (d4.success) { pbContactMap = d4.contactMap || {}; pbApplyFilters(); }
            // Compute prospect scores after contact data is available
            if (pbSettings.scorecard && pbSettings.scorecard.enabled) {
                pbComputeScores(pids);
            }
        });

        // Cross-query flags (needed for filtering)
        var hasFlags = (pbCurrentConfig.crossFlags || []).length > 0;
        if (hasFlags) {
            pbAjax({action: 'evaluate_cross_flags', people_ids: pids, flags_data: JSON.stringify(pbCurrentConfig.crossFlags)}, function(d5) {
                if (d5.success) { pbCrossFlags = d5.flagResults || {}; pbPopulateFlagFilter(); pbApplyFilters(); }
            });
        }
    });
}

// ============================================================
// FILTERING & SORTING
// ============================================================
function pbApplyFilters() {
    var search = (document.getElementById('pb-filter-search').value || '').toLowerCase();
    var status = document.getElementById('pb-filter-status').value;
    var flagFilter = document.getElementById('pb-filter-flag').value;
    var sort = document.getElementById('pb-filter-sort').value;

    pbFilteredProspects = pbProspects.filter(function(p) {
        var pid = p.PeopleId;
        // Search
        if (search && (p.Name2 || '').toLowerCase().indexOf(search) < 0 && (p.EmailAddress || '').toLowerCase().indexOf(search) < 0) return false;
        // Status
        if (status === 'pending' && (pbProcessedMap[pid] || pbDeferredSet[pid] || pbSkippedSet[pid])) return false;
        if (status === 'processed' && !pbProcessedMap[pid]) return false;
        if (status === 'deferred' && !pbDeferredSet[pid]) return false;
        if (status === 'skipped' && !pbSkippedSet[pid]) return false;
        // Flag filter
        if (flagFilter !== 'all') {
            var parts = flagFilter.split('_on');
            var wantOn = flagFilter.endsWith('_on');
            var flagId = wantOn ? flagFilter.replace('_on', '') : flagFilter.replace('_off', '');
            var flagPids = pbCrossFlags[flagId] || [];
            var hasFlag = flagPids.indexOf(pid) >= 0;
            if (wantOn && !hasFlag) return false;
            if (!wantOn && hasFlag) return false;
        }
        return true;
    });

    // Sort
    pbFilteredProspects.sort(function(a, b) {
        if (sort === 'name') return (a.Name2 || '').localeCompare(b.Name2 || '');
        if (sort === 'age') return (a.Age || 999) - (b.Age || 999);
        if (sort === 'enrolled') return (a.EnrollmentDate || '').localeCompare(b.EnrollmentDate || '');
        if (sort === 'contacts') {
            var cmA = pbContactMap[a.PeopleId];
            var cmB = pbContactMap[b.PeopleId];
            var ca = cmA ? (cmA.weightedTotal != null ? cmA.weightedTotal : cmA.total) : 0;
            var cb = cmB ? (cmB.weightedTotal != null ? cmB.weightedTotal : cmB.total) : 0;
            return cb - ca;
        }
        if (sort === 'score') {
            var sa = pbScoreMap[String(a.PeopleId)] ? pbScoreMap[String(a.PeopleId)].score : 0;
            var sb = pbScoreMap[String(b.PeopleId)] ? pbScoreMap[String(b.PeopleId)].score : 0;
            return sb - sa;
        }
        return 0;
    });

    pbUpdateProgress();
    pbRenderCurrentView();
}

function pbUpdateProgress() {
    var total = pbProspects.length;
    // Count only PIDs that exist in current prospect list
    var currentPids = {};
    for (var i = 0; i < pbProspects.length; i++) currentPids[String(pbProspects[i].PeopleId)] = true;

    var processed = 0, deferred = 0, skipped = 0;
    for (var pk in pbProcessedMap) { if (currentPids[String(pk)]) processed++; }
    for (var dk in pbDeferredSet) { if (currentPids[String(dk)]) deferred++; }
    for (var sk in pbSkippedSet) { if (currentPids[String(sk)]) skipped++; }

    var worked = processed + deferred + skipped;
    var pending = total - worked;
    var pct = total > 0 ? Math.round(worked / total * 100) : 0;

    // Populate the 4 workspace stat cards.
    function setStat(id, val) {
        var el = document.getElementById(id);
        if (el) el.textContent = (val || 0).toLocaleString();
    }
    setStat('pb-ws-stat-pending', pending);
    setStat('pb-ws-stat-processed', processed);
    setStat('pb-ws-stat-deferred', deferred);
    setStat('pb-ws-stat-skipped', skipped);

    var ptxt = document.getElementById('pb-progress-text');
    if (ptxt) ptxt.textContent = (total || 0).toLocaleString() + ' total';
    var ppct = document.getElementById('pb-progress-pct');
    if (ppct) ppct.textContent = pct + '% worked';
    var pbar = document.getElementById('pb-progress-bar');
    if (pbar) pbar.style.width = pct + '%';
}

// ============================================================
// VIEW RENDERING
// ============================================================
function pbSetView(mode) {
    pbViewMode = mode;
    document.querySelectorAll('.pb-view-btn').forEach(function(b) {
        b.classList.toggle('active', b.dataset.view === mode);
    });
    document.getElementById('pb-view-list').style.display = mode === 'list' ? '' : 'none';
    document.getElementById('pb-view-single').style.display = mode === 'single' ? '' : 'none';
    document.getElementById('pb-view-batch').style.display = mode === 'batch' ? '' : 'none';
    pbRenderCurrentView();
}

function pbRenderCurrentView() {
    if (pbViewMode === 'list') pbRenderListView();
    else if (pbViewMode === 'single') pbRenderSingleView();
    else if (pbViewMode === 'batch') pbRenderBatchView();
}

function pbGetFieldValue(prospect, field, asHtml) {
    var ft = field.fieldType;
    var src = field.sourceField;
    var pid = prospect.PeopleId;

    if (ft === 'person') {
        var pval = prospect[src] || '';
        if ((src === 'CellPhone' || src === 'HomePhone') && pval) return pbFmtPhone(pval);
        if (src === 'MaritalStatus' && pval === 'Widowed' && prospect.SpouseDeceasedDate) {
            return pval + ' <span class="pb-text-muted" style="font-size:0.9em;">(spouse deceased ' + prospect.SpouseDeceasedDate + ')</span>';
        }
        return pval;
    }
    if (ft === 'family') {
        var fam = pbFamilyData[String(pid)] || {};
        // Rich HTML rendering for FamilyDetail and FamilyDetailList
        if (src === 'FamilyDetail' && asHtml && fam.FamilyDetailList) {
            return pbRenderFamilyDetail(fam.FamilyDetailList);
        }
        if (src === 'FamilyDetail') {
            return (fam.FamilyDetail || '').replace(/\\n/g, '; ');
        }
        if (src === 'SpouseDetail') return fam.SpouseDetail || '';
        return fam[src] || '';
    }
    if (ft === 'involvement') {
        var invs = pbInvData[String(pid)] || [];
        if (!invs.length) return '<span class="pb-text-muted">No involvements</span>';
        if (asHtml) return pbRenderInvolvementDetail(invs);
        return invs.map(function(iv) { return iv.name + ' (' + iv.memberType + ')'; }).join(', ');
    }
    if (ft === 'extravalue') {
        var evs = pbExtraValues[String(pid)] || {};
        return evs[src] || '';
    }
    if (ft === 'regquestion') {
        var rd = pbRegData[String(pid)] || {};
        return rd[src] || '';
    }
    if (ft === 'medical') return prospect[src] || '';
    return '';
}

function pbRenderFamilyDetail(detailList) {
    if (!detailList || !detailList.length) return '<span class="pb-text-muted">No family members</span>';
    var noContactDays = (pbCurrentConfig && pbCurrentConfig.noContactDays) ? pbCurrentConfig.noContactDays : 90;
    var html = '<div style="font-size:0.85em;">';
    for (var i = 0; i < detailList.length; i++) {
        var m = detailList[i];
        var statusColor = m.statusId === 10 ? 'var(--pb-success)' : m.statusId === 20 ? 'var(--pb-warning)' : 'var(--pb-muted)';
        var posLabel = m.posLabel || 'Other';
        var indent = '46px';
        html += '<div style="padding:5px 0;border-bottom:1px solid #f0f0f0;display:flex;gap:8px;align-items:flex-start;">';
        // Photo
        if (m.photoUrl) {
            var famBigPhoto = m.bigPhotoUrl || m.photoUrl;
            html += '<img src="' + m.photoUrl + '" data-large="' + famBigPhoto + '" onclick="event.stopPropagation();pbShowPhotoModal(this.dataset.large || this.src)" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0;cursor:pointer;" onerror="this.style.display=&quot;none&quot;" title="Click to enlarge">';
        } else {
            html += '<div style="width:36px;height:36px;border-radius:50%;background:var(--pb-light);display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:14px;color:var(--pb-muted);">&#128100;</div>';
        }
        html += '<div style="flex:1;min-width:0;">';
        // Row 1: Name, position, status, age, attendance
        html += '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">';
        html += '<span class="pb-badge pb-badge-muted" style="font-size:0.65em;">' + posLabel + '</span>';
        html += '<a href="/Person2/' + m.pid + '" target="_blank" style="color:var(--pb-primary);font-weight:600;font-size:0.95em;" onclick="event.stopPropagation()">' + m.name + '</a>';
        html += '<span class="pb-badge" style="background:' + statusColor + ';color:#fff;font-size:0.65em;">' + (m.status || '?') + '</span>';
        if (m.age) html += '<span class="pb-text-muted" style="font-size:0.8em;">age ' + m.age + '</span>';
        // Attendance with days-since logic
        if (m.lastAttend) {
            var ds = m.daysSince;
            if (ds !== null && ds !== undefined && ds > noContactDays) {
                html += '<span style="color:var(--pb-danger);font-size:0.8em;font-weight:600;">' + ds + 'd ago (' + m.lastAttend + ')</span>';
            } else {
                html += '<span class="pb-text-muted" style="font-size:0.8em;">attended ' + m.lastAttend + (ds !== null ? ' (' + ds + 'd)' : '') + '</span>';
            }
        } else {
            html += '<span style="color:var(--pb-danger);font-size:0.8em;font-weight:600;">no attendance</span>';
        }
        html += '</div>';
        // Row 2: Contact + involvements inline
        html += '<div style="display:flex;flex-wrap:wrap;gap:4px 10px;margin-top:2px;font-size:0.8em;">';
        if (m.phone) html += '<span class="pb-text-muted">&#9742; ' + pbFmtPhone(m.phone) + '</span>';
        if (m.email) html += '<span class="pb-text-muted">&#9993; ' + m.email + '</span>';
        if (m.invNames && m.invNames.length) {
            for (var j = 0; j < m.invNames.length; j++) {
                html += '<span class="pb-chip">' + m.invNames[j] + '</span>';
            }
        } else {
            html += '<span class="pb-text-muted" style="font-style:italic;">No involvements</span>';
        }
        html += '</div>';
        html += '</div></div>';
    }
    html += '</div>';
    return html;
}

function pbRenderInvolvementDetail(invs) {
    if (!invs || !invs.length) return '<span class="pb-text-muted">No involvements</span>';
    var html = '<div style="font-size:0.85em;">';
    // Group by program
    var byProg = {};
    for (var i = 0; i < invs.length; i++) {
        var prog = invs[i].program || 'Other';
        if (!byProg[prog]) byProg[prog] = [];
        byProg[prog].push(invs[i]);
    }
    for (var prog in byProg) {
        html += '<div style="margin-bottom:6px;"><div class="pb-text-muted pb-bold" style="font-size:0.8em;text-transform:uppercase;margin-bottom:2px;">' + prog + '</div>';
        var items = byProg[prog];
        for (var j = 0; j < items.length; j++) {
            var iv = items[j];
            var pctColor = iv.pct >= 75 ? 'var(--pb-success)' : iv.pct >= 40 ? 'var(--pb-warning)' : 'var(--pb-danger)';
            html += '<div style="padding:3px 0 3px 10px;border-left:3px solid ' + pctColor + ';margin-bottom:2px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">';
            html += '<a href="/Organization/' + iv.orgId + '" target="_blank" style="color:var(--pb-primary);font-weight:600;">' + iv.name + '</a>';
            html += '<span class="pb-badge pb-badge-muted">' + iv.memberType + '</span>';
            // Attendance stats
            if (iv.totalMeetings > 0) {
                html += '<span style="font-size:0.8em;color:' + pctColor + ';font-weight:700;">' + iv.attended + '/' + iv.totalMeetings + ' (' + iv.pct + '%)</span>';
            } else {
                html += '<span class="pb-text-muted" style="font-size:0.8em;">no meetings</span>';
            }
            if (iv.lastAttend) html += '<span class="pb-text-muted" style="font-size:0.8em;">last ' + iv.lastAttend + '</span>';
            if (iv.enrollDate) html += '<span class="pb-text-muted" style="font-size:0.8em;">enrolled ' + iv.enrollDate + '</span>';
            html += '</div>';
        }
        html += '</div>';
    }
    html += '</div>';
    return html;
}

function pbGetStatusClass(pid) {
    if (pbProcessedMap[pid]) return 'processed';
    if (pbDeferredSet[pid]) return 'deferred';
    if (pbSkippedSet[pid]) return 'skipped';
    return '';
}

function pbFmtPhone(ph) {
    if (!ph) return '';
    var d = ph.replace(/[^0-9]/g, '');
    if (d.length === 10) return '(' + d.substr(0,3) + ') ' + d.substr(3,3) + '-' + d.substr(6);
    if (d.length === 11 && d[0] === '1') return '(' + d.substr(1,3) + ') ' + d.substr(4,3) + '-' + d.substr(7);
    return ph;
}

function pbBuildContactBadges(pid) {
    var cm = pbContactMap[String(pid)];
    if (!cm) return '<span class="pb-text-muted pb-text-sm">No contacts</span>';
    var methods = pbSettings.contact_methods || [];
    var html = '<span class="pb-contact-badge">';
    for (var i = 0; i < methods.length; i++) {
        var code = methods[i].code;
        var count = cm.methods[code] || 0;
        html += '<span class="pb-contact-code' + (count > 0 ? ' has-count' : '') + '">' + code + '(' + count + ')</span>';
    }
    var other = cm.methods['O'] || 0;
    if (other > 0) {
        var ow = pbSettings.other_weight != null ? pbSettings.other_weight : 1;
        var dimStyle = ow < 1 ? ' style="opacity:' + Math.max(0.35, ow) + ';"' : '';
        html += '<span class="pb-contact-code has-count"' + dimStyle + '>O(' + other + ')</span>';
    }
    html += '</span>';
    return html;
}

function pbBuildFlagBadges(pid) {
    var flags = pbCurrentConfig.crossFlags || [];
    if (!flags.length) return '';
    var html = '';
    for (var i = 0; i < flags.length; i++) {
        var flagPids = pbCrossFlags[flags[i].id] || [];
        var isOn = flagPids.indexOf(pid) >= 0;
        html += '<span class="pb-flag-badge ' + (isOn ? 'pb-flag-on' : 'pb-flag-off') + '">' + (flags[i].label || flags[i].pb_type) + '</span>';
    }
    return html;
}

// Wraps pbBuildActionMenuSections into a native <details> element. Used
// in Single view where the menu lives inside a tall card and the absolute-
// positioned dropdown variant has positioning hassles. <details> handles
// open/close natively, no JS, no positioning.
function pbBuildActionPanel(pid) {
    var sections = pbBuildActionMenuSections(pid);
    return '<details class="pb-action-panel" style="display:inline-block;">'
         + '<summary style="cursor:pointer;list-style:none;padding:4px 10px;border:1px solid var(--pb-border);border-radius:var(--pb-radius);background:var(--pb-white);color:var(--pb-accent);font-size:0.8em;font-weight:600;display:inline-block;">Action &#9660;</summary>'
         + '<div class="pb-action-panel-body" style="position:absolute;right:0;margin-top:4px;min-width:280px;max-height:400px;overflow-y:auto;background:var(--pb-white);border:1px solid var(--pb-border);border-radius:var(--pb-radius);box-shadow:0 6px 20px rgba(0,0,0,0.15);z-index:600;">'
         + sections
         + '</div></details>';
}

// Extracted from pbBuildActionMenu so both the dropdown (List) and the
// inline <details> (Single) can render the same items.
function pbBuildActionMenuSections(pid) {
    var actions = pbCurrentConfig.targetActions || [];
    var html = '';
    var tags = actions.filter(function(a) { return a.pb_type === 'tag'; });
    if (tags.length) {
        html += '<div class="pb-action-menu-section">';
        html += '<div class="pb-action-menu-label">Tag</div>';
        for (var t = 0; t < tags.length; t++) {
            var ti = actions.indexOf(tags[t]);
            html += '<div class="pb-action-menu-item" onclick="event.stopPropagation();pbProcessSingle(' + pid + ',' + ti + ')">';
            html += '<span class="pb-ami-icon" style="color:var(--pb-accent);">&#127991;</span>';
            html += '<span>' + tags[t].label + '</span></div>';
        }
        html += '</div>';
    }
    var invs = actions.filter(function(a) { return a.pb_type === 'involvement'; });
    if (invs.length) {
        html += '<div class="pb-action-menu-section">';
        html += '<div class="pb-action-menu-label">Assign to Involvement</div>';
        for (var v = 0; v < invs.length; v++) {
            var vi = actions.indexOf(invs[v]);
            var hk = String(invs[v].orgId);
            var hd = pbDestHealth[hk];
            var healthHint = '';
            if (hd) {
                healthHint = hd.prospectCount + 'p';
                if (hd.staleProspects > 0) healthHint += ' / ' + hd.staleProspects + 's';
                healthHint += ' / ' + hd.totalMembers + 't';
            }
            // Show "as <MemberType>" inline so staff can see the role this
            // action will apply before they click. Critical signal because
            // older saved actions cached the wrong memberTypeId on churches
            // that customized lookup.MemberType. The label is the source of
            // truth -- if the cached id disagrees with the label, prefer
            // the label-derived value (matches the server's label-first
            // resolution in process_action).
            var _mtLabel = '';
            var _mtFromId = (invs[v].memberTypeId && pbMemberTypeLabels && pbMemberTypeLabels[String(invs[v].memberTypeId)])
                              ? pbMemberTypeLabels[String(invs[v].memberTypeId)] : '';
            var _mtFromLabel = '';
            if (invs[v].label) {
                var _lbl = invs[v].label;
                var _idx = _lbl.lastIndexOf(' as ');
                if (_idx >= 0) _mtFromLabel = _lbl.substring(_idx + 4).trim();
            }
            // Prefer label-derived (matches what the user actually saw when
            // configuring). Fall back to id-derived if label doesn't have "as".
            _mtLabel = _mtFromLabel || _mtFromId;
            var _mtSpan = _mtLabel
                ? ' <span style="font-size:0.78em;color:#107c10;font-weight:600;">as ' + _mtLabel + '</span>'
                : ' <span style="font-size:0.78em;color:#d13438;font-weight:600;" title="No member type configured -- person will land as default. Edit this action to set a role.">no role set</span>';
            html += '<div class="pb-action-menu-item" style="padding:0;">';
            html += '<span class="pb-ami-icon" style="color:var(--pb-success);cursor:pointer;padding:6px 4px;" onclick="event.stopPropagation();pbProcessSingle(' + pid + ',' + vi + ')" title="Assign to this group">&#10133;</span>';
            html += '<span style="flex:1;cursor:pointer;padding:6px 4px;" onclick="event.stopPropagation();pbShowDestDetail(' + invs[v].orgId + ')">' + (invs[v].orgName || invs[v].label) + _mtSpan + '</span>';
            if (healthHint) html += '<span class="pb-ami-health" style="padding:6px 4px;" title="prospects/stale/total">' + healthHint + '</span>';
            html += '</div>';
        }
        html += '</div>';
    }
    html += '<div class="pb-action-menu-section">';
    html += '<div class="pb-action-menu-item" onclick="event.stopPropagation();pbDeferProspect(' + pid + ')">';
    html += '<span class="pb-ami-icon" style="color:var(--pb-warning);">&#9200;</span><span>Defer</span></div>';
    html += '<div class="pb-action-menu-item" onclick="event.stopPropagation();pbSkipProspect(' + pid + ')">';
    html += '<span class="pb-ami-icon" style="color:var(--pb-muted);">&#10060;</span><span>Skip / Not a Fit</span></div>';
    html += '<div class="pb-action-menu-item" onclick="event.stopPropagation();pbShowContactModal(' + pid + ')">';
    html += '<span class="pb-ami-icon">&#128221;</span><span>Log Contact</span></div>';
    html += '</div>';
    if (pbGrpGroups.length > 0) {
        html += '<div class="pb-action-menu-section">';
        html += '<div class="pb-action-menu-label">Group Management</div>';
        for (var gi = 0; gi < pbGrpGroups.length; gi++) {
            var grp = pbGrpGroups[gi];
            var grpColor = grp.color || 'var(--pb-accent)';
            html += '<div class="pb-action-menu-item" data-gid="' + grp.id + '" data-pid="' + pid + '" onclick="event.stopPropagation();pbGrpQuickAssign(this.dataset.gid,parseInt(this.dataset.pid))">';
            html += '<span class="pb-ami-icon" style="color:' + grpColor + ';">&#128101;</span>';
            html += '<span>' + grp.name + '</span></div>';
        }
        html += '</div>';
    }
    return html;
}

function pbBuildActionMenu(pid, openUpward) {
    var arrow = openUpward ? '&#9650;' : '&#9660;';
    var html = '<div class="pb-action-wrap" style="display:inline-block;">';
    html += '<div class="pb-action-trigger" onclick="event.stopPropagation();pbToggleActionMenu(' + pid + ')">Action ' + arrow + '</div>';
    var upCls = openUpward ? ' pb-action-menu-up' : '';
    html += '<div class="pb-action-menu' + upCls + '" id="pb-amenu-' + pid + '">';
    html += pbBuildActionMenuSections(pid);
    html += '</div></div>';
    return html;
}

var pbOpenMenuPid = null;
function pbToggleActionMenu(pid) {
    // Close any open menu
    if (pbOpenMenuPid && pbOpenMenuPid !== pid) {
        var prev = document.getElementById('pb-amenu-' + pbOpenMenuPid);
        if (prev) prev.classList.remove('open');
    }
    var menu = document.getElementById('pb-amenu-' + pid);
    if (menu) {
        menu.classList.toggle('open');
        pbOpenMenuPid = menu.classList.contains('open') ? pid : null;
        // Load destination health on first open if needed
        if (menu.classList.contains('open') && !pbDestHealthLoaded) {
            pbLoadDestinationHealth();
        }
    }
}

// Close action menu when clicking outside
document.addEventListener('click', function(e) {
    if (pbOpenMenuPid && !e.target.closest('.pb-action-wrap')) {
        var menu = document.getElementById('pb-amenu-' + pbOpenMenuPid);
        if (menu) menu.classList.remove('open');
        pbOpenMenuPid = null;
    }
});

// LIST VIEW
function pbRenderListView() {
    var el = document.getElementById('pb-list-content');
    if (!pbFilteredProspects.length) {
        el.innerHTML = '<div class="pb-empty">No prospects match your filters.</div>';
        return;
    }
    var html = '';
    for (var i = 0; i < pbFilteredProspects.length; i++) {
        var p = pbFilteredProspects[i];
        var pid = p.PeopleId;
        var statusClass = pbGetStatusClass(pid);
        var isExpanded = pbExpandedPid === pid;
        html += '<div class="pb-prospect-card ' + statusClass + '" id="pb-card-' + pid + '" data-pid="' + pid + '">';
        // Row: photo + info + action dropdown
        html += '<div style="display:flex;gap:8px;align-items:center;">';
        if (p.PhotoUrl) {
            html += '<img src="' + p.PhotoUrl + '" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0;cursor:pointer;" onerror="this.style.display=&quot;none&quot;" onclick="pbToggleExpand(' + pid + ')">';
        }
        html += '<div style="flex:1;min-width:0;cursor:pointer;" onclick="pbToggleExpand(' + pid + ')">';
        html += '<span class="pb-bold"><a href="/Person2/' + pid + '" target="_blank" style="color:var(--pb-primary);text-decoration:none;" onclick="event.stopPropagation()">' + (p.Name2 || 'Unknown') + '</a></span>';
        if (statusClass) html += ' <span class="pb-badge pb-badge-' + (statusClass === 'processed' ? 'success' : statusClass === 'deferred' ? 'warning' : 'muted') + '" style="cursor:pointer;" onclick="event.stopPropagation();pbResetStatus(' + pid + ')" title="Click to undo">' + statusClass + ' &times;</span>';
        html += pbBuildGroupBadges(pid);
        html += ' <span class="pb-text-muted" style="font-size:0.8em;">';
        var inlineParts = [];
        if (p.CellPhone) inlineParts.push(pbFmtPhone(p.CellPhone));
        if (p.EmailAddress) inlineParts.push(p.EmailAddress);
        if (p.Age) inlineParts.push('Age ' + p.Age);
        if (p.MemberType) inlineParts.push(p.MemberType);
        html += inlineParts.join(' &middot; ');
        html += '</span>';
        html += pbBuildScoreBadge(pid);
        html += '</div>';
        // Right side: flags + contacts + action dropdown
        html += '<div style="display:flex;gap:6px;align-items:center;flex-shrink:0;">';
        var flagHtml = pbBuildFlagBadges(pid);
        if (flagHtml) html += '<span style="display:none;"></span>' + flagHtml;
        html += pbBuildContactBadges(pid);
        if (!pbProcessedMap[pid]) html += pbBuildActionMenu(pid);
        html += '</div>';
        html += '</div>';
        // Expandable detail area
        html += '<div id="pb-detail-' + pid + '" style="' + (isExpanded ? '' : 'display:none;') + 'margin-top:8px;border-top:1px solid var(--pb-border);padding-top:8px;">';
        if (isExpanded && pbDetailCache[pid]) {
            html += pbRenderPersonDetail(p, pbDetailCache[pid]);
        } else if (isExpanded) {
            html += '<div class="pb-loading"><span class="pb-spin"></span> Loading details...</div>';
        }
        html += '</div>';
        html += '</div>';
    }
    el.innerHTML = html;
    if (pbExpandedPid) {
        var card = document.getElementById('pb-card-' + pbExpandedPid);
        if (card) card.scrollIntoView({behavior: 'smooth', block: 'nearest'});
    }
}

function pbToggleExpand(pid) {
    if (pbExpandedPid === pid) {
        // Collapse
        pbExpandedPid = null;
        var detailEl = document.getElementById('pb-detail-' + pid);
        if (detailEl) detailEl.style.display = 'none';
        // Update arrow
        var card = document.getElementById('pb-card-' + pid);
        if (card) { var arrows = card.querySelectorAll('span'); /* last span has arrow */ }
        return;
    }
    // Collapse previous
    if (pbExpandedPid) {
        var prevDetail = document.getElementById('pb-detail-' + pbExpandedPid);
        if (prevDetail) prevDetail.style.display = 'none';
    }
    pbExpandedPid = pid;
    var detailEl = document.getElementById('pb-detail-' + pid);
    if (!detailEl) return;
    detailEl.style.display = '';

    // Check cache
    if (pbDetailCache[pid]) {
        var p = pbProspects.find(function(pr) { return pr.PeopleId === pid; });
        detailEl.innerHTML = pbRenderPersonDetail(p, pbDetailCache[pid]);
        return;
    }

    // Load detail
    detailEl.innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading details...</div>';
    var evFields = (pbCurrentConfig.displayFields || []).filter(function(f) { return f.fieldType === 'extravalue'; }).map(function(f) { return f.sourceField; });
    var orgId = (pbCurrentConfig.source && pbCurrentConfig.source.pb_type === 'involvement') ? pbCurrentConfig.source.orgId : 0;
    pbAjax({action: 'load_person_detail', people_id: pid, org_id: orgId, ev_fields: evFields.join('|')}, function(d) {
        if (d.success) {
            pbDetailCache[pid] = d.detail;
            // Also merge into global maps for single/batch views
            if (d.detail.family && d.detail.family.FamilyDetailList) {
                pbFamilyData[String(pid)] = d.detail.family;
            }
            pbInvData[String(pid)] = d.detail.involvements || [];
            if (d.detail.extraValues) pbExtraValues[String(pid)] = d.detail.extraValues;
            if (d.detail.regData) pbRegData[String(pid)] = d.detail.regData;

            var p = pbProspects.find(function(pr) { return pr.PeopleId === pid; });
            var el = document.getElementById('pb-detail-' + pid);
            if (el && p) el.innerHTML = pbRenderPersonDetail(p, d.detail);
        } else {
            var el = document.getElementById('pb-detail-' + pid);
            if (el) el.innerHTML = '<div class="pb-text-muted">Failed to load details.</div>';
        }
    });
}

// ============================================================
// DESTINATION HEALTH SIDE PANEL
// ============================================================
var pbDestPanelOpen = false;

function pbToggleDestPanel() {
    var panel = document.getElementById('pb-dest-panel');
    var btn = document.getElementById('pb-dest-panel-btn');
    pbDestPanelOpen = !pbDestPanelOpen;
    panel.style.display = pbDestPanelOpen ? '' : 'none';
    if (btn) btn.style.background = pbDestPanelOpen ? 'var(--pb-accent)' : '#e6f7ff';
    if (btn) btn.style.color = pbDestPanelOpen ? '#fff' : 'var(--pb-accent)';
    if (pbDestPanelOpen) pbLoadDestPanel();
}

function pbLoadDestPanel() {
    var actions = (pbCurrentConfig && pbCurrentConfig.targetActions) ? pbCurrentConfig.targetActions : [];
    var invActions = actions.filter(function(a) { return a.pb_type === 'involvement'; });
    var el = document.getElementById('pb-dest-panel-content');

    if (!invActions.length) {
        el.innerHTML = '<div class="pb-text-muted" style="text-align:center;padding:20px;">No involvement destinations configured.</div>';
        return;
    }

    el.innerHTML = '<div style="text-align:center;padding:15px;"><span class="pb-spin"></span> Loading...</div>';

    var orgIds = invActions.map(function(a) { return a.orgId; }).join(',');
    var ncd = (pbCurrentConfig && pbCurrentConfig.noContactDays) || 90;
    var prospectTypes = '';

    pbAjax({action: 'get_destination_health', org_ids: orgIds, no_contact_days: ncd, prospect_types: prospectTypes}, function(d) {
        if (!d.success) {
            el.innerHTML = '<div class="pb-text-muted">Failed to load.</div>';
            return;
        }
        var html = '';
        for (var i = 0; i < invActions.length; i++) {
            var orgId = invActions[i].orgId;
            var h = d.health ? d.health[String(orgId)] : null;
            if (!h) continue;

            var statusColor = h.healthStatus === 'Healthy' ? 'var(--pb-success)' : h.healthStatus === 'Warning' ? 'var(--pb-warning)' : 'var(--pb-danger)';
            html += '<div style="border-left:3px solid ' + statusColor + ';padding:8px 10px;margin-bottom:8px;background:var(--pb-light);border-radius:0 var(--pb-radius) var(--pb-radius) 0;cursor:pointer;" onclick="pbShowDestDetail(' + orgId + ')">';
            html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">';
            html += '<strong style="font-size:0.95em;">' + (invActions[i].orgName || invActions[i].label) + '</strong>';
            html += '<span style="background:' + statusColor + ';color:#fff;padding:1px 6px;border-radius:8px;font-size:10px;">' + (h.healthStatus || '') + '</span>';
            html += '</div>';

            // Compact stats
            html += '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:4px;margin-bottom:4px;">';
            html += '<div style="text-align:center;padding:3px;background:var(--pb-white);border-radius:3px;"><strong>' + (h.totalMembers || 0) + '</strong><br><span style="font-size:0.8em;color:var(--pb-muted);">Total</span></div>';
            html += '<div style="text-align:center;padding:3px;background:var(--pb-white);border-radius:3px;"><strong style="color:var(--pb-accent);">' + (h.prospectCount || 0) + '</strong><br><span style="font-size:0.8em;color:var(--pb-muted);">Prospects</span></div>';
            html += '<div style="text-align:center;padding:3px;background:var(--pb-white);border-radius:3px;"><strong style="color:var(--pb-warning);">' + (h.staleCount || 0) + '</strong><br><span style="font-size:0.8em;color:var(--pb-muted);">Stale</span></div>';
            html += '</div>';

            if (h.lastMeeting) html += '<div class="pb-text-muted" style="font-size:0.85em;">Last mtg: ' + h.lastMeeting + '</div>';
            if (h.avgAttendance) html += '<div class="pb-text-muted" style="font-size:0.85em;">Avg attend: ' + h.avgAttendance + '</div>';
            html += '</div>';
        }
        if (!html) html = '<div class="pb-text-muted" style="text-align:center;padding:15px;">No health data available.</div>';
        el.innerHTML = html;
    });
}

function pbToggleFamilyTimeline(pid) {
    var el = document.getElementById('pb-fam-tl-' + pid);
    var icon = document.getElementById('pb-fam-tl-icon-' + pid);
    if (!el) return;
    var hidden = el.style.display === 'none';
    el.style.display = hidden ? '' : 'none';
    if (icon) icon.className = hidden ? 'fa fa-caret-down' : 'fa fa-caret-right';
    // Load on first expand
    if (hidden && !el.dataset.loaded) {
        el.innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading family timeline...</div>';
        el.dataset.loaded = '1';
        pbAjax({action: 'get_family_journey', peopleId: pid}, function(d) {
            if (!d.success) {
                el.innerHTML = '<div class="pb-text-muted">' + (d.error || d.message || 'Failed to load') + '</div>';
                return;
            }
            el.innerHTML = pbRenderInlineFamilyTimeline(d);
        });
    }
}

function pbRenderInlineFamilyTimeline(data) {
    var family = data.family_info;
    var members = data.family_members;
    var html = '<div style="font-size:0.82em;">';

    // Family summary bar
    html += '<div style="display:flex;gap:12px;flex-wrap:wrap;padding:6px 10px;background:var(--pb-light);border-radius:var(--pb-radius);margin-bottom:8px;">';
    html += '<span><strong>' + family.total_members + '</strong> members</span>';
    html += '<span><strong>' + family.engaged_members + '</strong> engaged</span>';
    html += '<span>Avg score: <strong>' + family.avg_engagement + '</strong></span>';
    html += '</div>';

    // Each member
    for (var mi = 0; mi < members.length; mi++) {
        var member = members[mi];
        var person = member.person;
        var journey = member.journey;
        var ins = member.insights;

        html += '<div style="border:1px solid var(--pb-border);border-radius:var(--pb-radius);margin-bottom:6px;overflow:hidden;">';
        // Header
        html += '<div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#fff;padding:6px 10px;display:flex;justify-content:space-between;align-items:center;">';
        html += '<span><strong>' + person.name + '</strong> <small>' + (person.position || '') + ' &middot; Age ' + (person.age || '?') + '</small></span>';
        html += '<span style="background:rgba(255,255,255,0.2);padding:2px 8px;border-radius:10px;font-size:11px;">Score: ' + person.engagement_score + '</span>';
        html += '</div>';

        // Events
        html += '<div style="padding:6px 10px;">';
        if (journey.length === 0) {
            html += '<div class="pb-text-muted" style="text-align:center;padding:8px;">No events recorded</div>';
        } else {
            html += '<div style="position:relative;padding-left:14px;">';
            html += '<div style="position:absolute;left:4px;top:2px;bottom:2px;width:2px;background:#ddd;"></div>';
            for (var ei = 0; ei < journey.length; ei++) {
                var ev = journey[ei];
                html += '<div style="position:relative;margin-bottom:3px;padding-left:12px;display:flex;align-items:baseline;gap:6px;">';
                html += '<div style="position:absolute;left:-11px;top:4px;width:8px;height:8px;border-radius:50%;background:' + ev.color + ';"></div>';
                html += '<span style="font-weight:600;color:' + ev.color + ';font-size:0.95em;">' + ev.event + '</span>';
                html += '<span class="pb-text-muted" style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + ev.description + '</span>';
                html += '<span class="pb-text-muted" style="white-space:nowrap;">' + ev.date + '</span>';
                html += '</div>';
            }
            html += '</div>';
        }
        html += '</div></div>';
    }

    html += '</div>';
    return html;
}

function pbTogglePastInv(pid) {
    var el = document.getElementById('pb-past-inv-' + pid);
    var icon = document.getElementById('pb-past-inv-icon-' + pid);
    if (el) {
        var hidden = el.style.display === 'none';
        el.style.display = hidden ? '' : 'none';
        if (icon) icon.className = hidden ? 'fa fa-caret-down' : 'fa fa-caret-right';
    }
}

function pbFormatDuration(days) {
    if (!days || days <= 0) return '';
    var y = Math.floor(days / 365);
    var m = Math.floor((days % 365) / 30);
    if (y > 0 && m > 0) return y + 'y ' + m + 'm';
    if (y > 0) return y + 'y';
    if (m > 0) return m + 'm';
    return days + 'd';
}

function pbRenderPersonDetail(p, detail) {
    var pid = p.PeopleId;
    var fields = (pbCurrentConfig && pbCurrentConfig.displayFields) ? pbCurrentConfig.displayFields : [];
    var prof = detail.profile || {};
    var eng = detail.engagement || null;
    var mediumPhoto = prof.mediumPhoto || prof.largePhoto || p.PhotoUrl || '';
    var largePhoto = prof.largePhoto || mediumPhoto;
    var html = '';

    // === ROW 1: Photo + Contact/Demographics with inline engagement ===
    html += '<div style="display:flex;gap:10px;align-items:flex-start;font-size:0.85em;">';

    // Photo
    if (mediumPhoto) {
        html += '<img src="' + mediumPhoto + '" data-large="' + largePhoto + '" onclick="pbShowPhotoModal(this.dataset.large || this.src)" style="width:72px;height:72px;border-radius:8px;object-fit:cover;cursor:pointer;border:2px solid var(--pb-border);flex-shrink:0;" onerror="this.style.display=&quot;none&quot;" title="Click to enlarge">';
    }

    // Contact + Demographics + Score inline
    html += '<div style="flex:1;min-width:0;">';

    // First line: scorecard badge + contact info
    html += '<div style="display:flex;flex-wrap:wrap;gap:3px 12px;align-items:center;margin-bottom:3px;">';
    var scEntry = pbScoreMap[String(pid)];
    if (scEntry && pbSettings.scorecard && pbSettings.scorecard.enabled) {
        var scBadgeColor = pbGetScoreColor(scEntry.score);
        html += '<span data-pid="' + pid + '" onclick="pbShowScorePopup(parseInt(this.dataset.pid),event)" style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:12px;background:var(--pb-light);border:1px solid ' + scBadgeColor + ';font-size:0.9em;cursor:pointer;" title="Click for score breakdown">';
        html += '<strong style="color:var(--pb-primary);">' + scEntry.score + '</strong>';
        html += '<span style="background:' + scBadgeColor + ';width:30px;height:4px;border-radius:2px;display:inline-block;"></span>';
        // Show activity status from engagement data if available
        if (eng) {
            var statusColor = eng.status === 'Active' ? 'var(--pb-success)' : eng.status === 'Inactive' ? 'var(--pb-danger)' : 'var(--pb-warning)';
            html += '<span style="background:' + statusColor + ';color:#fff;padding:0 5px;border-radius:6px;font-size:10px;">' + eng.status + '</span>';
        }
        html += '</span>';
    } else if (eng) {
        // Fallback to engagement score if scorecard not enabled
        var barColor = eng.score >= 60 ? 'var(--pb-success)' : eng.score >= 40 ? 'var(--pb-warning)' : 'var(--pb-danger)';
        var statusColor = eng.status === 'Active' ? 'var(--pb-success)' : eng.status === 'Inactive' ? 'var(--pb-danger)' : 'var(--pb-warning)';
        html += '<span style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:12px;background:var(--pb-light);border:1px solid ' + barColor + ';font-size:0.9em;">';
        html += '<strong style="color:var(--pb-primary);">' + eng.score + '</strong>';
        html += '<span style="background:' + barColor + ';width:30px;height:4px;border-radius:2px;display:inline-block;"></span>';
        html += '<span style="background:' + statusColor + ';color:#fff;padding:0 5px;border-radius:6px;font-size:10px;">' + eng.status + '</span>';
        html += '</span>';
    }
    if (p.EmailAddress) html += '<span>&#9993; ' + p.EmailAddress + '</span>';
    if (p.CellPhone) html += '<span>&#9742; ' + pbFmtPhone(p.CellPhone) + '</span>';
    if (prof.homePhone) html += '<span>&#127968; ' + pbFmtPhone(prof.homePhone) + '</span>';
    if (prof.workPhone) html += '<span>&#128188; ' + pbFmtPhone(prof.workPhone) + '</span>';
    html += '</div>';

    if (prof.address) html += '<div class="pb-text-muted" style="margin-bottom:2px;">&#128205; ' + prof.address + '</div>';
    var demoParts = [];
    if (prof.employer) demoParts.push('Employer: ' + prof.employer);
    if (prof.occupation) demoParts.push('Occupation: ' + prof.occupation);
    if (prof.school) demoParts.push('School: ' + prof.school);
    if (demoParts.length) html += '<div class="pb-text-muted" style="margin-bottom:2px;">' + demoParts.join(' &middot; ') + '</div>';
    var personFields = fields.filter(function(f) { return f.fieldType === 'person' && ['Name2','EmailAddress','CellPhone','Age','MemberType'].indexOf(f.sourceField) < 0; });
    var evFields = fields.filter(function(f) { return f.fieldType === 'extravalue'; });
    var regFields = fields.filter(function(f) { return f.fieldType === 'regquestion'; });
    var medFields = fields.filter(function(f) { return f.fieldType === 'medical'; });
    var simpleFields = personFields.concat(evFields).concat(regFields).concat(medFields);
    if (simpleFields.length) {
        html += '<div style="display:flex;flex-wrap:wrap;gap:2px 12px;">';
        for (var i = 0; i < simpleFields.length; i++) {
            var val = pbGetFieldValue(p, simpleFields[i]);
            if (val) html += '<span><span class="pb-text-muted">' + simpleFields[i].label + ':</span> ' + val + '</span>';
        }
        html += '</div>';
    }
    html += '</div>';

    html += '</div>';

    // === ROW 2: Milestones + Engagement Breakdown side by side ===
    var milestones = detail.milestones || [];
    if (milestones.length || eng) {
        html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px;font-size:0.82em;">';

        // Milestones card
        html += '<div style="padding:8px 10px;border-left:3px solid var(--pb-accent);background:var(--pb-light);border-radius:var(--pb-radius);">';
        html += '<div class="pb-bold" style="color:var(--pb-accent);margin-bottom:4px;font-size:0.9em;"><i class="fa fa-flag"></i> Milestones</div>';
        if (milestones.length) {
            for (var mi = 0; mi < milestones.length; mi++) {
                var ms = milestones[mi];
                html += '<div style="display:flex;justify-content:space-between;padding:1px 0;">';
                html += '<span>' + ms.icon + ' ' + ms.label + '</span>';
                html += '<span class="pb-text-muted">' + ms.date + '</span>';
                html += '</div>';
            }
        } else {
            html += '<span class="pb-text-muted">None recorded</span>';
        }
        html += '</div>';

        // Engagement Breakdown card
        html += '<div style="padding:8px 10px;border-left:3px solid var(--pb-success);background:var(--pb-light);border-radius:var(--pb-radius);">';
        html += '<div class="pb-bold" style="color:var(--pb-success);margin-bottom:4px;font-size:0.9em;"><i class="fa fa-bar-chart"></i> Engagement Breakdown <span style="font-weight:400;color:var(--pb-text-muted,#94a3b8);font-size:10px;">(what&#39;s happening?)</span></div>';
        if (eng && eng.factors) {
            var fLabels = {attend_recency: "Attend Recency", attend_frequency: "Attend Frequency", group_involvement: "Groups", serving: "Serving"};
            var fIcons = {attend_recency: "&#128197;", attend_frequency: "&#128200;", group_involvement: "&#128101;", serving: "&#9997;"};
            var fDetails = {};
            if (eng.daysSince != null) fDetails.attend_recency = eng.daysSince + "d ago";
            fDetails.attend_frequency = (eng.attend90 || 0) + "x (90d)";
            fDetails.group_involvement = (eng.groupCount || 0) + " orgs";
            fDetails.serving = (eng.servingCount || 0) + " roles";
            var fKeys = ["attend_recency", "attend_frequency", "group_involvement", "serving"];
            for (var fi = 0; fi < fKeys.length; fi++) {
                var fk = fKeys[fi]; var ff = eng.factors[fk]; if (!ff) continue;
                var fs = ff.score; var fw = ff.weight;
                var fbc = fs >= 70 ? "var(--pb-success)" : fs >= 40 ? "var(--pb-warning)" : "var(--pb-danger)";
                html += '<div style="display:flex;align-items:center;gap:6px;padding:2px 0;font-size:11px">';
                html += '<span style="width:14px">' + fIcons[fk] + '</span>';
                html += '<span style="width:80px;color:var(--pb-text-muted,#475569)">' + fLabels[fk] + '</span>';
                html += '<div style="flex:1;height:6px;background:#e2e8f0;border-radius:3px;overflow:hidden"><div style="width:' + fs + '%;height:100%;background:' + fbc + ';border-radius:3px"></div></div>';
                html += '<span style="width:28px;text-align:right;font-weight:600;color:' + fbc + '">' + fs + '</span>';
                html += '<span style="width:55px;color:var(--pb-text-muted,#94a3b8);font-size:10px">' + ((fDetails[fk]) || "") + '</span>';
                html += '</div>';
            }
            if (eng.lastAttendance) html += '<div style="font-size:11px;color:var(--pb-text-muted,#64748b);margin-top:4px;border-top:1px solid var(--pb-border);padding-top:3px">Last: ' + eng.lastAttendance + '</div>';
        } else if (eng) {
            if (eng.lastAttendance) html += '<div style="padding:1px 0;"><span class="pb-text-muted">Last seen:</span> ' + eng.lastAttendance + '</div>';
            if (eng.attend90 !== undefined) html += '<div style="padding:1px 0;"><span class="pb-text-muted">90d:</span> ' + eng.attend90 + 'x</div>';
            if (eng.servingCount > 0) html += '<div style="padding:1px 0;"><span class="pb-text-muted">Serving:</span> ' + eng.servingCount + ' role(s)</div>';
        } else {
            html += '<span class="pb-text-muted">No data</span>';
        }
        html += '</div>';

        html += '</div>';
    }

    // === ROW 3: Journey Timeline (full width) ===
    var journeyEvents = detail.journeyEvents || [];
    if (journeyEvents.length) {
        html += '<div style="padding:10px 12px;border-left:3px solid #667eea;background:var(--pb-light);border-radius:var(--pb-radius);margin-top:8px;font-size:0.82em;">';
        html += '<div class="pb-bold" style="color:#667eea;margin-bottom:6px;font-size:0.9em;"><i class="fa fa-map-o"></i> Journey Timeline</div>';
        html += '<div style="position:relative;padding-left:16px;">';
        html += '<div style="position:absolute;left:5px;top:2px;bottom:2px;width:2px;background:#dee2e6;"></div>';
        for (var ji = 0; ji < journeyEvents.length; ji++) {
            var je = journeyEvents[ji];
            var isInv = (je.event === 'Joined' || je.event === 'Joined (Past)');
            html += '<div style="position:relative;margin-bottom:5px;padding-left:14px;">';
            html += '<div style="position:absolute;left:-12px;top:4px;width:10px;height:10px;border-radius:50%;background:' + je.color + ';"></div>';
            // Line 1: event type + program prefix + description
            html += '<div style="display:flex;align-items:baseline;gap:6px;">';
            html += '<span style="font-weight:600;color:' + je.color + ';white-space:nowrap;">' + je.event + '</span>';
            if (je.program) html += '<span class="pb-text-muted" style="white-space:nowrap;">' + je.program + ' &rsaquo;</span>';
            html += '<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + je.description + '</span>';
            html += '</div>';
            // Line 2: dates row with icons
            var dateParts = [];
            if (je.enrolled) dateParts.push('<span><i class="fa fa-sign-in" style="color:var(--pb-success);"></i> Enrolled ' + je.enrolled + '</span>');
            else if (je.date) dateParts.push('<span><i class="fa fa-calendar-o"></i> ' + je.date + '</span>');
            if (isInv) {
                if (je.lastAttended) dateParts.push('<span><i class="fa fa-calendar-check-o" style="color:var(--pb-accent);"></i> Last ' + je.lastAttended + '</span>');
                else dateParts.push('<span style="color:var(--pb-danger);"><i class="fa fa-calendar-times-o"></i> No meetings</span>');
            }
            if (je.inactive) dateParts.push('<span><i class="fa fa-sign-out" style="color:var(--pb-danger);"></i> Dropped ' + je.inactive + '</span>');
            if (dateParts.length) {
                html += '<div class="pb-text-muted" style="font-size:0.9em;margin-top:1px;display:flex;gap:14px;flex-wrap:wrap;">' + dateParts.join('') + '</div>';
            }
            html += '</div>';
        }
        html += '</div>';
        html += '</div>';
    }

    // === ROW 4: Journey Insights ===
    if (journeyEvents.length && eng) {
        html += '<div style="padding:8px 12px;background:#e6f7ff;border-radius:var(--pb-radius);border-left:3px solid var(--pb-accent);margin-top:8px;font-size:0.82em;">';
        html += '<div class="pb-bold" style="color:var(--pb-accent);margin-bottom:4px;font-size:0.9em;"><i class="fa fa-lightbulb-o"></i> Journey Insights</div>';
        html += '<div style="display:flex;flex-wrap:wrap;gap:4px 16px;color:var(--pb-text);">';
        html += '<span><strong>Events:</strong> ' + journeyEvents.length + '</span>';
        // Entry point: first non-system event
        var entryPoint = '';
        for (var ep = 0; ep < journeyEvents.length; ep++) {
            if (journeyEvents[ep].type !== 'system') { entryPoint = journeyEvents[ep].description; break; }
        }
        if (entryPoint) html += '<span><strong>Entry:</strong> ' + entryPoint + '</span>';
        // Show scorecard score if enabled, else engagement score
        var jiScore = (scEntry && pbSettings.scorecard && pbSettings.scorecard.enabled) ? scEntry.score : eng.score;
        html += '<span><strong>Status:</strong> ' + eng.level + ' (' + jiScore + ')</span>';
        // Check for patterns
        var hasAtt = journeyEvents.some(function(e) { return e.type === 'attendance'; });
        var hasSrv = journeyEvents.some(function(e) { return e.type === 'serving' || e.type === 'leader'; });
        var hasSg = journeyEvents.some(function(e) { return e.type === 'smallgroup'; });
        if (hasAtt && hasSrv) html += '<span><strong>Active:</strong> Attendance &rarr; Serving</span>';
        if (hasSg) html += '<span><strong>Small Group:</strong> Connected</span>';
        if (eng.daysSince > 90) html += '<span style="color:var(--pb-danger);"><strong>Re-engage:</strong> 90+ days absent</span>';
        html += '</div></div>';
    }

    // === ROW 4b: Prospect Scorecard ("What should I do?") ===
    var scData = pbScoreMap[String(pid)] || pbScoreMap[pid];
    if (scData && pbSettings.scorecard && pbSettings.scorecard.enabled) {
        var scColor = pbGetScoreColor(scData.score);
        html += '<div style="padding:8px 12px;background:#f0f4ff;border-radius:var(--pb-radius);border-left:3px solid ' + scColor + ';margin-top:8px;font-size:0.82em;">';
        html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">';
        html += '<span class="pb-bold" style="color:var(--pb-primary);font-size:0.9em;"><i class="fa fa-bullseye"></i> Priority Scorecard <span style="font-weight:400;color:var(--pb-text-muted,#94a3b8);font-size:10px;">(what should I do?)</span></span>';
        html += '<span style="background:' + scColor + ';color:#fff;padding:2px 10px;border-radius:10px;font-weight:700;font-size:1.05em;">' + scData.score + ' / 100</span>';
        html += '</div>';
        html += pbBuildScoreDetail(scData);
        html += '</div>';
    }

    // === ROW 5: Family ===
    var famList = (detail.family && detail.family.FamilyDetailList) ? detail.family.FamilyDetailList : [];
    if (famList.length) {
        html += '<div style="margin-top:6px;"><div class="pb-bold pb-text-sm" style="color:var(--pb-primary);margin-bottom:2px;">Family (' + famList.length + ')</div>' + pbRenderFamilyDetail(famList) + '</div>';
    }

    // === ROW 6: Family Timeline (collapsible, loaded on demand) ===
    html += '<div style="margin-top:6px;">';
    html += '<div class="pb-bold pb-text-sm" style="color:#667eea;cursor:pointer;" onclick="pbToggleFamilyTimeline(' + pid + ')">';
    html += '<i class="fa fa-caret-right" id="pb-fam-tl-icon-' + pid + '"></i> <i class="fa fa-users"></i> Family Journey Timeline</div>';
    html += '<div id="pb-fam-tl-' + pid + '" style="display:none;margin-top:4px;"></div>';
    html += '</div>';

    return html;
}

function pbShowPhotoModal(src) {
    var overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.7);z-index:20000;display:flex;align-items:center;justify-content:center;cursor:pointer;';
    overlay.onclick = function() { document.body.removeChild(overlay); };
    var img = document.createElement('img');
    img.src = src;
    img.style.cssText = 'max-width:90vw;max-height:85vh;border-radius:12px;box-shadow:0 8px 40px rgba(0,0,0,0.5);';
    overlay.appendChild(img);
    document.body.appendChild(overlay);
}

// SINGLE VIEW
function pbRenderSingleView() {
    if (!pbFilteredProspects.length) {
        document.getElementById('pb-single-content').innerHTML = '<div class="pb-empty">No prospects to display.</div>';
        return;
    }
    if (pbCurrentIndex >= pbFilteredProspects.length) pbCurrentIndex = pbFilteredProspects.length - 1;
    if (pbCurrentIndex < 0) pbCurrentIndex = 0;

    var p = pbFilteredProspects[pbCurrentIndex];
    var pid = p.PeopleId;
    document.getElementById('pb-single-index').textContent = (pbCurrentIndex + 1) + ' of ' + pbFilteredProspects.length;

    var statusClass = pbGetStatusClass(pid);

    var html = '<div class="pb-prospect-card ' + statusClass + '" style="border-width:2px;">';

    // Top banner: name on the left, Action menu top-right (uses the dead
    // space alongside the name and stays in view even when the card grows
    // long). pbRenderPersonDetail handles photo/contact/score below.
    html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--pb-border);">';
    html += '<div style="flex:1;min-width:0;">';
    html += '<h3 style="margin:0 0 4px 0;"><a href="/Person2/' + pid + '" target="_blank" style="color:var(--pb-primary);text-decoration:none;">' + (p.Name2 || 'Unknown') + '</a></h3>';
    if (statusClass) html += '<span class="pb-badge pb-badge-' + (statusClass === 'processed' ? 'success' : statusClass === 'deferred' ? 'warning' : 'muted') + '" style="font-size:0.85em;">' + statusClass + '</span> ';
    // Contact-effort badges only render when there's actual history; the
    // muted "No contacts" placeholder is just noise here.
    var cm = pbContactMap[String(pid)];
    if (cm) html += '<div class="pb-mt-sm">' + pbBuildContactBadges(pid) + '</div>';
    var fb = pbBuildFlagBadges(pid);
    if (fb) html += '<div class="pb-mt-sm">' + fb + '</div>';

    // Quick contact-log buttons (Phone / Email / Text / Visit / Mail / Notes)
    // -- same shape used in Group Management's expand. Lets the user log a
    // contact attempt without having to drop into the Action menu.
    var methods = (pbSettings && pbSettings.contact_methods) ? pbSettings.contact_methods : [];
    if (methods.length || p.PeopleId) {
        var cmColors = {P:'#0078d4',E:'#107c10',T:'#6b69d6',V:'#d83b01',M:'#8764b8'};
        html += '<div class="pb-mt-sm" style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">';
        html += '<span style="font-weight:600;font-size:0.85em;color:#666;">Log contact:</span>';
        for (var cmi = 0; cmi < methods.length; cmi++) {
            var m = methods[cmi];
            var col = cmColors[m.code] || '#666';
            html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthQuickLog(' + pid + ',&quot;'
                  + (m.keyword || '').replace(/"/g, '') + '&quot;)" style="padding:3px 10px;'
                  + 'font-size:0.8em;background:' + col + ';color:#fff;" title="Log '
                  + (m.label || m.code) + '">' + (m.code || '?') + ' ' + (m.label || '') + '</button>';
        }
        html += '<button class="pb-btn pb-btn-sm" onclick="pbHealthLogContact(' + pid + ')" '
              + 'style="padding:3px 10px;font-size:0.8em;">&#128221; Notes</button>';
        html += '</div>';
    }

    html += '</div>';  // end left column

    // Right side: Action menu. Uses the native <details> variant in Single
    // -- the absolute-positioned dropdown was getting eaten by the tall
    // card's flex layout, and <details> handles open/close natively.
    if (!pbProcessedMap[pid]) {
        html += '<div style="flex-shrink:0;position:relative;">' + pbBuildActionPanel(pid) + '</div>';
    }
    html += '</div>';  // end banner

    // Rich detail block -- exact same renderer the List view's expand uses.
    // It includes the configured display fields inline + engagement +
    // scorecard + milestones + journey + family + family-timeline. Lazy-
    // loaded so Prev/Next stays snappy.
    html += '<div id="pb-single-rich">';
    if (pbDetailCache[pid]) {
        html += pbRenderPersonDetail(p, pbDetailCache[pid]);
    } else {
        html += '<div class="pb-loading" style="text-align:center;padding:18px;"><span class="pb-spin"></span> Loading prospect detail...</div>';
    }
    html += '</div>';

    // (Action menu now lives top-right of the header -- no second copy here.)

    html += '</div>';
    document.getElementById('pb-single-content').innerHTML = html;

    // Lazy-load detail when we don't have it cached. Once it lands, both
    // Single and List share the same pbDetailCache so re-visits are instant.
    if (!pbDetailCache[pid]) {
        var evFields = (pbCurrentConfig.displayFields || []).filter(function(f) { return f.fieldType === 'extravalue'; }).map(function(f) { return f.sourceField; });
        var orgId = (pbCurrentConfig.source && pbCurrentConfig.source.pb_type === 'involvement') ? pbCurrentConfig.source.orgId : 0;
        pbAjax({action: 'load_person_detail', people_id: pid, org_id: orgId, ev_fields: evFields.join('|')}, function(d) {
            if (!d.success) {
                var er = document.getElementById('pb-single-rich');
                if (er) er.innerHTML = '<div class="pb-text-muted">Failed to load prospect detail.</div>';
                return;
            }
            pbDetailCache[pid] = d.detail;
            // Hydrate shared maps that other views rely on.
            if (d.detail.family && d.detail.family.FamilyDetailList) pbFamilyData[String(pid)] = d.detail.family;
            pbInvData[String(pid)] = d.detail.involvements || [];
            if (d.detail.extraValues) pbExtraValues[String(pid)] = d.detail.extraValues;
            if (d.detail.regData) pbRegData[String(pid)] = d.detail.regData;
            // Guard against user navigating away before the ajax landed.
            if (pbViewMode !== 'single' || pbFilteredProspects[pbCurrentIndex]
                && pbFilteredProspects[pbCurrentIndex].PeopleId !== pid) return;
            var richEl = document.getElementById('pb-single-rich');
            if (richEl) richEl.innerHTML = pbRenderPersonDetail(p, d.detail);
        });
    }
}

function pbSinglePrev() { if (pbCurrentIndex > 0) { pbCurrentIndex--; pbRenderSingleView(); } }
function pbSingleNext() { if (pbCurrentIndex < pbFilteredProspects.length - 1) { pbCurrentIndex++; pbRenderSingleView(); } }

// BATCH VIEW
function pbRenderBatchView() {
    var el = document.getElementById('pb-batch-content');
    if (!pbFilteredProspects.length) {
        el.innerHTML = '<div class="pb-empty">No prospects match your filters.</div>';
        return;
    }
    var scorecardEnabled = !!(pbSettings && pbSettings.scorecard && pbSettings.scorecard.enabled);
    var html = '<div class="pb-table-wrap"><table class="pb-table"><thead><tr>';
    html += '<th style="width:30px;"><input type="checkbox" id="pb-batch-hdr-cb" onchange="pbToggleSelectAll(this.checked)"></th>';
    html += '<th style="width:18px;"></th>';  // chevron column
    html += '<th>Name</th>';
    if (scorecardEnabled) html += '<th style="width:60px;text-align:center;">Score</th>';
    html += '<th>Contact</th><th>Flags</th><th>Status</th>';
    html += '</tr></thead><tbody>';

    for (var i = 0; i < pbFilteredProspects.length; i++) {
        var p = pbFilteredProspects[i];
        var pid = p.PeopleId;
        var checked = pbBatchSelected[pid] ? ' checked' : '';
        var statusClass = pbGetStatusClass(pid);

        // Main row -- click anywhere except the checkbox / name link / score
        // chip toggles the detail. Stop-propagation guards live on those
        // child elements.
        html += '<tr class="pb-batch-row" data-pid="' + pid + '" style="cursor:pointer;" onclick="pbBatchToggleDetail(' + pid + ')">';
        html += '<td onclick="event.stopPropagation()"><input type="checkbox" class="pb-batch-cb" data-pid="' + pid + '"' + checked + ' onchange="pbToggleBatchItem(' + pid + ',this.checked)"></td>';
        html += '<td style="text-align:center;"><span class="pb-batch-chev" id="pb-batch-chev-' + pid + '" style="color:#888;font-weight:700;display:inline-block;transition:transform 0.18s ease;">&#10095;</span></td>';
        html += '<td><a href="/Person2/' + pid + '" target="_blank" onclick="event.stopPropagation()" style="color:var(--pb-primary);">' + (p.Name2 || 'Unknown') + '</a>';
        if (p.Age) html += ' <span class="pb-text-muted">(Age ' + p.Age + ')</span>';
        html += '</td>';
        if (scorecardEnabled) {
            var scData = pbScoreMap[String(pid)] || pbScoreMap[pid];
            if (scData) {
                var scColor = pbGetScoreColor(scData.score);
                html += '<td style="text-align:center;"><span title="Priority Scorecard -- '
                      + scData.score + ' / 100" style="display:inline-block;background:'
                      + scColor + ';color:#fff;padding:1px 8px;border-radius:10px;'
                      + 'font-weight:700;font-size:0.85em;">' + scData.score + '</span></td>';
            } else {
                html += '<td style="text-align:center;color:#bbb;font-size:0.85em;">-</td>';
            }
        }
        html += '<td>' + pbBuildContactBadges(pid) + '</td>';
        html += '<td>' + pbBuildFlagBadges(pid) + '</td>';
        html += '<td>';
        if (statusClass) html += '<span class="pb-badge pb-badge-' + (statusClass === 'processed' ? 'success' : statusClass === 'deferred' ? 'warning' : 'muted') + '">' + statusClass + '</span>';
        else html += '<span class="pb-badge pb-badge-primary">pending</span>';
        html += '</td></tr>';

        // Hidden detail row; lazy-loaded by pbBatchToggleDetail.
        var colSpan = scorecardEnabled ? 7 : 6;
        html += '<tr class="pb-batch-detail-row" id="pb-batch-detail-' + pid + '" style="display:none;">'
              + '<td colspan="' + colSpan + '" style="padding:0;border-bottom:2px solid #e0e0e0;background:#fafafa;">'
              + '<div id="pb-batch-detail-content-' + pid + '" style="padding:12px 16px;">'
              + '</div></td></tr>';
    }
    html += '</tbody></table></div>';
    el.innerHTML = html;
    pbUpdateBatchCount();
}

function pbBatchToggleDetail(pid) {
    var row = document.getElementById('pb-batch-detail-' + pid);
    var chev = document.getElementById('pb-batch-chev-' + pid);
    if (!row) return;
    if (row.style.display !== 'none') {
        row.style.display = 'none';
        if (chev) { chev.style.transform = ''; chev.style.color = '#888'; }
        return;
    }
    row.style.display = '';
    if (chev) { chev.style.transform = 'rotate(90deg)'; chev.style.color = 'var(--pb-accent)'; }

    // Lazy-load and render the rich detail. Shared cache with List + Single
    // so re-opening the same person here is instant.
    var contentEl = document.getElementById('pb-batch-detail-content-' + pid);
    if (!contentEl || contentEl.dataset.loaded === '1') return;
    contentEl.dataset.loaded = '1';
    var p = pbFilteredProspects.find(function(pr) { return pr.PeopleId === pid; })
         || pbProspects.find(function(pr) { return pr.PeopleId === pid; });
    if (!p) { contentEl.innerHTML = '<div class="pb-text-muted">Prospect not found.</div>'; return; }
    if (pbDetailCache[pid]) {
        contentEl.innerHTML = pbRenderPersonDetail(p, pbDetailCache[pid]);
        return;
    }
    contentEl.innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading details...</div>';
    var evFields = (pbCurrentConfig.displayFields || []).filter(function(f) { return f.fieldType === 'extravalue'; }).map(function(f) { return f.sourceField; });
    var orgId = (pbCurrentConfig.source && pbCurrentConfig.source.pb_type === 'involvement') ? pbCurrentConfig.source.orgId : 0;
    pbAjax({action: 'load_person_detail', people_id: pid, org_id: orgId, ev_fields: evFields.join('|')}, function(d) {
        if (!d.success) { contentEl.innerHTML = '<div class="pb-text-muted">Failed to load.</div>'; return; }
        pbDetailCache[pid] = d.detail;
        if (d.detail.family && d.detail.family.FamilyDetailList) pbFamilyData[String(pid)] = d.detail.family;
        pbInvData[String(pid)] = d.detail.involvements || [];
        if (d.detail.extraValues) pbExtraValues[String(pid)] = d.detail.extraValues;
        if (d.detail.regData) pbRegData[String(pid)] = d.detail.regData;
        contentEl.innerHTML = pbRenderPersonDetail(p, d.detail);
    });
}

function pbToggleSelectAll(checked) {
    pbBatchSelected = {};
    if (checked) {
        for (var i = 0; i < pbFilteredProspects.length; i++) {
            pbBatchSelected[pbFilteredProspects[i].PeopleId] = true;
        }
    }
    document.querySelectorAll('.pb-batch-cb').forEach(function(cb) { cb.checked = checked; });
    pbUpdateBatchCount();
}

function pbToggleBatchItem(pid, checked) {
    if (checked) pbBatchSelected[pid] = true;
    else delete pbBatchSelected[pid];
    pbUpdateBatchCount();
}

function pbUpdateBatchCount() {
    document.getElementById('pb-batch-selected-count').textContent = Object.keys(pbBatchSelected).length + ' selected';
}

function pbExecuteBatchAction() {
    var actionVal = document.getElementById('pb-batch-action').value;
    var selectedPids = Object.keys(pbBatchSelected).map(Number);
    if (!selectedPids.length) { alert('No prospects selected.'); return; }
    if (!actionVal) { alert('Please choose an action.'); return; }

    if (actionVal === '__processed') {
        selectedPids.forEach(function(pid) { pbProcessedMap[pid] = true; delete pbDeferredSet[pid]; delete pbSkippedSet[pid]; });
        pbApplyFilters();
    } else if (actionVal === '__deferred') {
        selectedPids.forEach(function(pid) { pbDeferredSet[pid] = true; delete pbProcessedMap[pid]; delete pbSkippedSet[pid]; });
        pbApplyFilters();
    } else if (actionVal === '__skipped') {
        selectedPids.forEach(function(pid) { pbSkippedSet[pid] = true; delete pbProcessedMap[pid]; delete pbDeferredSet[pid]; });
        pbApplyFilters();
    } else if (actionVal.indexOf('grp_') === 0) {
        var grpId = actionVal.replace('grp_', '');
        var grp = pbGrpFindGroup(grpId);
        var grpName = grp ? grp.name : 'group';
        if (!confirm('Assign ' + selectedPids.length + ' prospects to "' + grpName + '"?')) return;
        pbAjax({action: 'assign_to_group', group_id: grpId, people_ids: selectedPids.join(','), notes: 'Batch assigned from workspace'}, function(d) {
            if (d.success) {
                alert(d.assigned + ' prospect(s) assigned to "' + grpName + '".');
            }
        });
    } else if (actionVal.indexOf('action_') === 0) {
        var idx = parseInt(actionVal.replace('action_', ''));
        var target = pbCurrentConfig.targetActions[idx];
        if (!confirm('Apply "' + target.label + '" to ' + selectedPids.length + ' prospects?')) return;
        pbAjax({action: 'process_action', target_data: JSON.stringify(target), people_ids: selectedPids.join(',')}, function(d) {
            if (d.success) {
                selectedPids.forEach(function(pid) { pbProcessedMap[pid] = true; delete pbDeferredSet[pid]; delete pbSkippedSet[pid]; });
                alert('Processed ' + d.processed + ' prospects.' + (d.errors && d.errors.length ? ' Errors: ' + d.errors.join('; ') : ''));
                pbApplyFilters();
            } else {
                alert('Error: ' + (d.message || 'Failed'));
            }
        });
    }
    pbBatchSelected = {};
}

// ============================================================
// PROSPECT ACTIONS
// ============================================================
function pbProcessSingle(pid, actionIdx) {
    // Close action menu
    if (pbOpenMenuPid) {
        var menu = document.getElementById('pb-amenu-' + pbOpenMenuPid);
        if (menu) menu.classList.remove('open');
        pbOpenMenuPid = null;
    }
    var target = pbCurrentConfig.targetActions[actionIdx];
    var personName = pbGetProspectName(pid);
    pbAjax({action: 'process_action', target_data: JSON.stringify(target), people_ids: String(pid)}, function(d) {
        if (d.success) {
            pbProcessedMap[pid] = true;
            delete pbDeferredSet[pid];
            delete pbSkippedSet[pid];
            pbLogActivity(pid, personName, 'processed', target.label);
            pbApplyFilters();
            pbAutoSaveWorkState();
            if (pbViewMode === 'single') pbSingleNext();
        } else {
            alert('Error: ' + (d.message || 'Failed'));
        }
    });
}

function pbDeferProspect(pid) {
    if (pbOpenMenuPid) { var m = document.getElementById('pb-amenu-' + pbOpenMenuPid); if (m) m.classList.remove('open'); pbOpenMenuPid = null; }
    pbDeferredSet[pid] = true;
    delete pbProcessedMap[pid];
    delete pbSkippedSet[pid];
    pbLogActivity(pid, pbGetProspectName(pid), 'deferred');
    pbApplyFilters();
    pbAutoSaveWorkState();
    if (pbViewMode === 'single') pbSingleNext();
}

function pbSkipProspect(pid) {
    if (pbOpenMenuPid) { var m = document.getElementById('pb-amenu-' + pbOpenMenuPid); if (m) m.classList.remove('open'); pbOpenMenuPid = null; }
    pbSkippedSet[pid] = true;
    delete pbProcessedMap[pid];
    delete pbDeferredSet[pid];
    pbLogActivity(pid, pbGetProspectName(pid), 'skipped');
    pbApplyFilters();
    pbAutoSaveWorkState();
    if (pbViewMode === 'single') pbSingleNext();
}

function pbResetStatus(pid) {
    delete pbProcessedMap[pid];
    delete pbDeferredSet[pid];
    delete pbSkippedSet[pid];
    pbLogActivity(pid, pbGetProspectName(pid), 'reset');
    pbApplyFilters();
    pbAutoSaveWorkState();
}

function pbGetProspectName(pid) {
    for (var i = 0; i < pbProspects.length; i++) {
        if (pbProspects[i].PeopleId === pid) return pbProspects[i].Name2 || 'Unknown';
    }
    return 'PID ' + pid;
}

function pbLogActivity(pid, personName, actionType, actionDetail) {
    var entry = {
        configId: pbCurrentConfig ? pbCurrentConfig.id : '',
        configName: pbCurrentConfig ? pbCurrentConfig.name : '',
        peopleId: pid,
        personName: personName,
        actionType: actionType,
        actionDetail: actionDetail || ''
    };
    pbAjax({action: 'log_activity', log_data: JSON.stringify(entry)}, function(d) {
        // Silent - no UI feedback
    });
}

// Auto-save/restore work state per config
var pbAutoSaveTimer = null;
function pbAutoSaveWorkState() {
    clearTimeout(pbAutoSaveTimer);
    pbAutoSaveTimer = setTimeout(function() {
        if (!pbCurrentConfig || !pbCurrentConfig.id) return;
        var state = {
            processedMap: pbProcessedMap,
            deferredSet: pbDeferredSet,
            skippedSet: pbSkippedSet,
            viewMode: pbViewMode,
            currentIndex: pbCurrentIndex,
            savedAt: new Date().toISOString()
        };
        pbAjax({action: 'save_work_state', config_id: pbCurrentConfig.id, state_data: JSON.stringify(state)}, function(d) {
            // Silent save - no UI feedback needed
        });
    }, 500);  // Debounce 500ms so rapid clicks don't spam saves
}

function pbLoadWorkState(configId, callback) {
    pbAjax({action: 'load_work_state', config_id: configId}, function(d) {
        if (d.success && d.state) {
            pbProcessedMap = d.state.processedMap || {};
            pbDeferredSet = d.state.deferredSet || {};
            pbSkippedSet = d.state.skippedSet || {};
            if (d.state.viewMode) pbViewMode = d.state.viewMode;
            if (d.state.currentIndex) pbCurrentIndex = d.state.currentIndex;
        }
        if (callback) callback();
    });
}

// ============================================================
// CONTACT LOGGING
// ============================================================
function pbShowContactModal(pid) {
    document.getElementById('pb-contact-pid').value = pid;
    document.getElementById('pb-contact-notes').value = '';
    var sel = document.getElementById('pb-contact-keyword');
    sel.innerHTML = '<option value="">-- Select Method --</option>';
    var methods = pbSettings.contact_methods || [];
    for (var i = 0; i < methods.length; i++) {
        sel.innerHTML += '<option value="' + methods[i].keyword + '">' + methods[i].code + ' - ' + methods[i].label + '</option>';
    }
    sel.innerHTML += '<option value="Other">Other</option>';
    document.getElementById('pb-contact-modal').classList.add('active');
}

function pbCloseContactModal() { document.getElementById('pb-contact-modal').classList.remove('active'); }

function pbSubmitContact() {
    var pid = document.getElementById('pb-contact-pid').value;
    var keyword = document.getElementById('pb-contact-keyword').value;
    var notes = document.getElementById('pb-contact-notes').value;
    if (!notes) { alert('Please enter notes.'); return; }

    pbAjax({action: 'log_contact', people_id: pid, keyword: keyword, note_text: notes}, function(d) {
        if (d.success) {
            pbCloseContactModal();
            // Refresh contact efforts
            pbAjax({action: 'load_contact_efforts', people_ids: pid}, function(d2) {
                if (d2.success) {
                    var newMap = d2.contactMap || {};
                    for (var k in newMap) pbContactMap[k] = newMap[k];
                    pbApplyFilters();
                }
            });
        } else {
            alert('Error: ' + (d.message || 'Failed'));
        }
    });
}

// ============================================================
// SESSION MANAGEMENT
// ============================================================
function pbSaveCurrentSession() {
    if (!pbCurrentConfig) { alert('No configuration loaded. Open a configuration first.'); return; }
    var name = prompt('Session name:', pbCurrentConfig.name + ' - ' + new Date().toLocaleDateString());
    if (!name) return;

    var session = {
        id: pbCurrentSession ? pbCurrentSession.id : '',
        name: name,
        configId: pbCurrentConfig.id,
        configName: pbCurrentConfig.name,
        processedMap: pbProcessedMap,
        deferredSet: pbDeferredSet,
        skippedSet: pbSkippedSet,
        currentIndex: pbCurrentIndex,
        viewMode: pbViewMode
    };

    pbAjax({action: 'save_session', session_data: JSON.stringify(session)}, function(d) {
        if (d.success) {
            pbCurrentSession = d.session;
            alert('Session saved!');
        } else {
            alert('Error: ' + (d.message || 'Failed'));
        }
    });
}

function pbLoadActivityLog() {
    var configFilter = document.getElementById('pb-log-config-filter').value;
    var sourceFilter = document.getElementById('pb-log-source-filter').value;
    var groupFilter = document.getElementById('pb-log-group-filter').value;
    var el = document.getElementById('pb-activity-log');
    el.innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading...</div>';
    pbAjax({action: 'load_activity_log', config_id: configFilter, source_filter: sourceFilter, group_id: groupFilter}, function(d) {
        if (!d.success) { el.innerHTML = '<div class="pb-text-muted">Failed to load.</div>'; return; }
        var log = d.log || [];
        if (!log.length) {
            el.innerHTML = '<div class="pb-empty">'
                + '<div class="pb-empty-icon" style="font-size:3em;">&#128214;</div>'
                + '<h3 style="margin:8px 0 4px;color:var(--pb-primary);">A clean page. The first chapter is up to you.</h3>'
                + '<p style="max-width:520px;margin:6px auto 16px;line-height:1.5;">'
                + 'Every action your team takes shows up here: processing a prospect, logging a call, assigning to a group, sending a touch. '
                + 'It\\'s your audit trail and "what did I just do" rewind button.</p>'
                + '<button class="pb-btn pb-btn-primary" onclick="pbSwitchTab(\\'configs\\')">Open Prospect Management to start</button>'
                + '</div>';
            return;
        }
        var actionIcons = {processed: '&#10004;', deferred: '&#9200;', skipped: '&#10060;', reset: '&#8634;', contact: '&#128221;', group_effort: '&#128222;', assign_group: '&#128101;'};
        var actionColors = {processed: 'var(--pb-success)', deferred: 'var(--pb-warning)', skipped: 'var(--pb-muted)', reset: 'var(--pb-accent)', contact: 'var(--pb-primary)', group_effort: '#9b59b6', assign_group: '#3498db'};
        var html = '<div style="font-size:0.85em;">';
        var lastDate = '';
        for (var i = 0; i < log.length; i++) {
            var e = log[i];
            var day = (e.timestamp || '').substring(0, 10);
            if (day !== lastDate) {
                html += '<div style="padding:6px 0;font-weight:700;color:var(--pb-primary);border-bottom:1px solid var(--pb-border);margin-top:8px;">' + day + '</div>';
                lastDate = day;
            }
            var icon = actionIcons[e.actionType] || '&#8226;';
            var color = actionColors[e.actionType] || 'var(--pb-muted)';
            var badgeClass = e.actionType === 'processed' ? 'success' : e.actionType === 'deferred' ? 'warning' : e.actionType === 'group_effort' ? 'accent' : 'muted';
            html += '<div style="display:flex;gap:8px;align-items:center;padding:4px 0;border-bottom:1px solid #f5f5f5;flex-wrap:wrap;">';
            html += '<span style="color:' + color + ';font-size:1.1em;width:20px;text-align:center;">' + icon + '</span>';
            html += '<span style="min-width:50px;color:var(--pb-muted);font-size:0.8em;">' + (e.timestamp || '').substring(11, 16) + '</span>';
            html += '<span style="min-width:90px;font-weight:600;font-size:0.9em;">' + (e.userName || 'Unknown') + '</span>';
            html += '<span class="pb-badge pb-badge-' + badgeClass + '" style="font-size:0.7em;">' + (e.actionType || '').replace('_', ' ') + '</span>';
            html += '<a href="/Person2/' + e.peopleId + '" target="_blank" style="color:var(--pb-primary);">' + (e.personName || '') + '</a>';
            if (e.actionDetail) html += '<span class="pb-text-muted" style="font-size:0.85em;">&rarr; ' + e.actionDetail + '</span>';
            // Source chip
            if (e.configName) html += '<span class="pb-chip" style="margin-left:auto;">' + e.configName + '</span>';
            if (e.groupName) html += '<span class="pb-chip" style="margin-left:' + (e.configName ? '4px' : 'auto') + ';background:#9b59b622;color:#9b59b6;border:1px solid #9b59b644;">' + e.groupName + '</span>';
            html += '</div>';
        }
        html += '</div>';
        if (log.length >= 200) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:6px;">Showing last 200 entries</div>';
        el.innerHTML = html;
    });
}

function pbClearActivityLog() {
    if (!confirm('Clear all activity log entries? This cannot be undone.')) return;
    pbAjax({action: 'clear_activity_log'}, function(d) {
        if (d.success) pbLoadActivityLog();
    });
}

// ============================================================
// SETTINGS
// ============================================================
var pbKeywords = [];
var pbContactMethods = [];

function pbLoadSettings() {
    // Load keywords from DB first, then settings
    pbAjax({action: 'get_keywords'}, function(kd) {
        if (kd.success) pbKeywords = kd.keywords || [];
        pbAjax({action: 'load_settings'}, function(d) {
            if (d.success) {
                pbSettings = d.settings || {};
                pbContactMethods = pbSettings.contact_methods || [];
                pbRenderContactMethods();
                // Set other weight dropdown
                var owSel = document.getElementById('pb-other-weight');
                if (owSel) owSel.value = String(pbSettings.other_weight != null ? pbSettings.other_weight : 1);
                // Default message body for senders
                var dmbEl = document.getElementById('pb-default-message-body');
                if (dmbEl) dmbEl.value = pbSettings.default_message_body || '';
                // Render scorecard settings
                pbRenderScorecardSettings();
                // Pre-fill the scheduled-tasks toggle. Default ON when the
                // setting hasn't been persisted yet so the script behaves
                // the way it always has.
                var schedEl = document.getElementById('pb-sched-enabled');
                if (schedEl) {
                    var enabled = (pbSettings.sched_enabled !== false);
                    pbSettings.sched_enabled = enabled;
                    schedEl.checked = enabled;
                }
                // Refresh install status badges on both UI surfaces.
                pbCheckSchedInstall();
                // Re-render prospects if already loaded (settings may arrive after prospects)
                if (pbProspects.length) pbRenderCurrentView();
            }
        });
    });
}

// ---- Scheduled tasks install/check helpers ------------------------------
function pbSetSchedStatusText(installed, scriptName, slot) {
    var labels = [
        document.getElementById('pb-sched-status'),
        document.getElementById('pb-sched-status-settings')
    ];
    var msg = installed
        ? '✓ Installed in ' + (slot || 'ScheduledTasks') + ' as ' + scriptName
        : 'Not installed';
    var color = installed ? 'var(--pb-success)' : 'var(--pb-muted)';
    for (var i = 0; i < labels.length; i++) {
        if (!labels[i]) continue;
        labels[i].textContent = msg;
        labels[i].style.color = color;
    }
    // Toggle the install vs remove buttons on both surfaces.
    var pairs = [
        ['pb-sched-install-btn',  'pb-sched-uninstall-btn'],
        ['pb-sched-install-btn2', 'pb-sched-uninstall-btn2']
    ];
    for (var p = 0; p < pairs.length; p++) {
        var ins = document.getElementById(pairs[p][0]);
        var uni = document.getElementById(pairs[p][1]);
        if (ins) ins.style.display = installed ? 'none' : '';
        if (uni) uni.style.display = installed ? '' : 'none';
    }
}

function pbShowSchedResult(success, msg) {
    var boxes = [
        document.getElementById('pb-sched-result'),
        document.getElementById('pb-sched-result-settings')
    ];
    var bg = success ? '#e6f7ec' : '#fdecea';
    var fg = success ? 'var(--pb-success)' : 'var(--pb-danger)';
    for (var i = 0; i < boxes.length; i++) {
        if (!boxes[i]) continue;
        boxes[i].style.display = '';
        boxes[i].style.background = bg;
        boxes[i].style.color = fg;
        boxes[i].textContent = msg;
    }
}

function pbCheckSchedInstall() {
    pbAjax({action: 'check_sched_install'}, function(d) {
        if (!d || !d.success) return;
        pbSetSchedStatusText(!!d.installed, d.scriptName || 'TPxi_ProspectBuilder', d.contentSlot || 'ScheduledTasks');
    });
}

function pbInstallSched() {
    pbAjax({action: 'install_sched'}, function(d) {
        if (!d) { pbShowSchedResult(false, 'No response from server'); return; }
        pbShowSchedResult(!!d.success, d.message || (d.success ? 'Installed' : 'Failed'));
        if (d.success) pbCheckSchedInstall();
    });
}

function pbUninstallSched() {
    if (!confirm('Remove the Prospect Builder block from ScheduledTasks?\\n\\nOther scripts in ScheduledTasks will be untouched.')) return;
    pbAjax({action: 'uninstall_sched'}, function(d) {
        if (!d) { pbShowSchedResult(false, 'No response from server'); return; }
        pbShowSchedResult(!!d.success, d.message || (d.success ? 'Removed' : 'Failed'));
        if (d.success) pbCheckSchedInstall();
    });
}

function pbRenderContactMethods() {
    var tbody = document.getElementById('pb-method-rows');
    if (!pbContactMethods.length) {
        var html = '<tr><td colspan="4" style="padding:12px;color:var(--pb-muted);text-align:center;">No contact methods configured. Click "+ Add Contact Method" to get started.</td></tr>';
        html += '<tr style="border-bottom:1px solid #eee;background:var(--pb-light);opacity:0.7;">';
        html += '<td style="padding:6px 8px;"><span style="font-weight:700;color:var(--pb-muted);">O</span></td>';
        html += '<td style="padding:6px 8px;color:var(--pb-muted);">Other</td>';
        html += '<td style="padding:6px 8px;color:var(--pb-muted);font-style:italic;font-size:0.85em;">Catch-all for unmatched task/note activity</td>';
        html += '<td></td></tr>';
        tbody.innerHTML = html;
        return;
    }
    var html = '';
    for (var i = 0; i < pbContactMethods.length; i++) {
        var m = pbContactMethods[i];
        html += '<tr style="border-bottom:1px solid #eee;">';
        html += '<td style="padding:6px 8px;"><input class="pb-input" style="width:50px;text-align:center;font-weight:700;" value="' + (m.code || '') + '" data-idx="' + i + '" data-field="code" onchange="pbContactMethods[parseInt(this.dataset.idx)].code=this.value" maxlength="2"></td>';
        html += '<td style="padding:6px 8px;"><input class="pb-input" style="width:120px;" value="' + (m.label || '') + '" data-idx="' + i + '" data-field="label" onchange="pbContactMethods[parseInt(this.dataset.idx)].label=this.value"></td>';
        html += '<td style="padding:6px 8px;"><select class="pb-select" data-idx="' + i + '" onchange="pbSetKeyword(parseInt(this.dataset.idx),this)">';
        html += '<option value="">-- Select Keyword --</option>';
        for (var j = 0; j < pbKeywords.length; j++) {
            var kw = pbKeywords[j];
            var sel = (m.keywordId && parseInt(m.keywordId) === kw.keywordId) ? ' selected' : '';
            html += '<option value="' + kw.keywordId + '"' + sel + '>' + kw.description + '</option>';
        }
        html += '</select></td>';
        html += '<td style="padding:6px 8px;text-align:center;"><button class="pb-btn pb-btn-danger pb-btn-sm" data-idx="' + i + '" onclick="pbRemoveContactMethod(parseInt(this.dataset.idx))">Remove</button></td>';
        html += '</tr>';
    }
    // Fixed O=Other row
    html += '<tr style="border-bottom:1px solid #eee;background:var(--pb-light);opacity:0.7;">';
    html += '<td style="padding:6px 8px;"><span style="font-weight:700;color:var(--pb-muted);">O</span></td>';
    html += '<td style="padding:6px 8px;color:var(--pb-muted);">Other</td>';
    html += '<td style="padding:6px 8px;color:var(--pb-muted);font-style:italic;font-size:0.85em;">Catch-all for unmatched task/note activity</td>';
    html += '<td></td></tr>';
    tbody.innerHTML = html;
}

function pbSetKeyword(idx, sel) {
    var opt = sel.options[sel.selectedIndex];
    if (opt.value) {
        pbContactMethods[idx].keywordId = parseInt(opt.value);
        pbContactMethods[idx].keyword = opt.text;
    } else {
        pbContactMethods[idx].keywordId = null;
        pbContactMethods[idx].keyword = '';
    }
}

function pbAddContactMethod() {
    pbContactMethods.push({code: '', label: '', keyword: '', keywordId: null});
    pbRenderContactMethods();
}

function pbRemoveContactMethod(idx) {
    pbContactMethods.splice(idx, 1);
    pbRenderContactMethods();
}

function pbSaveSettings() {
    pbSettings.contact_methods = pbContactMethods;
    pbCollectScorecardSettings();
    var dmbEl = document.getElementById('pb-default-message-body');
    if (dmbEl) pbSettings.default_message_body = dmbEl.value;
    pbAjax({action: 'save_settings', settings_data: JSON.stringify(pbSettings)}, function(d) {
        var el = document.getElementById('pb-settings-status');
        if (d.success) {
            el.textContent = 'Saved!';
            el.style.color = 'var(--pb-success)';
        } else {
            el.textContent = 'Error: ' + (d.message || 'Failed');
            el.style.color = 'var(--pb-danger)';
        }
        setTimeout(function() { el.textContent = ''; }, 3000);
    });
}

// ============================================================
// PROSPECT SCORECARD
// ============================================================
var pbScorecardDefaults = {
    enabled: false,
    factors: {
        contact_efforts:    {enabled: true, weight: 20, label: 'Contact Efforts', description: 'Weighted outreach efforts (phone, email, visit, etc.)'},
        attend_recency:     {enabled: true, weight: 20, label: 'Attendance Recency', description: 'How recently they last attended (7d=100, 30d=70, 90d+=low)'},
        attend_frequency:   {enabled: true, weight: 15, label: 'Attendance Frequency', description: 'Number of attendances in last 90 days'},
        involvements:       {enabled: true, weight: 10, label: 'Involvements', description: 'Number of active organization memberships'},
        enrollment_recency: {enabled: true, weight: 10, label: 'Enrollment Recency', description: 'How recently enrolled (newer = higher priority)'},
        family_engaged:     {enabled: true, weight: 10, label: 'Family Engaged', description: 'Another family member is already active in an organization'},
        member_status:      {enabled: true, weight: 10, label: 'Member Status', description: 'Prospect/Visitor score higher than existing Members'},
        serving_roles:      {enabled: true, weight: 15, label: 'Serving Roles', description: 'Number of active serving/leadership roles (volunteer, leader, etc.)'},
        tasknote_activity:  {enabled: true, weight: 5,  label: 'TaskNote Activity', description: 'Recent TaskNote entries about this person'}
    }
};

function pbGetScorecard() {
    if (!pbSettings.scorecard) pbSettings.scorecard = JSON.parse(JSON.stringify(pbScorecardDefaults));
    // Ensure all default factors exist
    var defs = pbScorecardDefaults.factors;
    for (var k in defs) {
        if (!pbSettings.scorecard.factors[k]) {
            pbSettings.scorecard.factors[k] = JSON.parse(JSON.stringify(defs[k]));
        }
        // Ensure label and description exist
        if (!pbSettings.scorecard.factors[k].label) pbSettings.scorecard.factors[k].label = defs[k].label;
        if (!pbSettings.scorecard.factors[k].description) pbSettings.scorecard.factors[k].description = defs[k].description;
    }
    return pbSettings.scorecard;
}

function pbRenderScorecardSettings() {
    var sc = pbGetScorecard();
    var el = document.getElementById('pb-scorecard-enabled');
    if (el) el.checked = sc.enabled;
    var tbody = document.getElementById('pb-scorecard-rows');
    if (!tbody) return;
    var html = '';
    var factorOrder = ['contact_efforts','attend_recency','attend_frequency','involvements','serving_roles','enrollment_recency','family_engaged','member_status','tasknote_activity'];
    for (var i = 0; i < factorOrder.length; i++) {
        var key = factorOrder[i];
        var f = sc.factors[key];
        if (!f) continue;
        html += '<tr style="border-bottom:1px solid var(--pb-border);">';
        html += '<td style="padding:6px 8px;text-align:center;"><input type="checkbox" data-factor="' + key + '"' + (f.enabled ? ' checked' : '') + ' onchange="pbScUpdateFactor(this.dataset.factor,&quot;enabled&quot;,this.checked)"></td>';
        html += '<td style="padding:6px 8px;font-weight:600;">' + f.label + '</td>';
        html += '<td style="padding:6px 8px;"><input type="number" class="pb-input" style="width:60px;text-align:center;" min="0" max="100" value="' + (f.weight || 0) + '" data-factor="' + key + '" onchange="pbScUpdateFactor(this.dataset.factor,&quot;weight&quot;,parseInt(this.value)||0)"></td>';
        html += '<td style="padding:6px 8px;color:var(--pb-muted);font-size:0.8em;">' + f.description + '</td>';
        html += '</tr>';
    }
    tbody.innerHTML = html;
}

function pbScUpdateFactor(factorKey, field, value) {
    var sc = pbGetScorecard();
    if (sc.factors[factorKey]) sc.factors[factorKey][field] = value;
}

function pbCollectScorecardSettings() {
    var el = document.getElementById('pb-scorecard-enabled');
    if (el && pbSettings.scorecard) {
        pbSettings.scorecard.enabled = el.checked;
    }
}

function pbGetScoreColor(score) {
    if (score >= 70) return 'var(--pb-success)';
    if (score >= 40) return 'var(--pb-warning)';
    return 'var(--pb-danger)';
}

function pbBuildScoreDetail(scoreData) {
    if (!scoreData || !scoreData.breakdown) return '';
    var sc = pbGetScorecard();
    var enabled = [];
    var totalWeight = 0;
    var factorOrder = ['contact_efforts','attend_recency','attend_frequency','involvements','serving_roles','enrollment_recency','family_engaged','member_status','tasknote_activity'];
    for (var i = 0; i < factorOrder.length; i++) {
        var k = factorOrder[i];
        var f = sc.factors[k];
        if (f && f.enabled && scoreData.breakdown[k] != null) {
            enabled.push({key: k, label: f.label || k, weight: f.weight || 0, sub: scoreData.breakdown[k]});
            totalWeight += (f.weight || 0);
        }
    }
    if (!totalWeight) totalWeight = 1;
    var html = '<table style="width:100%;border-collapse:collapse;font-size:0.8em;">';
    html += '<tr style="border-bottom:2px solid #ddd;"><th style="padding:3px 6px;text-align:left;">Factor</th><th style="padding:3px 6px;text-align:center;">Score</th><th style="padding:3px 6px;text-align:center;">Weight</th><th style="padding:3px 6px;text-align:center;">Contribution</th></tr>';
    for (var j = 0; j < enabled.length; j++) {
        var e = enabled[j];
        var pct = Math.round(e.weight / totalWeight * 100);
        var contrib = Math.round(e.sub * e.weight / totalWeight);
        var barColor = e.sub >= 70 ? 'var(--pb-success)' : e.sub >= 40 ? 'var(--pb-warning)' : 'var(--pb-danger)';
        html += '<tr style="border-bottom:1px solid #eee;">';
        html += '<td style="padding:3px 6px;">' + e.label + '</td>';
        html += '<td style="padding:3px 6px;text-align:center;"><div style="background:#eee;border-radius:3px;overflow:hidden;height:14px;position:relative;min-width:50px;"><div style="position:absolute;left:0;top:0;height:100%;width:' + e.sub + '%;background:' + barColor + ';border-radius:3px;"></div><span style="position:relative;font-size:0.85em;font-weight:700;line-height:14px;">' + e.sub + '</span></div></td>';
        html += '<td style="padding:3px 6px;text-align:center;color:#888;">' + pct + '%</td>';
        html += '<td style="padding:3px 6px;text-align:center;font-weight:700;">+' + contrib + '</td>';
        html += '</tr>';
    }
    html += '<tr style="border-top:2px solid #333;"><td style="padding:4px 6px;font-weight:700;" colspan="3">Total Score</td><td style="padding:4px 6px;text-align:center;font-weight:700;font-size:1.1em;">' + scoreData.score + '</td></tr>';
    html += '</table>';
    html += '<div style="margin-top:6px;font-size:0.75em;color:#888;">Each factor scores 0-100. The weight determines what % of the total it contributes. Contribution = Score x Weight%.</div>';
    return html;
}

function pbShowScorePopup(pid, evt) {
    evt.stopPropagation();
    var s = pbScoreMap[String(pid)];
    if (!s) return;
    // Remove any existing popup
    var old = document.getElementById('pb-score-popup');
    if (old) old.remove();
    var popup = document.createElement('div');
    popup.id = 'pb-score-popup';
    popup.style.cssText = 'position:fixed;z-index:10000;background:#fff;border:1px solid #ccc;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,0.2);padding:12px;max-width:380px;min-width:300px;';
    // Position near click
    var x = Math.min(evt.clientX, window.innerWidth - 400);
    var y = Math.min(evt.clientY + 10, window.innerHeight - 300);
    popup.style.left = x + 'px';
    popup.style.top = y + 'px';
    var color = pbGetScoreColor(s.score);
    popup.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;"><span style="font-weight:700;font-size:1em;">Score Breakdown</span><span style="background:' + color + ';color:#fff;padding:2px 10px;border-radius:10px;font-weight:700;">' + s.score + ' / 100</span></div>' + pbBuildScoreDetail(s) + '<div style="text-align:right;margin-top:8px;"><button onclick="document.getElementById(&quot;pb-score-popup&quot;).remove()" style="padding:3px 12px;border:1px solid #ccc;border-radius:4px;background:#f5f5f5;cursor:pointer;font-size:0.85em;">Close</button></div>';
    document.body.appendChild(popup);
    // Close on outside click
    setTimeout(function() {
        document.addEventListener('click', function closePopup(e) {
            var p = document.getElementById('pb-score-popup');
            if (p && !p.contains(e.target)) { p.remove(); document.removeEventListener('click', closePopup); }
        });
    }, 100);
}

function pbBuildScoreBadge(pid) {
    if (!pbSettings.scorecard || !pbSettings.scorecard.enabled) return '';
    var s = pbScoreMap[String(pid)];
    if (!s) return '';
    var score = s.score;
    var color = pbGetScoreColor(score);
    return ' <span data-pid="' + pid + '" onclick="event.stopPropagation();pbShowScorePopup(parseInt(this.dataset.pid),event)" style="display:inline-block;background:' + color + ';color:#fff;padding:1px 7px;border-radius:10px;font-size:0.75em;font-weight:700;cursor:pointer;">' + score + '</span>';
}

function pbBuildScoreTooltip(scoreData) {
    if (!scoreData || !scoreData.breakdown) return '';
    var sc = pbGetScorecard();
    var enabled = [];
    var totalWeight = 0;
    var factorOrder = ['contact_efforts','attend_recency','attend_frequency','involvements','serving_roles','enrollment_recency','family_engaged','member_status','tasknote_activity'];
    for (var i = 0; i < factorOrder.length; i++) {
        var k = factorOrder[i];
        var f = sc.factors[k];
        if (f && f.enabled && scoreData.breakdown[k] != null) {
            enabled.push({label: f.label || k, weight: f.weight || 0, sub: scoreData.breakdown[k]});
            totalWeight += (f.weight || 0);
        }
    }
    if (!totalWeight) totalWeight = 1;
    var parts = [];
    for (var j = 0; j < enabled.length; j++) {
        var e = enabled[j];
        var pct = Math.round(e.weight / totalWeight * 100);
        var contrib = Math.round(e.sub * e.weight / totalWeight);
        parts.push(e.label + ': ' + e.sub + ' x ' + pct + '% = +' + contrib);
    }
    parts.push('Total: ' + scoreData.score);
    return parts.join('&#10;');
}

function pbComputeScores(pids, callback) {
    var sc = pbGetScorecard();
    if (!sc.enabled) return;
    // Build member data from loaded prospects (workspace or group)
    var memberData = {};
    var allProspects = pbProspects.length > 0 ? pbProspects : pbGrpDetailProspects;
    for (var i = 0; i < allProspects.length; i++) {
        var p = allProspects[i];
        var pk = String(p.PeopleId);
        var cm = pbContactMap[pk] || pbGrpContactMap[pk] || {};
        memberData[pk] = {
            memberStatus: p.MemberStatus || '',
            weightedTotal: cm.weightedTotal || cm.total || 0,
            taskNoteTotal: cm.total || 0
        };
    }
    pbAjax({
        action: 'compute_scores',
        people_ids: pids,
        scorecard_config: JSON.stringify(sc),
        member_data: JSON.stringify(memberData)
    }, function(d) {
        if (d.success) {
            var newScores = d.scores || {};
            for (var k in newScores) pbScoreMap[k] = newScores[k];
            pbApplyFilters();
            if (callback) callback();
        }
    });
}

// ============================================================
// PRINT (Popup Window)
// ============================================================
window.pbPrintReport = function() {
    if (!pbCurrentConfig || !pbFilteredProspects.length) { alert('No prospects to print.'); return; }
    var css = '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}';
    css += 'body{margin:0;padding:20px;font-family:Segoe UI,sans-serif;color:#333;font-size:12px}';
    css += 'h2{color:#2c3e50;border-bottom:2px solid #2c3e50;padding-bottom:5px}';
    css += 'table{width:100%;border-collapse:collapse;margin-top:10px}';
    css += 'th{background:#2c3e50;color:#fff;padding:6px 8px;text-align:left;font-size:11px}';
    css += 'td{padding:5px 8px;border-bottom:1px solid #ddd;font-size:11px}';
    css += 'tr:nth-child(even){background:#f5f6fa}';
    css += '.badge{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700}';
    css += '.processed{background:#d5f5e3;color:#27ae60}';
    css += '.pending{background:#d6eaf8;color:#2980b9}';

    var fields = pbCurrentConfig.displayFields || [];
    var htmlContent = '<h2>' + pbCurrentConfig.name + ' - Prospect Report</h2>';
    htmlContent += '<p>Generated: ' + new Date().toLocaleString() + ' | Total: ' + pbFilteredProspects.length + '</p>';
    htmlContent += '<table><thead><tr><th>#</th>';
    for (var i = 0; i < fields.length; i++) htmlContent += '<th>' + fields[i].label + '</th>';
    htmlContent += '<th>Status</th></tr></thead><tbody>';
    for (var j = 0; j < pbFilteredProspects.length; j++) {
        var p = pbFilteredProspects[j];
        var status = pbProcessedMap[p.PeopleId] ? 'processed' : pbDeferredSet[p.PeopleId] ? 'deferred' : 'pending';
        htmlContent += '<tr><td>' + (j+1) + '</td>';
        for (var k = 0; k < fields.length; k++) htmlContent += '<td>' + pbGetFieldValue(p, fields[k]) + '</td>';
        htmlContent += '<td><span class="badge ' + status + '">' + status + '</span></td></tr>';
    }
    htmlContent += '</tbody></table>';

    var pw = window.open('', '_blank');
    if (!pw) { alert('Popup blocked - please allow popups for this site.'); return; }
    pw.document.write('<!DOCTYPE html><html><head><title>Prospect Report</title><style>' + css + '</style></head><body>');
    pw.document.write(htmlContent);
    pw.document.write('</body></html>');
    pw.document.close();
    pw.focus();
    setTimeout(function() { pw.print(); }, 300);
};

// ============================================================
// JOURNEY HELPERS
// ============================================================
function pbGetEngagementLevel(score) {
    if (score >= 80) return 'Highly engaged';
    if (score >= 60) return 'Engaged';
    if (score >= 40) return 'Moderately engaged';
    if (score >= 20) return 'Low engagement';
    return 'Not engaged';
}

// NOTE: pbRenderJourneyTimeline and pbRenderFamilyTimeline removed -
// Journey is now inline in the detail view via pbRenderPersonDetail
// and pbRenderInlineFamilyTimeline

// ============================================================
// PROSPECT GROUPS
// ============================================================
var pbGrpEffortLabels = {
    'phone': 'Phone Call', 'email': 'Email', 'text': 'Text', 'visit': 'Home Visit',
    'event_invite': 'Event Invite', 'meeting': 'Meeting', 'other': 'Other'
};
var pbGrpResultLabels = {
    'connected': 'Connected', 'no_answer': 'No Answer', 'left_message': 'Left Message',
    'interested': 'Interested', 'declined': 'Declined', 'follow_up': 'Follow-Up', 'other': 'Other'
};
var pbGrpEffortIcons = {
    'phone': '&#128222;', 'email': '&#9993;', 'text': '&#128172;', 'visit': '&#127968;',
    'event_invite': '&#127881;', 'meeting': '&#129309;', 'other': '&#128196;'
};

function pbGrpInitTab() {
    // Only pre-populate the dropdown if we already have group data loaded
    // (e.g. from pbInit's pbGrpLoadAll). Otherwise wait for pbGrpLoadData's
    // AJAX callback -- pre-populating an empty array shows "-- Select Group --"
    // and looks broken until the user changes source.
    if (pbGrpGroups && pbGrpGroups.length > 0
        && typeof pbGrpHealthPopulateGroups === 'function') {
        pbGrpHealthPopulateGroups();
    }
    pbGrpLoadData();
    if (pbGrpHealthPrograms.length === 0 && typeof pbGrpHealthLoadPrograms === 'function') pbGrpHealthLoadPrograms();
}

// Select a group in the By Involvement source dropdown and load its health.
// Bound to group card clicks now that the detail view is gone.
function pbGrpLoadHealthForGroup(groupId) {
    var srcSel = document.getElementById('pb-grp-health-source');
    var grpSel = document.getElementById('pb-grp-health-group');
    if (!srcSel || !grpSel) return;
    srcSel.value = 'group';
    pbGrpHealthSourceChanged();
    // Ensure the group dropdown is populated before selecting
    if (!grpSel.options.length || grpSel.options.length === 1) pbGrpHealthPopulateGroups();
    grpSel.value = groupId;
    pbGrpLoadHealth();
    // Scroll the health view into view
    var healthHdr = document.getElementById('pb-grp-health-summary');
    if (healthHdr && healthHdr.scrollIntoView) healthHdr.scrollIntoView({behavior: 'smooth', block: 'start'});
}

function pbGrpLoadData() {
    // Show loading state immediately so user knows data is being fetched
    var grid = document.getElementById('pb-grp-grid');
    var empty = document.getElementById('pb-grp-empty');
    if (grid) { grid.style.display = ''; grid.innerHTML = '<div class="pb-loading" style="grid-column:1/-1;text-align:center;padding:30px;"><span class="pb-spin"></span> Loading groups...</div>'; }
    if (empty) empty.style.display = 'none';
    pbAjax({action: 'load_prospect_groups'}, function(d) {
        if (!d.success) return;
        pbGrpGroups = d.groups || [];
        pbGrpAssignments = d.assignments || {};
        pbGrpEfforts = d.efforts || [];
        pbGrpChanges = d.changeLog || [];
        // Stats omitted from group config cards (used to be slow on large groups).
        // Click a card to load the health view for full prospect detail.
        pbGrpRenderOverview();
    });
}

function pbGrpRenderOverview() {
    // Refresh dropdowns + button state (still used in Manual mode).
    if (typeof pbGrpHealthPopulateGroups === 'function') pbGrpHealthPopulateGroups();
    pbGrpUpdateInlineButtons();

    var noGroups = !pbGrpGroups || pbGrpGroups.length === 0;

    var empty    = document.getElementById('pb-grp-empty-state');
    var cardWrap = document.getElementById('pb-grp-card-grid-wrap');
    var srcArea  = document.getElementById('pb-grp-source-area');

    if (empty)    empty.style.display    = noGroups ? '' : 'none';
    if (cardWrap) cardWrap.style.display = noGroups ? 'none' : '';
    // Source/Group dropdowns hidden by default when card grid is visible.
    // Toggle to reveal via the "Manual source >>" button (pbGrpToggleAdvanced).
    if (srcArea)  srcArea.style.display  = 'none';

    // Hide summary/content when empty so the page reads cleanly.
    var ids = ['pb-grp-health-summary', 'pb-grp-health-content'];
    for (var i = 0; i < ids.length; i++) {
        var el = document.getElementById(ids[i]);
        if (el) el.style.display = noGroups ? 'none' : '';
    }

    if (!noGroups) pbGrpRenderCards();
}

function pbGrpRenderCards() {
    var el = document.getElementById('pb-grp-card-grid');
    if (!el) return;
    // Lift the existing config-grid pattern. Click opens the group's health view.
    var html = '';
    for (var i = 0; i < pbGrpGroups.length; i++) {
        var g = pbGrpGroups[i];
        var name = g.name || 'Untitled group';
        var desc = g.description || '';
        var gid = g.id || '';
        html += '<div class="pb-grp-card" data-gid="' + gid + '" ' +
                'onclick="pbGrpLoadHealthForGroup(this.dataset.gid)" ' +
                'style="cursor:pointer;">';
        html +=   '<div class="pb-grp-card-title">' + pbEsc(name) + '</div>';
        if (desc) html += '<div class="pb-grp-card-desc">' + pbEsc(desc) + '</div>';
        html +=   '<div class="pb-grp-stats" style="margin-top:6px;">';
        html +=     '<button class="pb-btn pb-btn-sm pb-btn-primary" data-gid="' + gid + '" ' +
                      'onclick="event.stopPropagation();pbGrpLoadHealthForGroup(this.dataset.gid)">' +
                      'Open</button>';
        html +=     '<button class="pb-btn pb-btn-sm" data-gid="' + gid + '" ' +
                      'onclick="event.stopPropagation();pbGrpShowModal(pbGrpFindGroup(this.dataset.gid))">' +
                      'Edit</button>';
        html +=   '</div>';
        html += '</div>';
    }
    el.innerHTML = html;
}

function pbGrpToggleAdvanced() {
    var src = document.getElementById('pb-grp-source-area');
    var btn = document.getElementById('pb-grp-advanced-btn');
    if (!src) return;
    var shown = (src.style.display !== 'none');
    src.style.display = shown ? 'none' : '';
    if (btn) btn.textContent = shown ? 'Manual source »' : 'Hide manual source';
}

// Enable/disable the Edit button based on dropdown selection. Delete moved
// inside the Edit modal so users can't misclick it from the toolbar.
function pbGrpUpdateInlineButtons() {
    var sel = document.getElementById('pb-grp-health-group');
    var editBtn = document.getElementById('pb-grp-inline-edit');
    if (!sel || !editBtn) return;
    var hasSelection = !!sel.value;
    editBtn.disabled = !hasSelection;
}

function pbGrpInlineEdit() {
    var sel = document.getElementById('pb-grp-health-group');
    if (!sel || !sel.value) return;
    var g = pbGrpFindGroup(sel.value);
    if (!g) { alert('Group not found.'); return; }
    pbGrpShowModal(g);
}

// Delete from inside the edit modal. Confirms, then routes to the existing
// pbGrpDelete so the server-side action stays identical. After a successful
// delete the modal closes and the health view refreshes.
function pbGrpModalDelete() {
    var id = document.getElementById('pb-grp-edit-id').value;
    if (!id) return;
    pbGrpDelete(id, function() {
        pbGrpCloseModal();
    });
}

function pbGrpFindGroup(id) {
    for (var i = 0; i < pbGrpGroups.length; i++) {
        if (pbGrpGroups[i].id === id) return pbGrpGroups[i];
    }
    return null;
}

// --- Group CRUD ---
var pbGrpPrograms = [];
var pbGrpDivisions = [];
var pbGrpSearchTimer = null;

function pbGrpShowModal(group) {
    var modal = document.getElementById('pb-grp-modal');
    document.getElementById('pb-grp-modal-title').textContent = group ? 'Edit Group' : 'New Group';
    document.getElementById('pb-grp-edit-id').value = group ? (group.id || '') : '';
    document.getElementById('pb-grp-edit-name').value = group ? (group.name || '') : '';
    document.getElementById('pb-grp-edit-desc').value = group ? (group.description || '') : '';
    // Reveal the Delete button only when editing an existing group; new
    // groups have nothing to delete yet.
    var delBtn = document.getElementById('pb-grp-modal-delete');
    if (delBtn) delBtn.style.display = group ? 'inline-block' : 'none';

    // Ministry level
    var level = group ? (group.level || 'program') : 'program';
    document.getElementById('pb-grp-edit-level').value = level;
    document.getElementById('pb-grp-edit-org-id').value = group ? (group.orgId || '') : '';
    document.getElementById('pb-grp-inv-selected').textContent = group && group.orgName ? group.orgName + ' (#' + group.orgId + ')' : '';
    document.getElementById('pb-grp-edit-min-days').value = group ? (group.minEnrollDays || 0) : 0;
    document.getElementById('pb-grp-edit-stale-days').value = group ? (group.minStaleDays || 0) : 0;

    // Set color
    var color = group ? (group.color || '#3498db') : '#3498db';
    var swatches = document.querySelectorAll('.pb-grp-color-swatch');
    for (var i = 0; i < swatches.length; i++) {
        swatches[i].classList.toggle('selected', swatches[i].dataset.color === color);
    }

    // Load member type checkboxes
    var selectedMTs = group ? (group.memberTypes || []) : [];
    pbGrpRenderMemberTypes(selectedMTs);

    // Load target actions
    pbGrpConfigActions = group ? (group.targetActions || []).slice() : [];
    pbGrpRenderActions();
    document.getElementById('pb-grp-add-action-type').value = 'tag';
    document.getElementById('pb-grp-action-tag-fields').style.display = '';
    document.getElementById('pb-grp-action-inv-fields').style.display = 'none';

    // Load programs then set values
    pbGrpLoadPrograms(function() {
        if (group && group.programId) {
            document.getElementById('pb-grp-edit-program').value = group.programId;
            if (level === 'division' || level === 'involvement') {
                pbGrpProgramChanged(function() {
                    if (group.divisionId) {
                        document.getElementById('pb-grp-edit-division').value = group.divisionId;
                    }
                });
            }
        }
        pbGrpLevelChanged();
    });

    modal.classList.add('active');
}
function pbGrpCloseModal() { document.getElementById('pb-grp-modal').classList.remove('active'); }
function pbGrpSelectColor(el) {
    document.querySelectorAll('.pb-grp-color-swatch').forEach(function(s) { s.classList.remove('selected'); });
    el.classList.add('selected');
}

function pbGrpLevelChanged() {
    var level = document.getElementById('pb-grp-edit-level').value;
    document.getElementById('pb-grp-div-row').style.display = (level === 'division' || level === 'involvement') ? '' : 'none';
    document.getElementById('pb-grp-inv-row').style.display = level === 'involvement' ? '' : 'none';
}

function pbGrpLoadPrograms(callback) {
    if (pbGrpPrograms.length) {
        pbGrpRenderProgramDropdown();
        if (callback) callback();
        return;
    }
    pbAjax({action: 'get_programs'}, function(d) {
        if (d.success) {
            pbGrpPrograms = d.programs || [];
            pbGrpRenderProgramDropdown();
        }
        if (callback) callback();
    });
}

function pbGrpRenderProgramDropdown() {
    var sel = document.getElementById('pb-grp-edit-program');
    var current = sel.value;
    sel.innerHTML = '<option value="">-- Select Program --</option>';
    for (var i = 0; i < pbGrpPrograms.length; i++) {
        var p = pbGrpPrograms[i];
        sel.innerHTML += '<option value="' + p.id + '">' + p.name + '</option>';
    }
    if (current) sel.value = current;
}

function pbGrpProgramChanged(callback) {
    var progId = document.getElementById('pb-grp-edit-program').value;
    var divSel = document.getElementById('pb-grp-edit-division');
    divSel.innerHTML = '<option value="">-- Select Division --</option>';
    if (!progId) { if (callback) callback(); return; }
    pbAjax({action: 'get_divisions', program_id: progId}, function(d) {
        if (d.success) {
            pbGrpDivisions = d.divisions || [];
            for (var i = 0; i < pbGrpDivisions.length; i++) {
                divSel.innerHTML += '<option value="' + pbGrpDivisions[i].id + '">' + pbGrpDivisions[i].name + '</option>';
            }
        }
        if (callback) callback();
    });
}

function pbGrpSearchInvolvements(term) {
    clearTimeout(pbGrpSearchTimer);
    if (term.length < 2) { document.getElementById('pb-grp-inv-results').style.display = 'none'; return; }
    pbGrpSearchTimer = setTimeout(function() {
        pbAjax({action: 'search_involvements', search_term: term}, function(d) {
            if (!d.success) return;
            var el = document.getElementById('pb-grp-inv-results');
            var html = '';
            for (var i = 0; i < (d.involvements || []).length; i++) {
                var inv = d.involvements[i];
                html += '<div class="pb-search-item" data-orgid="' + inv.orgId + '" data-name="' + (inv.name || '').replace(/"/g, '&quot;') + '" onclick="pbGrpSelectInvolvement(parseInt(this.dataset.orgid), this.dataset.name)">';
                html += '<strong>' + inv.name + '</strong> (#' + inv.orgId + ')';
                html += '<br><span class="pb-text-muted">' + (inv.program || '') + ' / ' + (inv.division || '') + '</span>';
                html += '</div>';
            }
            el.innerHTML = html || '<div class="pb-search-item pb-text-muted">No results</div>';
            el.style.display = '';
        });
    }, 300);
}

function pbGrpSelectInvolvement(orgId, name) {
    document.getElementById('pb-grp-edit-org-id').value = orgId;
    document.getElementById('pb-grp-inv-selected').textContent = name + ' (#' + orgId + ')';
    document.getElementById('pb-grp-inv-results').style.display = 'none';
    document.getElementById('pb-grp-edit-inv-search').value = '';
}

// --- Group Target Actions ---
var pbGrpConfigActions = [];
var pbGrpActionSearchTimer = null;

function pbGrpToggleActionFields() {
    var type = document.getElementById('pb-grp-add-action-type').value;
    document.getElementById('pb-grp-action-tag-fields').style.display = type === 'tag' ? '' : 'none';
    document.getElementById('pb-grp-action-inv-fields').style.display = type === 'involvement' ? '' : 'none';
}

function pbGrpAddTagAction() {
    var name = document.getElementById('pb-grp-action-tag-name').value.trim();
    if (!name) { alert('Enter a tag name.'); return; }
    pbGrpConfigActions.push({pb_type: 'tag', tagName: name, label: 'Tag: ' + name});
    pbGrpRenderActions();
    document.getElementById('pb-grp-action-tag-name').value = '';
}

function pbGrpActionSearchInv(term) {
    clearTimeout(pbGrpActionSearchTimer);
    if (term.length < 2) { document.getElementById('pb-grp-action-inv-results').innerHTML = ''; return; }
    pbGrpActionSearchTimer = setTimeout(function() {
        pbAjax({action: 'search_involvements', search_term: term}, function(d) {
            if (!d.success) return;
            var el = document.getElementById('pb-grp-action-inv-results');
            var html = '';
            for (var i = 0; i < (d.involvements || []).length; i++) {
                var inv = d.involvements[i];
                html += '<div class="pb-search-item" data-orgid="' + inv.orgId + '" data-name="' + (inv.name || '').replace(/"/g, '&quot;') + '" onclick="pbGrpAddInvAction(parseInt(this.dataset.orgid), this.dataset.name)">';
                html += '<strong>' + inv.name + '</strong> (#' + inv.orgId + ')';
                html += '<br><span class="pb-text-muted">' + (inv.program || '') + ' / ' + (inv.division || '') + '</span>';
                html += '</div>';
            }
            el.innerHTML = html || '<div class="pb-search-item pb-text-muted">No results</div>';
        });
    }, 300);
}

function pbGrpAddInvAction(orgId, name) {
    var mtId = document.getElementById('pb-grp-action-member-type').value;
    var mtLabel = document.getElementById('pb-grp-action-member-type').options[document.getElementById('pb-grp-action-member-type').selectedIndex].text;
    pbGrpConfigActions.push({pb_type: 'involvement', orgId: orgId, orgName: name, memberTypeId: parseInt(mtId), label: name + ' as ' + mtLabel});
    pbGrpRenderActions();
    document.getElementById('pb-grp-action-inv-search').value = '';
    document.getElementById('pb-grp-action-inv-results').innerHTML = '';
}

function pbGrpRenderActions() {
    var el = document.getElementById('pb-grp-actions-list');
    if (!pbGrpConfigActions.length) {
        el.innerHTML = '<div class="pb-text-muted pb-text-sm">No target actions defined yet.</div>';
        return;
    }
    var html = '';
    for (var i = 0; i < pbGrpConfigActions.length; i++) {
        var a = pbGrpConfigActions[i];
        html += '<div style="display:flex;justify-content:space-between;align-items:center;padding:4px 8px;border:1px solid var(--pb-border);border-radius:var(--pb-radius);margin-bottom:3px;font-size:0.85em;">';
        html += '<span>' + (a.pb_type === 'tag' ? '&#127991; ' : '&#128101; ') + a.label + '</span>';
        html += '<span style="cursor:pointer;color:var(--pb-danger);font-weight:700;" data-idx="' + i + '" onclick="pbGrpRemoveAction(parseInt(this.dataset.idx))">&times;</span>';
        html += '</div>';
    }
    el.innerHTML = html;
}

function pbGrpRemoveAction(idx) {
    pbGrpConfigActions.splice(idx, 1);
    pbGrpRenderActions();
}

function pbGrpRenderMemberTypes(selectedIds) {
    var el = document.getElementById('pb-grp-mt-list');
    if (!pbMemberTypesFromDB.length) {
        pbAjax({action: 'get_member_types'}, function(d) {
            if (d.success) {
                pbMemberTypesFromDB = d.memberTypes || [];
                pbGrpRenderMemberTypes(selectedIds);
            }
        });
        return;
    }
    var html = '';
    for (var i = 0; i < pbMemberTypesFromDB.length; i++) {
        var mt = pbMemberTypesFromDB[i];
        var checked = selectedIds.indexOf(mt.id) >= 0 ? ' checked' : '';
        html += '<label class="pb-checkbox"><input type="checkbox" class="pb-grp-mt" value="' + mt.id + '"' + checked + '> ' + mt.description + '</label>';
    }
    el.innerHTML = html;
}

function pbGrpSaveGroup() {
    var name = document.getElementById('pb-grp-edit-name').value.trim();
    if (!name) { alert('Please enter a group name.'); return; }
    var level = document.getElementById('pb-grp-edit-level').value;
    var progSel = document.getElementById('pb-grp-edit-program');
    var programId = progSel.value ? parseInt(progSel.value) : 0;
    var programName = progSel.options[progSel.selectedIndex] ? progSel.options[progSel.selectedIndex].text : '';
    var divSel = document.getElementById('pb-grp-edit-division');
    var divisionId = divSel.value ? parseInt(divSel.value) : 0;
    var divisionName = divSel.options[divSel.selectedIndex] ? divSel.options[divSel.selectedIndex].text : '';
    var orgId = parseInt(document.getElementById('pb-grp-edit-org-id').value) || 0;
    var orgName = document.getElementById('pb-grp-inv-selected').textContent || '';

    if (!programId) { alert('Please select a program.'); return; }
    if ((level === 'division' || level === 'involvement') && !divisionId) { alert('Please select a division.'); return; }
    if (level === 'involvement' && !orgId) { alert('Please select an involvement.'); return; }

    // Collect selected member types
    var memberTypes = [];
    document.querySelectorAll('.pb-grp-mt:checked').forEach(function(cb) {
        memberTypes.push(parseInt(cb.value));
    });

    var selectedSwatch = document.querySelector('.pb-grp-color-swatch.selected');
    var group = {
        id: document.getElementById('pb-grp-edit-id').value || '',
        name: name,
        description: document.getElementById('pb-grp-edit-desc').value.trim(),
        level: level,
        programId: programId,
        programName: programName,
        divisionId: divisionId,
        divisionName: divisionName,
        orgId: orgId,
        orgName: orgName,
        memberTypes: memberTypes,
        targetActions: pbGrpConfigActions,
        minEnrollDays: parseInt(document.getElementById('pb-grp-edit-min-days').value) || 0,
        minStaleDays: parseInt(document.getElementById('pb-grp-edit-stale-days').value) || 0,
        color: selectedSwatch ? selectedSwatch.dataset.color : '#3498db'
    };
    pbAjax({action: 'save_prospect_group', group_data: JSON.stringify(group)}, function(d) {
        if (d.success) {
            pbGrpCloseModal();
            pbGrpLoadData();
        } else {
            alert('Error: ' + (d.message || 'Save failed'));
        }
    });
}

function pbGrpDelete(id, onSuccess) {
    var g = pbGrpFindGroup(id);
    var name = g ? g.name : 'this group';
    if (!confirm('Delete "' + name + '" and all its assignments/efforts?')) return;
    pbAjax({action: 'delete_prospect_group', group_id: id}, function(d) {
        if (d.success) {
            pbGrpLoadData();
            if (typeof onSuccess === 'function') onSuccess();
        }
    });
}

// --- Detail View ---
function pbGrpViewDetail(groupId) {
    pbGrpCurrentGroup = pbGrpFindGroup(groupId);
    if (!pbGrpCurrentGroup) return;
    pbGrpDetailTab = 'prospects';

    document.getElementById('pb-grp-overview').style.display = 'none';
    document.getElementById('pb-grp-detail').style.display = '';
    document.getElementById('pb-grp-detail-name').textContent = pbGrpCurrentGroup.name;
    document.getElementById('pb-grp-detail-desc').textContent = pbGrpCurrentGroup.description || '';

    // Update change badge
    var unacked = 0;
    for (var i = 0; i < pbGrpChanges.length; i++) {
        if (pbGrpChanges[i].groupId === groupId && !pbGrpChanges[i].acknowledged) unacked++;
    }
    var badge = document.getElementById('pb-grp-change-badge');
    if (unacked > 0) { badge.textContent = unacked; badge.style.display = ''; }
    else { badge.style.display = 'none'; }

    // Reset sub-tab state
    document.querySelectorAll('.pb-grp-detail-tab').forEach(function(t) { t.classList.remove('active'); });
    document.querySelector('.pb-grp-detail-tab[data-dtab="prospects"]').classList.add('active');
    document.getElementById('pb-grp-detail-prospects').style.display = '';
    document.getElementById('pb-grp-detail-efforts').style.display = 'none';
    document.getElementById('pb-grp-detail-changes').style.display = 'none';

    pbGrpLoadDetailProspects();
}

function pbGrpBackToOverview() {
    pbGrpCurrentGroup = null;
    document.getElementById('pb-grp-detail').style.display = 'none';
    document.getElementById('pb-grp-overview').style.display = '';
    pbGrpLoadData();
}

function pbGrpSwitchDetailTab(tab) {
    pbGrpDetailTab = tab;
    document.querySelectorAll('.pb-grp-detail-tab').forEach(function(t) { t.classList.remove('active'); });
    document.querySelector('.pb-grp-detail-tab[data-dtab="' + tab + '"]').classList.add('active');
    document.getElementById('pb-grp-detail-prospects').style.display = tab === 'prospects' ? '' : 'none';
    document.getElementById('pb-grp-detail-efforts').style.display = tab === 'efforts' ? '' : 'none';
    document.getElementById('pb-grp-detail-changes').style.display = tab === 'changes' ? '' : 'none';

    if (tab === 'efforts') pbGrpRenderEfforts();
    if (tab === 'changes') pbGrpRenderChanges();
}

var pbGrpContactMap = {};  // Contact efforts from TouchPoint TaskNotes for group prospects

function pbGrpLoadDetailProspects() {
    if (!pbGrpCurrentGroup) return;
    var el = document.getElementById('pb-grp-detail-prospects');
    el.innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading prospects...</div>';

    pbAjax({action: 'get_group_prospect_details', group_id: pbGrpCurrentGroup.id}, function(d) {
        if (!d.success) {
            el.innerHTML = '<div class="pb-text-muted" style="padding:20px;text-align:center;">Failed to load prospects.</div>';
            return;
        }
        pbGrpDetailProspects = d.prospects || [];
        pbGrpRenderDetailProspects();

        // Also load TouchPoint contact efforts for these prospects
        if (pbGrpDetailProspects.length > 0) {
            var pids = pbGrpDetailProspects.map(function(p) { return p.PeopleId; }).join(',');
            pbAjax({action: 'load_contact_efforts', people_ids: pids}, function(d2) {
                if (d2.success) {
                    pbGrpContactMap = d2.contactMap || {};
                    pbGrpRenderDetailProspects();  // Re-render with contact data
                }
                // Compute scores for group prospects
                if (pbSettings.scorecard && pbSettings.scorecard.enabled) {
                    pbComputeScores(pids, function() { pbGrpRenderDetailProspects(); });
                }
            });
        }
    });
}

function pbGrpRenderDetailProspects() {
    var el = document.getElementById('pb-grp-detail-prospects');
    if (!pbGrpDetailProspects.length) {
        el.innerHTML = '<div class="pb-empty" style="padding:30px;"><div style="font-size:1.8em;margin-bottom:8px;">&#128100;</div><p>No prospects found for this group.<br>Check that the program/division/involvement has members.</p></div>';
        return;
    }

    var now = new Date();
    var thirtyDaysAgo = new Date(now - 30 * 24 * 60 * 60 * 1000).toISOString().substring(0, 10);

    var html = '<div style="overflow-x:auto;">';
    html += '<table class="pb-table" style="table-layout:fixed;width:100%;"><thead><tr>';
    html += '<th style="width:150px;">Name</th><th style="width:75px;">Type</th><th style="width:35px;">Age</th><th style="width:100px;">Status</th>';
    html += '<th style="width:170px;">Contact History</th><th style="width:50px;">Efforts</th><th style="width:50px;">Score</th><th style="width:80px;">Last Activity</th><th style="width:80px;">Actions</th>';
    html += '</tr></thead><tbody>';

    var methods = pbSettings.contact_methods || [];

    for (var i = 0; i < pbGrpDetailProspects.length; i++) {
        var p = pbGrpDetailProspects[i];
        var grpCm = pbGrpContactMap[String(p.PeopleId)];
        var grpWt = grpCm ? (grpCm.weightedTotal != null ? grpCm.weightedTotal : grpCm.total) : 0;
        var hasEfforts = p.effortCount > 0 || grpWt > 0;
        var isStale = p.isAssigned && (!p.lastEffort || p.lastEffort.substring(0, 10) < thirtyDaysAgo) && !hasEfforts;
        var isDropped = p.status === 'dropped';
        var rowClass = isDropped ? 'pb-grp-prospect-row dropped' : isStale ? 'pb-grp-prospect-row stale' : 'pb-grp-prospect-row';

        // Build contact effort badges from TouchPoint TaskNotes
        var cm = pbGrpContactMap[String(p.PeopleId)];
        var contactHtml = '';
        if (cm) {
            contactHtml = '<span class="pb-contact-badge">';
            for (var mi = 0; mi < methods.length; mi++) {
                var code = methods[mi].code;
                var cnt = (cm.methods && cm.methods[code]) || 0;
                contactHtml += '<span class="pb-contact-code' + (cnt > 0 ? ' has-count' : '') + '">' + code + '(' + cnt + ')</span>';
            }
            var otherCnt = (cm.methods && cm.methods['O']) || 0;
            if (otherCnt > 0) {
                var grpOw = pbSettings.other_weight != null ? pbSettings.other_weight : 1;
                var grpDim = grpOw < 1 ? ' style="opacity:' + Math.max(0.35, grpOw) + ';"' : '';
                contactHtml += '<span class="pb-contact-code has-count"' + grpDim + '>O(' + otherCnt + ')</span>';
            }
            contactHtml += '</span>';
            if (cm.lastDate) contactHtml += '<div class="pb-text-muted" style="font-size:0.75em;">Last: ' + cm.lastDate.substring(0, 10) + '</div>';
        } else {
            contactHtml = '<span class="pb-text-muted pb-text-sm">None</span>';
        }

        // Determine last activity date (most recent of contact effort or group effort)
        var lastActivity = '';
        if (p.lastEffort && cm && cm.lastDate) {
            lastActivity = p.lastEffort > cm.lastDate ? p.lastEffort.substring(0, 10) : cm.lastDate.substring(0, 10);
        } else if (p.lastEffort) {
            lastActivity = p.lastEffort.substring(0, 10);
        } else if (cm && cm.lastDate) {
            lastActivity = cm.lastDate.substring(0, 10);
        }

        html += '<tr class="' + rowClass + '" style="cursor:pointer;" data-pid="' + p.PeopleId + '" onclick="pbGrpToggleDetail(parseInt(this.dataset.pid))">';
        html += '<td>';
        html += '<span style="color:var(--pb-primary);font-weight:600;">' + p.Name2 + '</span>';
        html += ' <a href="/Person2/' + p.PeopleId + '" target="_blank" style="color:var(--pb-muted);font-size:0.7em;text-decoration:none;" title="Open in TouchPoint" onclick="event.stopPropagation()">&#8599;</a>';
        if (p.notes) html += '<div class="pb-text-muted" style="font-size:0.75em;">' + p.notes + '</div>';
        if (p.EnrollmentDate) html += '<div class="pb-text-muted" style="font-size:0.7em;">Enrolled: ' + p.EnrollmentDate + '</div>';
        html += '</td>';
        html += '<td><span class="pb-badge ' + (p.MemberType === 'Prospect' ? 'pb-badge-warning' : p.MemberType === 'Member' ? 'pb-badge-success' : 'pb-badge-muted') + '" style="font-size:0.7em;">' + (p.MemberType || '-') + '</span></td>';
        html += '<td>' + (p.Age || '-') + '</td>';
        html += '<td>' + (p.MemberStatus || '-') + '</td>';
        html += '<td>' + contactHtml + '</td>';
        html += '<td>' + (p.effortCount || 0) + '</td>';
        var grpScoreData = pbScoreMap[String(p.PeopleId)];
        if (pbSettings.scorecard && pbSettings.scorecard.enabled && grpScoreData) {
            var gs = grpScoreData.score;
            html += '<td><span data-pid="' + p.PeopleId + '" onclick="event.stopPropagation();pbShowScorePopup(parseInt(this.dataset.pid),event)" style="font-weight:700;color:' + pbGetScoreColor(gs) + ';cursor:pointer;">' + gs + '</span></td>';
        } else {
            html += '<td>-</td>';
        }
        html += '<td>' + (lastActivity || '<span class="pb-text-muted">None</span>') + '</td>';
        html += '<td style="white-space:nowrap;">';
        var safeName = (p.Name2 || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;');
        html += '<div class="pb-action-wrap">';
        html += '<div class="pb-action-trigger" data-pid="' + p.PeopleId + '" onclick="event.stopPropagation();pbGrpToggleActionMenu(parseInt(this.dataset.pid))">Action &#9660;</div>';
        html += '<div class="pb-action-menu" id="pb-grp-amenu-' + p.PeopleId + '">';
        // Target actions section
        var grpActions = (pbGrpCurrentGroup && pbGrpCurrentGroup.targetActions) ? pbGrpCurrentGroup.targetActions : [];
        if (grpActions.length) {
            html += '<div class="pb-action-menu-section">';
            html += '<div class="pb-action-menu-label">Assign To</div>';
            for (var ai = 0; ai < grpActions.length; ai++) {
                var act = grpActions[ai];
                var actIcon = act.pb_type === 'tag' ? '&#127991;' : '&#10133;';
                html += '<div class="pb-action-menu-item" data-aidx="' + ai + '" data-pid="' + p.PeopleId + '" onclick="event.stopPropagation();pbGrpProcessAction(parseInt(this.dataset.pid),parseInt(this.dataset.aidx))">';
                html += '<span class="pb-ami-icon" style="color:var(--pb-accent);">' + actIcon + '</span>';
                html += '<span>' + (act.label || act.tagName || act.orgName) + '</span></div>';
            }
            html += '</div>';
        }
        // Status/utility section
        html += '<div class="pb-action-menu-section">';
        html += '<div class="pb-action-menu-item" data-gid="' + pbGrpCurrentGroup.id + '" data-pid="' + p.PeopleId + '" data-name="' + safeName + '" onclick="event.stopPropagation();pbGrpShowEffortModal(this.dataset.gid,parseInt(this.dataset.pid),this.dataset.name)">';
        html += '<span class="pb-ami-icon">&#128221;</span><span>Log Effort</span></div>';
        html += '<div class="pb-action-menu-item" data-gid="' + pbGrpCurrentGroup.id + '" data-pid="' + p.PeopleId + '" onclick="event.stopPropagation();pbGrpRemoveProspect(this.dataset.gid,parseInt(this.dataset.pid))">';
        html += '<span class="pb-ami-icon" style="color:var(--pb-danger);">&#10060;</span><span>Remove</span></div>';
        html += '</div>';
        html += '</div></div>';
        html += '</td>';
        html += '</tr>';
        // Expandable detail row
        html += '<tr id="pb-grp-detail-row-' + p.PeopleId + '" style="display:none;"><td colspan="9" style="padding:12px 15px;background:var(--pb-light);border-bottom:2px solid var(--pb-border);">';
        html += '<div id="pb-grp-person-detail-' + p.PeopleId + '"><div class="pb-loading"><span class="pb-spin"></span> Loading...</div></div>';
        html += '</td></tr>';
    }

    html += '</tbody></table></div>';
    el.innerHTML = html;
}

// --- Person Detail Expand ---
var pbGrpExpandedPid = null;
var pbGrpDetailCache = {};

function pbGrpToggleDetail(pid) {
    // Collapse previous
    if (pbGrpExpandedPid && pbGrpExpandedPid !== pid) {
        var prevRow = document.getElementById('pb-grp-detail-row-' + pbGrpExpandedPid);
        if (prevRow) prevRow.style.display = 'none';
    }

    var detailRow = document.getElementById('pb-grp-detail-row-' + pid);
    if (!detailRow) return;

    if (pbGrpExpandedPid === pid) {
        // Collapse
        detailRow.style.display = 'none';
        pbGrpExpandedPid = null;
        return;
    }

    pbGrpExpandedPid = pid;
    detailRow.style.display = '';
    var contentEl = document.getElementById('pb-grp-person-detail-' + pid);

    // Check cache
    if (pbGrpDetailCache[pid]) {
        var p = pbGrpDetailProspects.find(function(pr) { return pr.PeopleId === pid; });
        contentEl.innerHTML = pbRenderPersonDetail(p, pbGrpDetailCache[pid]);
        return;
    }

    // Load detail via existing AJAX action
    contentEl.innerHTML = '<div class="pb-loading"><span class="pb-spin"></span> Loading details...</div>';
    pbAjax({action: 'load_person_detail', people_id: pid, org_id: 0, ev_fields: ''}, function(d) {
        if (d.success) {
            pbGrpDetailCache[pid] = d.detail;
            var p = pbGrpDetailProspects.find(function(pr) { return pr.PeopleId === pid; });
            var el = document.getElementById('pb-grp-person-detail-' + pid);
            if (el && p) el.innerHTML = pbRenderPersonDetail(p, d.detail);
        } else {
            var el = document.getElementById('pb-grp-person-detail-' + pid);
            if (el) el.innerHTML = '<div class="pb-text-muted">Failed to load details.</div>';
        }
    });
}

// --- Efforts Timeline ---
function pbGrpRenderEfforts() {
    var el = document.getElementById('pb-grp-detail-efforts');
    if (!pbGrpCurrentGroup) return;
    var gid = pbGrpCurrentGroup.id;
    var groupEfforts = pbGrpEfforts.filter(function(e) { return e.groupId === gid; });

    if (!groupEfforts.length) {
        el.innerHTML = '<div class="pb-empty" style="padding:30px;"><div style="font-size:1.8em;margin-bottom:8px;">&#128203;</div><p>No efforts logged yet for this group.</p></div>';
        return;
    }

    // Build person name lookup from loaded prospects
    var nameMap = {};
    for (var i = 0; i < pbGrpDetailProspects.length; i++) {
        nameMap[String(pbGrpDetailProspects[i].PeopleId)] = pbGrpDetailProspects[i].Name2;
    }

    var html = '';
    for (var i = 0; i < groupEfforts.length && i < 100; i++) {
        var e = groupEfforts[i];
        var icon = pbGrpEffortIcons[e.effortType] || '&#128196;';
        var typeLabel = pbGrpEffortLabels[e.effortType] || e.effortType;
        var resultLabel = pbGrpResultLabels[e.result] || e.result;
        var personName = nameMap[String(e.peopleId)] || ('Person #' + e.peopleId);
        var resultColor = e.result === 'connected' || e.result === 'interested' ? 'var(--pb-success)' :
                          e.result === 'declined' ? 'var(--pb-danger)' : 'var(--pb-muted)';

        html += '<div class="pb-grp-timeline-item">';
        html += '<div class="pb-grp-timeline-icon" style="background:rgba(52,152,219,0.1);color:var(--pb-accent);">' + icon + '</div>';
        html += '<div style="flex:1;">';
        html += '<div><strong>' + typeLabel + '</strong> with <a href="/Person2/' + e.peopleId + '" target="_blank" style="color:var(--pb-primary);">' + personName + '</a>';
        html += ' <span class="pb-badge" style="background:' + resultColor + ';color:#fff;font-size:0.7em;">' + resultLabel + '</span></div>';
        if (e.description) html += '<div class="pb-text-muted pb-text-sm" style="margin-top:2px;">' + e.description + '</div>';
        html += '<div class="pb-text-muted" style="font-size:0.75em;margin-top:2px;">' + (e.timestamp || '') + ' by ' + (e.userName || 'Unknown') + '</div>';
        html += '</div></div>';
    }
    el.innerHTML = html;
}

// --- Changes ---
function pbGrpRenderChanges() {
    var el = document.getElementById('pb-grp-detail-changes');
    if (!pbGrpCurrentGroup) return;
    var gid = pbGrpCurrentGroup.id;
    var groupChanges = pbGrpChanges.filter(function(c) { return c.groupId === gid; });

    if (!groupChanges.length) {
        el.innerHTML = '<div class="pb-empty" style="padding:30px;"><div style="font-size:1.8em;margin-bottom:8px;">&#128270;</div><p>No changes detected yet. Click "Detect Changes" to scan.</p></div>';
        return;
    }

    var nameMap = {};
    for (var i = 0; i < pbGrpDetailProspects.length; i++) {
        nameMap[String(pbGrpDetailProspects[i].PeopleId)] = pbGrpDetailProspects[i].Name2;
    }

    var html = '';
    for (var i = 0; i < groupChanges.length; i++) {
        var c = groupChanges[i];
        var personName = nameMap[String(c.peopleId)] || ('Person #' + c.peopleId);
        var changeIcon = c.changeType === 'attended' ? '&#9989;' : c.changeType === 'status_change' ? '&#128260;' : c.changeType === 'joined_group' ? '&#128101;' : '&#128161;';
        var acked = c.acknowledged;

        html += '<div class="pb-grp-change-item' + (acked ? '' : ' unacked') + '">';
        html += '<span style="font-size:1.2em;">' + changeIcon + '</span>';
        html += '<div style="flex:1;">';
        html += '<div><a href="/Person2/' + c.peopleId + '" target="_blank" style="color:var(--pb-primary);font-weight:600;">' + personName + '</a></div>';
        html += '<div class="pb-text-sm">' + (c.description || '') + '</div>';
        html += '<div class="pb-text-muted" style="font-size:0.75em;">' + (c.detectedAt || '') + '</div>';
        html += '</div>';
        if (!acked) {
            html += '<button class="pb-btn pb-btn-sm pb-btn-success" data-cid="' + c.id + '" onclick="pbGrpAcknowledge(this.dataset.cid)">Ack</button>';
        } else {
            html += '<span class="pb-badge pb-badge-muted">Acked</span>';
        }
        html += '</div>';
    }
    el.innerHTML = html;
}

function pbGrpDetectChanges() {
    if (!pbGrpCurrentGroup) return;
    var btn = document.querySelector('[onclick="pbGrpDetectChanges()"]');
    if (btn) { btn.disabled = true; btn.textContent = 'Scanning...'; }
    pbAjax({action: 'detect_group_changes', group_id: pbGrpCurrentGroup.id}, function(d) {
        if (btn) { btn.disabled = false; btn.textContent = 'Detect Changes'; }
        if (d.success) {
            var msg = d.newChanges > 0 ? d.newChanges + ' new change(s) detected!' : 'No new changes detected.';
            alert(msg);
            // Reload to get updated changelog
            pbGrpLoadData();
            setTimeout(function() {
                pbGrpViewDetail(pbGrpCurrentGroup.id);
                pbGrpSwitchDetailTab('changes');
            }, 500);
        }
    });
}

function pbGrpAcknowledge(changeId) {
    pbAjax({action: 'acknowledge_group_change', change_id: changeId}, function(d) {
        if (d.success) {
            // Update local state
            for (var i = 0; i < pbGrpChanges.length; i++) {
                if (pbGrpChanges[i].id === changeId) { pbGrpChanges[i].acknowledged = true; break; }
            }
            pbGrpRenderChanges();
            // Update badge
            var unacked = 0;
            for (var i = 0; i < pbGrpChanges.length; i++) {
                if (pbGrpChanges[i].groupId === pbGrpCurrentGroup.id && !pbGrpChanges[i].acknowledged) unacked++;
            }
            var badge = document.getElementById('pb-grp-change-badge');
            if (unacked > 0) { badge.textContent = unacked; badge.style.display = ''; }
            else { badge.style.display = 'none'; }
        }
    });
}

// --- Assignment ---
function pbGrpShowAssignModal(pids, names) {
    // Can be called from detail view (no args = show search) or from workspace (with pids)
    if (!pbGrpGroups.length) {
        alert('No prospect groups exist for this configuration. Create one first.');
        return;
    }
    if (pids) {
        document.getElementById('pb-grp-assign-pids').value = pids;
        document.getElementById('pb-grp-assign-names').textContent = names || (pids.split(',').length + ' prospect(s)');
    } else {
        document.getElementById('pb-grp-assign-pids').value = '';
        document.getElementById('pb-grp-assign-names').textContent = 'Assign prospects from the Workspace action menu.';
    }
    // Populate group dropdown
    var sel = document.getElementById('pb-grp-assign-target');
    sel.innerHTML = '';
    for (var i = 0; i < pbGrpGroups.length; i++) {
        sel.innerHTML += '<option value="' + pbGrpGroups[i].id + '">' + pbGrpGroups[i].name + '</option>';
    }
    document.getElementById('pb-grp-assign-notes').value = '';
    document.getElementById('pb-grp-assign-modal').classList.add('active');
}
function pbGrpCloseAssignModal() { document.getElementById('pb-grp-assign-modal').classList.remove('active'); }

function pbGrpSubmitAssign() {
    var pids = document.getElementById('pb-grp-assign-pids').value;
    var groupId = document.getElementById('pb-grp-assign-target').value;
    var notes = document.getElementById('pb-grp-assign-notes').value.trim();
    if (!pids || !groupId) { alert('Missing data.'); return; }

    pbAjax({action: 'assign_to_group', group_id: groupId, people_ids: pids, notes: notes}, function(d) {
        if (d.success) {
            pbGrpCloseAssignModal();
            alert(d.assigned + ' prospect(s) assigned.');
            if (pbGrpCurrentGroup && pbGrpCurrentGroup.id === groupId) {
                pbGrpLoadDetailProspects();
            }
            pbGrpLoadData();
        } else {
            alert('Error: ' + (d.message || 'Assign failed'));
        }
    });
}

// --- Effort Logging ---
function pbGrpShowEffortModal(groupId, pid, name) {
    document.getElementById('pb-grp-effort-gid').value = groupId;
    document.getElementById('pb-grp-effort-pid').value = pid;
    document.getElementById('pb-grp-effort-person').textContent = name;
    document.getElementById('pb-grp-effort-type').value = 'phone';
    document.getElementById('pb-grp-effort-result').value = 'connected';
    document.getElementById('pb-grp-effort-desc').value = '';
    document.getElementById('pb-grp-effort-modal').classList.add('active');
}
function pbGrpCloseEffortModal() { document.getElementById('pb-grp-effort-modal').classList.remove('active'); }

function pbGrpSubmitEffort() {
    var effort = {
        groupId: document.getElementById('pb-grp-effort-gid').value,
        peopleId: parseInt(document.getElementById('pb-grp-effort-pid').value),
        effortType: document.getElementById('pb-grp-effort-type').value,
        result: document.getElementById('pb-grp-effort-result').value,
        description: document.getElementById('pb-grp-effort-desc').value.trim()
    };
    if (!effort.groupId || !effort.peopleId) { alert('Missing data.'); return; }

    pbAjax({action: 'log_group_effort', effort_data: JSON.stringify(effort)}, function(d) {
        if (d.success) {
            pbGrpCloseEffortModal();
            // Add to local efforts
            if (d.effort) pbGrpEfforts.unshift(d.effort);
            // Refresh prospect list to update effort counts
            pbGrpLoadDetailProspects();
        } else {
            alert('Error: ' + (d.message || 'Failed'));
        }
    });
}

// --- Move ---
function pbGrpShowMoveModal(fromGroupId, pid, name) {
    document.getElementById('pb-grp-move-from').value = fromGroupId;
    document.getElementById('pb-grp-move-pids').value = String(pid);
    document.getElementById('pb-grp-move-names').textContent = name;
    var fromGroup = pbGrpFindGroup(fromGroupId);
    document.getElementById('pb-grp-move-from-label').textContent = 'Moving from: ' + (fromGroup ? fromGroup.name : 'Unknown');

    // Populate target dropdown (exclude current group)
    var sel = document.getElementById('pb-grp-move-target');
    sel.innerHTML = '';
    for (var i = 0; i < pbGrpGroups.length; i++) {
        if (pbGrpGroups[i].id !== fromGroupId) {
            sel.innerHTML += '<option value="' + pbGrpGroups[i].id + '">' + pbGrpGroups[i].name + '</option>';
        }
    }
    if (!sel.innerHTML) {
        alert('No other groups to move to. Create another group first.');
        return;
    }
    document.getElementById('pb-grp-move-notes').value = '';
    document.getElementById('pb-grp-move-modal').classList.add('active');
}
function pbGrpCloseMoveModal() { document.getElementById('pb-grp-move-modal').classList.remove('active'); }

function pbGrpSubmitMove() {
    var fromGroup = document.getElementById('pb-grp-move-from').value;
    var toGroup = document.getElementById('pb-grp-move-target').value;
    var pids = document.getElementById('pb-grp-move-pids').value;
    var notes = document.getElementById('pb-grp-move-notes').value.trim();
    if (!fromGroup || !toGroup || !pids) { alert('Missing data.'); return; }

    pbAjax({action: 'move_prospect_group', from_group_id: fromGroup, to_group_id: toGroup, people_ids: pids, notes: notes}, function(d) {
        if (d.success) {
            pbGrpCloseMoveModal();
            alert(d.moved + ' prospect(s) moved.');
            pbGrpLoadDetailProspects();
            pbGrpLoadData();
        }
    });
}

// --- Status Update & Remove ---
function pbGrpUpdateStatus(groupId, pid, newStatus) {
    var label = newStatus === 'dropped' ? 'Drop this prospect from the group?' : 'Reactivate this prospect?';
    if (!confirm(label)) return;
    pbAjax({action: 'update_group_assignment', group_id: groupId, people_id: pid, new_status: newStatus}, function(d) {
        if (d.success) pbGrpLoadDetailProspects();
    });
}

function pbGrpRemoveProspect(groupId, pid) {
    if (!confirm('Remove this prospect from the group entirely?')) return;
    pbAjax({action: 'remove_from_group', group_id: groupId, people_ids: String(pid)}, function(d) {
        if (d.success) pbGrpLoadDetailProspects();
    });
}

// Group action menu toggle
var pbGrpOpenMenuPid = null;
function pbGrpToggleActionMenu(pid) {
    if (pbGrpOpenMenuPid && pbGrpOpenMenuPid !== pid) {
        var prev = document.getElementById('pb-grp-amenu-' + pbGrpOpenMenuPid);
        if (prev) prev.classList.remove('open');
    }
    var menu = document.getElementById('pb-grp-amenu-' + pid);
    if (menu) {
        menu.classList.toggle('open');
        pbGrpOpenMenuPid = menu.classList.contains('open') ? pid : null;
    }
}
// Close group action menu when clicking outside
document.addEventListener('click', function(e) {
    if (pbGrpOpenMenuPid && !e.target.closest('.pb-action-wrap')) {
        var menu = document.getElementById('pb-grp-amenu-' + pbGrpOpenMenuPid);
        if (menu) menu.classList.remove('open');
        pbGrpOpenMenuPid = null;
    }
});

// Process a target action for a prospect in a group
function pbGrpProcessAction(pid, actionIdx) {
    if (!pbGrpCurrentGroup || !pbGrpCurrentGroup.targetActions) return;
    var target = pbGrpCurrentGroup.targetActions[actionIdx];
    if (!target) return;
    // Close menu
    if (pbGrpOpenMenuPid) {
        var m = document.getElementById('pb-grp-amenu-' + pbGrpOpenMenuPid);
        if (m) m.classList.remove('open');
        pbGrpOpenMenuPid = null;
    }
    if (!confirm('Apply "' + (target.label || 'action') + '" to this prospect?')) return;

    pbAjax({action: 'process_action', target_data: JSON.stringify(target), people_ids: String(pid)}, function(d) {
        if (d.success) {
            alert('Done! ' + (target.label || 'Action applied.'));
        } else {
            alert('Error: ' + (d.message || 'Failed'));
        }
    });
}

// Build group badge chips for a prospect in the list view
function pbBuildGroupBadges(pid) {
    if (!pbGrpGroups.length) return '';
    var pidKey = String(pid);
    var badges = '';
    for (var i = 0; i < pbGrpGroups.length; i++) {
        var grp = pbGrpGroups[i];
        var assignments = pbGrpAssignments[grp.id] || {};
        if (assignments[pidKey]) {
            var color = grp.color || '#3498db';
            badges += ' <span class="pb-chip" style="background:' + color + '22;color:' + color + ';border:1px solid ' + color + '44;font-weight:600;">' + grp.name + '</span>';
        }
    }
    return badges;
}

// Quick assign from workspace action menu
function pbGrpQuickAssign(groupId, pid) {
    pbAjax({action: 'assign_to_group', group_id: groupId, people_ids: String(pid), notes: ''}, function(d) {
        if (d.success) {
            // Close action menu
            if (pbOpenMenuPid) {
                var menu = document.getElementById('pb-amenu-' + pbOpenMenuPid);
                if (menu) menu.classList.remove('open');
                pbOpenMenuPid = null;
            }
            // Log activity
            var grp = pbGrpFindGroup(groupId);
            var grpName = grp ? grp.name : 'group';
            var personName = pbGetProspectName(pid);
            pbLogActivity(pid, personName, 'assign_group', 'Assigned to group: ' + grpName);
        }
    });
}

// Load all prospect groups (called when workspace launches for action menu)
function pbGrpLoadAll() {
    pbAjax({action: 'load_prospect_groups'}, function(d) {
        if (d.success) {
            pbGrpGroups = d.groups || [];
            pbGrpAssignments = d.assignments || {};
            // If the Group Management tab is already open when this returns,
            // populate the dropdown that pbGrpInitTab skipped during the race.
            var sel = document.getElementById('pb-grp-health-group');
            if (sel && typeof pbGrpHealthPopulateGroups === 'function') {
                pbGrpHealthPopulateGroups();
            }
        }
    });
}

// INITIALIZATION
// ============================================================
function pbInit() {
    pbApplyIntroVisibility();  // Show/hide per-tab help banners based on localStorage
    pbLoadConfigs();
    pbLoadSettings();
    pbGrpLoadAll();  // Pre-load prospect groups for workspace action menu
    // Pre-load member types from DB on startup
    pbAjax({action: 'get_member_types'}, function(d) {
        if (d.success) {
            pbMemberTypesFromDB = d.memberTypes || [];
            pbMemberTypeLabels = {};
            for (var i = 0; i < pbMemberTypesFromDB.length; i++) {
                pbMemberTypeLabels[String(pbMemberTypesFromDB[i].id)] = pbMemberTypesFromDB[i].description;
            }
        }
    });
    // Close search dropdowns when clicking outside
    document.addEventListener('click', function(e) {
        if (!e.target.closest('.pb-search-wrap')) {
            var drops = document.querySelectorAll('.pb-search-results');
            for (var i = 0; i < drops.length; i++) drops[i].style.display = 'none';
        }
    });
    // Dashboard is the default landing tab - kick off its data load now.
    if (typeof pbDashInitTab === 'function') pbDashInitTab();
    // Fire the auto-update version check last so the rest of the UI
    // renders without waiting on the external DisplayCache call.
    if (typeof checkForAppUpdate === 'function') checkForAppUpdate();
}
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', pbInit);
} else {
    pbInit();
}
</script>
'''

    model.Form = generate_ui()
