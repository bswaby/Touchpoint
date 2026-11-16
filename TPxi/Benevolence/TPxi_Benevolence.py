#Roles=Access
# -*- coding: utf-8 -*-
#####################################################################
# TPxi_Benevolence   (keep this line)
#####################################################################
# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub: https://github.com/bswaby/Touchpoint  (40+ free tools)
# ----------------------------------------------------------------
#
# Benevolence requests inside TouchPoint: a router gives each request from
# your request form to one committee member, who calls the person, decides
# with another member's affirmation, and sends it back; the router confirms
# the payouts (a check to the payee, or help set out at a pickup place).
#
# OPEN IT AT /PyScriptForm/TPxi_Benevolence (not /PyScript/).
#
# ACCESS: READ THIS BEFORE GIVING IT TO ANYONE
#   What people write on the form is confidential. Who gets in is decided
#   by three gates, and the tool checks every page and every action itself:
#   /PyScriptForm only checks that someone is signed in, and a #Roles line
#   (like the one at the top) does NOT keep anyone out of /PyScriptForm.
#
#   1. ROUTERS see every request. They need BOTH:
#        - their name in Settings, Setup, Forms and people (Routers), AND
#        - the ACCESS ROLE in TouchPoint (Admin, Roles). It is named
#          Benevolence unless you pick another in Settings, Setup (Access
#          role), e.g. a role your church already uses for this.
#      Take either away and they are out.
#   2. COMMITTEE MEMBERS see only the requests given to them, plus any a
#      teammate brought them in on (read, write in the case conversation,
#      add notes; no calling, deciding or handing off). The gate is the
#      committee involvement picked in Settings: anyone CURRENTLY on its
#      roster gets in, whatever their member type (pending and inactive
#      members do not). Drop someone from that involvement and they are out
#      at their next page. They do NOT need the access role, and
#      should not have it.
#   3. ADMINS (TouchPoint Admin role) get Settings only. Being an admin
#      opens no one's request; list yourself as a router for that.
#
#   TouchPoint setup to do once (Settings, Setup checks each of these and
#   says what is wrong; going live waits until the form is locked):
#     - Create a role named Benevolence (or pick one you have as the access
#       role); give it to the routers only.
#     - Lock each form's involvement (request form, documents form): open
#       it, Settings, General, Limit To Role = the access role (only admins
#       can set this). Without it, anyone who can look up involvements can
#       read every answer.
#     - Keep those involvements' member directory and chat off: everyone
#       who used the form is a member.
#     - Keep the committee involvement to real committee members.
#   Anyone with SQL report rights in TouchPoint can still read form answers
#   and extra values; keep that right to people you would trust with this.
#
# WHERE THE DATA LIVES
#   The form's own answers stay where TouchPoint keeps them (Registration,
#   RegPeople, RegAnswer). This tool's work on a request (who has it, the
#   call, the decision, payouts) is one JSON extra value on the REQUESTER,
#   named Benevolence, keyed by the submission's RegPeopleId (which survives
#   a merge). Never Special Content: TouchPoint serves any Special Content
#   by name with no sign-in.
#   The extra value is registered in StandardExtraValues2 with
#   VisibilityRoles so only the access role sees it on a profile.
#   Every page load checks that registration (the preflight); while it is
#   missing nothing is saved and admins are told. A merge where both people
#   have the value renames one Benevolence_mv (then _mv_mv), so those names
#   are registered too, and folded back in when the person is opened.
#   Settings (no personal data) are in Special Content TPxi_BenevolenceSettings.
#
# VERSION HISTORY
#   0.1.0  First cut: the router's home, a request (membership, family, give it
#          to a committee member in turn), Settings and the preflight.
#   0.1.1  A new request gives the router one Your decision card: Decline,
#          Assign or Help (Help when Settings, Decisions allows it).
#   0.1.2  What we will give: each line shows Pay to (payee, address, account
#          from the form, which can be added or corrected); each line says
#          Benevolence pays it or They pick it up, and the pickup place is only
#          asked when something is picked up; added help picks a category from
#          the form's needs. New form roles: payee address, account number.
#   0.1.3  Timeline: approvals and payouts list each line, with Details. A
#          ticked line with no amount is highlighted. Membership and family
#          shows each person's age and the badges on their TouchPoint record.
#   0.1.4  Membership and family keeps a Membership column. Someone at the
#          door searches everyone in TouchPoint and sorts each person by where
#          their request is or was (pickup, open and with whom, started and
#          not sent, documents only, a family member's, asked before, nothing).
#   0.1.5  Someone at the door lists only people who sent a request (with a
#          link to search everyone); a name opens the request. Each request
#          opens with Where it is and what has happened: the stage now and its
#          whole history, with what was given or paid and each Why.
#   0.2.0  Walk-ins: someone who asks without the form. Looked up in TouchPoint
#          as you type and linked if found; otherwise kept with what was typed
#          (on the router's record, hidden like the rest). A later form from
#          them shows a possible match to link. In the door search, Membership
#          and family, the timeline and Reports. Routers only.
#   0.2.1  Timeline: Back (to the request it was opened from), a button for
#          each of the family's requests, and every entry opens its request.
#   0.2.2  Reports by month: Members (a member on the day they asked, by join
#          date) and Joined later; Not sorted column title shortened. Saving
#          as you go: the router's note for the committee member now saves,
#          and anything still waiting to save is sent when leaving the page.
#   0.3.0  Updates like TPxi Texting: admins see when a newer version is
#          published and Update now saves it over this script; the version
#          shows at the bottom of every page.
#####################################################################

import re
import datetime

BEN_VERSION = '0.3.0'
# Updates come from the same TPxi script service as TPxi Texting and the other
# TPxi tools: the page asks DC_API_BASE which version is newest (admins only),
# and an admin's Update now pulls the new code from DC_API_WORKER and saves it
# over this script, under whatever name it was saved as. Saved data lives in
# extra values and Special Content under its own names, so it is untouched.
DC_SCRIPT_ID = 'TPxi_Benevolence'
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'
SETTINGS_CONTENT = 'TPxi_BenevolenceSettings'
EV_FIELD = 'Benevolence'
EV_ALIASES = ['Benevolence_mv', 'Benevolence_mv_mv']     # names a merge gives a second copy
TEAM_FIELD = 'BenevolenceTeam'                              # each team member's posts and what they have seen
TEAM_NAMES = [TEAM_FIELD, 'BenevolenceTeam_mv', 'BenevolenceTeam_mv_mv']
REQ_NAMES = [EV_FIELD] + EV_ALIASES
QUICK_FIELD = 'BenevolenceWalkIns'                          # people who asked without the form, on the router who wrote them down
QUICK_NAMES = [QUICK_FIELD, 'BenevolenceWalkIns_mv', 'BenevolenceWalkIns_mv_mv']
SEV_CONTENT = 'StandardExtraValues2'
TS = '%Y-%m-%d %H:%M:%S'
FIRST_NUMBER = 1001
model.Header = ''


# ---------------------------------------------------------------- JSON
# IronPython json.dumps breaks on non-ASCII; see CLAUDE.md.
def _json_escape_string(s):
    if not isinstance(s, (str, unicode)):
        s = unicode(s)
    parts = ['"']
    for ch in s:
        try:
            code = ord(ch)
        except Exception:
            parts.append('\\ufffd')
            continue
        if code == 0x22:
            parts.append('\\"')
        elif code == 0x5C:
            parts.append('\\\\')
        elif code == 0x0A:
            parts.append('\\n')
        elif code == 0x0D:
            parts.append('\\r')
        elif code == 0x09:
            parts.append('\\t')
        elif code < 0x20 or code >= 0x7F:
            parts.append('\\u%04x' % code)
        else:
            parts.append(chr(code))
    parts.append('"')
    return ''.join(parts)


def _json_encode(obj):
    if obj is None:
        return 'null'
    if obj is True:
        return 'true'
    if obj is False:
        return 'false'
    if isinstance(obj, (int, long)):
        return str(obj)
    if isinstance(obj, float):
        return 'null' if obj != obj else repr(obj)
    if isinstance(obj, (str, unicode)):
        return _json_escape_string(obj)
    if isinstance(obj, dict):
        return '{' + ','.join(_json_escape_string(unicode(k)) + ':' + _json_encode(v)
                              for k, v in obj.items()) + '}'
    if isinstance(obj, (list, tuple)):
        return '[' + ','.join(_json_encode(x) for x in obj) + ']'
    return _json_escape_string(unicode(obj))


def safe_json(obj):
    try:
        return _json_encode(obj)
    except Exception as e:
        return '{"success":false,"message":' + _json_escape_string('safe_json: ' + repr(e)) + '}'


def load_json(raw, default):
    import json
    try:
        v = json.loads(raw) if raw else default
        return v if isinstance(v, type(default)) else default
    except Exception:
        return default


# ---------------------------------------------------------------- helpers
def _param(name, default=''):
    try:
        v = getattr(model.Data, name)
        v = unicode(v).strip()
        return v if v else default
    except:
        return default


def now_str():
    return datetime.datetime.now().strftime(TS)


def esc(s):
    return (unicode(s if s is not None else '').replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def sql_str(s):
    return unicode(s or '').replace("'", "''")


def sql_ints(xs):
    xs = [int(x) for x in xs if unicode(x).lstrip('-').isdigit()]
    return ','.join(str(x) for x in xs) if xs else '0'


GUID = re.compile(r'^[0-9a-fA-F\-]{36}$')


def money(v):
    """'$1,250' from 1250; '' for None."""
    if v is None or v == '':
        return ''
    try:
        f = float(v)
    except Exception:
        return esc(v)
    return '${:,.0f}'.format(f) if f == int(f) else '${:,.2f}'.format(f)


def to_number(text):
    """A dollar amount typed as free text ('$700', '700.00', 'about 650') or None."""
    m = re.search(r'\d[\d,]*(\.\d+)?', unicode(text or ''))
    if not m:
        return None
    try:
        return float(m.group(0).replace(',', ''))
    except Exception:
        return None


def days_ago(ts):
    """'today', 'yesterday', '3 days ago' from 'YYYY-MM-DD HH:MM:SS'."""
    try:
        d = (datetime.date.today() - datetime.datetime.strptime(ts[:10], '%Y-%m-%d').date()).days
    except Exception:
        return ''
    return 'today' if d <= 0 else 'yesterday' if d == 1 else '%d days ago' % d


def nice_date(ts):
    try:
        d = datetime.datetime.strptime(ts[:10], '%Y-%m-%d')
    except Exception:
        return ts or ''
    return d.strftime('%b %d').replace(' 0', ' ') + ('' if d.year == datetime.date.today().year else ', %d' % d.year)


# Notes typed in the page are formatted text (bold, lists, links). Only these
# tags survive, and only http(s) and mailto links.
ALLOWED_TAGS = set(['b', 'strong', 'i', 'em', 'u', 'ul', 'ol', 'li', 'br', 'p', 'div', 'a'])


def clean_html(html):
    html = unicode(html or '')
    html = re.sub(r'(?is)<(script|style|iframe|object|embed)[^>]*>.*?</\1>', '', html)

    def tag(m):
        close, name, attrs = m.group(1), m.group(2).lower(), m.group(3) or ''
        if name not in ALLOWED_TAGS:
            return ''
        if close:
            return '</%s>' % name
        if name == 'a':
            h = re.search(r'''href\s*=\s*["']([^"']+)["']''', attrs, re.I)
            href = h.group(1) if h else ''
            if not re.match(r'^(https?://|mailto:)', href, re.I):
                return '<a>'
            return '<a href="%s" target="_blank" rel="noopener">' % esc(href)
        return '<%s>' % name
    return re.sub(r'<(/?)([a-zA-Z0-9]+)([^>]*)>', tag, html)[:20000]


def from_b64(v):
    """The page sends formatted notes base64-encoded (ASP.NET refuses '<' in a post)."""
    if not v:
        return u''
    import System
    try:
        return unicode(System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(v)))
    except Exception:
        return u''


# ---------------------------------------------------------------- settings
DEFAULT_SETTINGS = {
    'access_role': 'Benevolence',
    # Texts to the person who asked: send only, through one of TouchPoint's
    # texting groups. The text never holds details, only a link.
    'text_group': 0, 'text_church': '', 'text_phone': '', 'text_autosignin': False,
    'text_secure': '{church}: you have a secure message about your request for help. Read it here: {link} Replies to this number are not monitored.',
    'text_docs': '{church}: please send the documents for your request for help here: {link} Replies to this number are not monitored.',
    'form_org': 0,                  # the church's benevolence request form (a new-forms involvement)
    'docs_org': 0,                  # the documents form, if they have one
    'committee_org': 0,             # Benevolence Committee involvement
    'routers': [],                  # [{'pid': 123, 'main': True}]: main routers get the emails; backups
                                    # (main False) can do everything but are only emailed when something is stuck
    'merger_pid': 0,                # who merges duplicate records
    'go_live': '',                  # requests before this date are 'before this tool'
    'mode': 'testing',              # testing: nothing is emailed and admins can clear the data; live: the real thing
    # Emails (Settings, Emails). Sent only in Live mode, by ScheduledTasks.
    'email_on': False, 'email_time': '07:00', 'email_days': 'weekdays', 'email_backup_days': 3,
    'email_router': True, 'email_committee': True, 'email_secure': True, 'email_from': '', 'email_from_name': 'Benevolence',
    'email_tagline': "This mailbox is not monitored, so please don't reply. Open Benevolence to answer or ask a question.",
    # Right-away notices (Settings, Emails): one email per person at the next
    # ScheduledTasks run, listing what just happened that involves them.
    'notify_assigned': True, 'notify_shared': True, 'notify_decision': True, 'notify_handoff': True,
    'notify_docs': True, 'notify_paid': False,
    # When each goes out: 'run' (the next 15-minute run, batched) or 'now' (at the click).
    'notify_when': {},
    'email_secure_when': 'later',   # 'later': 10 to 25 minutes if unread; 'now': at once if they are not in the tool
    'away': {},                     # {'pid': 'YYYY-MM-DD'} away until
    'affirm_on': True, 'affirm_required': False, 'handoff_on': True,
    # What must be done before a decision can be sent (the always-on rules are in decision_missing).
    'req_why_help': True, 'req_why_decline': True, 'req_talked_help': True, 'affirm_over': 0,
    'req_pickup': False,            # pickup place, day and time set by the committee member before sending
    'earlier_on': True,             # whoever works a case sees the family's earlier decisions and call notes
    'router_decide': 'decline',     # on a new request the router may: 'none' (always give it out), 'decline', 'both' (decline or help)
    'stuck_days': 7,                # a request with the committee and no change this long shows on the router's list
    # Ways help is given. pickup: they come get it (place, day and time agreed);
    # otherwise it is paid to the payee. reqs: what they must bring or show.
    'help_kinds': [
        {'name': 'Check to the payee', 'pickup': False, 'reqs': []},
        {'name': 'Gas card', 'pickup': True, 'reqs': ["Driver's license", 'Proof of car insurance']},
        {'name': 'Walmart card', 'pickup': True, 'reqs': ['Photo ID']},
    ],
    'pickup_places': ['Church office'],
    # The call: one question per line; a line starting with # starts a topic;
    # [yes/no] at the end gives Yes / No / Didn't say buttons. Samples until
    # the team's own list is in.
    'call_questions': '''# Household
Who lives in the home?
Is anyone in the home working? [yes/no]
# Other help tried
Which agencies have you asked for help with this?
What did they say?
Is anyone in the family able to help? [yes/no]
# Income and bills
What is the household's monthly income?
What are the main monthly bills?
# The need
What happened that led to this need?
If we help this time, what changes so this does not happen again?
# Church and faith
Do you have a church home? Where?
Is there anything we can pray with you about?''',
    # How the tool reads the request form: phrases matched anywhere in a question,
    # ignoring case; several are separated by commas. Each church sets these to
    # its own form's wording (these match the form they were first built for).
    'form_words': {
        'need': 'Do you need',                  # a yes/no question per need; Yes = they asked for it
        'amount': 'Amount needed',              # follows a need: how much
        'payee': 'Payee name, Loan payee',      # follows a need: who it is paid to
        'address': 'Payee address',             # follows a need: where the payment goes
        'account': 'Account number, Loan number',   # follows a need: the account or loan; routers only
        'detail': 'Is the need for',            # follows a need: adds detail to its name
        'total': 'Total amount of assistance',
        'people': 'Number of people living',
        'elsewhere': 'Received assistance from another',
        'elsewhere_who': 'Name of the individual or organization',
        'private': 'Social security',           # shown to routers only
        'notice': 'Please note',                # an opening notice, not a section name
    },
    # Matching the form (Settings, Reading your form): what each question is,
    # by RegQuestionId, only where the admin changed what the tool worked out.
    'form_map': {},
    'form_seen': [],                # question ids the admin has looked at (new ones are flagged)
    'docs_map': {},                 # the documents form: show, private or hide, where the admin changed it
    'decline_reasons': ['Other help not tried first', 'Asked too recently', 'Not a need we cover',
                        'Could not reach them', 'Other'],
    'referrals': ['Salvation Army', 'County community services'],
}


def load_settings():
    s = load_json(model.TextContent(SETTINGS_CONTENT) or '', {})
    out = dict(DEFAULT_SETTINGS)
    out.update(s)
    # The first cut stored fixed kinds with a 'needs' string and no editor; swap
    # those for the editable list.
    if any('reqs' not in k for k in out.get('help_kinds') or []):
        out['help_kinds'] = DEFAULT_SETTINGS['help_kinds']
    words = dict(DEFAULT_SETTINGS['form_words'])
    words.update(out.get('form_words') or {})
    out['form_words'] = words
    return out


def save_settings(s):
    model.WriteContentText(SETTINGS_CONTENT, safe_json(s), "")


# ---------------------------------------------------------------- who is this
ME = model.UserPeopleId          # a property, not a method; None when not signed in
IS_ADMIN = model.UserIsInRole('Admin')


def committee_members(st):
    """[{pid, name}] currently on the committee involvement: the committee
    gate. Any member type counts; pending members and anyone past their
    inactive date do not. No role is needed."""
    org = int(st.get('committee_org') or 0)
    if not org:
        return []
    return [{'pid': r.PeopleId, 'name': r.Name} for r in q.QuerySql("""
        SELECT p.PeopleId, ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName Name
        FROM OrganizationMembers om JOIN People p ON p.PeopleId = om.PeopleId
        WHERE om.OrganizationId = %d AND ISNULL(om.Pending, 0) = 0 AND ISNULL(om.InactiveDate, '2100-01-01') > GETDATE()
        ORDER BY p.Name2""" % org)]


def access_role():
    """The TouchPoint role that marks routers and hides the data: Benevolence
    unless an admin picked another in Settings (a church may already have one,
    or call it something else)."""
    try:
        st = SETTINGS
    except NameError:
        st = load_settings()
    return (st.get('access_role') or 'Benevolence').strip() or 'Benevolence'


def role_b():
    """The access role's name, bold and escaped, safe inside a %-formatted string."""
    return '<b>%s</b>' % esc(access_role()).replace('%', '%%')


def who_am_i(st):
    if not ME:
        return 'none'
    # A router is listed in Settings AND has the access role, so taking the
    # role away in TouchPoint's user permissions also takes away access.
    if any(int(r.get('pid') or 0) == int(ME) for r in st.get('routers') or []) and model.UserIsInRole(access_role()):
        return 'router'
    if any(m['pid'] == int(ME) for m in committee_members(st)):
        return 'committee'
    return 'admin' if IS_ADMIN else 'none'


# ---------------------------------------------------------------- what each person may open
# One table for every page and action, checked before anything runs.
#   router     : listed in Settings AND has the access role; every request
#   committee  : currently on the committee involvement's roster (any member
#                type; not pending, not past an inactive date); only requests
#                given to them, or that a teammate brought them in on
#   admin      : TouchPoint Admin; Settings only. Being an admin does not
#                open anyone's request: add yourself as a router for that.
# The header at the top of this file explains these gates for a new church,
# and Settings, Setup shows them live (who_can_get_in, involvement_check).
PAGE_ACCESS = {'home': ('router',), 'all': ('router',), 'find': ('router',), 'older': ('router',),
               'req': ('router', 'committee'), 'timeline': ('router', 'committee'), 'compare': ('router', 'committee'),
               'mine': ('committee', 'router'), 'settings': ('admin',), 'help': ('router', 'committee', 'admin'),
               'call': ('committee', 'router'), 'team': ('committee', 'router'), 'reports': ('router',),
               'walkins': ('router',), 'walkin': ('router',)}
ACTION_ACCESS = {'protect': ('admin',), 'save_settings': ('admin',), 'apply_update': ('admin',), 'assign': ('router',), 'sort_old': ('router',),
                 'walkin_match': ('router',), 'walkin_save': ('router',), 'walkin_link': ('router',), 'assign_note': ('router',),
                 # A router can also be on the committee; these check the request is theirs.
                 'dec_save': ('committee', 'router'), 'dec_send': ('committee', 'router'),
                 'affirm_set': ('committee', 'router'), 'affirm_note': ('committee', 'router'),
                 'reset_one': ('admin',), 'reset_all': ('admin',), 'set_mode': ('admin',), 'email_preview': ('admin',),
                 'call_save': ('committee', 'router'), 'call_notes': ('committee', 'router'),
                 'contact_log': ('committee', 'router'),
                 'team_post': ('committee', 'router'), 'team_remove': ('committee', 'router'), 'req_msg': ('committee', 'router'), 'req_thread': ('committee', 'router'),
                 'share_add': ('committee', 'router'), 'share_remove': ('committee', 'router'),
                 'text_secure': ('committee', 'router'), 'text_docs': ('committee', 'router'),
                 'chat_poll': ('committee', 'router'),
                 'handoff_ask': ('committee', 'router'), 'handoff_cancel': ('committee', 'router'), 'handoff_keep': ('router',),
                 'pay_confirm': ('router',), 'pickup_done': ('router',), 'pickup_move': ('router',), 'docs_place': ('router',), 'form_preview': ('admin',),
                 'rdec_save': ('router',), 'rdec_send': ('router',), 'nudge': ('router',), 'take_over': ('router',),
                 'combine': ('router',), 'separate': ('router',), 'uncombine': ('router',)}


def has_access(role, allowed):
    return role in allowed or ('admin' in allowed and IS_ADMIN)


def can_open(role, r):
    """May this person open this request record?"""
    if role == 'router':
        return True
    if role == 'committee' and ME:
        return int((r or {}).get('owner') or 0) == int(ME) or int(ME) in [int(x) for x in (r or {}).get('shared') or []]
    return False


def names_for(pids):
    pids = [p for p in set(pids) if p]
    if not pids:
        return {}
    return dict((r.PeopleId, r.N) for r in q.QuerySql(
        "SELECT PeopleId, ISNULL(NULLIF(PreferredName, ''), FirstName) + ' ' + LastName N FROM People WHERE PeopleId IN (%s)"
        % sql_ints(pids)))


# ---------------------------------------------------------------- preflight: the extra value is hidden
def sev_xml():
    return model.TextContent(SEV_CONTENT) or ''


def protection_state(names=None):
    """Which of our extra value names are registered as hidden to the
    access role. Returns (ok, missing names, role exists)."""
    xml = sev_xml()
    missing = []
    for name in names or (REQ_NAMES + TEAM_NAMES + QUICK_NAMES):
        m = re.search(r'<Value\s+Name="%s"[^>]*>' % re.escape(name), xml)
        roles = re.search(r'VisibilityRoles="([^"]*)"', m.group(0)) if m else None
        if not roles or access_role().lower() not in [x.strip().lower() for x in roles.group(1).split(',')]:
            missing.append(name)
    role = q.QuerySqlTop1("SELECT COUNT(*) N FROM Roles WHERE RoleName = N'%s'" % sql_str(access_role()))
    return (not missing and bool(role and role.N)), missing, bool(role and role.N)


def role_holders(role):
    """[(pid, name)] of everyone signed up with a TouchPoint role."""
    return [(r.Pid, r.Name) for r in q.QuerySql("""
        SELECT DISTINCT p.PeopleId Pid, ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName Name
        FROM Roles r JOIN UserRole ur ON ur.RoleId = r.RoleId JOIN Users u ON u.UserId = ur.UserId
        JOIN People p ON p.PeopleId = u.PeopleId
        WHERE r.RoleName = N'%s' ORDER BY Name""" % sql_str(role))]


def involvement_check(st):
    """[(level, html)] for the involvements this tool reads; level is 'bad',
    'warn' or 'ok'. The request form's answers (and the documents form's
    files) are the confidential part: each form's involvement must be limited
    to a role (TouchPoint: the involvement's Settings, General, Limit To Role,
    which only admins can change), or anyone who can look up involvements can
    open it and read every answer."""
    out = []
    role = access_role()
    routers = set(int(r.get('pid') or 0) for r in st.get('routers') or [])
    if not int(st.get('form_org') or 0):
        out.append(('bad', 'No request form is picked yet (Forms and people, below).'))
    if not int(st.get('docs_org') or 0):
        out.append(('info', 'No documents form is picked. That is fine if you do not use one. If you do, add it under '
                            'Forms and people so it is checked here too: its uploads are as private as the request form.'))
    for key, what in (('form_org', 'request form'), ('docs_org', 'documents form')):
        org = int(st.get(key) or 0)
        if not org:
            continue
        o = q.QuerySqlTop1("SELECT OrganizationName N, ISNULL(LimitToRole, '') R, ISNULL(PublishDirectory, 0) D, "
                           "ISNULL(IsChatPublic, 0) C FROM Organizations WHERE OrganizationId = %d" % org)
        if not o:
            out.append(('bad', 'The %s\'s involvement (%d) was not found.' % (what, org)))
            continue
        nm = esc(o.N)
        if not o.R:
            out.append(('bad', '<b>%s (the %s) is not locked.</b> Anyone who can look up involvements can open it and read every answer people '
                               'wrote. In TouchPoint, open the involvement, then <b>Settings, General, Limit To Role</b>, pick <b>%s</b> '
                               'and save. Only admins can change that field.' % (nm, what, esc(role))))
        elif o.R.lower() != role.lower():
            holders = role_holders(o.R)
            out.append(('warn', '<b>%s (the %s) is limited to the %s role</b>, which %d %s. Everyone with it can read the answers. '
                                '%s is the safer pick: only the routers should have it.%s' % (
                                    nm, what, esc(o.R), len(holders), 'person has' if len(holders) == 1 else 'people have', esc(role),
                                    (' They include %s.' % esc(', '.join(n for p, n in holders[:12]) + (' and more' if len(holders) > 12 else '')))
                                    if holders else '')))
        else:
            out.append(('ok', '<b>%s (the %s) is locked</b> to the %s role. Only people with that role can open it in TouchPoint.' % (nm, what, esc(role))))
        if o.D:
            out.append(('warn', 'The member directory is published for <b>%s</b>. Everyone who used the %s is on it, so they may be able '
                                'to see each other. Turn the directory off in the involvement\'s settings.' % (nm, what)))
        if o.C:
            out.append(('warn', 'Chat is open to members of <b>%s</b>, which includes everyone who used the %s. Turn it off.' % (nm, what)))
    holders = role_holders(role)
    if holders:
        extra = [n for p, n in holders if p not in routers]
        if extra:
            names = ', '.join(extra[:-1]) + ' and ' + extra[-1] if len(extra) > 1 else extra[0]
            out.append(('warn', '<b>%s %s the %s role but %s.</b> The role opens every request and the form\'s answers '
                                'in TouchPoint. Committee members do not need it; take it off anyone who should not see everything.' % (
                                    esc(names), 'has' if len(extra) == 1 else 'have', esc(role),
                                    'is not a router' if len(extra) == 1 else 'are not routers')))
        else:
            out.append(('ok', 'Only the routers have the %s role (%s).' % (esc(role), esc(', '.join(n for p, n in holders)))))
    corg = int(st.get('committee_org') or 0)
    if corg:
        c = q.QuerySqlTop1("SELECT OrganizationName N, ISNULL(LimitToRole, '') R FROM Organizations WHERE OrganizationId = %d" % corg)
        if c:
            out.append(('ok', '<b>%s</b> (the committee) %s. It holds no request data, only who is on the committee.' % (
                esc(c.N), ('is limited to the %s role' % esc(c.R)) if c.R else 'is not locked, which is fine')))
    return out


def org_name_note(org, empty='not picked yet'):
    """The involvement's name beside its id, so a wrong id shows at once."""
    if not int(org or 0):
        return '<span class="muted">%s</span>' % esc(empty)
    o = q.QuerySqlTop1("SELECT OrganizationName N FROM Organizations WHERE OrganizationId = %d" % int(org))
    return ('<span class="muted">%s</span>' % esc(o.N)) if o else '<b style="color:#A12A22">no involvement has this id</b>'


def who_can_get_in(st):
    """The three gates, with names, for Settings: so a church that installs
    this can see who is in and why without reading the code."""
    holders = dict(role_holders(access_role()))
    routers = [int(r.get('pid') or 0) for r in st.get('routers') or []]
    rn = names_for(routers)
    rows = []
    for p in routers:
        has = p in holders
        rows.append('<div style="display:grid;grid-template-columns:22px 1fr;gap:8px"><b style="color:%s">%s</b><div>%s%s</div></div>' % (
            '#1E6B45' if has else '#A12A22', '&#10003;' if has else '&#10007;', esc(rn.get(p, '')),
            '' if has else ' · <b>does not have the %s role, so cannot get in.</b> Add it in TouchPoint (Admin, Roles).' % esc(access_role())))
    org = int(st.get('committee_org') or 0)
    o = q.QuerySqlTop1("SELECT OrganizationName N FROM Organizations WHERE OrganizationId = %d" % org) if org else None
    comm = committee_members(st)
    return ('<div class="bvc" id="gates"><h2>Who can get in</h2>'
            '<div class="muted">Three gates. Everyone else who opens the tool is turned away.</div>'
            '<div><b>1. Routers</b> see every request. They need their name under Routers below <b>and</b> the ' + role_b() + ' role in TouchPoint '
            '(the access role, set below).</div>%s'
            '<div><b>2. Committee members</b> see only the requests given to them, and any a teammate brings them in on. '
            'The gate is the committee involvement: anyone currently on its roster gets in, whatever their member type '
            '(not pending members or anyone past their inactive date). Drop someone from it and they are out. They do not need the ' + role_b() + ' role.</div>'
            '<div class="muted">%s</div>'
            '<div><b>3. Admins</b> get these Settings only. Being an admin opens no one\'s request; add yourself as a router for that.</div></div>') % (
        ''.join(rows) or '<div class="muted">No routers yet.</div>',
        ('Now: %d %s on <b>%s</b>%s.' % (len(comm), 'person' if len(comm) == 1 else 'people', esc(o.N),
                                        (': ' + esc(', '.join(m['name'] for m in comm[:20])) + (' and more' if len(comm) > 20 else '')) if comm else ''))
        if o else 'No committee involvement picked yet.')


def form_unlocked(st):
    return any(lv == 'bad' for lv, t in involvement_check(st))


def involvement_card(st):
    rows = involvement_check(st)
    icon = {'bad': '&#10007;', 'warn': '!', 'ok': '&#10003;', 'info': 'i'}
    color = {'bad': '#A12A22', 'warn': '#8A5A00', 'ok': '#1E6B45', 'info': '#1F4E8C'}
    worst = 'bad' if any(lv == 'bad' for lv, t in rows) else ('warn' if any(lv == 'warn' for lv, t in rows) else 'ok')
    return ('<div class="bvc %s" id="invchk"><h2>Who can see the form\'s answers</h2>'
            '<div class="muted">Checked each time this page opens. Fix anything marked here in TouchPoint, then reload.</div>%s</div>') % (
        {'bad': 'bad', 'warn': 'bvwarn', 'ok': ''}[worst],
        ''.join('<div style="display:grid;grid-template-columns:22px 1fr;gap:8px"><b style="color:%s">%s</b><div>%s</div></div>' % (
            color[lv], icon[lv], t) for lv, t in rows))


def protect_now():
    """Admin: register (or fix) the hidden extra value names in
    StandardExtraValues2, leaving every other entry as it was."""
    xml = sev_xml() or '<?xml version="1.0" encoding="utf-16"?>\n<Views>\n</Views>'
    for name in REQ_NAMES + TEAM_NAMES + QUICK_NAMES:
        line = '<Value Name="%s" Type="Text" VisibilityRoles="%s" />' % (name, esc(access_role()))
        pat = re.compile(r'<Value\s+Name="%s"[^>]*?/>' % re.escape(name))
        if pat.search(xml):
            xml = pat.sub(line, xml, 1)
            continue
        std = re.search(r'<View\s+Table="People"\s+Location="Standard"[^>]*>', xml)
        if std:
            xml = xml[:std.end()] + '\n    ' + line + xml[std.end():]
        else:
            xml = xml.replace('</Views>', '  <View Table="People" Location="Standard">\n    %s\n  </View>\n</Views>' % line)
    model.WriteContentText(SEV_CONTENT, xml, "")


# ---------------------------------------------------------------- the requester's record
def load_record(pid):
    """{'v': 1, 'requests': {RegPeopleId: {...}}} for one requester, with any
    copies a merge left under another name folded in (and removed)."""
    raw = model.ExtraValueText(int(pid), EV_FIELD) or ''
    rec = load_json(raw, {})
    rec.setdefault('v', 1)
    rec.setdefault('requests', {})
    extra = list(q.QuerySql("""
        SELECT Field, CAST(Data AS nvarchar(max)) Data FROM PeopleExtra
        WHERE PeopleId = %d AND Field IN (%s)""" % (int(pid), ','.join("'%s'" % a for a in EV_ALIASES))))
    if extra and PROTECTED:
        for r in extra:
            other = load_json(r.Data or '', {})
            for k, v in (other.get('requests') or {}).items():
                if k not in rec['requests']:
                    rec['requests'][k] = v
        save_record(pid, rec)
        for r in extra:
            model.DeleteExtraValue(int(pid), r.Field)
    return rec


def save_record(pid, rec):
    if not PROTECTED:
        raise Exception('Benevolence data is not protected yet, so nothing was saved. An admin can fix it in Settings.')
    model.AddExtraValueText(int(pid), EV_FIELD, safe_json(rec))


def all_records():
    """{pid: record} for everyone with one: one query."""
    out = {}
    for r in q.QuerySql("""
            SELECT PeopleId, CAST(Data AS nvarchar(max)) Data FROM PeopleExtra WITH (NOLOCK)
            WHERE Field = '%s' AND Data IS NOT NULL""" % EV_FIELD):
        out[r.PeopleId] = load_json(r.Data or '', {}).get('requests') or {}
    return out


def next_number(records):
    n = FIRST_NUMBER - 1
    for reqs in records.values():
        for r in reqs.values():
            try:
                n = max(n, int(r.get('num') or 0))
            except Exception:
                pass
    return n + 1


# What each right-away notice says (request number and what happened; never
# the person's name, needs or amounts), and who it goes to.
NOTICES = [
    ('assigned', 'A request is given to someone, or passed on to someone else', 'that committee member', '#{num} was given to you'),
    ('shared', 'Someone is brought in on a case, or asked to affirm a decision', 'that person', '#{num}: {by} {what}'),
    ('decision', 'A decision is sent back (help or decline)', 'the main routers', '#{num}: {by} sent their decision'),
    ('handoff', 'A member asks to hand a request off', 'the main routers', '#{num}: {by} asked to hand it off'),
    ('docs', 'Documents come in for a request', 'the case worker', '#{num}: documents came in'),
    ('paid', 'Payouts are confirmed, or it is picked up', 'the case worker', '#{num}: {what}'),
    ('nudge', 'The router checks on a request that has stalled', 'the case worker', '#{num}: {by} is checking on it'),
    ('resent', 'They send the form again and the router adds it to the open case', 'the case worker', '#{num}: they sent the form again; {by} added it to this case'),
    ('combined_away', 'A case you have is combined into someone else\'s', 'the case worker', '#{num}: {by} combined it into {what}; you can stop working it'),
]


def emails_live(st):
    return bool(st.get('email_on')) and st.get('mode') == 'live'


def notice_text(kind, num, by_name, what):
    say = dict((k, t) for k, _, _, t in NOTICES)
    return say[kind].replace('{num}', unicode(num or '?')).replace('{by}', by_name or 'Someone').replace('{what}', what or '')


def notify(r, kind, to, what='', rid=''):
    """Note a right-away notice on the request for these people. Sent at the
    next ScheduledTasks run, or at once when Settings, Emails says right away
    for this kind (then marked sent, so the run does not send it again)."""
    to = sorted(set(int(x) for x in to if x and int(x) != int(ME or 0)))
    if not to:
        return
    e = {'t': now_str(), 'k': kind, 'to': to, 'by': int(ME or 0), 'w': what}
    r.setdefault('events', []).append(e)
    r['events'] = r['events'][-40:]
    st = SETTINGS
    if not (emails_live(st) and st.get('notify_' + kind, kind != 'paid') and (st.get('notify_when') or {}).get(kind) == 'now'):
        return
    names = names_for(to + ([int(ME)] if ME else []))
    text = notice_text(kind, r.get('num'), names.get(int(ME or 0), ''), what)
    sent = []
    for pid in to:
        try:
            send_email(st, pid, 'Benevolence: ' + text, email_html('Hi %s,' % (names.get(pid) or '').split(' ')[0], [('Just now', [text])],
                                                                   'view=req&rid=%s' % rid if rid else ('view=home' if pid in router_ids() else 'view=mine'),
                                                                   'Open it' if rid else 'Open Benevolence'))
            sent.append(pid)
        except Exception as ex:
            email_log('notice failed: %s' % unicode(ex)[:120])
    if sent:
        e['sent'] = sent
        email_log('right-away notices sent at once: %d' % len(sent))


def main_routers(st):
    return [int(x['pid']) for x in st.get('routers') or [] if x.get('main')] or router_ids()


def add_history(r, what, note='', items=None):
    """items: what was given or paid, line by line, kept on the timeline as
    it stood then (never account numbers: the timeline is not routers only)."""
    h = {'t': now_str(), 'by': int(ME or 0), 'what': what, 'note': note}
    if items:
        h['items'] = items
    r.setdefault('history', []).append(h)


def given_lines(st, d):
    """The decision's help lines for the timeline."""
    return [{'need': x.get('need') or '', 'amount': to_number(x.get('amount')) or 0, 'how': x.get('how') or '',
             'deliv': item_deliv(st, d, x), 'payee': x.get('payee') or '', 'address': x.get('address') or ''}
            for x in d.get('items') or [] if x.get('help')]


def history_items_html(hst):
    """Under a timeline entry: a one-line summary of what was given or paid,
    and Details with each line."""
    items = hst.get('items') or []
    if not items:
        return ''
    summary = ' · '.join('%s %s' % (esc(x.get('need') or x.get('how') or ''), money(x.get('amount'))) for x in items)
    rows = []
    for x in items:
        if 'ref' in x or 'held' in x:          # a payout
            then = 'Held for pickup' if x.get('held') else 'Mailed or handed over'
            ref = ('Check ' + x['ref']) if x.get('ref') and x.get('how') and not x.get('card') else (x.get('ref') or '')
            rows.append('<tr><td><b>%s</b></td><td data-l="Amount">%s</td><td data-l="How">%s%s</td><td data-l="Then">%s</td></tr>' % (
                esc(x.get('need') or ''), money(x.get('amount')), esc(x.get('how') or ''), (' · ' + esc(ref)) if ref else '', then))
        else:
            to = 'They pick it up' if x.get('deliv') == 'pickup' else ('Benevolence pays %s' % esc(x['payee']) if x.get('payee') else 'Benevolence pays it')
            rows.append('<tr><td><b>%s</b></td><td data-l="Amount">%s</td><td data-l="How">%s</td><td data-l="Getting it to them">%s%s</td></tr>' % (
                esc(x.get('need') or ''), money(x.get('amount')), esc(x.get('how') or ''), to,
                ('<div class="muted">%s</div>' % esc(x['address'])) if x.get('address') and x.get('deliv') != 'pickup' else ''))
    return ('<div class="muted" style="margin-top:4px">%s</div><details style="margin-top:4px"><summary>Details</summary>'
            '<table><tbody>%s</tbody></table></details>') % (summary, ''.join(rows))


def shared_pids(r):
    """Teammates brought in on a case right now."""
    return [int(x) for x in (r or {}).get('shared') or []]


def share_add(r, who, why=''):
    """Bring a teammate in: they can read the case, join its conversation and
    add notes. 'shares' keeps every one, ended or not, for the router."""
    if who in shared_pids(r):
        return False
    r['shared'] = shared_pids(r) + [who]
    r.setdefault('shares', []).append({'pid': who, 't': now_str(), 'by': int(ME or 0), 'why': unicode(why or '')[:200]})
    return True


def share_end(r, who, why):
    r['shared'] = [x for x in shared_pids(r) if x != who]
    for x in r.get('shares') or []:
        if int(x.get('pid') or 0) == who and not x.get('ended'):
            x['ended'] = now_str()
            x['end_why'] = why


# ---------------------------------------------------------------- the form's answers
def submissions(st, where='', top=400):
    """Completed submissions to the request form, newest first."""
    return list(q.QuerySql("""
        SELECT TOP %d CAST(rp.RegPeopleId AS varchar(36)) Rid, rp.PeopleId Pid,
               CONVERT(varchar(19), ISNULL(rp.CompletedDate, r.CompletedDate), 120) T,
               ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName Name, p.FamilyId Fam
        FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
        JOIN People p ON p.PeopleId = rp.PeopleId
        WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL %s
        ORDER BY ISNULL(rp.CompletedDate, r.CompletedDate) DESC""" % (top, int(st['form_org']), where)))


def strip_tags(s):
    """Plain text from a TouchPoint label: tags out, and its coded characters
    (&quot; &amp; &nbsp; ...) turned back into text, so printing it with esc()
    does not show the codes."""
    t = re.sub(r'<[^>]+>', ' ', unicode(s or ''))
    t = re.sub(r'&#(\d+);', lambda m: unichr(int(m.group(1))) if int(m.group(1)) < 65536 else ' ', t)
    for k, v in (('&nbsp;', ' '), ('&quot;', '"'), ('&#39;', "'"), ('&apos;', "'"), ('&lt;', '<'), ('&gt;', '>'), ('&amp;', '&')):
        t = t.replace(k, v)
    return re.sub(r'\s+', ' ', t).strip()


def answer_value(raw):
    import json
    if raw is None:
        return ''
    try:
        v = json.loads(raw)
    except Exception:
        v = raw
    if isinstance(v, list):
        return ', '.join(unicode(x) for x in v)
    return unicode(v) if v is not None else ''


FORM_WORDS = [('need', 'A need question (answered Yes or No)'), ('amount', 'Amount for that need'),
              ('payee', 'Who that need is paid to'), ('address', 'Payee address for that need'),
              ('account', 'Account or loan number for that need (routers only)'), ('detail', 'More detail on that need'),
              ('total', 'Total asked for'), ('people', 'People in the home'),
              ('elsewhere', 'Asked anyone else for help?'), ('elsewhere_who', 'Who else they asked'),
              ('private', 'Private answers (routers only)'), ('notice', 'Opening notice, not a section')]


def word_test(words, key):
    """A function telling whether a question label has any of the phrases
    set for this key."""
    phrases = [x.strip().lower() for x in (words.get(key) or '').split(',') if x.strip()]
    return lambda label: bool(phrases) and any(x in (label or '').lower() for x in phrases)


def need_name(label, words):
    """'Do you need assistance for your rent expenses?' -> 'Rent'."""
    low = label.lower()
    for x in [x.strip() for x in (words.get('need') or '').split(',') if x.strip()]:
        i = low.find(x.lower())
        if i > -1:
            label = label[i + len(x):]
            break
    label = re.sub(r'(?i)^\s*(assistance|help)\s+(for|with)\s+(your\s+)?', '', label)
    label = re.sub(r'(?i)\s*(expenses?)?\s*\?*\s*$', '', label).strip()
    return label.capitalize() if label else 'Need'


def form_answers(st, rid):
    """Sections of (label, answer) for one submission, plus the summary
    pulled from them: needs [(need, amount, payee)], total, people, other help."""
    return answers_for(st, [rid]).get(rid.upper()) or parse_answers([], st.get('form_words'))


_QUESTIONS = {}


def answers_for(st, rids):
    """form_answers for many submissions in two queries: {RID: parsed}."""
    rids = [r.upper() for r in rids if GUID.match(r or '')]
    if not rids:
        return {}
    org = int(st['form_org'])
    if org not in _QUESTIONS:
        _QUESTIONS[org] = form_questions(org)
    ans = {}
    for i in range(0, len(rids), 200):
        for a in q.QuerySql("""
                SELECT CAST(RegPeopleId AS varchar(36)) Rid, CAST(RegQuestionId AS varchar(36)) Qid, AnswerValue A
                FROM RegAnswer WHERE RegPeopleId IN (%s)""" % ','.join("'%s'" % r for r in rids[i:i + 200])):
            ans.setdefault(a.Rid.upper(), {})[a.Qid.upper()] = a.A
    out = {}
    for rid in rids:
        mine = ans.get(rid, {})
        out[rid] = parse_answers([(qq.Qid, qq.T, qq.Label, mine.get(qq.Qid.upper()), qq.Opt) for qq in _QUESTIONS[org]],
                                 st.get('form_words'), st.get('form_map'))
    return out


def form_questions(org):
    """The form's questions in order, with their choices (for skip logic)."""
    return list(q.QuerySql("""
        SELECT CAST(RegQuestionId AS varchar(36)) Qid, [Order] O, QuestionTypeId T, Label, CAST(Options AS nvarchar(max)) Opt
        FROM RegQuestion WHERE OrganizationId = %d AND ISNULL(IsDisabled, 0) = 0 ORDER BY [Order]""" % int(org)))


# What a question on the request form is (Settings, Reading your form).
FORM_ROLES = [('need', 'A need (Yes or No)'), ('needs', 'A list of needs (each choice picked is one)'),
              ('amount', 'Amount for the need above'), ('payee', 'Who the need above is paid to'),
              ('address', 'Payee address for the need above'), ('account', 'Account or loan number for the need above (routers only)'),
              ('detail', 'More about the need above'), ('total', 'Total asked for'), ('people', 'People in the home'),
              ('elsewhere', 'Asked anyone else for help?'), ('elsewhere_who', 'Who else they asked'),
              ('private', 'Private (routers only)'), ('show', 'Just show it'), ('hide', 'Hide it')]
SECTION_ROLES = [('section', 'A section heading'), ('notice', 'A notice (not a section)')]
NEED_PARTS = ('amount', 'payee', 'address', 'account', 'detail', 'private')


def auto_role(t, label, is_):
    """What the tool works out a question is, from its type and wording."""
    if t == 4:
        return 'notice' if is_['notice'](label) else 'section'
    if t == 5:
        return 'person'
    for k in ('need', 'amount', 'account', 'address', 'payee', 'detail', 'total', 'people', 'elsewhere_who', 'elsewhere', 'private'):
        if is_[k](label):
            return k
    return 'amount' if t == 11 else 'show'


def skip_target(opts):
    """Where a Yes/No question's other answer (No) skips to: a RegQuestionId."""
    import json
    try:
        options = json.loads(opts or '[]') or []
    except Exception:
        return None
    for o in options:
        if (o.get('SkipToId') or '') and (o.get('Text') or o.get('Value') or '').strip().lower() not in ('yes', 'y'):
            return o['SkipToId'].upper()
    return None


def form_layout(questions, words=None, fmap=None):
    """questions: [(qid, type, label, options)] in form order. Returns one dict
    per question: role (the admin's choice, else the tool's), auto (the
    tool's), and owner (the index of the need it belongs to, or None).
    A need's questions are the ones its No answer skips past (the form's own
    logic); without skip logic, the ones up to the next need."""
    words = words or DEFAULT_SETTINGS['form_words']
    fmap = dict((k.upper(), v) for k, v in (fmap or {}).items())
    is_ = dict((k, word_test(words, k)) for k, _ in FORM_WORDS)
    out = []
    for qid, t, lab, opts in questions:
        label = strip_tags(lab)
        auto = auto_role(t, label, is_)
        role = fmap.get((qid or '').upper()) or auto
        if t in (4, 5) and role not in ('section', 'notice', 'person'):
            role = auto
        out.append({'qid': (qid or '').upper(), 't': t, 'label': label, 'role': role, 'auto': auto, 'owner': None, 'opts': opts})
    index = dict((x['qid'], i) for i, x in enumerate(out) if x['qid'])
    starts = [i for i, x in enumerate(out) if x['role'] in ('need', 'needs')]
    for n, i in enumerate(starts):
        nxt = starts[n + 1] if n + 1 < len(starts) else len(out)
        end = index.get(skip_target(out[i]['opts'])) if out[i]['role'] == 'need' else None
        end = end if end is not None and end > i else nxt
        for j in range(i + 1, min(end, nxt)):
            if out[j]['role'] in NEED_PARTS:
                out[j]['owner'] = i
    return out


def skips_taken(rows):
    """The questions this person's own answers skipped, as TouchPoint walks the
    form: a choice whose option says skip to question X jumps there ('end'
    ends the form). rows: [(qid, type, label, raw answer, options)].
    Returns the set of row indexes that were skipped."""
    import json
    index = dict(((r[0] or '').upper(), i) for i, r in enumerate(rows) if r[0])
    skipped, i = set(), 0
    while i < len(rows):
        qid, t, lab, raw, opts = rows[i]
        nxt = i + 1
        if opts and t in (2, 3, 6) and raw not in (None, ''):
            try:
                options = json.loads(opts) or []
            except Exception:
                options = []
            picked = set(x.lower() for x in answer_list(raw))
            for o in options:
                to = (o.get('SkipToId') or '').upper()
                if to and ((o.get('Text') or '').strip().lower() in picked or (o.get('Value') or '').strip().lower() in picked):
                    if to == 'END':
                        nxt = len(rows)
                    elif index.get(to, -1) > i:
                        nxt = index[to]
                    break
        skipped.update(range(i + 1, nxt))
        i = nxt
    return skipped


def answer_list(raw):
    """The choices picked on a multiple-choice answer."""
    import json
    try:
        v = json.loads(raw)
    except Exception:
        v = raw
    if isinstance(v, list):
        return [unicode(x).strip() for x in v if unicode(x).strip()]
    v = unicode(v or '').strip()
    return [x.strip() for x in v.split(',') if x.strip()] if v else []


def parse_answers(rows, words=None, fmap=None):
    """rows: [(qid, QuestionTypeId, Label, AnswerValue, Options)] in form
    order (or the older [(QuestionTypeId, Label, AnswerValue)]). words: the
    church's form phrases; fmap: what the admin said each question is."""
    words = words or DEFAULT_SETTINGS['form_words']
    rows = [r if len(r) == 5 else (None, r[0], r[1], r[2], None) for r in rows]
    lay = form_layout([(r[0], r[1], r[2], r[4]) for r in rows], words, fmap)
    sections, cur = [], {'name': 'Your information', 'items': []}
    needs, by_q = [], {}
    total, people, elsewhere, elsewhere_who = None, '', '', ''
    gone = skips_taken(rows)
    for i, (row, x) in enumerate(zip(rows, lay)):
        t, raw, label, role = row[1], row[3], x['label'], x['role']
        if i in gone and role not in ('section', 'notice'):
            continue                      # their own answers skipped it (an answer left from before they changed it)
        if role in ('section', 'notice'):
            if cur['items']:
                sections.append(cur)
            # The form opens with a notice ("Please Note: ..."), not a section name.
            name = 'Your information' if role == 'notice' else label[:60]
            cur = {'name': name or 'More', 'items': []}
            continue
        val = answer_value(raw).strip()
        if role == 'need' and val.lower() in ('yes', 'no', ''):
            need = {'need': need_name(label, words), 'yes': val.lower() == 'yes', 'amount': None, 'payee': '', 'address': '', 'account': ''}
            by_q[i] = need
            if need['yes']:
                needs.append(need)
        elif role == 'needs':
            picked = answer_list(raw)
            for v in picked:
                needs.append({'need': v[:60], 'yes': True, 'amount': None, 'payee': '', 'address': '', 'account': ''})
            # An amount or payee after a list belongs to it only when one need was picked.
            by_q[i] = needs[-1] if len(picked) == 1 else {'yes': False}
        need = by_q.get(x['owner']) if x['owner'] is not None else None
        if need and need.get('yes'):
            if role == 'amount':
                need['amount'] = to_number(val)
            elif role == 'payee':
                need['payee'] = val
            elif role == 'address':
                need['address'] = val
            elif role == 'account':
                need['account'] = val
            elif role == 'detail' and val:
                need['need'] = '%s (%s)' % (need['need'], val.lower())
        if role == 'total':
            total = to_number(val)
        elif role == 'people':
            people = val
        elif role == 'elsewhere':
            elsewhere = val
        elif role == 'elsewhere_who':
            elsewhere_who = val
        if not val or t == 5 or role in ('hide', 'person'):
            continue
        cur['items'].append({'q': label, 'a': val, 'private': role in ('private', 'account')})
    if cur['items']:
        sections.append(cur)
    return {'sections': sections, 'needs': needs, 'total': total, 'people': people,
            'elsewhere': elsewhere, 'elsewhere_who': elsewhere_who}


def need_summary(st, rid, cache=None):
    f = (cache or {}).get(rid.upper()) or form_answers(st, rid)
    return ', '.join(n['need'] for n in f['needs']) or 'Other', f['total']


# ---------------------------------------------------------------- family and membership
def record_badges(pids):
    """{pid: [(label, color)]}: the badges on each person's TouchPoint record,
    the way the profile header shows them: family position, then each status
    flag set to show as a badge that this user may see. Member status has
    its own column on the request page, so it is not repeated here."""
    pids = [int(x) for x in pids if x]
    out = dict((p, []) for p in pids)
    if not pids:
        return out
    ids = ','.join(str(p) for p in pids)
    ms_of = {}
    for x in q.QuerySql("""
            SELECT p.PeopleId Pid, fp.Description Pos, ms.Description Ms
            FROM People p LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
            LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
            WHERE p.PeopleId IN (%s)""" % ids):
        ms_of[x.Pid] = x.Ms
        if x.Pos:
            out[x.Pid].append((x.Pos, '#337AB7'))
    try:
        flags = list(q.QuerySql("""
            SELECT p.PeopleId Pid, sf.Name, sf.RoleName, sf.ColorCode C
            FROM People p CROSS APPLY dbo.StatusFlagsPerson(p.PeopleId) sf
            WHERE p.PeopleId IN (%s) AND sf.DisplayAsBadge = 1
            ORDER BY p.PeopleId, sf.RunOrder, sf.Name""" % ids))
    except Exception:
        flags = []                       # an older database without the function: positions and status still show
    for f in flags:
        if f.RoleName and not model.UserIsInRole(f.RoleName):
            continue
        if f.Name and f.Name not in [b[0] for b in out[f.Pid]] and f.Name != ms_of.get(f.Pid):
            out[f.Pid].append((f.Name, f.C if re.match(r'^#[0-9A-Fa-f]{3,8}$', f.C or '') else '#5CB85C'))
    return out


def badges_html(badges):
    return ' '.join('<span style="display:inline-block;background:%s;color:#fff;border-radius:4px;padding:1px 7px;margin:1px 2px 1px 0;'
                    'font-size:12px;font-weight:600;white-space:nowrap">%s</span>' % (c, esc(t)) for t, c in badges)


def family_of(pid):
    """Everyone in the requester's TouchPoint family, with membership."""
    return list(q.QuerySql("""
        SELECT p.PeopleId Pid, ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName Name, p.Age,
               ms.Description Ms, p.MemberStatusId MsId, CONVERT(varchar(10), p.JoinDate, 120) Joined,
               fp.Description Pos
        FROM People p JOIN People me ON me.FamilyId = p.FamilyId AND me.PeopleId = %d
        LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
        LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
        ORDER BY p.PositionInFamilyId, p.Age DESC""" % int(pid)))


def family_requests(st, pid):
    """Every submission by anyone in the family, oldest first."""
    return list(q.QuerySql("""
        SELECT CAST(rp.RegPeopleId AS varchar(36)) Rid, rp.PeopleId Pid,
               CONVERT(varchar(19), ISNULL(rp.CompletedDate, r.CompletedDate), 120) T
        FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
        JOIN People p ON p.PeopleId = rp.PeopleId JOIN People me ON me.FamilyId = p.FamilyId AND me.PeopleId = %d
        WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL
        ORDER BY ISNULL(rp.CompletedDate, r.CompletedDate)""" % (int(pid), int(st['form_org']))))


def outcome(r):
    """One line for a request's result from its record (or none)."""
    if not r:
        return ('before', 'Before this tool, outcome unknown')
    stt = r.get('status')
    if stt == 'declined':
        return ('declined', 'Declined' + (': ' + r['decision']['decline_reason'] if (r.get('decision') or {}).get('decline_reason') else ''))
    if stt == 'combined':
        return ('combined', 'Sent again, combined with another request')
    paid = sum(float(p.get('amount') or 0) for p in r.get('payouts') or [])
    if paid:
        return ('paid', 'Helped ' + money(paid) + (' (before this tool)' if r.get('older') else ''))
    return ('open', STATUS_LABEL.get(stt, 'Open'))


STATUS_LABEL = {'new': 'New: give it to a committee member', 'committee': 'With the committee',
                'requester': 'Waiting on the requester', 'approved': 'Approved: confirm payouts',
                'pickup': 'Waiting for pickup', 'done': 'Done', 'declined': 'Declined',
                'combined': 'Combined with another request'}


# ---------------------------------------------------------------- giving it out in turn
def turn_table(st, records):
    """Each committee member with open count, when they last got one, and how
    many this month; sorted so the next in turn is first."""
    month = datetime.date.today().strftime('%Y-%m')
    today = datetime.date.today().strftime('%Y-%m-%d')
    stats = {}
    for reqs in records.values():
        for r in reqs.values():
            for a in r.get('assigned') or []:
                s = stats.setdefault(int(a['pid']), {'open': 0, 'last': '', 'month': 0, 'helping': 0})
                s['last'] = max(s['last'], a.get('t') or '')
                if (a.get('t') or '')[:7] == month:
                    s['month'] += 1
            if r.get('status') in ('committee', 'requester') and r.get('owner'):
                stats.setdefault(int(r['owner']), {'open': 0, 'last': '', 'month': 0, 'helping': 0})['open'] += 1
                for p in shared_pids(r):
                    stats.setdefault(p, {'open': 0, 'last': '', 'month': 0, 'helping': 0})['helping'] += 1
    rows = []
    for m in committee_members(st):
        s = stats.get(m['pid'], {'open': 0, 'last': '', 'month': 0, 'helping': 0})
        away = (st.get('away') or {}).get(str(m['pid'])) or ''
        rows.append(dict(m, open=s['open'], last=s['last'], month=s['month'], helping=s['helping'],
                         away=away if away >= today else ''))
    rows.sort(key=lambda x: (bool(x['away']), x['last'] or '0000', x['open']))
    if rows and not rows[0]['away']:
        rows[0]['next'] = True
    return rows


# ================================================================ pages
CSS = """
<style>
.page-header{display:none!important}
#bv{font-family:'Atkinson Hyperlegible',Verdana,sans-serif;color:#1C2433;font-size:calc(16px*var(--z,1));line-height:1.4;max-width:1100px;margin:0 auto;text-align:left}
#bv h1,#bv h2,#bv th,#bv td,#bv label,#bv b,#bv span,#bv div,#bv p{color:inherit;text-align:inherit}
#bv h1,#bv h2{color:#1C2433!important}
#bv .sz{display:flex;gap:4px;align-items:center;font-size:calc(14px*var(--z,1))}
#bv .sz button{font:inherit;padding:3px 9px;border:1px solid #8A93A6;background:transparent;color:#fff;border-radius:6px;cursor:pointer}
#bv .sz button.on{background:#fff;color:#1C2433}
#bv *{box-sizing:border-box}
#bv, #bv main{max-width:100%}
#bv{overflow-wrap:break-word}
#bv main > *, #bv .bvc{min-width:0}
#bv a{color:#1F4E8C}
#bv .top{background:#1C2433;color:#fff;padding:16px 24px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
#bv .top a{color:#fff;font-weight:700}
#bv .top b{font-size:calc(21px*var(--z,1))}
#bv nav{display:flex;gap:6px;flex-wrap:wrap;border-bottom:2px solid #D9D4CA;padding:10px 12px 0}
#bv nav a{padding:12px 16px;color:#4A5468;text-decoration:none}
#bv nav a.on{color:#1C2433;font-weight:700;border-bottom:5px solid #1F4E8C}
#bv main{background:#F7F5F0;padding:18px;display:flex;flex-direction:column;gap:14px;border-radius:0 0 12px 12px}
#bv h1{margin:0;font-size:calc(23px*var(--z,1));font-weight:700;line-height:1.2}
#bv h2{margin:0;font-size:calc(18px*var(--z,1));font-weight:700}
#bv .bvc{background:#fff;border:2px solid #D9D4CA;border-radius:12px;padding:16px 20px;display:flex;flex-direction:column;gap:10px;margin:0}
#bv .bvc.bvblue{border:3px solid #1F4E8C}
#bv .bvc.bvwarn{background:#FFF8E8;border:3px solid #C98A12}
#bv .bvc.bvbad{background:#FDF0EE;border:3px solid #A12A22}
#bv .bvr{display:flex;justify-content:space-between;align-items:center;gap:14px;flex-wrap:wrap;margin:0}
#bv .muted{color:#4A5468;font-size:calc(15px*var(--z,1))}
#bv .pill{display:inline-block;padding:3px 12px;border-radius:999px;font-size:calc(15px*var(--z,1));font-weight:700}
#bv .p-new{background:#DCE8F7;color:#163A69}
#bv .p-ok{background:#DDF0E5;color:#1E6B45}
#bv .p-warn{background:#FFF8E8;color:#6E4300;border:2px solid #C98A12}
#bv .p-grey{background:#EEEAE2;color:#4A5468}
#bv .p-bad{background:#FDF0EE;color:#A12A22}
#bv .btn{display:inline-block;background:#1F4E8C;color:#fff;border:0;font:inherit;font-weight:700;padding:9px 18px;border-radius:9px;text-decoration:none;cursor:pointer;white-space:nowrap}
#bv .btn:hover{color:#fff;text-decoration:none}
#bv .btn.green{background:#1E6B45}
#bv .btn.ghost{background:#fff;color:#1C2433;border:2px solid #1C2433}
#bv .btn.red{background:#fff;color:#A12A22;border:2px solid #A12A22}
#bv .btn:disabled{opacity:.5;cursor:default}
#bv .boxes{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px}
#bv .boxes a{background:#fff;border:2px solid #D9D4CA;color:#1C2433;text-decoration:none;border-radius:12px;padding:10px 12px;display:flex;flex-direction:column}
#bv .boxes a.mine{background:#1F4E8C;color:#fff;border-color:#1F4E8C}
#bv .boxes .parts{font-size:calc(12px*var(--z,1));opacity:.85;margin-top:2px}
#bv .boxes a.on{outline:3px solid #E0A526;outline-offset:2px;font-weight:700}
#bv .boxes b{font-size:calc(20px*var(--z,1))}
#bv .boxes span{font-size:calc(14px*var(--z,1))}
#bv table{width:100%;border-collapse:collapse;font-size:calc(16px*var(--z,1));color:#1C2433}
#bv th{text-align:left;padding:8px 10px;border-bottom:2px solid #E6E1D8;color:#4A5468;font-weight:700}
#bv td{padding:10px;border-bottom:1px solid #E6E1D8;vertical-align:top}
#bv tr.hi td{background:#FFF8E8}
#bv details.annnew summary{display:inline-block;cursor:pointer;font-weight:700;color:#1F4E8C;padding:4px 0}#bv .chatlist{display:flex;flex-direction:column;gap:8px}#bv a.chatrow{display:block;background:#fff;border:1px solid #E6E1D8;border-radius:12px;padding:10px 14px;color:#1C2433;text-decoration:none}#bv a.chatrow.unread{border-left:5px solid #1F4E8C}#bv .pickgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(200px,100%),1fr));gap:4px;margin:8px 0}#bv .thread{display:flex;flex-direction:column;gap:8px}#bv .bub{max-width:80%;align-self:flex-start;background:#fff;border:1px solid #E6E1D8;border-radius:14px 14px 14px 4px;padding:8px 12px}#bv .bub.me{align-self:flex-end;background:#1F4E8C;color:#fff;border-color:#1F4E8C;border-radius:14px 14px 4px 14px}#bv .bub.me a{color:#fff}#bv .bub.ann{background:#EEF4FB;border-color:#1F4E8C}#bv .bub.me.ann{background:#163A69}#bv .bub.new{box-shadow:0 0 0 2px #E0A526}#bv .bubh{font-size:calc(13px*var(--z,1));opacity:.85;display:flex;gap:8px;flex-wrap:wrap}#bv .chatbox{position:sticky;bottom:0;padding:10px 12px}#bv .chatbox textarea{width:100%;font:inherit;padding:8px;border:2px solid #D9D4CA;border-radius:10px;box-sizing:border-box}#bv .post{background:#fff;border:1px solid #E6E1D8;border-radius:10px;padding:10px 14px;display:flex;flex-direction:column;gap:4px;margin-bottom:8px}#bv .post.ann{border-color:#1F4E8C;background:#EEF4FB}#bv .post.new{border-left:5px solid #1F4E8C}#bv .callgrid{display:grid;grid-template-columns:minmax(0,1fr);gap:14px}#bv .callmain{display:flex;flex-direction:column;gap:14px;min-width:0}#bv .callform{align-self:start;min-width:0}@media (min-width:1000px){#bv .callgrid{grid-template-columns:minmax(0,1.5fr) minmax(0,1fr)}#bv .callform{position:sticky;top:10px;max-height:calc(100vh - 20px);overflow:auto}}#bv .subnav{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px}#bv .subnav a{padding:6px 14px;border:1px solid #D9D4CA;border-radius:999px;text-decoration:none;color:#1C2433;background:#fff}#bv .subnav a.on{background:#1C2433;color:#fff;border-color:#1C2433}#bv .seg{display:inline-flex;border:2px solid #1C2433;border-radius:10px;overflow:hidden;flex-wrap:wrap}#bv .seg button{font:inherit;font-weight:700;border:0;padding:10px 16px;background:#fff;color:#1C2433;cursor:pointer}#bv .seg button.on{background:#1C2433;color:#fff}#bv .bar{height:10px;background:#E6E1D8;border-radius:999px;overflow:hidden;margin:6px 0}#bv .bar div{height:100%;width:0;background:#1F4E8C}#bv .chips{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0}#bv .chip{padding:3px 12px;border-radius:999px;background:#F2F0EB;color:#4A5468;text-decoration:none}#bv .chip.done{background:#DDF0E5;color:#1E6B45;font-weight:700}#bv .chip.done::after{content:' \\2713'}#bv .ct{display:flex;flex-direction:column;gap:14px;padding-bottom:10px;border-bottom:1px solid #E6E1D8;scroll-margin-top:12px}#bv .cq{display:flex;flex-direction:column;gap:6px}#bv .cq textarea{width:100%;font:inherit;padding:10px;border:2px solid #D9D4CA;border-radius:10px;box-sizing:border-box;resize:vertical}#bv .help{background:#EEF4FB;border:1px solid #BFD3EC;border-radius:10px;padding:10px 14px}#bv .help > b{color:#163A69}#bv .help ul{margin:6px 0;padding-left:20px}#bv .help li{margin:3px 0}#bv .bvbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap}#bv .bvbar .sp{flex:1}#bv .qbtn{width:30px;height:30px;border-radius:50%;border:2px solid #fff;background:transparent;color:#fff;font:inherit;font-weight:700;cursor:pointer;padding:0;flex:none;margin-left:8px}#bv .qbtn.on{background:#fff;color:#1C2433}#bv .chip-test,#bv .chip-view{border-radius:999px;padding:2px 10px;font-size:calc(13px*var(--z,1));font-weight:700;white-space:nowrap}#bv .ttl{display:flex;align-items:center;gap:8px;flex-wrap:wrap}#bv .tools{display:flex;align-items:center}#bv .help{margin-bottom:2px}#bv nav .navsw{margin-left:auto;align-self:center;font:inherit;font-size:calc(14px*var(--z,1));font-weight:700;color:#1F4E8C;background:#fff;border:2px solid #1F4E8C;border-radius:999px;padding:5px 14px;cursor:pointer;white-space:nowrap;margin-bottom:6px}#bv .chip-test{background:#FCEFD3;color:#6E4300;border:1px solid #E0A526}#bv .chip-view{background:#DCE8F7;color:#163A69;border:1px solid #1F4E8C}#bv ol.steps li{margin:8px 0}#bv table.cmp{table-layout:fixed}#bv table.kinds textarea{width:100%;min-width:180px}
#bv table.cmp td,#bv table.cmp th{overflow-wrap:anywhere}
#bv .fmtools{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}#bv .fmtools details.bvc{padding:8px 14px;margin:0;flex:1 1 280px}#bv .fmtools details[open]{flex-basis:100%}#bv details.fmsec{border:1px solid #E6E1D8;border-radius:10px;margin:6px 0;background:#fff}#bv details.fmsec>summary{display:flex;justify-content:flex-start;align-items:center;gap:10px;flex-wrap:wrap;text-align:left;padding:8px 12px;cursor:pointer;background:#F2F0EB;border-radius:10px}#bv details.fmsec[open]>summary{border-radius:10px 10px 0 0}#bv details.fmsec table{margin:0}#bv table.docsmap select{width:100%;max-width:100%;box-sizing:border-box}#bv details.fmsec select,#bv .fmprev select{max-width:100%;width:100%;box-sizing:border-box;min-width:0}#bv .fmprev #fmrid{width:auto}#bv .fmgrid{display:grid;grid-template-columns:minmax(0,1fr);gap:14px}#bv .fmlist{min-width:0}#bv .fmprev{min-width:0;align-self:start}@media (min-width:1100px){#bv .fmgrid{grid-template-columns:minmax(0,1.6fr) minmax(0,1fr)}#bv .fmprev{position:sticky;top:10px;max-height:calc(100vh - 20px);overflow:auto}}
#bv table.cmp th.sub{font-size:calc(13px*var(--z,1));font-weight:400}
#bv label.pick{display:flex;gap:10px;align-items:center;cursor:pointer}
#bv input[type=radio],#bv input[type=checkbox]{width:22px;height:22px}
#bv input[type=text],#bv select,#bv textarea{font:inherit;padding:10px 12px;border:2px solid #D9D4CA;border-radius:10px;background:#fff}
#bv details{border-bottom:1px solid #E6E1D8}#bv .sec{border-top:1px solid #E6E1D8;margin-top:12px;padding-top:12px}#bv .tcbtn{position:fixed;right:max(18px,env(safe-area-inset-right,0px));bottom:calc(18px + env(safe-area-inset-bottom,0px));z-index:100000;font:inherit;font-weight:700;font-size:calc(18px*var(--z,1));background:#1C2433;color:#fff;border:0;border-radius:999px;padding:14px 22px;box-shadow:0 4px 14px rgba(0,0,0,.25);cursor:pointer}#bv .tcbtn.hasnew{background:#B4570B}#bv .tcbtn.pulse{animation:tcpulse 1.2s ease-out 1}@keyframes tcpulse{0%{transform:scale(1)}30%{transform:scale(1.12)}100%{transform:scale(1)}}#bv .tcpanel{position:fixed;z-index:100001;right:max(18px,env(safe-area-inset-right,0px));bottom:calc(18px + env(safe-area-inset-bottom,0px));width:min(440px,calc(100vw - 36px));height:min(600px,calc(100vh - 160px));background:#fff;border:2px solid #1C2433;border-radius:14px;overflow:hidden;box-shadow:0 10px 32px rgba(0,0,0,.28);display:flex;flex-direction:column;font-size:calc(17px*var(--z,1))}@media (max-width:640px){#bv .tcpanel{top:0;right:0;bottom:0;left:0;width:auto;height:auto;border:0;border-radius:0}}#bv .tcpanel[hidden]{display:none}#bv .tchead{display:flex;align-items:center;gap:10px;padding:calc(10px + env(safe-area-inset-top,0px)) 14px 10px;border-bottom:2px solid #E6E1D8;background:#F7F5F0}#bv .tchead .t{flex:1;min-width:0}#bv .tcbody{flex:1;overflow:auto;padding:12px 16px;display:flex;flex-direction:column;gap:8px}#bv .tcfoot{border-top:2px solid #E6E1D8;padding:10px 16px calc(12px + env(safe-area-inset-bottom,0px));display:flex;flex-direction:column;gap:8px}#bv .tcfoot textarea{width:100%;font:inherit;padding:10px;border:2px solid #D9D4CA;border-radius:10px;box-sizing:border-box}#bv .tcclose{font:inherit;font-weight:700;font-size:calc(17px*var(--z,1));background:#fff;color:#1C2433;border:2px solid #1C2433;border-radius:10px;padding:8px 16px;cursor:pointer}#bv .backbar{position:sticky;top:env(safe-area-inset-top,0px);z-index:5;background:#F7F5F0;padding:8px 0;display:flex;gap:10px;align-items:center;flex-wrap:wrap}#bv .backbar .btn{font-size:calc(17px*var(--z,1))}#bv .rpgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(170px,100%),1fr));gap:10px}#bv .rpbox{border:2px solid #E6E1D8;border-radius:12px;padding:12px 16px}#bv .rpbox b{display:block;font-size:calc(26px*var(--z,1))}#bv #rpt h3{margin:18px 0 6px}#bv .tw{overflow-x:auto}#bv details.cm{display:block;border:2px solid #D9D4CA}#bv details.cm > summary{padding:2px 0;flex-wrap:wrap}#bv .cmhead{display:flex;gap:12px;align-items:baseline}#bv .cmhead .lnk{font-size:calc(14px*var(--z,1));font-weight:700;text-decoration:none}#bv .shform{background:#F7F5F0;border-radius:10px;padding:10px 12px;margin:6px 0;display:flex;flex-direction:column;gap:6px}#bv .cthread{display:flex;flex-direction:column;gap:8px;margin:6px 0;max-height:420px;overflow:auto}#bv .bub .to{font-size:calc(12px*var(--z,1));opacity:.8}#bv .bub .rep{font-size:calc(13px*var(--z,1))}#bv .bub.me .rep{display:none}#bv .replying{background:#EEF4FB;border-radius:8px;padding:4px 10px;font-size:calc(13px*var(--z,1))}#bv .cml{font-size:calc(12px*var(--z,1));letter-spacing:.06em;text-transform:uppercase;color:#6B7385;font-weight:700;margin-bottom:6px}#bv .cmrow{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:3px 0}#bv .lnk{background:none;border:0;color:#1F4E8C;font:inherit;text-decoration:underline;cursor:pointer;padding:0}#bv details.sec{border-bottom:0;padding-top:0}#bv details.flat{border-bottom:0}
#bv details summary{cursor:pointer;font-weight:700;padding:10px 0;list-style:none;display:flex;gap:10px}
#bv details summary::-webkit-details-marker{display:none}
#bv details summary::before{content:'\\25B8';color:#1F4E8C}
#bv details[open] summary::before{content:'\\25BE'}
#bv .qa{display:grid;grid-template-columns:1fr 1fr;gap:4px 16px;font-size:calc(15px*var(--z,1));padding:4px 0 12px 22px}
#bv .qa .a{font-weight:700}
#bv .rt-bar{display:flex;gap:6px;flex-wrap:wrap;background:#F2F0EB;border:2px solid #D9D4CA;border-bottom:0;border-radius:10px 10px 0 0;padding:8px;align-items:center}
#bv .rt-bar button{font:inherit;font-size:calc(16px*var(--z,1));padding:5px 11px;border:2px solid #D9D4CA;background:#fff;border-radius:8px;cursor:pointer}
#bv .rt-bar .saved{margin-left:auto;font-size:calc(15px*var(--z,1));color:#1E6B45;font-weight:700}
#bv .rt-bar .saved.not{color:#A12A22}
#bv .rt{border:2px solid #D9D4CA;border-radius:0 0 10px 10px;padding:12px;min-height:90px;background:#fff;outline:none;resize:vertical;overflow:auto}
#bv .rt-wrap.full{position:fixed;inset:0;z-index:2000;background:#fff;padding:20px;display:flex;flex-direction:column}
#bv .rt-wrap.full .rt{flex:1;overflow:auto}
#bv .msg{padding:10px 14px;border-radius:10px;font-weight:700}
#bv .msg.ok{background:#DDF0E5;color:#1E6B45}
#bv .msg.bad{background:#FDF0EE;color:#A12A22}
@media (max-width:700px){
  #bv{font-size:calc(17px*var(--z,1))}
  #bv .top{padding:8px 12px;flex-wrap:nowrap;border-radius:0}
  #bv .top b{font-size:calc(18px*var(--z,1))}
  #bv .top .who{display:none}
  #bv .sz .szl{display:none}
  #bv .sz button{padding:2px 7px}
  #bv nav{flex-wrap:nowrap;overflow-x:auto;-webkit-overflow-scrolling:touch;padding:4px 6px 0;gap:0;scrollbar-width:none}
  #bv nav a{padding:10px 12px;white-space:nowrap}
  #bv .bvr{flex-wrap:wrap}
  #bv .bvr label{flex-wrap:wrap;min-width:0}
  #bv input, #bv select, #bv textarea{max-width:100%}
  #bv .btn{white-space:normal;max-width:100%;text-align:center}
  #bv main{padding:12px}
  #bv .qa{grid-template-columns:1fr}
  #bv table, #bv thead, #bv tbody, #bv tr, #bv td{display:block;width:100%}
  #bv thead{display:none}
  #bv tr{border-bottom:2px solid #E6E1D8;padding:6px 0}
  #bv td{border:0;padding:3px 6px}
  #bv td[data-l]{display:inline-block;width:auto;margin-right:14px}
  #bv td[data-l]::before{content:attr(data-l) ': ';color:#4A5468}
}
</style>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Atkinson+Hyperlegible:wght@400;700&display=swap">
"""

# Small shared page script: posting actions, and the formatted-notes box
# (bold, italic, underline, lists, link, clear; saved as you type; full page).
PAGE_JS = r"""
<script>
(function(){
  var base = location.pathname.replace('/PyScript/', '/PyScriptForm/');
  window.bvPost = function(data, cb){
    var x = new XMLHttpRequest();
    x.open('POST', base, true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){ var r; try { r = JSON.parse(x.responseText); } catch(e) {
      r = {success: false, message: x.status == 200 ? 'TouchPoint answered with a page, not data: you may need to sign in again.' : 'TouchPoint error ' + x.status}; }
      cb(r); };
    x.onerror = function(){ cb({success: false, message: 'Could not reach TouchPoint.'}); };
    x.send(Object.keys(data).map(function(k){ return k + '=' + encodeURIComponent(data[k]); }).join('&'));
  };
  window.bvB64 = function(s){ return btoa(unescape(encodeURIComponent(s || ''))); };
  // Leaving the page: send what is still waiting to save, so the last few
  // words typed are not lost (a normal post is cut off as the page goes).
  window.bvBeacon = function(data){
    var body = Object.keys(data).map(function(k){ return k + '=' + encodeURIComponent(data[k]); }).join('&');
    try { if (navigator.sendBeacon && navigator.sendBeacon(base, new Blob([body], {type: 'application/x-www-form-urlencoded'}))) return; } catch(e) {}
    bvPost(data, function(){});
  };
  window.bvGo = function(q){ location.href = base + q; };
  // Text size: each person's choice, kept in this browser.
  // Text size scales the type only (--z), never the page box, so the width stays put.
  function zoom(z){ var bv = document.getElementById('bv'); if (!bv) return; bv.style.setProperty('--z', z);
    Array.prototype.forEach.call(bv.querySelectorAll('.sz button'), function(b){ b.className = b.getAttribute('data-z') == String(z) ? 'on' : ''; }); }
  var z0 = '1'; try { z0 = localStorage.getItem('bv.size') || '1'; } catch(e) {}
  zoom(z0);
  Array.prototype.forEach.call(document.querySelectorAll('#bv .sz button'), function(b){
    b.onclick = function(){ var z = b.getAttribute('data-z'); try { localStorage.setItem('bv.size', z); } catch(e) {} zoom(z); }; });
  // Formatted notes: every .rt-wrap holds a toolbar and a contenteditable box.
  // data-save="action:key" posts the text a second after typing stops, and
  // keeps it in this browser until TouchPoint has it.
  function wire(w){
    var box = w.querySelector('.rt'), saved = w.querySelector('.saved'), save = w.getAttribute('data-save');
    var local = 'bv.draft.' + (save || ''), timer = null;
    try { var kept = localStorage.getItem(local); if (save && kept && kept != box.innerHTML) { box.innerHTML = kept; push(); } } catch(e) {}
    Array.prototype.forEach.call(w.querySelectorAll('[data-cmd]'), function(b){
      b.onmousedown = function(e){ e.preventDefault(); };
      b.onclick = function(){
        var c = b.getAttribute('data-cmd');
        if (c == 'full') { w.classList.toggle('full'); b.textContent = w.classList.contains('full') ? 'Done' : 'Full page'; box.focus(); return; }
        if (c == 'createLink') { var u = prompt('Link address (https://...)'); if (u) document.execCommand('createLink', false, u); }
        else document.execCommand(c, false, null);
        box.focus(); changed();
      };
    });
    function changed(){
      if (!save) return;
      try { localStorage.setItem(local, box.innerHTML); } catch(e) {}
      if (saved) { saved.textContent = 'Saving...'; saved.className = 'saved'; }
      clearTimeout(timer); timer = setTimeout(push, 1000);
    }
    window.addEventListener('pagehide', function(){
      if (!save || !timer) return;
      clearTimeout(timer); timer = null;
      var p = save.split(':');
      bvBeacon({action: p[0], rid: p[1], field: p[2] || '', html: bvB64(box.innerHTML)});
    });
    function push(){
      if (!save) return;
      var p = save.split(':');
      bvPost({action: p[0], rid: p[1], field: p[2] || '', html: bvB64(box.innerHTML)}, function(r){
        if (r.success) { try { localStorage.removeItem(local); } catch(e) {} if (saved) { saved.textContent = 'Saved'; saved.className = 'saved'; } }
        else if (saved) { saved.textContent = 'Not saved yet: sign in again'; saved.className = 'saved not'; }
      });
    }
    box.addEventListener('input', changed);
    w.bvValue = function(){ return box.innerHTML; };
  }
  Array.prototype.forEach.call(document.querySelectorAll('#bv .rt-wrap'), wire);
})();
</script>
"""


def rich_box(html, save='', minh=90):
    """The formatted-notes box. save='action:rid:field' autosaves."""
    return ('<div class="rt-wrap" data-save="%s"><div class="rt-bar">'
            '<button type="button" data-cmd="bold"><b>B</b></button><button type="button" data-cmd="italic"><i>I</i></button>'
            '<button type="button" data-cmd="underline"><u>U</u></button><button type="button" data-cmd="insertUnorderedList">&bull; List</button>'
            '<button type="button" data-cmd="insertOrderedList">1. List</button><button type="button" data-cmd="createLink">Link</button>'
            '<button type="button" data-cmd="removeFormat">Clear</button>%s'
            '<button type="button" data-cmd="full">Full page</button></div>'
            '<div class="rt" contenteditable="true" style="min-height:%dpx">%s</div></div>') % (
        esc(save), '<span class="saved">Saved</span>' if save else '<span class="saved"></span>', minh, clean_html(html))


def shell(role, active, body, title='Benevolence'):
    me = names_for([ME]).get(ME, '') if ME else ''
    tabs = []
    if role == 'router':
        tabs = [('home', 'Waiting on you'), ('all', 'All requests'), ('walkins', 'Walk-ins'), ('reports', 'Reports')]
        if IS_MEMBER:
            tabs.append(('mine', "People I'm helping"))
    if role == 'committee':
        tabs = [('mine', "People I'm helping")]
    if role in ('router', 'committee'):
        try:
            n = team_unread()
        except Exception:
            n = 0
        tabs.append(('team', 'Messages' + (' (%d new)' % n if n else '')))
    if IS_ADMIN and not VIEW_AS:
        tabs.append(('settings', 'Settings'))
    tabs.append(('help', 'How it works'))
    nav = ''.join('<a href="?view=%s" class="%s">%s</a>' % (k, 'on' if k == active else '', esc(l)) for k, l in tabs)
    banner = ''
    if IS_ADMIN and not VIEW_AS and active != 'settings':
        try:
            unlocked = form_unlocked(SETTINGS)
        except Exception:
            unlocked = False
        if unlocked:
            banner = ('<div class="bvc bad"><b>The request form\'s answers are not locked down.</b>'
                      '<span>Admins: <a href="?view=settings&amp;sec=setup#invchk">see Settings, Setup</a> for what to fix.</span></div>')
    if not PROTECTED:
        banner = ('<div class="bvc bvbad"><b>Benevolence data is not protected, so nothing can be saved.</b>'
                  '<span>%s</span></div>' % ('The access role ' + role_b() + ' does not exist in TouchPoint. Create it (Admin, Roles), or pick a role you already have in Settings, Setup; then reload.' if IS_ADMIN
                                              else 'Ask an admin to open Benevolence once; it protects the data by itself.'))
    size = ('<span class="sz"><span class="szl">Text size</span> <button type="button" data-z="0.85" title="Smaller text">A-</button>'
            '<button type="button" data-z="1" title="Standard text">A</button><button type="button" data-z="1.15" title="Larger text">A+</button></span>')
    # No extra rows: Testing and Committee view are tags beside the title, the
    # ? sits with the text size buttons, and the view switch ends the tab row.
    tags = ''
    if SETTINGS.get('mode') != 'live':
        tags += ('<span class="chip-test" title="Nothing is emailed, and admins can clear this data in Settings. '
                 'An admin switches to Live in Settings when the team is ready.">Testing</span>')
    if VIEW_AS:
        tags += ('<span class="chip-view" title="You are seeing exactly what a committee member sees, with your own assignments.">'
                 'Committee view</span>')
    q = ('<button type="button" class="qbtn" id="helpq" title="How this page works" aria-label="How this page works" '
         'aria-expanded="false">?</button>') if HELP_BOX[0] else ''
    if VIEW_AS:
        nav += '<button type="button" class="navsw" data-as="">Back to router view</button>'
    elif REAL_ROLE == 'router' and IS_MEMBER:
        nav += '<button type="button" class="navsw" data-as="committee">See the committee view</button>'
    banner += HELP_BOX[0]
    banner += ('<script>Array.prototype.forEach.call(document.querySelectorAll("#bv [data-as]"),function(b){b.onclick=function(){'
               'b.disabled=true; bvPost({action:"view_as",as:b.getAttribute("data-as")},function(r){if(r.success) location.href="?view="+(b.getAttribute("data-as")?"mine":"home");'
               ' else {b.disabled=false; alert(r.message);}});};});</script>')
    return (CSS + '<div id="bv"><div class="top"><span class="ttl"><b>%s</b>%s</span><span class="tools">%s%s</span>'
            '<span class="who">Signed in as %s</span></div><nav>%s</nav><main>%s%s%s<div class="muted" style="text-align:right;margin-top:18px">TPxi Benevolence %s</div></main></div>'
            % (esc(title), tags, size, q, esc(me), nav, update_check_js(), banner, body, esc(BEN_VERSION)) + PAGE_JS)


def dc_headers():
    """Sent with the update download: this church's TouchPoint address and the
    script and version, so the TPxi script service can say who is on what. No
    person, nothing from the database. The admin setting TPxiSendHost = 0
    turns it off (shared with the other TPxi scripts)."""
    h = {}
    try:
        if str(model.Setting("TPxiSendHost", "1")).strip() in ("0", "false", "False"):
            return h
    except Exception:
        pass
    try:
        host = str(model.CmsHost or "").strip()
    except Exception:
        host = ""
    if host:
        h["X-CmsHost"] = host
    h["X-TPxi-Script"] = DC_SCRIPT_ID
    h["X-TPxi-Version"] = BEN_VERSION
    return h


def apply_update():
    """Admin only (ACTION_ACCESS). Saves the newest TPxi Benevolence over this
    script, under whatever name it was saved as."""
    name = re.sub(r'[^\w.\- ]', '', _param('script_name') or '') or DC_SCRIPT_ID
    try:
        code = model.RestGet(DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID, dc_headers()) or ''
    except Exception as e:
        return {'success': False, 'message': 'Could not download the update: %s' % e}
    code = unicode(code)
    # Never save something that is not this script: an error page, an empty
    # reply or another tool would break the page for everyone.
    if len(code) < 5000 or 'BEN_VERSION' not in code or 'def page_request' not in code:
        return {'success': False, 'message': 'The download did not look like TPxi Benevolence, so nothing was changed.'}
    m = re.search(r"BEN_VERSION = '([\d.]+)'", code)
    model.WriteContentPython(name, code)
    return {'success': True, 'version': m.group(1) if m else '', 'message': 'Updated %s. Reloading.' % name}


def update_check_js():
    """Admins: ask the script service which version is newest, in the browser
    (TouchPoint's own server calls are often challenged by Cloudflare), and
    offer Update now only when it is NEWER (compared number by number, so
    0.10.0 beats 0.9.0; a copy ahead of what is published says nothing)."""
    if not IS_ADMIN:
        return ''
    return ('<div id="bvupd"></div><script>(function(){var MINE=%s, SID=%s;'
            'function cmp(a,b){var x=String(a||"").split("."),y=String(b||"").split(".");'
            ' for(var i=0;i<Math.max(x.length,y.length);i++){var n=parseInt(x[i],10)||0,m=parseInt(y[i],10)||0; if(n!==m) return n>m?1:-1;} return 0;}'
            'function show(v){var el=document.getElementById("bvupd"); if(!el) return;'
            ' el.innerHTML=\'<div class="bvc bvblue"><div class="bvr"><span><b>A newer Benevolence is out: \'+v+\'</b> <span class="muted">You have \'+MINE+\'. '
            'Updating keeps every request, setting and note.</span></span><span><button type="button" class="btn" id="bvupdgo">Update now</button> '
            '<span id="bvupdst" class="muted"></span></span></div></div>\';'
            ' document.getElementById("bvupdgo").onclick=function(){var b=this,st=document.getElementById("bvupdst"); b.disabled=true; st.textContent="Updating...";'
            '  var parts=location.pathname.split("/"), name=""; for(var i=0;i<parts.length;i++){ if(parts[i]=="PyScriptForm"||parts[i]=="PyScript") name=decodeURIComponent(parts[i+1]||""); }'
            '  bvPost({action:"apply_update",script_name:name},function(r){st.textContent=r.message||""; if(r.success) setTimeout(function(){location.reload();},800); else b.disabled=false;});};}'
            'try{var x=new XMLHttpRequest(); x.open("GET",%s,true); x.timeout=5000;'
            ' x.onreadystatechange=function(){ if(x.readyState!==4||x.status!==200) return; try{var v=JSON.parse(x.responseText)[SID]; if(v&&cmp(v,MINE)>0) show(v);}catch(e){}}; x.send();}catch(e){}'
            '})();</script>') % (safe_json(BEN_VERSION), safe_json(DC_SCRIPT_ID), safe_json(DC_API_BASE + '/script-versions'))


def page_denied():
    return CSS + ('<div id="bv"><main><div class="bvc"><h1>Benevolence</h1><p>This is for the benevolence routers and '
                  'committee. If you should have access, ask an admin to add you.</p></div></main></div>')


# ---------------------------------------------------------------- the router's home
DEFAULT_CUTOFF_DAYS = 30


def cutoff(st):
    """Requests from before this date are 'older': not on the router's home unless
    he brings one in. The go-live date in Settings, else the last 30 days."""
    return st.get('go_live') or (datetime.date.today() - datetime.timedelta(days=DEFAULT_CUTOFF_DAYS)).strftime('%Y-%m-%d')


def current_subs(st, records, top=300):
    """Submissions since the cutoff, plus older ones the router brought in."""
    cut = cutoff(st)
    subs = submissions(st, "AND ISNULL(rp.CompletedDate, r.CompletedDate) >= '%s'" % sql_str(cut), top)
    brought = [rid for reqs in records.values() for rid, r in reqs.items() if r.get('brought_in')]
    have = set(x.Rid for x in subs)
    brought = [b for b in brought if b not in have and GUID.match(b)]
    if brought:
        subs += submissions(st, "AND rp.RegPeopleId IN (%s)" % ','.join("'%s'" % b for b in brought), len(brought))
    return subs


def waiting_parts(by_status, today):
    """'41 new · 2 payouts · 1 message': what the Waiting on you count is made of."""
    open_ = by_status.get('committee', []) + by_status.get('requester', [])
    parts = [(len(by_status.get('new', [])), 'new', 'new'),
             (sum(1 for x in open_ if (x[1] or {}).get('handoff')), 'handoff ask', 'handoff asks'),
             (sum(1 for x in open_ if not (x[1] or {}).get('handoff') and req_unread(x[0].Rid, x[1])), 'message', 'messages'),
             (len(by_status.get('approved', [])), 'payout', 'payouts'),
             (sum(1 for x in by_status.get('pickup', []) if pickup_due(x[1]) and pickup_due(x[1]) <= today), 'pickup due', 'pickups due')]
    return ' · '.join('%d %s' % (n, one if n == 1 else many) for n, one, many in parts if n) or 'nothing right now'


def waiting_on_router(by_status, today):
    """Everything the router needs to act on: new requests, handoff asks,
    new messages, payouts to confirm, and pickups due today or missed."""
    return by_status.get('new', []) + [x for x in by_status.get('committee', []) + by_status.get('requester', [])
                                       if (x[1] or {}).get('handoff') or req_unread(x[0].Rid, x[1]) or is_stuck(SETTINGS, x[0], x[1])] + by_status.get('approved', []) + [
        x for x in by_status.get('pickup', []) if pickup_due(x[1]) and pickup_due(x[1]) <= today]


def stage_boxes(by_status, active, waiting=None, parts=''):
    """The stage boxes. They stay on top: clicking one lists that stage under
    them, so you can click from box to box. active: the stage being shown
    ('' on Waiting on you)."""
    month = datetime.date.today().strftime('%Y-%m')
    n = lambda k: len(by_status.get(k, []))
    done_month = sum(1 for s, r in by_status.get('done', []) if ((r or {}).get('closed') or '')[:7] == month)
    declined_month = sum(1 for s, r in by_status.get('declined', []) if ((r or {}).get('closed') or '')[:7] == month)
    boxes = [('', 'Waiting on you', waiting, True), ('new', 'New: to give out', n('new'), True), ('committee', 'With the committee', n('committee'), False),
             ('requester', 'Waiting on requester', n('requester'), False), ('approved', 'Approved: confirm payouts', n('approved'), True),
             ('pickup', 'Waiting for pickup', n('pickup'), False), ('done', 'Helped this month', done_month, False),
             ('declined', 'Declined this month', declined_month, False)]
    return ('<div class="boxes">%s</div><div class="muted"><b>Waiting on you</b> is your to-do list. The other boxes count every '
            'request by stage; blue ones need you. Click any box to see who is in it.</div>') % ''.join(
        '<a href="%s" class="%s%s"><b>%s</b><span>%s</span>%s</a>' % (
            '?view=home' if not k else '?view=home&amp;stage=%s' % k, 'mine' if mine else '', ' on' if k == active else '',
            '&#8962;' if c is None else c, esc(l), ('<span class="parts">%s</span>' % esc(parts)) if not k and parts else '')
        for k, l, c, mine in boxes)


STAGE_TITLES = {'new': 'New: to give out', 'committee': 'With the committee', 'requester': 'Waiting on requester',
                'approved': 'Approved: confirm payouts', 'pickup': 'Waiting for pickup', 'done': 'Helped this month',
                'declined': 'Declined this month'}


def request_card(s, r, need, total, tag, fam_count, sub, member='', button=''):
    """One request as a card: the Waiting on you look, used for every stage."""
    num = ('#%s · ' % r['num']) if r and r.get('num') else ''
    if member:
        tag += ' <span class="pill %s">%s</span>' % ('p-ok' if member.lower() == 'member' else 'p-grey', esc(member))
    rep = '<span class="pill p-warn">Family asked %d times this year</span>' % fam_count if fam_count > 1 else ''
    return ('<div class="bvc"><div class="bvr"><div style="display:flex;flex-direction:column;gap:6px">'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><b style="font-size:calc(23px*var(--z,1))">%s%s</b>%s%s</div>'
            '<div>%s · <b>%s requested</b></div><div class="muted">%s</div></div>'
            '%s</div></div>') % (
        num, esc(s.Name), tag, rep, esc(need), money(total) or 'amount not given', sub,
        button or '<a class="btn" href="?view=req&amp;rid=%s">Open request</a>' % s.Rid)


def stage_cards(st, records, by_status, stage):
    """Everyone in one stage, as cards, with what matters at that stage."""
    month = datetime.date.today().strftime('%Y-%m')
    today = datetime.date.today().strftime('%Y-%m-%d')
    items = list(by_status.get(stage, []))
    if stage in ('done', 'declined'):
        items = [x for x in items if ((x[1] or {}).get('closed') or '')[:7] == month]
        items.sort(key=lambda x: (x[1] or {}).get('closed') or '', reverse=True)
    elif stage in ('committee', 'requester'):
        items.sort(key=lambda x: last_touch(x[1], x[0]))
    elif stage == 'pickup':
        items.sort(key=lambda x: pickup_due(x[1]) or '9999')
    h = ['<h1>%s (%d)</h1>' % (esc(STAGE_TITLES[stage]), len(items))]
    if not items:
        h.append('<div class="bvc"><b>No one here right now.</b></div>')
        return ''.join(h)
    nexts = turn_table(st, records)
    nxt = nexts[0]['name'] if nexts and nexts[0].get('next') else ''
    owners = names_for([int((r or {}).get('owner') or 0) for s, r in items])
    fam_counts = family_counts(st, list(set(s.Fam for s, r in items)))
    cache = answers_for(st, [s.Rid for s, r in items])
    ms = membership_of([s.Pid for s, r in items])
    week_ago = (datetime.date.today() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
    for s, r in items:
        r = r or {}
        need, total = need_summary(st, s.Rid, cache)
        sub = 'Came in %s' % days_ago(s.T)
        tag = ''
        if stage == 'new':
            tag = '<span class="pill p-new">New: pick a committee member</span>'
            sub += (' · %s is next in turn' % esc(nxt)) if nxt else ''
        elif stage in ('committee', 'requester'):
            who, where = with_whom(r, owners)
            lt = last_touch(r, s)
            tag = '<span class="pill %s">With %s</span>' % ('p-warn' if lt[:10] <= week_ago else 'p-new', esc(who))
            sub = '%s · last change %s · came in %s' % (esc(where[:1].upper() + where[1:]), days_ago(lt), days_ago(s.T))
            if shared_pids(r):
                sub += ' · also on it: %s' % esc(', '.join(sorted(names_for(shared_pids(r)).values())))
        elif stage == 'approved':
            d = r.get('decision') or {}
            tag = '<span class="pill p-ok">Approved %s</span>' % money(sum(to_number(x.get('amount')) or 0 for i, x in given_items(r)))
            sub = 'Approved by %s %s' % (esc(names_for([d.get('by')]).get(int(d.get('by') or 0), '')), days_ago(d.get('sent') or ''))
        elif stage == 'pickup':
            d = r.get('decision') or {}
            late = pickup_due(r) and pickup_due(r) < today
            tag = '<span class="pill %s">%s</span>' % ('p-warn' if late else 'p-new', 'Missed pickup' if late else
                                                    ('Pickup today' if pickup_due(r) == today else 'Pickup ' + esc(pickup_when(d) or 'not set')))
            sub = '%s, %s · %s' % (esc(d.get('pickup') or ''), esc(pickup_when(d) or 'no time set'),
                                   ', '.join('%s %s' % (esc(x.get('how') or ''), money(to_number(paid_amount(r, i, x)))) for i, x in pickup_items(st, r)))
        elif stage == 'done':
            paid = sum(float(p.get('amount') or 0) for p in r.get('payouts') or [])
            tag = '<span class="pill p-ok">Helped %s</span>' % money(paid)
            sub = 'Finished %s' % nice_date(r.get('closed'))
        elif stage == 'declined':
            d = r.get('decision') or {}
            tag = '<span class="pill p-grey">Declined</span>'
            sub = '%s%s · %s' % (esc(d.get('decline_reason') or ''), (', sent to ' + esc(d['referral'])) if d.get('referral') else '',
                                 nice_date(r.get('closed')))
        h.append(request_card(s, r, need, total, tag, fam_counts.get(s.Fam, 0), sub, ms.get(s.Pid, '')))
    return ''.join(h)


def given_note(records):
    """'#12 is with Pat.' after the router gives one out and lands back here."""
    rid = _param('given')
    if not re.match(r'^[0-9A-Fa-f\-]{36}$', rid or ''):
        return ''
    for reqs in records.values():
        for k, r in reqs.items():
            if k.upper() == rid.upper() and r.get('owner'):
                return ('<div class="bvc" style="border-left:4px solid #1E6B45"><b>#%s is with %s.</b> <span class="muted">They get it on their list. '
                        '<a href="?view=req&amp;rid=%s">Open it again</a></span></div>') % (
                    r.get('num'), esc(names_for([r['owner']]).get(int(r['owner']), '')), k)
    return ''


def page_home(st, role):
    docs_sweep(st)
    records = all_records()
    subs = current_subs(st, records, 200)
    counts = {}
    by_status = {}
    for s in subs:
        r = (records.get(s.Pid) or {}).get(s.Rid)
        stt = (r or {}).get('status') or 'new'
        counts[stt] = counts.get(stt, 0) + 1
        by_status.setdefault(stt, []).append((s, r))
    stage = _param('stage') if _param('stage') in STAGE_TITLES else ''
    today = datetime.date.today().strftime('%Y-%m-%d')
    # Sent the form twice, both new: one card (the newer), and counted once.
    reps = repeat_pairs(st, [x for v in by_status.values() for x in v])
    news = set(x[0].Rid.upper() for x in by_status.get('new', []))
    hide = set(k for k, v in reps.items() if v[2] == 'newer' and ((v[1] or {}).get('status') or 'new') == 'new' and k in news)
    shown = dict(by_status, new=[x for x in by_status.get('new', []) if x[0].Rid.upper() not in hide])
    waiting = waiting_on_router(shown, today)
    # Sent twice and both already given out: the newer one waits on the router to compare.
    gone = lambda x: ((x or {}).get('status') or 'new') != 'new'
    inw = set(x[0].Rid.upper() for x in waiting)
    twice = [x for v in [by_status.get('committee', []), by_status.get('requester', [])] for x in v
             if reps.get(x[0].Rid.upper()) and reps[x[0].Rid.upper()][2] == 'older' and gone(reps[x[0].Rid.upper()][1])]
    # first, so two people are not working the same need any longer than needed
    waiting = [x for x in twice if x[0].Rid.upper() not in inw] + waiting
    parts = waiting_parts(shown, today)
    if twice:
        parts = ('%d sent twice' % len(twice)) + ('' if parts == 'nothing right now' else ' · ' + parts)
    h = [given_note(records), announcements_box(can_post=True), stage_boxes(by_status, stage, len(waiting), parts),
         older_note(st, records), unsorted_docs_card(st) if role == 'router' else '',
         '<form class="bvr" onsubmit="bvGo(\'?view=find&amp;q=\' + encodeURIComponent(this.q.value)); return false;">'
         '<label style="flex:1;display:flex;gap:10px;align-items:center"><b style="white-space:nowrap">Someone at the door?</b>'
         '<input type="text" name="q" placeholder="Name or phone number" style="flex:1"></label><button class="btn ghost">Find</button>'
         '<a class="btn ghost" href="?view=walkin">+ Someone without a form</a></form>']
    if stage:
        return ''.join(h) + stage_cards(st, records, by_status, stage)
    h.append('<h1>Waiting on you</h1>')
    nexts = turn_table(st, records)
    nxt = nexts[0]['name'] if nexts and nexts[0].get('next') else ''
    mine = waiting[:60]
    if not mine:
        h.append('<div class="bvc"><b>Nothing is waiting on you.</b></div>')
    fams = {}
    for s, r in mine:
        fams.setdefault(s.Fam, 0)
    fam_counts = family_counts(st, list(fams.keys()))
    cache = answers_for(st, [s.Rid for s, r in mine] + [v[0].Rid for v in reps.values()])
    ms = membership_of([s.Pid for s, r in mine])
    for s, r in mine:
        need, total = need_summary(st, s.Rid, cache)
        stt = (r or {}).get('status') or 'new'
        tag = ('<span class="pill p-new">New: pick a committee member</span>' if stt == 'new'
               else '<span class="pill p-warn">%s asked to hand it off</span>' % esc(names_for([r.get('owner')]).get(int(r.get('owner') or 0), 'A member'))
               if stt in ('committee', 'requester') and r.get('handoff')
               else '<span class="pill p-new">New message from %s</span>' % esc(names_for([req_unread(s.Rid, r)[-1]['by']]).get(int(req_unread(s.Rid, r)[-1]['by']), 'the team'))
               if stt in ('committee', 'requester') and req_unread(s.Rid, r)
               else '<span class="pill p-warn">No change in %d days</span>' % is_stuck(st, s, r)
               if stt in ('committee', 'requester')
               else '<span class="pill p-ok">Approved: confirm payouts</span>' if stt == 'approved'
               else '<span class="pill p-warn">%s</span>' % ('Pickup today' if pickup_due(r) == today else 'Missed pickup: call them'))
        rp = reps.get(s.Rid.upper()) if stt == 'new' else None
        both_out = reps.get(s.Rid.upper()) if stt in ('committee', 'requester') and (s, r) in twice else None
        if both_out:
            ow = names_for([int(r.get('owner') or 0), int((both_out[1] or {}).get('owner') or 0)])
            a_, b_ = ow.get(int(r.get('owner') or 0), 'a member'), ow.get(int((both_out[1] or {}).get('owner') or 0), 'a member')
            tag = '<span class="pill p-warn">Sent twice, both given out: %s</span>' % (esc(a_) if a_ == b_ else '%s and %s' % (esc(a_), esc(b_)))
        if rp:
            o_open = ((rp[1] or {}).get('status') or 'new') != 'new'
            if o_open:
                tag = '<span class="pill p-warn">Sent again: already with %s</span>' % esc(names_for([rp[1].get('owner')]).get(int(rp[1].get('owner') or 0), 'a member'))
            else:
                need = merge_needs(need, need_summary(st, rp[0].Rid, cache)[0])
                tag = '<span class="pill p-warn">Sent the form twice, %d days apart</span>' % days_apart(s.T, rp[0].T)
        h.append(request_card(s, r, need, total, tag, fam_counts.get(s.Fam, 0),
                              'Came in %s%s' % (days_ago(s.T), (' · %s is next in turn' % esc(nxt)) if nxt and stt == 'new' else ''),
                              ms.get(s.Pid, ''),
                              # a new message opens straight into the chat
                              ('<a class="btn" href="?view=req&amp;rid=%s#repeat">Compare the two</a>' % s.Rid if both_out
                               else '<a class="btn" href="?view=req&amp;rid=%s#rmsg">Read the message</a>' % s.Rid if req_unread(s.Rid, r)
                               else '<a class="btn" href="?view=req&amp;rid=%s#moving">Check on it</a>' % s.Rid)
                              if stt in ('committee', 'requester') and not r.get('handoff') else ''))
    h.append(pickups_section(st, by_status.get('pickup', [])))
    others = []
    for stt in ('committee', 'requester'):
        others += by_status.get(stt, [])
    if others:
        owners = names_for([int((r or {}).get('owner') or 0) for s, r in others])
        others.sort(key=lambda x: last_touch(x[1], x[0]))
        trs = []
        for s, r in others:
            who, stage = with_whom(r, owners)
            lt = last_touch(r, s)
            old = lt[:10] <= (datetime.date.today() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
            trs.append('<tr class="%s"><td><a href="?view=req&amp;rid=%s"><b>#%s %s</b></a></td><td data-l="With">%s</td>'
                       '<td data-l="Where">%s</td><td data-l="Last change">%s</td></tr>' % (
                           'hi' if old else '', s.Rid, r.get('num') or '', esc(s.Name), esc(who), esc(stage), days_ago(lt)))
        load = {}
        for s, r in others:
            if (r or {}).get('status') in ('committee', 'requester'):
                k = int(r.get('owner') or 0)
                n, oldest = load.get(k, (0, ''))
                lt = last_touch(r, s)
                load[k] = (n + 1, min(oldest, lt) if oldest else lt)
        chips = ''.join('<span class="pill %s">%s: %d open, oldest change %s</span>' % (
            'p-warn' if v[1][:10] <= (datetime.date.today() - datetime.timedelta(days=7)).strftime('%Y-%m-%d') else 'p-new',
            esc(owners.get(k, 'Unknown')), v[0], days_ago(v[1])) for k, v in sorted(load.items(), key=lambda x: x[1][1]))
        if chips:
            h.append('<h2>Who has what</h2><div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">%s</div>' % chips)
        h.append('<h2>With someone else (%d)</h2><div class="bvc"><table><thead><tr><th>Request</th><th>With</th><th>Where</th>'
                 '<th>Last change</th></tr></thead><tbody>%s</tbody></table><div class="muted">Oldest first. Nothing has happened '
                 'on the highlighted ones for a week or more.</div></div>' % (len(others), ''.join(trs)))
    return ''.join(h)


def older_note(st, records):
    """A line on home: how many older requests are not sorted yet."""
    cut = cutoff(st)
    n = q.QuerySqlTop1("""
        SELECT COUNT(*) N FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
        WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL AND ISNULL(rp.CompletedDate, r.CompletedDate) < '%s'""" % (
        int(st['form_org']), sql_str(cut)))
    sorted_n = sum(1 for reqs in records.values() for r in reqs.values() if r.get('older'))
    left = max(0, (n.N if n else 0) - sorted_n)
    set_it = '' if st.get('go_live') else (' <b>No go-live date is set, so only the last %d days show here.</b>%s' % (
        DEFAULT_CUTOFF_DAYS, ' Set it in Settings.' if IS_ADMIN else ' Ask an admin to set it.'))
    return ('<div class="bvc"><div class="bvr"><span>Requests from before %s: <b>%d not sorted</b>. Sorting them is optional; '
            'it makes family history complete.%s</span><a class="btn ghost" href="?view=older">Go through older requests</a></div></div>') % (
        nice_date(cut), left, set_it)


def page_older(st, role):
    """The router works through requests from before the cutoff, at their own pace."""
    records = all_records()
    cut = cutoff(st)
    show = _param('show') or 'unsorted'
    page = int(_param('page', '0') or 0) if _param('page', '0').isdigit() else 0
    subs = submissions(st, "AND ISNULL(rp.CompletedDate, r.CompletedDate) < '%s'" % sql_str(cut), 2000)
    rows = []
    for s in subs:
        r = (records.get(s.Pid) or {}).get(s.Rid)
        done = bool(r and (r.get('older') or r.get('brought_in')))
        if show == 'unsorted' and done:
            continue
        rows.append((s, r))
    per = 25
    chunk = rows[page * per:(page + 1) * per]
    h = ['<h1>Older requests</h1>',
         '<div class="muted">Everything that came in before %s, newest first. Mark what happened if you know it; '
         'bring one in if it still needs help. Anything left alone shows as "before this tool, outcome unknown".</div>' % nice_date(cut),
         '<div class="bvr" style="justify-content:flex-start"><a class="btn %s" href="?view=older&amp;show=unsorted">Not sorted (%d)</a>'
         '<a class="btn %s" href="?view=older&amp;show=all">All (%d)</a></div>' % (
             '' if show == 'unsorted' else 'ghost', sum(1 for s in subs if not ((records.get(s.Pid) or {}).get(s.Rid) or {}).get('older')
                                                         and not ((records.get(s.Pid) or {}).get(s.Rid) or {}).get('brought_in')),
             'ghost' if show == 'unsorted' else '', len(subs))]
    cache = answers_for(st, [s.Rid for s, r in chunk])
    for s, r in chunk:
        need, total = need_summary(st, s.Rid, cache)
        if r and r.get('brought_in'):
            state = '<span class="pill p-new">Brought in: %s</span>' % esc(STATUS_LABEL.get(r.get('status'), ''))
        elif r and r.get('older'):
            o = outcome(r)
            state = '<span class="pill %s">%s</span>' % ('p-bad' if o[0] == 'declined' else 'p-ok', esc(o[1]))
        else:
            state = ''
        h.append('<div class="bvc" data-rid="%s"><div class="bvr"><div><b style="font-size:calc(19px*var(--z,1))">%s</b> %s'
                 '<div>%s · %s · <span class="muted">%s</span></div></div>'
                 '<a href="?view=timeline&amp;pid=%d">Timeline</a></div>'
                 '<div class="bvr" style="justify-content:flex-start">'
                 '<label>Helped $<input type="text" class="amt" size="7" value="%s"></label>'
                 '<button class="btn green" data-o="helped">Helped</button><button class="btn red" data-o="declined">Declined</button>'
                 '<button class="btn ghost" data-o="bring">Bring into the new process</button>'
                 '%s<span class="st muted"></span></div></div>' % (
                     s.Rid, esc(s.Name), state, esc(need), money(total) or 'amount not given', nice_date(s.T), s.Pid,
                     ('%g' % total) if total else '',
                     '<button class="btn ghost" data-o="clear">Unsort</button>' if r and (r.get('older') or r.get('brought_in')) else ''))
    pages = (len(rows) + per - 1) // per
    if pages > 1:
        h.append('<div class="bvr" style="justify-content:flex-start">%s<span class="muted">Page %d of %d</span>%s</div>' % (
            ('<a class="btn ghost" href="?view=older&amp;show=%s&amp;page=%d">Newer</a>' % (show, page - 1)) if page > 0 else '',
            page + 1, pages,
            ('<a class="btn ghost" href="?view=older&amp;show=%s&amp;page=%d">Older</a>' % (show, page + 1)) if page + 1 < pages else ''))
    if not chunk:
        h.append('<div class="bvc"><b>Nothing left to sort.</b></div>')
    h.append('<script>Array.prototype.forEach.call(document.querySelectorAll("[data-rid] button[data-o]"),function(b){'
             'b.onclick=function(){var c=b.closest("[data-rid]"),st=c.querySelector(".st");st.textContent="Saving...";'
             'bvPost({action:"sort_old",rid:c.getAttribute("data-rid"),outcome:b.getAttribute("data-o"),amount:c.querySelector(".amt").value},'
             'function(r){if(r.success){st.textContent=r.message;if(%s)c.style.opacity=".45";}else st.textContent=r.message;});};});</script>'
             % ('true' if show == 'unsorted' else 'false'))
    return ''.join(h)


def act_sort_old(st, role):
    """Record what happened to a request from before the cutoff, or bring it
    into the new process as a new request."""
    rid = _param('rid')
    s = find_submission(st, rid)
    if not s:
        raise Exception('Request not found.')
    if s.T >= cutoff(st):
        raise Exception('This request is not from before the cutoff.')
    o = _param('outcome')
    rec = load_record(s.Pid)
    r = rec['requests'].get(rid) or {}
    if r.get('brought_in') and r.get('status') not in (None, 'new') and o != 'clear':
        raise Exception('This one is already being worked. Open it from All requests.')
    if o == 'clear':
        if r.get('status') not in (None, 'new', 'done', 'declined') or len(r.get('history') or []) > 1:
            raise Exception('This one has been worked since; it cannot be unsorted.')
        rec['requests'].pop(rid, None)
        save_record(s.Pid, rec)
        return {'success': True, 'message': 'Unsorted.'}
    r = {'came': s.T, 'closed': s.T}
    if o == 'helped':
        amt = to_number(_param('amount'))
        if not amt:
            raise Exception('Type the amount they were helped with.')
        r.update({'status': 'done', 'older': True, 'payouts': [{'amount': amt, 'how': 'before this tool'}]})
        msg = 'Marked helped, %s.' % money(amt)
    elif o == 'declined':
        r.update({'status': 'declined', 'older': True, 'decision': {'reason': 'Before this tool'}})
        msg = 'Marked declined.'
    elif o == 'bring':
        r = {'came': s.T, 'status': 'new', 'brought_in': True}
        msg = 'Brought in: it is on Waiting on you now.'
    else:
        raise Exception('Pick what happened.')
    add_history(r, {'helped': 'Sorted: helped before this tool', 'declined': 'Sorted: declined before this tool',
                    'bring': 'Brought into the new process'}[o])
    rec['requests'][rid] = r
    save_record(s.Pid, rec)
    return {'success': True, 'message': msg}


def membership_of(pids):
    """{PeopleId: member status} for these people."""
    if not pids:
        return {}
    return dict((r.PeopleId, r.Ms) for r in q.QuerySql("""
        SELECT p.PeopleId, ms.Description Ms FROM People p LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
        WHERE p.PeopleId IN (%s)""" % sql_ints(pids)))


def family_counts(st, fams):
    """{FamilyId: requests in the last 12 months} for these families."""
    if not fams:
        return {}
    return dict((r.F, r.N) for r in q.QuerySql("""
        SELECT p.FamilyId F, COUNT(*) N FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
        JOIN People p ON p.PeopleId = rp.PeopleId
        WHERE r.OrganizationId = %d AND r.CompletedDate >= DATEADD(month, -12, GETDATE()) AND p.FamilyId IN (%s)%s
        GROUP BY p.FamilyId""" % (int(st['form_org']), sql_ints(fams),
                                   (' AND rp.RegPeopleId NOT IN (%s)' % ','.join("'%s'" % k for k in combined_map())) if combined_map() else '')))


def last_touch(r, s=None):
    """When anything last happened on a request."""
    ts = [h.get('t') or '' for h in (r or {}).get('history') or []] + [((r or {}).get('decision') or {}).get('saved') or '']
    ts = [t for t in ts if t]
    return max(ts) if ts else ((r or {}).get('came') or (s.T if s else ''))


def committee_stage(r):
    """Where a request with the committee is, in a few words."""
    d = r.get('decision') or {}
    if r.get('handoff'):
        return 'asked to hand it off'
    aff = r.get('affirm') or {}
    if aff.get('pid') and not aff.get('note') and r.get('shared'):
        return 'shared for affirmation'
    if d.get('saved'):
        return 'decision started'
    if call_started(r) or reached(r):
        return 'called, not decided'
    if r.get('contacts'):
        n = len(r['contacts'])
        return 'trying to reach them (%d %s)' % (n, 'try' if n == 1 else 'tries')
    return 'not started'


def with_whom(r, names):
    """(name of whoever it waits on, stage) or ('', '')."""
    stt = (r or {}).get('status')
    if stt in ('committee', 'requester'):
        nm = names.get(int(r.get('owner') or 0), '')
        return (nm, committee_stage(r) if stt == 'committee' else 'waiting on the requester')
    if stt == 'approved':
        return ('Router', 'confirm payouts')
    if stt == 'pickup':
        d = r.get('decision') or {}
        return (d.get('pickup') or 'Pickup', ('pickup ' + pickup_when(d)) if pickup_when(d) else 'waiting for pickup')
    return ('', '')


def where_it_is(st, s, r):
    """(key, label) for a request's place in the process."""
    if r and r.get('status'):
        if r.get('older'):
            return ('older', outcome(r)[1])
        return (r['status'], STATUS_LABEL.get(r['status'], r['status']))
    if s.T < cutoff(st):
        return ('before', 'Before this tool, not sorted')
    return ('new', STATUS_LABEL['new'])


def page_all(st, role):
    records = all_records()
    want = _param('status')
    term = _param('q')
    per = 25
    page = int(_param('page', '0')) if _param('page', '0').isdigit() else 0
    where = ''
    if term:
        digits = re.sub(r'\D', '', term)
        if len(digits) >= 4:
            where = "AND (p.CellPhone LIKE '%%%s%%' OR p.HomePhone LIKE '%%%s%%')" % (digits, digits)
        else:
            words = [w for w in re.split(r'\s+', re.sub(r"[^\w' -]", ' ', term, flags=re.U)) if w][:3]
            where = 'AND ' + (' AND '.join(
                "(p.FirstName LIKE N'%%{0}%%' OR ISNULL(p.PreferredName,'') LIKE N'%%{0}%%' OR p.LastName LIKE N'%%{0}%%')".format(
                    w.replace("'", "''")) for w in words) or '1=1')
    subs = submissions(st, where, 3000)
    picked = []
    for s in subs:
        r = (records.get(s.Pid) or {}).get(s.Rid)
        key, label = where_it_is(st, s, r)
        if want and key != want:
            continue
        picked.append((s, r, label))
    chunk = picked[page * per:(page + 1) * per]
    cache = answers_for(st, [s.Rid for s, r, l in chunk])
    rows = []
    owners = names_for([int((r or {}).get('owner') or 0) for s, r, l in chunk])
    ms = membership_of([s.Pid for s, r, l in chunk])
    for s, r, label in chunk:
        need, total = need_summary(st, s.Rid, cache)
        who, stage = with_whom(r, owners)
        where = esc(label)
        if who:
            where += '<div class="muted">%s · %s · %s</div>' % (esc(who), esc(stage), days_ago(last_touch(r, s)).replace(' ago', ''))
        rows.append('<tr><td><a href="?view=req&amp;rid=%s"><b>%s</b></a></td><td data-l="Came in">%s</td><td data-l="Need">%s</td>'
                    '<td data-l="Asked">%s</td><td data-l="Member">%s</td><td data-l="Where">%s</td></tr>' % (
                        s.Rid, esc(('#%s ' % r['num'] if r and r.get('num') else '') + s.Name), nice_date(s.T), esc(need),
                        money(total), esc(ms.get(s.Pid, '')), where))
    filters = [('', 'Everything'), ('new', 'New'), ('committee', 'With the committee'), ('requester', 'Waiting on requester'),
               ('approved', 'Approved'), ('pickup', 'Waiting for pickup'), ('done', 'Done'), ('declined', 'Declined'), ('combined', 'Combined (sent twice)'),
               ('before', 'Before this tool, not sorted'), ('older', 'Before this tool, sorted')]
    opts = ''.join('<option value="%s"%s>%s</option>' % (k, ' selected' if k == want else '', esc(l)) for k, l in filters)
    pages = max(1, (len(picked) + per - 1) // per)

    def link(pg, text):
        return '<a class="btn ghost" href="?view=all&amp;status=%s&amp;q=%s&amp;page=%d">%s</a>' % (
            esc(want), esc(term).replace(' ', '+'), pg, text)
    return ('<h1>%s</h1>' % ({'new': 'New: to give out', 'committee': 'With the committee', 'requester': 'Waiting on requester',
                              'approved': 'Approved: confirm payouts', 'pickup': 'Waiting for pickup', 'done': 'Helped',
                              'declined': 'Declined'}.get(want) or 'All requests') +
            '<form class="bvr" style="justify-content:flex-start" method="get">'
            '<input type="hidden" name="view" value="all">'
            '<input type="text" name="q" value="%s" placeholder="Name or phone number" style="flex:1;min-width:200px">'
            '<select name="status" onchange="this.form.submit()">%s</select><button class="btn">Search</button>'
            '%s</form>'
            '<div class="muted">%d request%s%s</div>'
            '<table><thead><tr><th>Who</th><th>Came in</th><th>Need</th><th>Asked</th><th>Membership</th><th>Where it is</th></tr></thead>'
            '<tbody>%s</tbody></table>'
            '<div class="bvr" style="justify-content:flex-start">%s<span class="muted">Page %d of %d</span>%s</div>') % (
        esc(term), opts, ('<a href="?view=all">Clear</a>' if term or want else ''),
        len(picked), '' if len(picked) == 1 else 's', (' matching "%s"' % esc(term)) if term else '',
        ''.join(rows) or '<tr><td colspan="6">None.</td></tr>',
        link(page - 1, 'Newer') if page > 0 else '', page + 1, pages, link(page + 1, 'Older') if page + 1 < pages else '')


# ---------------------------------------------------------------- a request
def find_submission(st, rid):
    if not GUID.match(rid or ''):
        return None
    rows = submissions(st, "AND rp.RegPeopleId = '%s'" % sql_str(rid), 1)
    return rows[0] if rows else None


def page_request(st, role):
    rid = _param('rid')
    s = find_submission(st, rid)
    if not s:
        return '<div class="bvc">That request was not found.</div>'
    rec = load_record(s.Pid)
    r = rec['requests'].get(rid) or {}
    if not can_open(role, r):
        return '<div class="bvc">This request was not given to you.</div>'
    note_here('open:' + s.Rid.upper())
    ups = docs_by_request(st, [s]).get(s.Rid.upper()) or [] if int(st.get('docs_org') or 0) else []
    changed = docs_note(r, ups) if r.get('status') in ('committee', 'requester') else False
    if ME and int(r.get('owner') or 0) == int(ME) and (r.get('docs_in') or '') > (r.get('docs_looked') or ''):
        r['docs_looked'] = r['docs_in']   # the case worker has now seen what came in
        changed = True
    if changed:
        save_record(s.Pid, rec)
    f = form_answers(st, rid)
    fam = family_of(s.Pid)
    me_row = [x for x in fam if x.Pid == s.Pid]
    ms = me_row[0] if me_row else None
    stt = r.get('status') or 'new'
    talk = ''
    if r.get('status') and r.get('status') not in ('new', 'combined') and can_message(role, r):
        nu = len(req_unread(s.Rid, r))
        talk = ' · <a href="#rmsg"><b>Team chat</b><b data-chatn>%s</b></a>' % ((' (%d new)' % nu) if nu else '')
    dseen = my_team()['seen'].get('DOCS:' + s.Rid.upper()) or ''
    dnew = len(set(u['rid'] for u in ups if u['t'] > dseen))
    if ups and stt != 'new':
        talk += ' · <a href="#docs"><b>Documents</b>%s</a>' % ((' <b>(%d new)</b>' % dnew) if dnew else (' (%d)' % len(set(u['rid'] for u in ups))))
    h = ['<div class="bvr"><div><h1>%s%s</h1><div class="muted">Came in %s · <a href="?view=timeline&amp;pid=%d&amp;from=%s">See the timeline</a>%s</div></div>'
         '<span class="pill p-new">%s</span></div>' % (('#%s · ' % r['num']) if r.get('num') else '', esc(s.Name),
                                                       nice_date(s.T), s.Pid, s.Rid, talk, esc(STATUS_LABEL.get(stt, stt)))]

    # Membership and family, with every family member's requests.
    records = all_records()
    fam_all = family_requests(st, s.Pid)
    rec_of = lambda fr: (records.get(fr.Pid) or {}).get(fr.Rid) or (records.get(fr.Pid) or {}).get(fr.Rid.upper())
    fam_reqs = [fr for fr in fam_all if fr.Rid.upper() == rid.upper() or (rec_of(fr) or {}).get('status') != 'combined']
    names = dict((x.Pid, x.Name) for x in fam)
    by_person = {}
    for fr in fam_reqs:
        by_person.setdefault(fr.Pid, []).append(fr)
    year_ago = (datetime.date.today() - datetime.timedelta(days=365)).strftime('%Y-%m-%d')
    others_this_year = [fr for fr in fam_reqs if fr.Rid != rid and fr.T >= year_ago]
    under_names = set(fr.Pid for fr in fam_reqs if fr.T >= year_ago)
    rows = []
    badges = record_badges([x.Pid for x in fam])
    wk_by = {}
    if role == 'router' and QUICK_PROTECTED:
        for e in all_walkins():
            if e.get('pid'):
                wk_by.setdefault(int(e['pid']), []).append(e)
    for x in fam:
        reqs = [walkin_lines_for(e) for e in wk_by.get(int(x.Pid), [])]
        for fr in by_person.get(x.Pid, []):
            if fr.Rid == rid:
                continue
            o = outcome((records.get(fr.Pid) or {}).get(fr.Rid))
            reqs.append('<a href="?view=req&amp;rid=%s">%s</a> · <span class="%s">%s</span> · <a href="?view=compare&amp;rid=%s&amp;other=%s">Compare</a>' % (
                fr.Rid, nice_date(fr.T), 'pill p-bad' if o[0] == 'declined' else '', esc(o[1]), rid, fr.Rid))
        mem = x.MsId == 10
        rows.append('<tr class="%s"><td><b>%s</b>%s%s</td><td data-l="Membership"><span class="pill %s">%s</span></td>'
                    '<td data-l="On their record">%s</td><td data-l="Other requests">%s</td></tr>' % (
            'hi' if reqs and x.Pid != s.Pid else '', esc(x.Name), (', %d' % x.Age) if x.Age is not None else '',
            ' <span class="muted">(asking now)</span>' if x.Pid == s.Pid else '',
            'p-ok' if mem else 'p-grey', esc((x.Ms or 'Unknown') + (' since %s' % x.Joined[:4] if mem and x.Joined else '')),
            badges_html(badges.get(x.Pid) or []) or '<span class="muted">None</span>', '<br>'.join(reqs) or '<span class="muted">None</span>'))
    member = ms and ms.MsId == 10
    h.append('<div class="bvc"><div class="bvr"><h2>Membership and family</h2><span class="pill %s">%s</span></div>'
             '<table><thead><tr><th>In the family (age)</th><th>Membership</th><th>On their record</th><th>Their other requests</th></tr></thead><tbody>%s</tbody></table>%s</div>' % (
                 'p-ok' if member else 'p-grey',
                 esc((ms.Ms or 'Unknown') + (' since %s' % ms.Joined[:4] if member and ms.Joined else '')) if ms else 'Unknown',
                 ''.join(rows),
                 ('<div class="bvc bvwarn"><div class="bvr"><b>This family has asked %d times in the past year%s.</b>'
                  '<a class="btn ghost" href="?view=compare&amp;rid=%s&amp;other=%s">Compare with the last one</a></div></div>' % (
                     len(others_this_year) + 1, (', under %d names' % len(under_names)) if len(under_names) > 1 else '',
                     rid, others_this_year[-1].Rid))
                 if others_this_year else
                 ('<div class="bvr" style="justify-content:flex-start"><a class="btn ghost" href="?view=compare&amp;rid=%s&amp;other=%s">'
                  'Compare with an earlier request</a></div>' % (rid, [fr for fr in fam_reqs if fr.Rid != rid and fr.T < s.T][-1].Rid))
                 if [fr for fr in fam_reqs if fr.Rid != rid and fr.T < s.T] else ''))
    h.append(earlier_help_card(st, s, r, role, fam_reqs, records))
    if stt == 'combined':
        h.insert(1, combined_card(st, s, r, role))
    elif role == 'router':
        rp = repeat_pairs(st, [(fr, rec_of(fr)) for fr in fam_all if int(fr.Pid) == int(s.Pid)]).get(rid.upper())
        if rp:
            h.insert(1, repeat_card(st, s, r, rp))
    elif int(ME or 0) == int(r.get('owner') or 0) or int(ME or 0) in shared_pids(r):
        rp = repeat_pairs(st, [(fr, rec_of(fr)) for fr in fam_all if int(fr.Pid) == int(s.Pid)]).get(rid.upper())
        if rp:
            h.insert(1, resent_note(s, rp))
    if r.get('combined'):
        h.append(also_sent_card(st, s, r, role))

    # How to reach them: routers, and the committee member it was given to.
    working = bool(ME) and int(r.get('owner') or 0) == int(ME) and stt in ('committee', 'requester')
    if (role == 'router' or (ME and int(r.get('owner') or 0) == int(ME))) and not working:
        h.append(contact_card(s))

    # What they asked for.
    need_rows = ''.join('<div class="bvr"><span><b>%s</b>%s</span><b>%s</b></div>' % (
        esc(n['need']), (' · ' + esc(n['payee'])) if n['payee'] else '', money(n['amount'])) for n in f['needs'])
    tried = f['elsewhere']
    h.append('<div class="bvc"><div class="bvr"><div><div class="muted">Total requested</div><b style="font-size:calc(30px*var(--z,1))">%s</b></div>'
             '<div><div class="muted">People in the home</div><b style="font-size:calc(30px*var(--z,1))">%s</b></div>'
             '<div><div class="muted">Help asked for elsewhere</div><b style="font-size:calc(24px*var(--z,1));color:%s">%s</b></div></div>%s</div>' % (
                 money(f['total']) or 'Not given', esc(f['people'] or '?'),
                 '#1E6B45' if tried.lower() == 'yes' else '#A12A22',
                 esc(('Yes: ' + f['elsewhere_who']) if tried.lower() == 'yes' and f['elsewhere_who'] else (tried or 'Not answered')),
                 need_rows))

    # Give it to a committee member.
    owner_me = bool(ME) and int(r.get('owner') or 0) == int(ME)
    shared_me = bool(ME) and int(ME) in [int(x) for x in r.get('shared') or []]
    REQ_RELATION[0] = 'committee' if owner_me else ('router' if role == 'router' else ('shared' if shared_me else 'other'))
    if role == 'router' and stt == 'new':
        h.append(router_choice_card(st, s, r, f, rid, records))
    elif stt in ('approved', 'declined', 'pickup', 'done') and (r.get('decision') or {}).get('sent'):
        h.append(decision_summary(st, r))
        if role == 'router':
            h.append(people_card(st, s, r, role))
        if role == 'router' and stt == 'approved':
            h.append(payout_card(st, s, r))
        elif role == 'router' and stt == 'pickup':
            h.append(pickup_card(st, s, r))
        if stt in ('pickup', 'done'):
            h.append(payouts_list(r))
    elif owner_me and stt in ('committee', 'requester'):
        if r.get('handoff_kept') and not r.get('handoff'):
            hk = r['handoff_kept']
            h.append('<div class="bvc"><b>The router asked you to keep this one</b>%s</div>' % (('<div>' + esc(hk['note']) + '</div>') if hk.get('note') else ''))
        h.insert(1, people_card(st, s, r, role, handoff_member_card(st, s, r, True)))   # top: Case management
        h.append(call_card(st, s, r))
        h.append(docs_card(st, s, r, ups, role))
        h.append(decision_card(st, s, r, f, role == 'router'))
    elif shared_me and stt in ('committee', 'requester'):
        h.insert(1, people_card(st, s, r, role))
        h.append(call_summary(st, r))
        h.append(docs_card(st, s, r, ups, role))
        h.append(plan_so_far(st, r))
    elif role == 'router' and stt in ('committee', 'requester'):
        h.append(moving_card(st, s, r))
        h.append(handoff_router_card(st, s, r, records))
        h.insert(1, people_card(st, s, r, role))
        h.append(call_summary(st, r))
        h.append(docs_card(st, s, r, ups, role))
        h.append(plan_so_far(st, r))
    elif r.get('owner'):
        nm = names_for([r['owner']]).get(int(r['owner']), '')
        h.append('<div class="bvc"><b>With %s</b><span class="muted">Given %s</span></div>' % (
            esc(nm), days_ago((r.get('assigned') or [{}])[-1].get('t', ''))))

    if not working:
        h.append(texts_card(st, s, r, role))
    if r.get('status') and r.get('status') not in ('new', 'combined') and can_message(role, r):
        h.append(chat_widget(s, r))
    h.append('<div class="bvc"><h2>Everything they wrote on the form</h2>%s</div>' % form_sections(f, role == 'router'))
    if IS_ADMIN and r and st.get('mode') != 'live':
        h.append('<div class="bvr" style="justify-content:flex-start"><button class="btn red" id="reset1">Start this request over</button>'
                 '<span class="muted" id="reset1st">Admins, testing only: erases what this tool saved for it; the form and files stay.</span></div>'
                 '<script>document.getElementById("reset1").onclick=function(){var st=document.getElementById("reset1st");st.textContent="Clearing...";'
                 'bvPost({action:"reset_one",rid:"%s"},function(r){if(r.success) location.reload(); else st.textContent=r.message;});};</script>' % rid)
    h.insert(1, request_history_card(st, s, r, collapsed=working))
    h.insert(1, walkin_match_card(st, s, role))
    return ''.join(h)


def where_now(st, r):
    """Where a request stands, in one line."""
    stt = r.get('status') or 'new'
    d = r.get('decision') or {}
    owner = int(r.get('owner') or 0)
    nm = names_for([owner]).get(owner, '') if owner else ''
    if stt == 'new':
        return 'New: waiting for a router to give it to a committee member or decide it.'
    if stt in ('committee', 'requester'):
        return '%s with <b>%s</b>, given %s.' % ('Waiting on the requester,' if stt == 'requester' else 'With the committee:', esc(nm or 'a committee member'),
                                              days_ago((r.get('assigned') or [{}])[-1].get('t', '')))
    total = sum(to_number(x.get('amount')) or 0 for x in d.get('items') or [] if x.get('help'))
    if stt == 'approved':
        return 'Approved %s: waiting for the router to confirm the payouts.' % money(total)
    if stt == 'pickup':
        return 'Paid out; waiting for pickup at <b>%s</b>%s.' % (esc(d.get('pickup') or 'a place not set'), (', ' + esc(pickup_when(d))) if pickup_when(d) else '')
    if stt == 'declined':
        return 'Declined%s%s.' % ((': ' + esc(d['decline_reason'])) if d.get('decline_reason') else '', (' on ' + nice_date(r['closed'])) if r.get('closed') else '')
    if stt == 'done':
        paid = sum(float(x.get('amount') or 0) for x in r.get('payouts') or [])
        return 'Done: helped %s%s.' % (money(paid or total), (', closed ' + nice_date(r['closed'])) if r.get('closed') else '')
    return esc(STATUS_LABEL.get(stt, stt))


def request_history_card(st, s, r, collapsed=False):
    """At the top of a request: where it is now, then everything that
    happened to it, newest first, with what was given or paid and each Why."""
    hist = sorted(r.get('history') or [], key=lambda x: x.get('t') or '', reverse=True)
    who = names_for([x.get('by') for x in hist])
    rows = ''.join(
        '<div style="display:grid;grid-template-columns:110px 1fr;gap:12px;padding:8px 0;border-bottom:1px solid #E6E1D8">'
        '<span class="muted">%s</span><div><b>%s</b>%s%s%s</div></div>' % (
            nice_date(x.get('t')), esc(x.get('what') or ''), (' · ' + esc(who.get(x.get('by'), ''))) if who.get(x.get('by')) else '',
            history_items_html(x),
            ('<div style="background:#F7F5F0;border-radius:10px;padding:8px 12px;margin-top:6px">%s</div>' % clean_html(x['note'])) if x.get('note') else '')
        for x in hist)
    rows += ('<div style="display:grid;grid-template-columns:110px 1fr;gap:12px;padding:8px 0">'
             '<span class="muted">%s</span><div><b>%s sent the request form</b></div></div>') % (nice_date(s.T), esc(s.Name))
    return ('<details class="bvc" id="rhist"%s><summary><b>Where it is and what has happened</b> <span class="muted">%d %s</span></summary>'
            '<div class="shform" style="margin:8px 0"><b>Now:</b> %s</div>%s'
            '<div class="muted" style="margin-top:6px">The whole family: <a href="?view=timeline&amp;pid=%d&amp;from=%s">See the timeline</a></div></details>') % (
        '' if collapsed else ' open', len(hist) + 1, 'step' if not hist else 'steps', where_now(st, r), rows, int(s.Pid), s.Rid)


def contact_card(s):
    """Phone, email and address from their TouchPoint record, plus what they
    typed on the form when it differs."""
    p = q.QuerySqlTop1("""
        SELECT p.CellPhone Cell, p.HomePhone Home, p.EmailAddress Email, p.PrimaryAddress A1, p.PrimaryAddress2 A2,
               p.PrimaryCity City, p.PrimaryState St, p.PrimaryZip Zip
        FROM People p WHERE p.PeopleId = %d""" % int(s.Pid))
    f = q.QuerySqlTop1("SELECT Phone, Email, Address, Zip FROM RegPeople WHERE RegPeopleId = '%s'" % sql_str(s.Rid))

    def phone(v):
        d = re.sub(r'\D', '', v or '')
        return '(%s) %s-%s' % (d[:3], d[3:6], d[6:10]) if len(d) == 10 else (v or '')
    rows = []
    if p:
        if p.Cell:
            rows.append(('Cell', phone(p.Cell)))
        if p.Home:
            rows.append(('Home', phone(p.Home)))
        if p.Email:
            rows.append(('Email', p.Email))
        addr = ', '.join(x for x in [p.A1, p.A2, ' '.join(x for x in [p.City, p.St, (p.Zip or '')[:5]] if x)] if x)
        if addr:
            rows.append(('Address', addr))
    if f:
        if f.Phone and re.sub(r'\D', '', f.Phone)[-10:] not in [re.sub(r'\D', '', (p.Cell if p else '') or ''), re.sub(r'\D', '', (p.Home if p else '') or '')]:
            rows.append(('Phone on the form', phone(f.Phone)))
        if f.Email and (not p or (f.Email or '').lower() != (p.Email or '').lower()):
            rows.append(('Email on the form', f.Email))
    body = ''.join('<div><span class="muted">%s</span> <b>%s</b></div>' % (esc(k), esc(v)) for k, v in rows) or '<div class="muted">No phone or email on file.</div>'
    return ('<div class="bvc"><h2>How to reach them</h2>'
            '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(min(240px,100%%),1fr));gap:6px 18px">%s</div></div>') % body


def form_sections(f, show_private, open_all=False):
    """The form's answers, one collapsible block per section."""
    secs = []
    for sec in f['sections']:
        items = ''.join('<span>%s</span><span class="a">%s</span>' % (esc(i['q']), esc(i['a']) if (show_private or not i['private']) else '<i>private</i>')
                        for i in sec['items'])
        secs.append('<details%s><summary>%s</summary><div class="qa">%s</div></details>' % (' open' if open_all else '', esc(sec['name']), items))
    return ''.join(secs)


def compare_rows(f):
    """{section: [(key, label, answer, private)]} with repeated labels in a
    section told apart by their place (several "Amount Needed")."""
    out = []
    for sec in f['sections']:
        seen, items = {}, []
        for i in sec['items']:
            n = seen.get(i['q'], 0) + 1
            seen[i['q']] = n
            items.append(('%s|%d' % (i['q'], n), i['q'], i['a'], i['private']))
        out.append((sec['name'], items))
    return out


def page_compare(st, role):
    """This request beside an earlier one from the same family: what changed."""
    rid, other = _param('rid'), _param('other')
    a, b = find_submission(st, rid), find_submission(st, other)
    if not a or not b:
        return '<div class="bvc">Request not found.</div>'
    records = all_records()
    if not can_open(role, (records.get(a.Pid) or {}).get(a.Rid)):
        return '<div class="bvc">This request was not given to you.</div>'
    if b.Rid.upper() not in [x.Rid.upper() for x in family_requests(st, a.Pid)]:
        return '<div class="bvc">Those two requests are not from the same family.</div>'
    fa, fb = form_answers(st, a.Rid), form_answers(st, b.Rid)
    ob = outcome((records.get(b.Pid) or {}).get(b.Rid))
    only = _param('all') != '1'
    show_private = role == 'router'

    def needs(f):
        return '<br>'.join('%s %s' % (esc(n['need']), money(n['amount'])) for n in f['needs']) or 'None listed'

    def cell(v, priv):
        return '<i>private</i>' if priv and not show_private else esc(v)
    summary = [('Came in', nice_date(a.T), nice_date(b.T)), ('Who asked', esc(a.Name), esc(b.Name)),
               ('Total requested', money(fa['total']) or 'Not given', money(fb['total']) or 'Not given'),
               ('Needs', needs(fa), needs(fb)), ('People in the home', esc(fa['people']), esc(fb['people'])),
               ('Help asked for elsewhere', esc(fa['elsewhere'] + (': ' + fa['elsewhere_who'] if fa['elsewhere_who'] else '')),
                esc(fb['elsewhere'] + (': ' + fb['elsewhere_who'] if fb['elsewhere_who'] else ''))),
               ('How it ended', 'Open now', esc(ob[1]))]
    h = ['<div class="bvr"><h1>Compare requests</h1><a class="btn ghost" href="?view=req&amp;rid=%s">Back to this request</a></div>' % a.Rid,
         ('<div class="bvc"><table class="cmp">' + '<colgroup><col style="width:30%%"><col style="width:35%%"><col style="width:35%%"></colgroup>' + '<thead><tr><th></th><th>This request</th><th>Earlier request</th></tr></thead><tbody>%s</tbody></table></div>') % ''.join(
             '<tr class="%s"><td><b>%s</b></td><td data-l="This one">%s</td><td data-l="Earlier">%s</td></tr>' % (
                 'hi' if x != y and k not in ('Came in', 'How it ended') else '', k, x, y) for k, x, y in summary)]
    ra, rb = dict(compare_rows(fa)), compare_rows(fb)
    order = [n for n, _ in compare_rows(fa)] + [n for n, _ in rb if n not in ra]
    rbd = dict(rb)
    changed_total = 0
    secs = []
    for name in order:
        ia = dict((k, (lab, v, pv)) for k, lab, v, pv in ra.get(name, []))
        ib = dict((k, (lab, v, pv)) for k, lab, v, pv in rbd.get(name, []))
        keys = [k for k, _, _, _ in ra.get(name, [])] + [k for k, _, _, _ in rbd.get(name, []) if k not in ia]
        trs = []
        for k in keys:
            la = ia.get(k) or ib.get(k)
            va, vb = (ia.get(k) or ('', '', False))[1], (ib.get(k) or ('', '', False))[1]
            diff = va.strip().lower() != vb.strip().lower()
            changed_total += 1 if diff else 0
            if only and not diff:
                continue
            trs.append('<tr class="%s"><td>%s</td><td data-l="This one">%s</td><td data-l="Earlier">%s</td></tr>' % (
                'hi' if diff and not only else '', esc(la[0]), cell(va, la[2]) or '<span class="muted">blank</span>', cell(vb, la[2]) or '<span class="muted">blank</span>'))
        if trs:
            secs.append(('<details open><summary>%s</summary><table class="cmp">' + '<colgroup><col style="width:30%%"><col style="width:35%%"><col style="width:35%%"></colgroup>' +
                         '<thead><tr><th class="sub"></th><th class="sub">This request (%s)</th><th class="sub">Earlier request (%s)</th></tr></thead>'
                         '<tbody>%s</tbody></table></details>') % (esc(name), nice_date(a.T), nice_date(b.T), ''.join(trs)))
    h.append('<div class="bvc"><div class="bvr"><h2>Every answer</h2><a href="?view=compare&amp;rid=%s&amp;other=%s&amp;all=%s">%s</a></div>'
             '<div class="muted">%d answers changed.%s</div>%s</div>' % (
                 a.Rid, b.Rid, '1' if only else '0', 'Show every answer' if only else 'Show only what changed',
                 changed_total, '' if only else ' Changed answers are highlighted.', ''.join(secs) or '<div>Nothing changed.</div>'))
    return ''.join(h)


def assign_card(st, rid, records, exclude=0, draft=''):
    """Pick a committee member. exclude: the current owner, when handing off.
    draft: the note typed so far (saved as they type, on a new request)."""
    rows = [m for m in turn_table(st, records) if m['pid'] != int(exclude or 0)]
    if exclude:
        for m in rows:
            m.pop('next', None)
        free = [m for m in rows if not m['away']]
        if free:
            free[0]['next'] = True
    if not rows:
        return ('<div class="bvc bvblue"><h2>Give it to a committee member</h2><p>No committee members yet. '
                'Pick the committee involvement in Settings.</p></div>')
    trs = []
    for i, m in enumerate(rows):
        dis = ' disabled' if m['away'] else ''
        trs.append('<tr><td><label class="pick"><input type="radio" name="who" value="%d"%s%s><b>%s</b>%s</label></td>'
                   '<td data-l="Open now">%d</td><td data-l="Helping on">%s</td><td data-l="Last given">%s</td><td data-l="This month">%d</td></tr>' % (
                       m['pid'], ' checked' if m.get('next') else '', dis, esc(m['name']),
                       ' <span class="pill p-new">Next in turn</span>' if m.get('next') else
                       (' <span class="pill p-grey">Away until %s</span>' % nice_date(m['away']) if m['away'] else ''),
                       m['open'], m['helping'] or '', ('given ' + days_ago(m['last'])) if m['last'] else 'never yet', m['month']))
    return ('<div class="bvc bvblue" id="assign"><h2>%s</h2>'
            '<div class="muted">Taking turns, from the committee involvement. The next in turn is picked; choose anyone. '
            '<b>Helping on</b> counts open cases a teammate brought them in on.%s</div>'
            '<table><thead><tr><th>Member</th><th>Open now</th><th>Helping on</th><th>Last given one</th><th>This month</th></tr></thead><tbody>%s</tbody></table>'
            '<div><b>Note for them (optional)</b>%s</div>'
            '<div class="bvr" style="justify-content:flex-start"><button class="btn" id="assignbtn">Give it to them</button>'
            '<span id="assignst"></span></div></div>'
            '<script>document.getElementById("assignbtn").onclick=function(){'
            'var w=document.querySelector("#assign input[name=who]:checked"); if(!w) return;'
            'var b=this; b.disabled=true; document.getElementById("assignst").textContent="Saving...";'
            'bvPost({action:"assign",rid:"%s",reassign:"%s",who:w.value,html:bvB64(document.querySelector("#assign .rt").innerHTML)},function(r){'
            'if(r.success) location.href=r.next||location.href; else {b.disabled=false; document.getElementById("assignst").textContent=r.message;}});};</script>') % (
        'Give it to someone else' if exclude else 'Give it to a committee member',
        ' What has been done so far (tries, call notes, a decision draft) goes with it.' if exclude else '',
        ''.join(trs), rich_box(draft, '' if exclude else 'assign_note:%s:' % rid, 70), rid, '1' if exclude else '')


# ---------------------------------------------------------------- timeline
def page_timeline(st, role):
    try:
        pid = int(_param('pid', '0'))
    except Exception:
        pid = 0
    fam = family_of(pid) if pid else []
    if not fam:
        return '<div class="bvc">Person not found.</div>'
    records = all_records()
    names = dict((x.Pid, x.Name) for x in fam)
    fam_reqs = family_requests(st, pid)
    mine = [fr for fr in fam_reqs if can_open(role, (records.get(fr.Pid) or {}).get(fr.Rid))]
    walks = [e for e in all_walkins() if e.get('pid') and int(e['pid']) in names] if role == 'router' and QUICK_PROTECTED else []
    if not mine and not walks:
        return '<div class="bvc">Person not found.</div>'
    events = []
    req_btns = []
    for fr in mine:
        r = (records.get(fr.Pid) or {}).get(fr.Rid) or {}
        need, total = need_summary(st, fr.Rid)
        label = ('#%s' % r['num']) if r.get('num') else nice_date(fr.T)
        link = ' · <a href="?view=req&amp;rid=%s">%s</a>' % (fr.Rid, label)
        req_btns.append('<a class="btn %s" href="?view=req&amp;rid=%s">%s · %s · %s</a>' % (
            'ghost' if outcome(r or None)[0] != 'open' else '', fr.Rid, label, esc(names.get(fr.Pid, '')), esc(outcome(r or None)[1])))
        url = '?view=req&amp;rid=%s' % fr.Rid
        events.append((fr.T, '<b>%s asked</b> · %s · %s%s' % (esc(names.get(fr.Pid, '')), esc(need), money(total), link), '', url))
        who = names_for([h.get('by') for h in r.get('history') or []])
        for hst in r.get('history') or []:
            events.append((hst['t'], '<b>%s</b> · %s%s%s' % (esc(hst['what']), esc(who.get(hst.get('by'), '')), link, history_items_html(hst)), hst.get('note') or '', url))
        mwho = names_for([m.get('by') for m in r.get('messages') or []])
        for m in r.get('messages') or []:
            if msg_visible(r, m, ME):
                events.append((m.get('t') or '', '<b>Message</b> · %s%s%s' % (esc(mwho.get(int(m.get('by') or 0), '')),
                                                                           (' · ' + MSG_TO[m['to']]) if MSG_TO.get(m.get('to')) else '', link), text_html(m.get('text')), url))
    if walks:
        for e in walks:
                give = [{'need': g.get('need'), 'amount': to_number(g.get('amount')) or 0, 'how': g.get('how'), 'ref': g.get('ref') or '', 'held': False,
                         'card': needs_pickup(st, g.get('how'))} for g in e.get('gives') or []]
                events.append((e.get('t') or '', '<b>Walk-in, no form: %s</b> · %s · asked for %s · <a href="?view=walkin&amp;id=%s">%s</a>%s' % (
                    esc(walkin_outcome(e)[1]), esc(names.get(int(e['pid']), '')), esc(walkin_asked(e)), esc(e['id']), esc(e.get('num') or ''),
                    history_items_html({'items': give})), e.get('reason') or '', '?view=walkin&amp;id=%s' % esc(e['id'])))
    events.sort(key=lambda e: e[0], reverse=True)
    for e in walks:
        req_btns.append('<a class="btn ghost" href="?view=walkin&amp;id=%s">%s · %s · %s</a>' % (
            esc(e['id']), esc(e.get('num') or ''), esc(names.get(int(e['pid']), '')), esc(walkin_outcome(e)[1])))
    back = _param('from')
    h = ['<div class="bvr"><div><h1>%s\'s family</h1><div class="muted">%s</div></div>%s</div>' % (
             esc(names.get(pid, '')), esc(', '.join(x.Name for x in fam)),
             ('<a class="btn ghost" href="?view=req&amp;rid=%s">&larr; Back to the request</a>' % esc(back)) if GUID.match(back or '') else
             '<a class="btn ghost" href="?view=home" onclick="if(history.length>1&&document.referrer.indexOf(location.host)>-1){history.back();return false;}">&larr; Back</a>'),
         '<div class="bvc"><b>Their requests</b> <span class="muted">Open one to see where it is and work on it.</span>'
         '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap;gap:8px;margin-top:6px">%s</div></div>' % ''.join(req_btns),
         '<div class="bvc">']
    for t, line, note, url in events:
        h.append('<div class="tlrow" data-href="%s" title="Open it" style="display:grid;grid-template-columns:120px 1fr;gap:16px;padding:10px 8px;'
                 'border-bottom:1px solid #E6E1D8;cursor:pointer;border-radius:8px">'
                 '<b class="muted">%s</b><div>%s%s</div></div>' % (
                     url, nice_date(t), line, ('<div style="background:#F7F5F0;border-radius:10px;padding:10px 14px;margin-top:6px">%s</div>' % clean_html(note)) if note else ''))
    h.append('</div><style>.tlrow:hover{background:#F2EFE8}</style>'
             '<script>Array.prototype.forEach.call(document.querySelectorAll(".tlrow"),function(r){r.onclick=function(e){'
             'if(e.target.closest("a,summary,details,button")) return; location.href=r.getAttribute("data-href");};});</script>')
    return ''.join(h)


FIND_GROUPS = [('pickup', 'Waiting for pickup', 'bvwarn'), ('open', 'Has an open request', 'bvblue'),
               ('walkin', 'Walk-ins not in TouchPoint', ''),
               ('started', 'Started the form but did not send it', ''), ('docs', 'Sent documents, but no request', ''),
               ('family', 'Someone in their family has asked', ''), ('before', 'Asked before', ''),
               ('none', 'Nothing in Benevolence', '')]


def page_find(st, role):
    """Someone at the door: search everyone in TouchPoint by name or phone,
    then sort each person by where their request is, or was: waiting for
    pickup, open (and with whom), a form started and not sent, documents with
    no request, a family member's request, asked before, or nothing at all."""
    term = _param('q')
    digits = re.sub(r'\D', '', term)
    if len(term) < 2:
        return '<div class="bvc">Type at least 2 letters or 4 digits.</div>'
    form, docs = int(st['form_org']), int(st.get('docs_org') or 0)
    if len(digits) >= 4:
        cond = ("(p.CellPhone LIKE '%%{0}%%' OR p.HomePhone LIKE '%%{0}%%' OR p.PeopleId IN (SELECT rp.PeopleId FROM RegPeople rp "
                "JOIN Registration r ON r.RegistrationId = rp.RegistrationId WHERE r.OrganizationId = {1} "
                "AND REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(rp.Phone, ''), '-', ''), '(', ''), ')', ''), ' ', '') LIKE '%%{0}%%'))").format(digits, form)
    else:
        words = [w for w in re.split(r'\s+', re.sub(r"[^\w' -]", ' ', term, flags=re.U)) if w][:3]
        cond = ' AND '.join("(p.FirstName LIKE N'%%{0}%%' OR ISNULL(p.PreferredName,'') LIKE N'%%{0}%%' OR p.LastName LIKE N'%%{0}%%')"
                            .format(w.replace("'", "''")) for w in words) or '1=0'
    everyone = _param('all') == '1'
    walks = all_walkins() if QUICK_PROTECTED else []
    linked = sorted(set(int(e['pid']) for e in walks if e.get('pid')))
    if not everyone:     # only people who sent the request form, or were written down as a walk-in
        cond += (" AND (p.PeopleId IN (SELECT rp.PeopleId FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId "
                 "WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL)%s)" % (
                     form, (' OR p.PeopleId IN (%s)' % ','.join(str(x) for x in linked)) if linked else ''))
    # Walk-ins not in TouchPoint, by name, phone or email.
    tl = term.strip().lower()
    loose = [e for e in walks if not e.get('pid') and (
        (len(digits) >= 4 and digits in re.sub(r'\D', '', e.get('phone') or '')) or
        (len(digits) < 4 and all(w.lower() in ('%s %s %s' % (e.get('first') or '', e.get('last') or '', e.get('email') or '')).lower()
                                 for w in re.split(r'\s+', tl) if w)))]
    walk_by = {}
    for e in walks:
        if e.get('pid'):
            walk_by.setdefault(int(e['pid']), []).append(e)
    people = list(q.QuerySql("""
        SELECT TOP 200 p.PeopleId Pid, ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName N, p.Age, p.FamilyId Fam,
               p.PrimaryCity City
        FROM People p WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0 AND %s
        ORDER BY p.LastName, p.FirstName""" % cond))
    wider = ('<a href="?view=find&amp;all=1&amp;q=%s">Search everyone in TouchPoint</a>' % esc(term.replace('&', '%26').replace(' ', '%20'))
             if not everyone else '<a href="?view=find&amp;q=%s">Only people who sent a request</a>' % esc(term.replace('&', '%26').replace(' ', '%20')))
    if not people and not loose:
        return ('<h1>Someone at the door</h1><div class="bvc">No one %s matches <b>%s</b>. Try part of the name, '
                'or the last 4 digits of their phone. %s · <a href="?view=walkin">+ Write them down as a walk-in</a></div>') % (
                    'in TouchPoint' if everyone else 'who has sent a request', esc(term), wider)
    fams = sorted(set(int(x.Fam) for x in people if x.Fam))
    pids = [int(x.Pid) for x in people]
    subs = submissions(st, 'AND p.FamilyId IN (%s)' % ','.join(str(f) for f in fams), 2000) if fams else []
    started = dict((x.Pid, x.T) for x in q.QuerySql("""
        SELECT x.Pid, CONVERT(varchar(10), MAX(x.T), 120) T FROM (
            SELECT ISNULL(rp.PeopleId, r.PeopleId) Pid, r.CreatedDate T
            FROM Registration r LEFT JOIN RegPeople rp ON rp.RegistrationId = r.RegistrationId
            WHERE r.OrganizationId = %d AND r.CompletedDate IS NULL) x
        WHERE x.Pid IN (%s) GROUP BY x.Pid""" % (form, ','.join(str(p) for p in pids)))) if pids else {}
    sent_docs = dict((x.Pid, x.T) for x in q.QuerySql("""
        SELECT rp.PeopleId Pid, CONVERT(varchar(10), MAX(ISNULL(rp.CompletedDate, r.CompletedDate)), 120) T
        FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
        WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL AND rp.PeopleId IN (%s)
        GROUP BY rp.PeopleId""" % (docs, ','.join(str(p) for p in pids)))) if docs and pids else {}
    records = all_records()
    owners = names_for([((records.get(x.Pid) or {}).get(x.Rid) or {}).get('owner') for x in subs])
    groups = dict((k, []) for k, _, _ in FIND_GROUPS)
    for x in people:
        mine = [fs for fs in subs if fs.Pid == x.Pid]                       # newest first
        fam = [fs for fs in subs if fs.Fam == x.Fam and fs.Pid != x.Pid]
        rec = lambda fs: (records.get(fs.Pid) or {}).get(fs.Rid) or {}
        lines, cats = [], set()
        for fs in mine + fam:
            r = rec(fs)
            stt = r.get('status') or ('new' if r or fs.T >= (st.get('go_live') or '9999') else '')
            who = '' if fs.Pid == x.Pid else '%s: ' % esc(fs.Name)
            link = '<a href="?view=req&amp;rid=%s">%s</a>' % (fs.Rid, ('#%s' % r['num']) if r.get('num') else nice_date(fs.T))
            if stt == 'pickup':
                d = r.get('decision') or {}
                lines.append('%sWaiting for pickup at <b>%s</b>%s · %s' % (who, esc(d.get('pickup') or 'a place not set'),
                                                                        (', ' + esc(pickup_when(d))) if pickup_when(d) else '', link))
                cats.add('pickup')
            elif stt in ('new', 'committee', 'requester', 'approved'):
                with_ = (' with <b>%s</b>' % esc(owners.get(int(r['owner']), ''))) if r.get('owner') and stt in ('committee', 'requester') else ''
                lines.append('%s%s%s · came in %s · %s' % (who, esc(STATUS_LABEL.get(stt, 'Open')), with_, nice_date(fs.T), link))
                cats.add('open' if fs.Pid == x.Pid else 'family')
            else:
                o = outcome(r or None)
                lines.append('%s%s · %s · %s' % (who, esc(o[1]), nice_date(fs.T), link))
                cats.add('before' if fs.Pid == x.Pid else ('family' if not mine else ''))
        for e in walk_by.get(int(x.Pid), []):
            lines.append(walkin_lines_for(e))
            cats.add('before')
        if x.Pid in started:
            lines.insert(0, 'Started the request form %s and did not send it. They can finish it from the same link, or start again.' % nice_date(started[x.Pid]))
            cats.add('started')
        if x.Pid in sent_docs and not mine:
            lines.insert(0, 'Sent the documents form %s, but has no request. See <b>Documents that match no request</b> on Waiting on you.' % nice_date(sent_docs[x.Pid]))
            cats.add('docs')
        where = [k for k, _, _ in FIND_GROUPS if k in cats] or ['none']
        best = ([fs for fs in mine + fam if (rec(fs).get('status') or '') == 'pickup'] +
                [fs for fs in mine + fam if (rec(fs).get('status') or 'new') in ('new', 'committee', 'requester', 'approved') and (rec(fs) or fs.T >= (st.get('go_live') or '9999'))] +
                mine + fam)
        wk = walk_by.get(int(x.Pid)) or []
        groups[where[0]].append((x, lines, ('?view=req&amp;rid=%s' % best[0].Rid) if best else (('?view=walkin&amp;id=%s' % esc(wk[0]['id'])) if wk else '')))
    for e in loose:
        groups['walkin'].append((None, e, '?view=walkin&amp;id=%s' % esc(e['id'])))
    h = ['<h1>Someone at the door</h1><div class="muted">%d %s %s match <b>%s</b>, sorted by where their request is.%s %s</div>' % (
        len(people) + len(loose), 'person' if len(people) + len(loose) == 1 else 'people', 'in TouchPoint' if everyone else 'who sent a request', esc(term),
        ' Showing the first 200; type more to narrow it.' if len(people) == 200 else '', wider + ' · <a href="?view=walkin">+ Someone without a form</a>')]
    for k, title, cls in FIND_GROUPS:
        g = groups[k]
        if not g:
            continue
        shown = g if k != 'none' else g[:40]
        if k == 'walkin':
            rows = ''.join('<div style="padding:8px 0;border-bottom:1px solid #E6E1D8"><div class="bvr"><span><a href="%s"><b>%s</b></a>%s</span>'
                           '<span><a class="btn" href="%s">Open</a></span></div><div class="muted">%s</div></div>' % (
                               url, esc(walkin_name(e)), ''.join(' · ' + esc(v) for v in [e.get('phone'), e.get('email')] if v), url, walkin_lines_for(e))
                           for x, e, url in shown)
        else:
            rows = ''.join('<div style="padding:8px 0;border-bottom:1px solid #E6E1D8"><div class="bvr"><span>%s%s%s</span><span>%s</span></div>%s</div>' % (
                ('<a href="%s"><b>%s</b></a>' % (best, esc(x.N))) if best else '<b>%s</b>' % esc(x.N),
                (', %d' % x.Age) if x.Age is not None else '', (' · ' + esc(x.City)) if x.City else '',
                ('<a class="btn" href="%s">Open</a> <a class="btn ghost" href="?view=timeline&amp;pid=%d">Family timeline</a>' % (best, x.Pid)) if best
                else '<a class="btn ghost" href="/Person2/%d" target="_blank">TouchPoint record</a>' % x.Pid,
                ''.join('<div class="muted" style="margin-top:2px">%s</div>' % l for l in lines)) for x, lines, best in shown)
        more = ('<div class="muted">And %d more with nothing here. Type more of the name to narrow it.</div>' % (len(g) - 40)) if len(g) > 40 else ''
        tip = ('<div class="muted">No request from them or anyone in their family. If they need help, point them to the request form.</div>'
               if k == 'none' else ('<div class="muted">Asked without the form and not linked to anyone in TouchPoint. Open one to link it.</div>' if k == 'walkin' else ''))
        h.append('<div class="bvc %s"><h2>%s <span class="muted">(%d)</span></h2>%s%s%s</div>' % (cls, title, len(g), tip, rows, more))
    return ''.join(h)


# ---------------------------------------------------------------- walk-ins: help asked for without the form
# People who ask in person or on the phone, often not in TouchPoint. A router
# writes down who they are, what they asked for and what we decided. Saved on
# the router's own record (the person may have none), in an extra value
# hidden like the rest. Linked to a TouchPoint person when there is one; a
# form that comes in later is checked against the ones not linked yet.
def digits_of(v):
    d = re.sub(r'\D', '', v or '')
    return d[-10:] if len(d) >= 10 else d


def all_walkins():
    """[entry] from every holder, each with _holder and _field (where it is saved)."""
    out = []
    for r in q.QuerySql("""
            SELECT PeopleId, Field, CAST(Data AS nvarchar(max)) Data FROM PeopleExtra WITH (NOLOCK)
            WHERE Field IN (%s) AND Data IS NOT NULL""" % ','.join("'%s'" % f for f in QUICK_NAMES)):
        for k, e in (load_json(r.Data or '', {}).get('entries') or {}).items():
            e = dict(e, id=k, _holder=int(r.PeopleId), _field=r.Field)
            out.append(e)
    out.sort(key=lambda e: e.get('t') or '', reverse=True)
    return out


def find_walkin(wid):
    for e in all_walkins():
        if e['id'] == wid:
            return e
    raise Exception('That walk-in was not found. Reload the page.')


def save_walkin(e):
    """Save an entry on its holder's record (a new one on the router saving
    it). An entry found under a merge copy moves to the main name."""
    if not QUICK_PROTECTED:
        raise Exception('Walk-ins are not protected yet, so nothing was saved. An admin opening the tool fixes it.')
    holder, field = e.get('_holder') or int(ME), e.get('_field') or QUICK_FIELD
    clean = dict((k, v) for k, v in e.items() if not k.startswith('_') and k != 'id')
    if field != QUICK_FIELD:
        old = load_json(model.ExtraValueText(holder, field) or '', {})
        (old.get('entries') or {}).pop(e['id'], None)
        if old.get('entries'):
            model.AddExtraValueText(holder, field, safe_json(old))
        else:
            model.DeleteExtraValue(holder, field)
    data = load_json(model.ExtraValueText(holder, QUICK_FIELD) or '', {})
    data.setdefault('v', 1)
    data.setdefault('entries', {})[e['id']] = clean
    model.AddExtraValueText(holder, QUICK_FIELD, safe_json(data))
    e['_holder'], e['_field'] = holder, QUICK_FIELD


def walkin_name(e):
    return ('%s %s' % (e.get('first') or '', e.get('last') or '')).strip() or 'No name given'


def walkin_outcome(e):
    total = sum(to_number(g.get('amount')) or 0 for g in e.get('gives') or [])
    if e.get('choice') == 'help':
        return ('paid', 'Helped %s' % money(total) if total else 'Helped')
    if e.get('choice') == 'decline':
        return ('declined', 'Declined' + (': ' + e['decline_reason'] if e.get('decline_reason') else ''))
    return ('open', 'Not decided yet')


def walkin_asked(e):
    return ', '.join('%s%s' % (a.get('need') or 'Other', (' ' + money(to_number(a.get('amount')))) if to_number(a.get('amount')) else '')
                     for a in e.get('asks') or []) or 'Not written down'


def walkins_matching(people, walkins=None):
    """{pid: [unlinked entries that may be them]} by phone, email or the
    same first and last name. people: rows with Pid, First, Last, Cell,
    Home, Email (and optional Phone2, Email2 from their form)."""
    walkins = [e for e in (walkins if walkins is not None else all_walkins()) if not e.get('pid')]
    out = {}
    for p in people:
        phones = set(x for x in [digits_of(getattr(p, k, '') or '') for k in ('Cell', 'Home', 'Phone2')] if len(x) >= 7)
        emails = set(x.strip().lower() for x in [getattr(p, k, '') or '' for k in ('Email', 'Email2')] if x and '@' in x)
        first, last = (p.First or '').strip().lower(), (p.Last or '').strip().lower()
        for e in walkins:
            if int(p.Pid) in [int(x) for x in e.get('not') or []]:
                continue
            ep, ee = digits_of(e.get('phone')), (e.get('email') or '').strip().lower()
            why = []
            if ep and len(ep) >= 7 and any(ep[-7:] == x[-7:] for x in phones):
                why.append('phone')
            if ee and ee in emails:
                why.append('email')
            if last and (e.get('last') or '').strip().lower() == last and first and (e.get('first') or '').strip().lower()[:3] == first[:3]:
                why.append('name')
            if why:
                out.setdefault(int(p.Pid), []).append(dict(e, _why=why))
    return out


def walkin_people(term_first, term_last, phone, email):
    """TouchPoint people who may be the walk-in: by phone, email or name."""
    conds = []
    d = digits_of(phone)
    if len(d) >= 7:
        conds.append("(REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(p.CellPhone,''),'-',''),'(',''),')',''),' ','') LIKE '%%%s' OR "
                     "REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(p.HomePhone,''),'-',''),'(',''),')',''),' ','') LIKE '%%%s')" % (d[-7:], d[-7:]))
    if email and '@' in email:
        conds.append("(p.EmailAddress = N'%s' OR p.EmailAddress2 = N'%s')" % (sql_str(email.strip()), sql_str(email.strip())))
    if term_last and len(term_last.strip()) >= 2:
        c = "p.LastName LIKE N'%s%%'" % sql_str(term_last.strip())
        if term_first and term_first.strip():
            c += " AND (p.FirstName LIKE N'%s%%' OR ISNULL(p.PreferredName,'') LIKE N'%s%%' OR p.NickName LIKE N'%s%%')" % (
                sql_str(term_first.strip()[:3]), sql_str(term_first.strip()[:3]), sql_str(term_first.strip()[:3]))
        conds.append('(%s)' % c)
    if not conds:
        return []
    return list(q.QuerySql("""
        SELECT TOP 12 p.PeopleId Pid, ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName N, p.Age, p.PrimaryCity City,
               p.PrimaryAddress Addr, p.CellPhone Cell, p.EmailAddress Email,
               (SELECT COUNT(*) FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
                WHERE rp.PeopleId = p.PeopleId AND r.OrganizationId = %d AND r.CompletedDate IS NOT NULL) Sent
        FROM People p WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0 AND (%s)
        ORDER BY p.LastName, p.FirstName""" % (int(SETTINGS['form_org']), ' OR '.join(conds))))


def act_walkin_match(st, role):
    import json
    f = load_json(from_b64(_param('data')) or '{}', {})
    rows = walkin_people(f.get('first'), f.get('last'), f.get('phone'), f.get('email'))
    me = _param('id')
    seen = []
    for e in all_walkins():
        if e['id'] == me:
            continue
        ep, ee = digits_of(e.get('phone')), (e.get('email') or '').lower()
        d, em = digits_of(f.get('phone')), (f.get('email') or '').strip().lower()
        if (len(d) >= 7 and ep[-7:] == d[-7:]) or (em and em == ee) or (
                (f.get('last') or '').strip() and (e.get('last') or '').lower() == (f.get('last') or '').strip().lower()
                and (e.get('first') or '').lower()[:3] == (f.get('first') or '').strip().lower()[:3]):
            seen.append(e)
    html = ''
    if rows:
        html += '<div><b>In TouchPoint</b></div>' + ''.join(
            '<div class="bvr" style="padding:6px 0;border-bottom:1px solid #E6E1D8"><span><b>%s</b>%s%s%s<span class="muted">%s</span></span>'
            '<button type="button" class="btn ghost w-pick" data-pid="%d" data-name="%s">This is them</button></div>' % (
                esc(x.N), (', %d' % x.Age) if x.Age is not None else '', (' · ' + esc(x.City)) if x.City else '',
                (' · <b>sent %d request%s</b>' % (x.Sent, '' if x.Sent == 1 else 's')) if x.Sent else '',
                ''.join(' · ' + esc(v) for v in [x.Addr, x.Cell, x.Email] if v), x.Pid, esc(x.N)) for x in rows)
    if seen:
        html += '<div style="margin-top:8px"><b>Seen before as a walk-in</b></div>' + ''.join(
            '<div style="padding:4px 0"><a href="?view=walkin&amp;id=%s">%s %s</a> · %s · %s · %s</div>' % (
                esc(e['id']), esc(e.get('num') or ''), esc(walkin_name(e)), nice_date(e.get('t')), esc(walkin_asked(e)), esc(walkin_outcome(e)[1]))
            for e in seen[:8])
    if not html:
        html = '<div class="muted">No match in TouchPoint yet. Keep going: they are saved as a walk-in with what you type here.</div>'
    return {'success': True, 'html': html}


def clean_walkin(st, raw, e):
    d = load_json(from_b64(raw) or '{}', {})
    for k, n in (('first', 40), ('last', 40), ('phone', 30), ('email', 80), ('address', 160)):
        e[k] = unicode(d.get(k) or '').strip()[:n]
    if not (e['first'] or e['last'] or e['phone']):
        raise Exception('Type at least a name or a phone number.')
    asks = []
    for a in (d.get('asks') or [])[:12]:
        need = added_need(a.get('cat'), a.get('detail'))[:80]
        if need or to_number(a.get('amount')):
            asks.append({'cat': unicode(a.get('cat') or '')[:60], 'detail': unicode(a.get('detail') or '').strip()[:80],
                         'need': need or 'Other', 'amount': unicode(a.get('amount') or '')[:20]})
    e['asks'] = asks
    e['choice'] = d.get('choice') if d.get('choice') in ('help', 'decline') else ''
    gives = []
    for g in (d.get('gives') or [])[:12]:
        need = added_need(g.get('cat'), g.get('detail'))[:80]
        if need or to_number(g.get('amount')):
            gives.append({'cat': unicode(g.get('cat') or '')[:60], 'detail': unicode(g.get('detail') or '').strip()[:80],
                          'need': need or 'Other', 'amount': unicode(g.get('amount') or '')[:20], 'how': unicode(g.get('how') or '')[:60],
                          'ref': unicode(g.get('ref') or '').strip()[:40]})
    e['gives'] = gives if e['choice'] == 'help' else []
    e['decline_reason'] = unicode(d.get('decline_reason') or '')[:80] if e['choice'] == 'decline' else ''
    e['referral'] = unicode(d.get('referral') or '')[:80] if e['choice'] == 'decline' else ''
    e['reason'] = clean_html(d.get('reason') or '')
    try:
        pid = int(d.get('pid') or 0)
    except Exception:
        pid = 0
    if pid != int(e.get('pid') or 0):
        e['pid'] = pid or None
        e.setdefault('history', []).append({'t': now_str(), 'by': int(ME), 'what': 'Linked to their TouchPoint record' if pid else 'Unlinked from TouchPoint'})
    if e['choice'] == 'help':
        for g in gives:
            if not to_number(g['amount']):
                raise Exception('Type the amount given for %s.' % g['need'])
            if not g['how']:
                raise Exception('Pick how for %s.' % g['need'])
        if not gives:
            raise Exception('Add what we gave, or pick Declined or Not decided yet.')
    if e['choice'] == 'decline' and not e['decline_reason']:
        raise Exception('Pick why it was declined.')
    return e


def next_walkin_number(walkins):
    n = 0
    for e in walkins:
        try:
            n = max(n, int((e.get('num') or 'W0')[1:]))
        except Exception:
            pass
    return 'W%d' % (n + 1)


def act_walkin_save(st, role):
    wid = _param('id')
    if wid:
        e = find_walkin(wid)
        before = (e.get('choice'), walkin_outcome(e)[1])
    else:
        ws = all_walkins()
        e = {'id': 'w%s%d' % (datetime.datetime.now().strftime('%Y%m%d%H%M%S'), int(ME)), 'num': next_walkin_number(ws),
             't': now_str(), 'by': int(ME), 'history': []}
        before = None
    clean_walkin(st, _param('data'), e)
    if before is None:
        e['history'].append({'t': now_str(), 'by': int(ME), 'what': 'Written down: ' + walkin_outcome(e)[1]})
    elif before[1] != walkin_outcome(e)[1]:
        e['history'].append({'t': now_str(), 'by': int(ME), 'what': walkin_outcome(e)[1]})
    else:
        e['history'].append({'t': now_str(), 'by': int(ME), 'what': 'Edited'})
    e['history'] = e['history'][-40:]
    save_walkin(e)
    return {'success': True, 'id': e['id']}


def act_walkin_link(st, role):
    """From a request: this walk-in is the person who sent it (or is not)."""
    e = find_walkin(_param('id'))
    pid = int(_param('pid', '0') or 0)
    if not pid:
        raise Exception('No person to link.')
    if _param('no') == '1':
        e['not'] = list(set([int(x) for x in e.get('not') or []] + [pid]))
        save_walkin(e)
        return {'success': True}
    e['pid'] = pid
    e.setdefault('history', []).append({'t': now_str(), 'by': int(ME), 'what': 'Linked to their TouchPoint record from a request they sent'})
    save_walkin(e)
    return {'success': True}


def walkin_rows_html(st, rows, prefix, cats=True, hows=False):
    """Lines of what they asked for (or what we gave) on the walk-in page."""
    out = []
    for a in rows or [{}]:
        out.append('<tr class="%s-row"><td data-l="What"><select class="w-cat">%s</select> '
                   '<input type="text" class="w-detail" value="%s" placeholder="What (more detail)" size="16"></td>'
                   '<td data-l="Amount"><input type="text" class="w-amt" value="%s" placeholder="$" size="8"></td>%s'
                   '<td><button type="button" class="btn red w-del" title="Remove">&times;</button></td></tr>' % (
                       prefix, cat_options(st, a.get('cat') if a else ''), esc((a or {}).get('detail') or ('' if (a or {}).get('cat') else (a or {}).get('need') or '')),
                       esc((a or {}).get('amount') or ''),
                       ('<td data-l="How"><select class="w-how">%s</select></td><td data-l="Check or card"><input type="text" class="w-ref" value="%s" size="12" placeholder="Check number"></td>' % (
                           help_kind_options(st, (a or {}).get('how') or first_paid(st)), esc((a or {}).get('ref') or ''))) if hows else ''))
    return ''.join(out)


def page_walkin(st, role):
    wid = _param('id')
    e = find_walkin(wid) if wid else {}
    if not QUICK_PROTECTED:
        return '<div class="bvc bvbad"><b>Walk-ins are not protected yet, so nothing can be saved.</b><span>An admin opening the tool fixes it.</span></div>'
    linked = int(e.get('pid') or 0)
    lname = names_for([linked]).get(linked, '') if linked else ''
    reasons = ''.join('<option%s>%s</option>' % (' selected' if x == e.get('decline_reason') else '', esc(x)) for x in st.get('decline_reasons') or [])
    refs = '<option value="">(none)</option>' + ''.join('<option%s>%s</option>' % (' selected' if x == e.get('referral') else '', esc(x)) for x in st.get('referrals') or [])
    hist = ''
    if e:
        who = names_for([x.get('by') for x in e.get('history') or []] + [e.get('by')])
        hist = ('<details class="bvc"><summary><b>What has happened</b></summary>%s</details>' % ''.join(
            '<div style="padding:4px 0"><span class="muted">%s</span> · <b>%s</b> · %s</div>' % (nice_date(x.get('t')), esc(x.get('what') or ''), esc(who.get(x.get('by'), '')))
            for x in reversed(e.get('history') or [])))
    blank_ask = walkin_rows_html(st, [{}], 'ask')
    blank_give = walkin_rows_html(st, [{}], 'give', hows=True)
    return ('<div class="bvr"><div><h1>%s</h1><div class="muted">%s</div></div><a class="btn ghost" href="?view=walkins">All walk-ins</a></div>'
            '<div class="bvc" id="wk" data-id="%s">'
            '<h2>Who they are</h2><div class="muted">Type what they give you. We look for them in TouchPoint as you type.</div>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap;gap:8px">'
            '<label>First name <input type="text" id="wfirst" value="%s" size="14"></label><label>Last name <input type="text" id="wlast" value="%s" size="16"></label>'
            '<label>Phone <input type="tel" id="wphone" value="%s" size="14"></label><label>Email <input type="email" id="wemail" value="%s" size="22"></label></div>'
            '<label>Address <input type="text" id="waddr" value="%s" size="50" style="max-width:100%%"></label>'
            '<input type="hidden" id="wpid" value="%s">'
            '<div class="shform" id="wlinked"%s><b>Linked to</b> <a id="wlname" href="/Person2/%s" target="_blank">%s</a> in TouchPoint. '
            '<button type="button" class="btn ghost" id="wunlink">Not them</button></div>'
            '<div class="shform" id="wmatch"%s><div class="muted">Matches show here as you type.</div></div>'
            '<h2>What they asked for</h2><table><thead><tr><th>What</th><th>Amount</th><th></th></tr></thead><tbody id="wasks">%s</tbody></table>'
            '<button type="button" class="btn ghost" id="waskadd">+ Add a need</button>'
            '<h2>What we decided</h2><div class="bvr" style="justify-content:flex-start">'
            '<label class="pick"><input type="radio" name="wchoice" value="help"%s> Helped</label>'
            '<label class="pick"><input type="radio" name="wchoice" value="decline"%s> Declined</label>'
            '<label class="pick"><input type="radio" name="wchoice" value=""%s> Not decided yet</label></div>'
            '<div id="whelp"><b>What we gave</b><table><thead><tr><th>What</th><th>Amount</th><th>How</th><th>Check or card</th><th></th></tr></thead><tbody id="wgives">%s</tbody></table>'
            '<button type="button" class="btn ghost" id="wgiveadd">+ Add</button></div>'
            '<div id="wdecline"><label>Why <select id="wreason">%s</select></label> <label>Send them to <select id="wref">%s</select></label></div>'
            '<div><b>Why (what you learned and why)</b>%s</div>'
            '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn green" id="wsave">Save</button><span id="wst" class="muted"></span></div></div>%s'
            '<script>(function(){var c=document.getElementById("wk"), id=c.getAttribute("data-id"), t=null, $=function(x){return document.getElementById(x);};'
            'var BA=%s, BG=%s;'
            'function rows(tb){return Array.prototype.map.call(document.querySelectorAll("#"+tb+" tr"),function(tr){var o={cat:tr.querySelector(".w-cat").value,'
            ' detail:tr.querySelector(".w-detail").value, amount:tr.querySelector(".w-amt").value};'
            ' if(tr.querySelector(".w-how")){o.how=tr.querySelector(".w-how").value; o.ref=tr.querySelector(".w-ref").value;} return o;});}'
            'function who(){return {first:$("wfirst").value,last:$("wlast").value,phone:$("wphone").value,email:$("wemail").value};}'
            'function val(){var w=who(), ch=c.querySelector("input[name=wchoice]:checked"); w.address=$("waddr").value; w.pid=$("wpid").value;'
            ' w.asks=rows("wasks"); w.choice=ch?ch.value:""; w.gives=rows("wgives"); w.decline_reason=$("wreason").value; w.referral=$("wref").value;'
            ' w.reason=c.querySelector(".rt").innerHTML; return w;}'
            'function show(){var ch=(c.querySelector("input[name=wchoice]:checked")||{}).value||""; $("whelp").style.display=ch=="help"?"":"none";'
            ' $("wdecline").style.display=ch=="decline"?"":"none"; var l=!!$("wpid").value; $("wlinked").style.display=l?"":"none"; $("wmatch").style.display=l?"none":"";}'
            'function wire(tr){tr.querySelector(".w-del").onclick=function(){tr.parentNode.removeChild(tr);};}'
            'Array.prototype.forEach.call(c.querySelectorAll("#wasks tr,#wgives tr"),wire);'
            'function add(tb,html){var tmp=document.createElement("tbody"); tmp.innerHTML=html; var tr=tmp.firstChild; $(tb).appendChild(tr); wire(tr); tr.querySelector(".w-cat").focus();}'
            '$("waskadd").onclick=function(){add("wasks",BA);}; $("wgiveadd").onclick=function(){add("wgives",BG);};'
            'function look(){clearTimeout(t); t=setTimeout(function(){var w=who(); if(!(w.last.trim().length>1||w.phone.replace(/\\D/g,"").length>=7||w.email.indexOf("@")>0)) return;'
            ' bvPost({action:"walkin_match",id:id,data:bvB64(JSON.stringify(w))},function(r){ if(!r.success) return; $("wmatch").innerHTML=r.html;'
            '  Array.prototype.forEach.call($("wmatch").querySelectorAll(".w-pick"),function(b){b.onclick=function(){$("wpid").value=b.getAttribute("data-pid");'
            '   $("wlname").textContent=b.getAttribute("data-name"); $("wlname").href="/Person2/"+b.getAttribute("data-pid"); show();};});});},500);}'
            '["wfirst","wlast","wphone","wemail"].forEach(function(k){$(k).addEventListener("input",look);});'
            '$("wunlink").onclick=function(){$("wpid").value=""; show(); look();};'
            'c.addEventListener("change",show);'
            '$("wsave").onclick=function(){var b=this; b.disabled=true; $("wst").textContent="Saving...";'
            ' bvPost({action:"walkin_save",id:id,data:bvB64(JSON.stringify(val()))},function(r){ if(r.success) location.href="?view=walkin&id="+r.id+"&saved=1";'
            '  else {b.disabled=false; $("wst").textContent=r.message;}});};'
            'show(); if(!$("wpid").value) look();})();</script>') % (
        ('%s · %s' % (esc(e.get('num') or ''), esc(walkin_name(e)))) if e else 'Someone without a form',
        ('Written down %s by %s%s' % (nice_date(e.get('t')), esc(names_for([e.get('by')]).get(e.get('by'), '')), ' · <b>Saved</b>' if _param('saved') else ''))
        if e else 'Someone who asked in person or on the phone. Write down who they are, what they asked for and what we decided.',
        esc(e.get('id') or ''), esc(e.get('first') or ''), esc(e.get('last') or ''), esc(e.get('phone') or ''), esc(e.get('email') or ''),
        esc(e.get('address') or ''), linked or '', '' if linked else ' style="display:none"', linked or '', esc(lname),
        ' style="display:none"' if linked else '',
        walkin_rows_html(st, e.get('asks') or [{}], 'ask'),
        ' checked' if e.get('choice') == 'help' else '', ' checked' if e.get('choice') == 'decline' else '', ' checked' if not e.get('choice') else '',
        walkin_rows_html(st, e.get('gives') or [{}], 'give', hows=True), reasons, refs, rich_box(e.get('reason') or '', '', 80), hist,
        safe_json(blank_ask), safe_json(blank_give))


def page_walkins(st, role):
    ws = all_walkins()
    only = _param('show')
    if only == 'unlinked':
        ws = [e for e in ws if not e.get('pid')]
    names = names_for([e.get('pid') for e in ws if e.get('pid')])
    rows = ''.join(
        '<tr><td><a href="?view=walkin&amp;id=%s"><b>%s %s</b></a><div class="muted">%s</div></td><td data-l="When">%s</td>'
        '<td data-l="Asked for">%s</td><td data-l="Decided"><span class="pill %s">%s</span></td><td data-l="TouchPoint">%s</td></tr>' % (
            esc(e['id']), esc(e.get('num') or ''), esc(walkin_name(e)), esc(' · '.join(x for x in [e.get('phone'), e.get('email')] if x)),
            nice_date(e.get('t')), esc(walkin_asked(e)), {'paid': 'p-ok', 'declined': 'p-bad', 'open': 'p-warn'}[walkin_outcome(e)[0]], esc(walkin_outcome(e)[1]),
            ('<a href="?view=timeline&amp;pid=%d">%s</a>' % (int(e['pid']), esc(names.get(int(e['pid']), 'Linked')))) if e.get('pid') else '<span class="muted">Not in TouchPoint</span>')
        for e in ws)
    return ('<div class="bvr"><h1>Walk-ins</h1><a class="btn" href="?view=walkin">+ Someone without a form</a></div>'
            '<div class="muted">People who asked in person or on the phone, without the request form. When one of them sends the form, '
            'their request shows a possible match to link.</div>'
            '<div class="bvr" style="justify-content:flex-start"><a href="?view=walkins" class="%s">All (%d)</a> · '
            '<a href="?view=walkins&amp;show=unlinked" class="%s">Not in TouchPoint</a></div>'
            '<div class="bvc"><table><thead><tr><th>Who</th><th>When</th><th>Asked for</th><th>Decided</th><th>TouchPoint</th></tr></thead><tbody>%s</tbody></table></div>') % (
        'on' if not only else '', len(all_walkins()) if only else len(ws), 'on' if only == 'unlinked' else '',
        rows or '<tr><td colspan="5" class="muted">None yet.</td></tr>')


def walkin_lines_for(e):
    """One line for a walk-in in a list of someone's requests."""
    return 'Walk-in %s · %s · asked for %s · <span class="%s">%s</span>' % (
        '<a href="?view=walkin&amp;id=%s">%s</a>' % (esc(e['id']), nice_date(e.get('t'))), esc(e.get('num') or ''), esc(walkin_asked(e)),
        'pill p-bad' if walkin_outcome(e)[0] == 'declined' else '', esc(walkin_outcome(e)[1]))


def walkin_match_card(st, s, role):
    """On a request: walk-ins not linked yet that may be the person who sent
    it, by phone, email or name (their record and what they typed on the form)."""
    if role != 'router' or not QUICK_PROTECTED:
        return ''
    p = q.QuerySqlTop1("""
        SELECT p.PeopleId Pid, p.FirstName First, p.LastName Last, p.CellPhone Cell, p.HomePhone Home, p.EmailAddress Email,
               rp.Phone Phone2, rp.Email Email2
        FROM People p LEFT JOIN RegPeople rp ON rp.RegPeopleId = '%s' WHERE p.PeopleId = %d""" % (sql_str(s.Rid), int(s.Pid)))
    if not p:
        return ''
    hits = walkins_matching([p]).get(int(s.Pid)) or []
    if not hits:
        return ''
    rows = ''.join(
        '<div class="bvr" style="padding:6px 0;border-bottom:1px solid #E6E1D8" data-w="%s"><span>%s<div class="muted">%s · same %s</div></span>'
        '<span><button type="button" class="btn ghost w-yes">Same person: link it</button> <button type="button" class="btn ghost w-no">Not them</button></span></div>' % (
            esc(e['id']), walkin_lines_for(e), esc(' · '.join(x for x in [walkin_name(e), e.get('phone'), e.get('email')] if x)), ' and '.join(e['_why'])) for e in hits)
    return ('<div class="bvc bvwarn" id="wmatchreq"><h2>They may have asked before without the form</h2>'
            '<div class="muted">A walk-in that is not linked to anyone yet looks like this person. Link it and it shows with their other requests.</div>%s'
            '<span id="wmst" class="muted"></span></div>'
            '<script>(function(){Array.prototype.forEach.call(document.querySelectorAll("#wmatchreq [data-w]"),function(row){'
            'function go(no){document.getElementById("wmst").textContent="Saving..."; bvPost({action:"walkin_link",id:row.getAttribute("data-w"),pid:"%d",no:no?"1":""},'
            ' function(r){ if(r.success) location.reload(); else document.getElementById("wmst").textContent=r.message;});}'
            'row.querySelector(".w-yes").onclick=function(){go(false);}; row.querySelector(".w-no").onclick=function(){go(true);};});})();</script>') % (rows, int(s.Pid))


# ---------------------------------------------------------------- settings
def setup_guide(st):
    """Getting started, step by step: what each part is for in the process,
    whether it is done, and where to do it. Returns (html, done, of, the next
    step's title or '', finished: live with every step done)."""
    ok, missing, role_ok = protection_state()
    form = int(st.get('form_org') or 0)
    form_row = q.QuerySqlTop1("SELECT ISNULL(LimitToRole, '') R FROM Organizations WHERE OrganizationId = %d" % form) if form else None
    docs = int(st.get('docs_org') or 0)
    docs_row = q.QuerySqlTop1("SELECT ISNULL(LimitToRole, '') R FROM Organizations WHERE OrganizationId = %d" % docs) if docs else None
    routers = [int(r.get('pid') or 0) for r in st.get('routers') or []]
    holders = set(p for p, n in role_holders(access_role())) if role_ok else set()
    committee = committee_members(st) if int(st.get('committee_org') or 0) else []
    live = st.get('mode') == 'live'
    sec = lambda k, l: '<a href="?view=settings&amp;sec=%s">%s</a>' % (k, l)
    below = sec('setup', 'Setup')
    # (title, why it matters, done: True / False / None for optional and not used, where)
    steps = [
        ('Make the access role and protect the data',
         'One TouchPoint role opens every request. Routers get it; committee members do not need it.',
         bool(ok), below + ', Access role; and Admin, Roles in TouchPoint'),
        ('Pick your request form and lock it',
         'People ask for help on an ordinary TouchPoint form. Locking its involvement to the access role keeps their answers private.',
         bool(form_row and form_row.R), below + ', Forms and people; then the involvement\'s Settings, General, Limit To Role'),
        ('Check how the tool reads your form',
         'The tool finds each need, its amount and payee, and the total from your form. Look it over once, with the preview, and save.',
         bool(st.get('form_seen')), sec('form', 'Reading your form')),
        ('Your documents form (optional)',
         'If people send bills or an ID on a second form, the tool puts each upload on the right request.',
         bool(docs_row and docs_row.R) if docs else None,
         below + ', Forms and people; then ' + sec('docs', 'Documents form')),
        ('Pick the committee',
         'Everyone on the committee involvement can be given a request to call about and decide on. Nobody else can.',
         bool(committee), below + ', Forms and people'),
        ('Choose your routers',
         'Routers see every request, give each one to a committee member, and confirm the payouts.',
         bool(routers) and all(p in holders for p in routers), below + ', Forms and people'),
        ('Write the call questions',
         'What committee members ask on the phone. The samples are a start; make them yours.',
         (st.get('call_questions') or '') != DEFAULT_SETTINGS['call_questions'], sec('call', 'The call')),
        ('Set the kinds of help and how decisions work',
         'Checks, gift cards and pickup places, what a decision needs before it is sent, and when a teammate must affirm it.',
         st.get('help_kinds') != DEFAULT_SETTINGS['help_kinds'] or st.get('pickup_places') != DEFAULT_SETTINGS['pickup_places'],
         sec('decisions', 'Decisions')),
        ('Turn on the emails',
         'A short morning email tells each person what is waiting on them. No names or needs are in it.',
         bool(st.get('email_on') and st.get('email_from')), sec('emails', 'Emails')),
        ('Texting them (optional)',
         'Send the person a link to a secure message, or to the documents form, through one of TouchPoint\'s texting groups.',
         True if int(st.get('text_group') or 0) else None, sec('texts', 'Texting them')),
        ('Try it, then go live',
         'In Testing nothing is emailed and you can clear what you tried. Going live starts the real thing.',
         live, below + ', Mode'),
    ]
    need = [x for x in steps if x[2] is not None]
    done = sum(1 for x in need if x[2])
    nxt = next((i for i, x in enumerate(steps) if x[2] is False), None)
    rows = []
    for i, (title, why, d, where) in enumerate(steps):
        mark = ('<b style="color:#1E6B45">&#10003;</b>' if d else '<span class="muted">optional</span>' if d is None
                else '<span class="pill p-new">Next</span>' if i == nxt else '<b style="color:#8A5A00">&#9675;</b>')
        rows.append('<li style="display:grid;grid-template-columns:70px 1fr;gap:10px;padding:8px 0;border-top:1px solid #E6E1D8%s">'
                    '<div>%s</div><div><b>%d. %s</b><div class="muted">%s</div><div style="font-size:.95em">Where: %s</div></div></li>' % (
                        ';background:#F7F5F0' if i == nxt else '', mark, i + 1, esc(title), esc(why), where))
    flow = ('<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap;gap:6px;margin:8px 0">%s</div>' % ' <span class="muted">&rarr;</span> '.join(
        '<span class="chip">%s</span>' % esc(x) for x in ('Someone fills out your form', 'A router gives it to a committee member',
                                                           'They call and decide', 'A teammate affirms it', 'The router pays or sets out pickup',
                                                           'Reports')))
    html = ('<div class="bvc" id="guide"><h2>Getting started <span class="muted" style="font-weight:400">%d of %d done%s</span></h2>'
            '<div class="muted">How a request moves through the tool:</div>%s'
            '<ol style="list-style:none;padding:0;margin:8px 0 0">%s</ol>'
            '<div class="muted"><a href="?view=help">How the whole process works</a></div></div>') % (
        done, len(need), '' if done < len(need) else ': ready', flow, ''.join(rows))
    return html, done, len(need), (steps[nxt][0] if nxt is not None else ''), live and done == len(need)


def guide_banner(guide):
    """One line on Setup pointing to Getting started while it is not finished."""
    html, done, of, nxt, finished = guide
    if finished:
        return ''
    return ('<a class="bvc" href="?view=settings&amp;sec=start" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;'
            'text-decoration:none;color:inherit;border-color:#1F4E8C;background:#EEF4FB"><b>Getting started: %d of %d done</b>%s'
            '<span style="margin-left:auto" class="btn ghost">Open the steps</span></a>') % (
        done, of, (' <span>Next: %s</span>' % esc(nxt)) if nxt else '')


def page_settings(st, role):
    if not IS_ADMIN:
        return '<div class="bvc">Settings are for admins.</div>'
    ok, missing, role_ok = protection_state()
    prot = ('<div class="bvc %s"><h2>Protecting the data</h2>%s</div>' % (
        '' if ok else 'bad',
        ('<div>Benevolence data is hidden from everyone without the ' + role_b() + ' role. This is checked, and fixed if needed, '
         'every time an admin opens the tool.</div>')
        if ok else ('<b>Not protected yet: nothing can be saved.</b><div>%s</div>' % (
            'The access role ' + role_b() + ' does not exist in TouchPoint. Create it (Admin, Roles), or pick a role you already have under <b>Access role</b> below, '
            'and give it to the routers only (committee members do not need it). Then reload this page; '
            'the tool protects the data by itself.' if not role_ok else
            'The tool could not update StandardExtraValues2: %s' % esc(AUTO_PROTECT_ERROR or 'reload the page to try again.')))))
    routers = names_for([r['pid'] for r in st.get('routers') or []])
    comm = committee_members(st)
    merger = names_for([st.get('merger_pid')]).get(int(st.get('merger_pid') or 0), '')
    guide = setup_guide(st)
    sec = _param('sec') or ('setup' if guide[4] else 'start')   # opens on Getting started until it is finished
    secs = [('start', 'Getting started'), ('setup', 'Setup'), ('form', 'Reading your form'), ('docs', 'Documents form'), ('call', 'The call'), ('decisions', 'Decisions'), ('emails', 'Emails'),
            ('texts', 'Texting them')]
    sub = '<div class="subnav">%s</div>' % ''.join('<a href="?view=settings&amp;sec=%s" class="%s">%s</a>' % (
        k, 'on' if k == sec else '', esc(l)) for k, l in secs)
    save = ('<div class="bvr" style="justify-content:flex-start"><button class="btn" id="savest">Save</button><span id="savestst"></span></div></div>'
            '<script>document.getElementById("savest").onclick=function(){var f=document.getElementById("setf"),d={action:"save_settings",s_section:f.getAttribute("data-sec")};'
            'var kv=document.getElementById("kindsv"); if(kv) kv.value=JSON.stringify(Array.prototype.map.call(document.querySelectorAll("#kinds tr"),function(r){'
            ' return {name:r.querySelector(".kname").value, pickup:r.querySelector(".kpick").value=="1", reqs:r.querySelector(".kreqs").value.split("\\n")};}));'
            'var pv=document.getElementById("placesv"); if(pv) pv.value=Array.prototype.map.call(document.querySelectorAll("#places .place span"),function(x){return x.textContent;}).join("\\n");'
            'Array.prototype.forEach.call(f.querySelectorAll("input[name],textarea[name],select[name]"),function(i){d["s_"+i.name]=i.type=="checkbox"?(i.checked?"1":"0"):i.value;});'
            'var st=document.getElementById("savestst"); st.textContent="Saving...";'
            'bvPost(d,function(r){st.textContent=r.success?"Saved.":r.message; if(r.success&&f.getAttribute("data-reload")) location.reload();});};</script>')
    if sec == 'start':
        return sub + guide[0]
    if sec == 'emails':
        return sub + page_emails_section(st) + save[save.index('<script>'):]
    if sec == 'texts':
        return sub + page_texts_section(st) + save
    if sec == 'form':
        return sub + ('<div class="bvc" id="setf" data-sec="form" data-reload="1">%s' % form_match_card(st)) + save
    if sec == 'docs':
        return sub + ('<div class="bvc" id="setf" data-sec="docs" data-reload="1">%s' % (docs_match_card(st) or (
            '<b>Documents form</b><div class="muted">No documents form is picked. If you use one (a TouchPoint form where people send '
            'bills, notices or an ID), add its involvement id in <a href="?view=settings&amp;sec=setup">Setup</a>, Forms and people.</div>'))) + save
    if sec == 'call':
        return sub + ('<div class="bvc" id="setf" data-sec="call" data-reload="1">%s' % call_questions_card(st)) + save
    if sec == 'decisions':
        lines = lambda k: esc('\n'.join(st.get(k) or []))
        return sub + ''.join([
            '<div class="bvc" id="setf" data-sec="decisions"><h2>Decisions</h2>',
            '<div><b>Kinds of help</b><div class="muted">Each shows in the How list on a decision. Picked-up kinds ask for a place, '
            'day and time, and list their requirements for the person handing it out. Removing one does not change decisions already made.</div>'
            '<table class="kinds"><thead><tr><th>Name</th><th>How it is given</th><th>Requirements (one per line)</th><th></th></tr></thead>'
            '<tbody id="kinds">%s</tbody></table><button type="button" class="btn ghost" id="addkind">+ Add a kind of help</button>'
            '<input type="hidden" name="help_kinds" id="kindsv"></div>' % ''.join(kind_row(k) for k in st.get('help_kinds') or []),
            '<div><b>Pickup places</b><div class="muted">Where gift cards and other help are picked up. Removing one does not change requests already using it.</div>'
            '<div id="places">%s</div><div class="bvr" style="justify-content:flex-start"><input type="text" id="newplace" placeholder="New place" size="24">'
            '<button type="button" class="btn ghost" id="addplace">Add</button></div><input type="hidden" name="pickup_places" id="placesv"></div>' % ''.join(
                '<div class="bvr place" style="justify-content:flex-start"><span>%s</span><button type="button" class="btn red rmplace">Remove</button></div>' % esc(x)
                for x in st.get('pickup_places') or []),
            '<div><b>Reasons for declining</b> <span class="muted">one per line</span><textarea name="decline_reasons" rows="5" style="width:100%%">%s</textarea></div>' % lines('decline_reasons'),
            '<div><b>Places to send them</b> <span class="muted">other agencies, one per line</span><textarea name="referrals" rows="4" style="width:100%%">%s</textarea></div>' % lines('referrals'),
            '<label class="pick"><input type="checkbox" name="affirm_on"%s> Affirmation: members can share a request with another member</label>' % (' checked' if st.get('affirm_on') else ''),
            '<label class="pick"><input type="checkbox" name="affirm_required"%s> Affirmation is required before it goes to the router</label>' % (' checked' if st.get('affirm_required') else ''),
            '<label class="pick"><input type="checkbox" name="handoff_on"%s> Handoff: a member can ask the router to give their request to someone else</label>' % (' checked' if st.get('handoff_on') else ''),
            ('<div class="shform"><b>Before a decision can be sent</b>'
             '<div class="muted">Always: Help needs something ticked, with an amount and how on each line. '
             'Decline needs a reason picked. These depend on the choice, so a Decline never trips on help lines.</div>'
             '<label class="pick"><input type="checkbox" name="req_why_help"%s> Help: the Why is written</label>'
             '<label class="pick"><input type="checkbox" name="req_talked_help"%s> Help: they talked with the person (call notes saved, or a try logged as talked with them)</label>'
             '<label class="pick"><input type="checkbox" name="req_pickup"%s> Help that is picked up: the pickup place, day and time are set '
             '<span class="muted">(off: the committee member may set them, and the router sets them with the payouts)</span></label>'
             '<label class="pick"><input type="checkbox" name="req_why_decline"%s> Decline: the Why is written</label>'
             '<label>Help over $ <input type="text" name="affirm_over" value="%s" size="7"> needs a teammate to affirm it <span class="muted">(0 means never; the required switch above covers every decision)</span></label></div>'
            '<div class="shform"><b>The router</b>'
            '<label>On a new request, a router can decide %s</label>'
            '<label>Show a request on the router\'s list when there is no change for <input type="text" name="stuck_days" value="%s" size="3"> days</label>'
            '<div class="muted">A router\'s decision follows the Why rules above; talking with them first and a teammate\'s affirmation '
            'are for committee decisions. On any open request the router can check on it, take it over, or give it to someone else.</div></div>'
            '<div class="shform"><b>Earlier help</b>'
            '<label class="pick"><input type="checkbox" name="earlier_on"%s> Whoever works a case sees the family\'s earlier decisions, '
            'the Why and the call notes <span class="muted">(read only; off: each case is judged on its own, and only routers see earlier ones)</span></label></div>') % (
                ' checked' if st.get('req_why_help', True) else '', ' checked' if st.get('req_talked_help', True) else '',
                ' checked' if st.get('req_pickup') else '', ' checked' if st.get('req_why_decline', True) else '', esc(('%g' % float(st.get('affirm_over') or 0))),
                when_select('router_decide', st.get('router_decide', 'decline'), ROUTER_DECIDE), stuck_days(st),
                ' checked' if st.get('earlier_on', True) else ''),
            save,
            '<script>(function(){var tb=document.getElementById("kinds");'
            'function wire(r){r.querySelector(".rmkind").onclick=function(){r.parentNode.removeChild(r);};}'
            'Array.prototype.forEach.call(tb.querySelectorAll("tr"),wire);'
            'document.getElementById("addkind").onclick=function(){var t=document.createElement("tbody"); t.innerHTML=%s;'
            ' var r=t.firstChild; tb.appendChild(r); wire(r); r.querySelector(".kname").focus();};})();</script>' % safe_json(kind_row({})),
            '<script>(function(){var box=document.getElementById("places");'
            'function wire(r){r.querySelector(".rmplace").onclick=function(){r.parentNode.removeChild(r);};}'
            'Array.prototype.forEach.call(box.querySelectorAll(".place"),wire);'
            'document.getElementById("addplace").onclick=function(){var i=document.getElementById("newplace"),v=i.value.trim(); if(!v) return;'
            ' var r=document.createElement("div"); r.className="bvr place"; r.style.justifyContent="flex-start";'
            ' var sp=document.createElement("span"); sp.textContent=v; r.appendChild(sp);'
            ' var b=document.createElement("button"); b.type="button"; b.className="btn red rmplace"; b.textContent="Remove"; r.appendChild(b);'
            ' box.appendChild(r); wire(r); i.value="";};})();</script>'])
    live = st.get('mode') == 'live'
    mode = ('<div class="bvc %s"><div class="bvr"><h2>Mode: %s</h2>%s</div><div>%s</div>%s</div>'
            '<script>(function(){var b=document.getElementById("modego"); if(!b) return; b.onclick=function(){'
            'var st=document.getElementById("modest"), c=document.getElementById("modec"); st.textContent="Saving...";'
            'bvPost({action:"set_mode",mode:b.getAttribute("data-mode"),confirm:c?c.value:""},function(r){if(r.success) location.reload(); else st.textContent=r.message;});};})();</script>') % (
        '' if live else 'bvwarn', 'Live' if live else 'Testing',
        '<span class="pill p-ok">Live since %s</span>' % nice_date(st.get('live_since') or st.get('go_live')) if live else '<span class="pill p-warn">Testing</span>',
        'The team is using it for real. Emails go out and data cannot be cleared.' if live else
        'Nothing is emailed, a Testing tag shows at the top of every page, and admins can clear the data below. '
        'Going live sets the go-live date to today if none is set, and turns off clearing.',
        ('<div class="bvr" style="justify-content:flex-start"><input type="text" id="modec" placeholder="Type TESTING" size="14">'
         '<button class="btn ghost" id="modego" data-mode="testing">Back to testing</button><span id="modest" class="muted">'
         'Brings back the buttons that erase data.</span></div>') if live else
        '<div class="bvr" style="justify-content:flex-start"><button class="btn green" id="modego" data-mode="live">Go live</button><span id="modest"></span></div>')
    h = [guide_banner(guide), mode, prot, who_can_get_in(st), involvement_card(st), '<div class="bvc" id="setf" data-sec="setup" data-reload="1"><h2>Forms and people</h2>',
         '<label>Request form involvement id <input type="text" name="form_org" value="%s" size="8"> %s</label>' % (
             st.get('form_org') or '', org_name_note(st.get('form_org'))),
         '<label>Documents form involvement id <input type="text" name="docs_org" value="%s" size="8"> %s</label>' % (
             st.get('docs_org') or '', org_name_note(st.get('docs_org'), 'optional: the form people use to send bills and IDs')),
         '<label>Committee involvement id <input type="text" name="committee_org" value="%s" size="8"> '
         '<span class="muted">%s</span></label>' % (st.get('committee_org') or '', esc(', '.join(m['name'] for m in comm)) or 'none yet'),
         '<label>Main routers (PeopleIds) <input type="text" name="routers_main" value="%s" size="24"> '
         '<span class="muted">get new requests and the emails</span></label>' % (
             ', '.join(str(r['pid']) for r in st.get('routers') or [] if r.get('main'))),
         '<label>Backup routers (PeopleIds) <input type="text" name="routers_backup" value="%s" size="24"> '
         '<span class="muted">can do everything; only emailed when something is stuck</span></label>' % (
             ', '.join(str(r['pid']) for r in st.get('routers') or [] if not r.get('main'))),
         access_role_field(st),
         router_check(st, routers),
         '<label>Merges duplicate records (PeopleId) <input type="text" name="merger_pid" value="%s" size="10"> <span class="muted">%s</span></label>' % (
             st.get('merger_pid') or '', esc(merger)),
         '<label>Go-live date (earlier requests count as before this tool) <input type="text" name="go_live" value="%s" size="12" placeholder="YYYY-MM-DD"></label>' % esc(st.get('go_live') or ''),
         save,
         '' if live else '<div class="bvc bvbad"><h2>Clear test data</h2><div>Erases everything this tool has saved: who each request was given to, '
         'decisions, notes, affirmations, messages and sorted older requests. Every request goes back to New and numbers start again at #%d. '
         'The form submissions and their files are not touched. <b>Once the team is using it for real, this would erase their work.</b></div>'
         '<div class="bvr" style="justify-content:flex-start"><input type="text" id="resetc" placeholder="Type RESET" size="12">'
         '<button class="btn red" id="resetgo">Clear all Benevolence data</button><span id="resetst"></span></div></div>'
         '<script>document.getElementById("resetgo").onclick=function(){var st=document.getElementById("resetst");st.textContent="Clearing...";'
         'bvPost({action:"reset_all",confirm:document.getElementById("resetc").value},function(r){st.textContent=r.success?r.message:r.message;});};</script>' % FIRST_NUMBER,
         ]
    return sub + ''.join(h)


# A request form to build in TouchPoint (new forms), for a church starting
# from scratch. Shaped the way this tool reads best: one Yes or No question
# per kind of need, whose No skips past that need's own questions.
STARTER_FORM = [
    ('Instructions', 'A short note: what you can and cannot help with, and that answers are kept private.', ''),
    ('Person info', 'Their name, address, phone and email (TouchPoint fills these in).', ''),
    ('Text', 'What is the best way and time to reach you?', ''),
    ('Text', 'Who referred you to us?', ''),
    ('Single choice', 'Do you attend a church? (Yes, No)', 'No skips the next question'),
    ('Text', 'Which church?', ''),
    ('Single choice', 'Has our church helped you before? (Yes, No)', ''),
    ('Instructions', 'Your household', ''),
    ('Text', 'How many people live in your home?', 'People in the home'),
    ('Dropdown', 'Your housing (Own, Rent, Staying with someone, No stable housing)', ''),
    ('Single choice', 'Does anyone else live in the home with you? (Yes, No)', 'No skips to the next heading'),
    ('Text', 'Their names, ages and how they are related to you', ''),
    ('Instructions', 'Income and bills', ''),
    ('Text', 'Job status, and monthly income from work', ''),
    ('Text', 'Other monthly income (benefits, support, pension)', ''),
    ('Text', 'Main monthly bills (housing, car, utilities, phone, food)', ''),
    ('Instructions', 'What you need help with', ''),
    ('Single choice', 'Do you need help with rent or mortgage? (Yes, No)', 'A need; No skips to the next need'),
    ('Text', 'Amount needed', 'Amount for the need above'),
    ('Text', 'Payee name', 'Who the need above is paid to'),
    ('Text', 'Account or loan number', 'Private (routers only)'),
    ('Single choice', 'Do you need help with electric? (Yes, No)', 'A need; the same three questions follow'),
    ('Single choice', 'Do you need help with water, gas or another utility? (Yes, No)', 'A need'),
    ('Single choice', 'Do you need help with a car payment or repair? (Yes, No)', 'A need'),
    ('Single choice', 'Do you need help with medical bills or prescriptions? (Yes, No)', 'A need; amount and payee follow'),
    ('Single choice', 'Do you need food? (Yes, No)', 'A need (no amount needed)'),
    ('Single choice', 'Do you need help with something else? (Yes, No)', 'A need; No skips its three questions'),
    ('Text', 'What is it?', 'More about the need above'),
    ('Text', 'Amount needed', 'Amount for the need above'),
    ('Text', 'Payee name', 'Who the need above is paid to'),
    ('Text', 'Total amount of help you are asking for', 'Total asked for'),
    ('Instructions', 'Other help', ''),
    ('Single choice', 'Have you asked anyone else for help with this? (Yes, No)', 'Asked anyone else; No skips ahead'),
    ('Text', 'Who did you ask?', 'Who else they asked'),
    ('Text', 'What did they give, and how much?', ''),
    ('Instructions', 'We may ask for a bill, a notice or an ID. If we do, you can send it on our documents form.', ''),
    ('Instructions', 'Agreement', ''),
    ('Multiple choice', 'I agree that the church may contact the people and agencies named here. (Yes, I agree)', ''),
    ('Text', 'Signature and date', ''),
]


def starter_form_card():
    rows = ''.join('<tr><td data-l="Type">%s</td><td data-l="Question">%s</td><td data-l="The tool reads it as" class="muted">%s</td></tr>' % (
        esc(t), esc(qn), esc(role)) for t, qn, role in STARTER_FORM)
    return ('<details class="bvc"><summary><b>Starting a request form from scratch?</b> <span class="muted">A form this tool reads with no changes</span></summary>'
            '<div class="muted">A form to build in TouchPoint (Involvement, Registration, new forms) that this tool reads with no changes. '
            'Word the questions your own way; what matters is the shape: one Yes or No question for each kind of need, '
            'with the amount and payee right after it, and the No answer set to skip to the next need.</div>'
            '<table class="cmp"><colgroup><col style="width:18%%"><col style="width:52%%"><col></colgroup>'
            '<thead><tr><th>Type</th><th>Question</th><th>The tool reads it as</th></tr></thead><tbody>%s</tbody></table></details>') % rows


DOCS_ROLES = [('show', 'Show it'), ('private', 'Private (routers only)'), ('hide', 'Hide it')]


def docs_auto_role(t, label, st):
    if t == 7:
        return 'show'
    return 'private' if word_test(st.get('form_words') or DEFAULT_SETTINGS['form_words'], 'private')(label) else 'show'


def docs_match_card(st):
    """The documents form's questions: show, private or hide on the case page."""
    if not int(st.get('docs_org') or 0):
        return ''
    fmap = dict((k.upper(), v) for k, v in (st.get('docs_map') or {}).items())
    rows = []
    for i, x in enumerate(form_questions(st['docs_org'])):
        label = strip_tags(x.Label)
        if x.T == 5:
            rows.append('<tr><td colspan="3" class="muted">%d. Their name and contact details (TouchPoint fills these in)</td></tr>' % (i + 1))
            continue
        if x.T == 4:
            rows.append('<tr style="background:#F2F0EB"><td colspan="3"><b>%d. %s</b> <span class="muted">(a heading or note on the form; '
                        'not shown on the case page)</span></td></tr>' % (i + 1, esc(label[:90] or '(heading)')))
            continue
        auto = docs_auto_role(x.T, label, st)
        cur = fmap.get(x.Qid.upper(), auto)
        roles = [r for r in DOCS_ROLES if not (x.T == 7 and r[0] == 'private')]
        rows.append('<tr><td data-l="Question">%d. %s%s</td><td data-l="Kind" class="muted">%s</td>'
                    '<td data-l="On the case page"><select name="dm_%s">%s</select></td></tr>' % (
                        i + 1, esc(label[:140]), ' <span class="muted">(you set this)</span>' if cur != auto else '',
                        'A file' if x.T == 7 else 'Their note', x.Qid.upper().replace('-', ''),
                        ''.join('<option value="%s"%s>%s</option>' % (k, ' selected' if k == cur else '', esc(l)) for k, l in roles)))
    out = ('<div><b>Your documents form</b><div class="muted">What they send on the documents form shows on the request it '
            'belongs to: the family\'s newest request made before they sent it. Files get an Open button; anything they typed shows as '
            'a note. Mark any question private (routers only) or hide it. Documents that match no request show on the router\'s '
            'list to attach by hand.</div><table class="cmp"><colgroup><col style="width:55%%"><col style="width:15%%"><col></colgroup>'
            '<thead><tr><th>Question</th><th>Kind</th><th>On the case page</th></tr></thead><tbody>%s</tbody></table></div>') % ''.join(rows)
    return out.replace('<table class="cmp">', '<table class="cmp docsmap">', 1)


def posted_form_map(st, words):
    """The form map from the Reading your form page as it stands (saved or
    not): only the questions set differently from what the tool works out.
    Returns (map, every question id)."""
    qs = form_questions(st['form_org'])
    auto = dict((x['qid'], x['auto']) for x in form_layout([(x.Qid, x.T, x.Label, x.Opt) for x in qs], words, {}))
    valid = set(k for k, _ in FORM_ROLES + SECTION_ROLES)
    fmap = dict((k.upper(), v) for k, v in (st.get('form_map') or {}).items() if k.upper() in auto)
    for qid in auto:
        v = _param('s_fm_' + qid.replace('-', ''), None)   # field names carry no hyphens
        if v is None:
            continue
        if v in valid and v != auto[qid]:
            fmap[qid] = v
        else:
            fmap.pop(qid, None)
    return fmap, list(auto)


def form_preview_html(st, rid):
    """One request read with st's form settings: what the lists, cards and
    Reports take from it, and what a committee member sees on the case page."""
    sub = find_submission(st, rid)
    if not sub:
        return '<div class="muted">That request was not found.</div>'
    _QUESTIONS.pop(int(st['form_org']), None)
    f = answers_for(st, [sub.Rid]).get(sub.Rid.upper()) or parse_answers([], st.get('form_words'))
    needs = ''.join('<li>%s%s%s%s</li>' % (esc(n['need']), (' · paid to ' + esc(n['payee'])) if n['payee'] else '',
                                           (' at ' + esc(n['address'])) if n.get('address') else '',
                                           ((' · ' + money(n['amount'])) if n['amount'] else ' · <span class="muted">no amount</span>') +
                                           (' · account on file' if n.get('account') else '')) for n in f['needs'])
    priv = sum(1 for sec in f['sections'] for i in sec['items'] if i['private'])
    return ('<div class="muted">%s, %s</div>'
            '<h3 style="margin:10px 0 4px">On the lists, the cards and in Reports</h3>'
            '<div>Asked for:</div><ul style="margin:4px 0">%s</ul>'
            '<div>Total asked: <b>%s</b> · People in the home: <b>%s</b> · Asked others for help: <b>%s</b>%s</div>'
            '<h3 style="margin:14px 0 4px">On the case page</h3>'
            '<div class="muted">What a committee member sees. %d private answer%s show only to routers.</div>%s') % (
        esc(sub.Name), nice_date(sub.T), needs or '<li><b>No needs found.</b> Mark the questions that ask about a need.</li>',
        money(f['total']) or 'not found', esc(f['people'] or 'not found'), esc(f['elsewhere'] or 'not found'),
        (' (' + esc(f['elsewhere_who']) + ')') if f['elsewhere_who'] else '', priv, '' if priv == 1 else 's',
        form_sections(f, False) or '<div class="muted">No answers to show.</div>')


def act_form_preview(st):
    """The live preview on Reading your form, with the choices on the page
    (not saved). Routers only: it shows a real request."""
    if not (ME and int(ME) in router_ids()):
        return {'success': True, 'html': '<div class="muted">Only routers see a preview from real requests.</div>'}
    if not int(st.get('form_org') or 0):
        return {'success': True, 'html': ''}
    words = dict(st.get('form_words') or {})
    for k, _ in FORM_WORDS:
        v = _param('s_fw_' + k, None)
        if v is not None:
            words[k] = v[:300]
    fmap, _ = posted_form_map(st, words)
    return {'success': True, 'html': form_preview_html(dict(st, form_words=words, form_map=fmap), _param('rid'))}


def form_match_card(st):
    """Reading your form: every question, what the tool takes it to be (worked
    out from the form's question types, its skip logic and the words below),
    a choice to change it, and the latest request read that way."""
    words = st.get('form_words') or {}
    phrase_rows = ''.join('<tr><td data-l="Looks for">%s</td><td><input type="text" name="fw_%s" value="%s" style="width:100%%"></td></tr>' % (
        esc(label), k, esc(words.get(k) or '')) for k, label in FORM_WORDS)
    phrases = ('<details class="bvc"><summary><b>Words used to guess</b> <span class="muted">(a fallback when the form itself does not say)</span></summary>'
               '<div class="muted">Words the tool looks for in a question, ignoring case; separate several with commas.</div>'
               '<table class="cmp"><colgroup><col style="width:34%%"><col></colgroup><tbody>%s</tbody></table></details>') % phrase_rows
    if not int(st.get('form_org') or 0):
        return ('<div><b>Reading your form</b><div class="muted">Set the request form above to see what the tool reads from it.</div>'
                '<div style="display:flex;flex-direction:column;gap:8px;margin:10px 0">%s%s</div></div>' % (starter_form_card(), phrases))
    qs = form_questions(st['form_org'])
    lay = form_layout([(x.Qid, x.T, x.Label, x.Opt) for x in qs], words, st.get('form_map'))
    # Real answers only for routers: an admin who is not one sees the form's
    # shape, never a request (admins get Settings, not anyone's request).
    see = bool(ME) and int(ME) in router_ids()
    latest = submissions(st, '', 1) if see else []
    example = {}
    if latest:
        for a in q.QuerySql("""SELECT CAST(RegQuestionId AS varchar(36)) Qid, AnswerValue A FROM RegAnswer
                               WHERE RegPeopleId = '%s'""" % sql_str(latest[0].Rid)):
            example[a.Qid.upper()] = answer_value(a.A).strip()
    seen = set(x.upper() for x in st.get('form_seen') or [])
    new = [x for x in lay if seen and x['qid'] not in seen]
    num = dict((x['qid'], i + 1) for i, x in enumerate(lay))

    def skips(x):
        """'No skips to 88. Do you need ... Electric' for a choice question with skip logic."""
        import json
        try:
            options = json.loads(x['opts'] or '[]') or []
        except Exception:
            return ''
        out = []
        for o in options:
            to = (o.get('SkipToId') or '').upper()
            if to:
                where = 'the end' if to == 'END' else ('%d. %s' % (num[to], lay[num[to] - 1]['label'][:50]) if to in num else 'a question no longer on the form')
                out.append('%s skips to %s' % (esc(o.get('Text') or o.get('Value') or '?'), esc(where)))
        return '<div class="muted">%s</div>' % '; '.join(out) if out else ''
    groups = [{'head': None, 'rows': [], 'open': False}]
    for i, x in enumerate(lay):
        badge = ' <span class="pill p-new">New</span>' if x in new else ''
        changed = ' <span class="muted">(you set this)</span>' if x['role'] != x['auto'] else ''
        if x['t'] == 4:
            opts = ''.join('<option value="%s"%s>%s</option>' % (k, ' selected' if k == x['role'] else '', esc(l)) for k, l in SECTION_ROLES)
            groups.append({'head': (i + 1, x['label'][:90] or '(heading)', badge + changed,
                                    '<select name="fm_%s">%s</select>' % (x['qid'].replace('-', ''), opts)),
                           'rows': [], 'open': bool(badge or changed)})
            continue
        g = groups[-1]
        if x['role'] == 'person':
            g['rows'].append('<tr><td colspan="%d" class="muted">%d. Their name and contact details (TouchPoint fills these in)</td></tr>' % (
                3 if see else 2, i + 1))
            continue
        opts = ''.join('<option value="%s"%s>%s</option>' % (k, ' selected' if k == x['role'] else '', esc(l)) for k, l in FORM_ROLES)
        part = ''
        if x['owner'] is not None:
            part = '<div class="muted">part of: %s</div>' % esc(need_name(lay[x['owner']]['label'], words))
        sk = skips(x)
        # A section opens by itself when there is something to check in it.
        if x['role'] != 'show' or badge or changed:
            g['open'] = True
        ex = example.get(x['qid'], '')
        ex = '(private)' if ex and x['role'] == 'private' else ex[:80]
        g['rows'].append('<tr><td data-l="Question"%s>%d. %s%s%s%s%s</td>%s<td data-l="This is"><select name="fm_%s">%s</select></td></tr>' % (
            ' style="padding-left:28px"' if x['owner'] is not None else '', i + 1, esc(x['label'][:140]), badge, changed, part, sk,
            ('<td data-l="Latest answer" class="muted">%s</td>' % esc(ex)) if see else '', x['qid'].replace('-', ''), opts))
    cols = ('<colgroup><col style="width:48%%"><col style="width:24%%"><col></colgroup>' if see
            else '<colgroup><col style="width:66%%"><col></colgroup>')
    sect = []
    for g in groups:
        if not g['rows'] and not g['head']:
            continue
        if g['head']:
            num_, name, marks, sel = g['head']
            title = '<span><b>%d. %s</b>%s <span class="muted">(%d question%s)</span></span>' % (
                num_, esc(name), marks, len(g['rows']), '' if len(g['rows']) == 1 else 's')
            g['rows'].insert(0, '<tr><td colspan="%d" class="muted"><span style="display:inline-flex;gap:8px;align-items:center">This heading is: %s</span></td></tr>' % (
                3 if see else 2, sel.replace('<select ', '<select style="width:auto" ')))
        else:
            title = '<span><b>Before the first heading</b> <span class="muted">(%d)</span></span>' % len(g['rows'])
        sect.append('<details class="fmsec"%s><summary>%s</summary><table class="cmp">%s<tbody>%s</tbody></table></details>' % (
            ' open' if g['open'] else '', title, cols, ''.join(g['rows'])))
    table = ('<div class="muted" style="margin:8px 0">Sections with something to check are open; the rest are just shown as asked. '
             'Open any section to look.</div>%s') % ''.join(sect)
    news = ('<div class="bvc" style="border-color:#C98A12"><b>%d new question%s on your form since you last saved here.</b> '
            'Check what the tool takes %s to be, then Save.</div>') % (len(new), '' if len(new) == 1 else 's', 'it' if len(new) == 1 else 'them') if new else ''
    preview = ''
    if see and latest:
        recent = submissions(st, '', 12)
        pick = ''.join('<option value="%s">%s, %s</option>' % (x.Rid, esc(x.Name), nice_date(x.T)) for x in recent)
        preview = ('<div class="fmprev bvc"><div class="bvr" style="justify-content:space-between;flex-wrap:wrap"><b>Preview</b>'
                   '<select id="fmrid" title="Request to preview">%s</select></div>'
                   '<div class="muted">Updates as you change the choices, before you save. Nothing here is saved.</div>'
                   '<div id="fmprev">%s</div></div>'
                   '<script>(function(){var f=document.getElementById("setf"), out=document.getElementById("fmprev"), t=null;'
                   'function go(){var d={action:"form_preview",rid:document.getElementById("fmrid").value};'
                   ' Array.prototype.forEach.call(f.querySelectorAll("select[name^=fm_],input[name^=fw_]"),function(i){d["s_"+i.name]=i.value;});'
                   ' out.style.opacity=".5"; bvPost(d,function(r){out.style.opacity="1"; if(r.success) out.innerHTML=r.html;});}'
                   'function soon(){clearTimeout(t); t=setTimeout(go,250);}'
                   'f.addEventListener("change",function(e){if(e.target.name&&(e.target.name.indexOf("fm_")==0||e.target.name.indexOf("fw_")==0)) soon();});'
                   'document.getElementById("fmrid").onchange=go;})();</script>') % (pick, form_preview_html(st, latest[0].Rid))
    return ('<div><b>Reading your form</b><div class="muted">What the tool takes each question on your request form to be. '
            'It works this out from the form itself (question types, sections and skip logic), so most churches only check it. '
            'Change any that are wrong, then Save. If your form asks for needs as one list (check all that apply), mark that '
            'question <b>A list of needs</b>. The documents form has its own tab: '
            '<a href="?view=settings&amp;sec=docs">Documents form</a>.%s</div>'
            '<div class="fmtools">%s%s</div>%s<div class="fmgrid"><div class="fmlist">%s</div>%s</div></div>') % (
        '' if see else ' Only routers see answers from real requests here.',
        phrases, starter_form_card(), news, table, preview)


def kind_row(k):
    return ('<tr><td data-l="Name"><input type="text" class="kname" value="%s" placeholder="Gas card" size="16"></td>'
            '<td data-l="Given"><select class="kpick"><option value="0"%s>Paid to the payee</option><option value="1"%s>Picked up</option></select></td>'
            '<td data-l="Requirements"><textarea class="kreqs" rows="2" placeholder="Driver\'s license">%s</textarea></td>'
            '<td><button type="button" class="btn red rmkind">Remove</button></td></tr>') % (
        esc(k.get('name') or ''), '' if k.get('pickup') else ' selected', ' selected' if k.get('pickup') else '', esc('\n'.join(k.get('reqs') or [])))


def access_role_field(st):
    """Pick the access role from the roles TouchPoint has, with how many people
    hold each, so a church can use a role it already has."""
    cur = access_role()
    rows = list(q.QuerySql("""
        SELECT r.RoleName N, COUNT(DISTINCT ur.UserId) U FROM Roles r LEFT JOIN UserRole ur ON ur.RoleId = r.RoleId
        GROUP BY r.RoleName ORDER BY r.RoleName"""))
    names = [r.N for r in rows]
    opts = ''.join('<option value="%s"%s>%s (%d %s)</option>' % (
        esc(r.N), ' selected' if r.N.lower() == cur.lower() else '', esc(r.N), r.U, 'person' if r.U == 1 else 'people') for r in rows)
    if cur.lower() not in [n.lower() for n in names]:
        opts = '<option value="%s" selected>%s (does not exist yet)</option>' % (esc(cur), esc(cur)) + opts
    return ('<label>Access role <select name="access_role">%s</select> '
            '<span class="muted">Marks the routers and hides the data. Give it to the routers only. Benevolence unless your church uses another; '
            'when you change it, the data is re-hidden under the new role the next time an admin opens the tool.</span></label>') % opts


def router_check(st, routers):
    """Who the routers are, whether each has the access role, and a way
    for the admin to add themselves."""
    pids = [int(r['pid']) for r in st.get('routers') or []]
    with_role = set()
    if pids:
        with_role = set(r.PeopleId for r in q.QuerySql("""
            SELECT DISTINCT u.PeopleId FROM Users u JOIN UserRole ur ON ur.UserId = u.UserId JOIN Roles ro ON ro.RoleId = ur.RoleId
            WHERE ro.RoleName = N'%s' AND u.PeopleId IN (%s)""" % (sql_str(access_role()), sql_ints(pids))))
    kind = dict((int(r['pid']), 'Main' if r.get('main') else 'Backup') for r in st.get('routers') or [])
    items = ''.join('<div><b>%s</b> %s %s</div>' % (kind.get(p, ''), esc(routers.get(p, 'PeopleId %d' % p)),
                                          '<span class="pill p-ok">has the %s role</span>' % esc(access_role()) if p in with_role else
                                          '<span class="pill p-bad">needs the %s role in TouchPoint to get in</span>' % esc(access_role())) for p in pids)
    me = '' if not ME or int(ME) in pids else (
        '<div class="muted">Your PeopleId is %d. <button type="button" class="btn ghost" onclick="var i=document.querySelector(\'#setf input[name=routers_main]\');'
        'i.value=(i.value?i.value+\', \':\'\')+\'%d\';">Add me</button> then Save settings.</div>' % (int(ME), int(ME)))
    return ('<div style="margin-left:4px">%s%s<div class="muted">Routers see every request. Being an admin opens only this page.</div></div>'
            % (items or '<div class="muted">No routers yet, so no one can see requests.</div>', me))


def act_save_settings(st):
    """Each Settings section saves only its own fields. TouchPoint returns ''
    for a field that was not sent, so the section names itself (s_section)
    rather than relying on which fields arrived; an emptied field then still
    counts as sent."""
    if not IS_ADMIN:
        raise Exception('Settings are for admins.')
    SECTION_FIELDS = {'setup': ['routers_main'], 'form': ['fw_' + k for k, _ in FORM_WORDS], 'call': ['call_questions'],
                      'decisions': ['affirm_on', 'affirm_required', 'handoff_on', 'req_why_help', 'req_talked_help', 'req_pickup', 'req_why_decline', 'earlier_on', 'router_decide', 'stuck_days', 'affirm_over', 'decline_reasons', 'referrals',
                                    'pickup_places', 'help_kinds'],
                      'emails': ['email_on', 'email_router', 'email_committee', 'email_secure', 'notify_assigned', 'notify_shared',
                                 'notify_decision', 'notify_handoff', 'notify_docs', 'notify_paid', 'notify_nudge', 'notify_resent', 'email_secure_when', 'email_time', 'email_days',
                                 'email_backup_days', 'email_from', 'email_from_name', 'email_tagline'],
                      'texts': ['text_group', 'text_church', 'text_phone', 'text_autosignin', 'text_secure', 'text_docs']}
    sent = set(SECTION_FIELDS.get(_param('s_section'), []))
    has = lambda k: k in sent or _param('s_' + k, None) is not None

    def num(k):
        v = _param('s_' + k)
        return int(v) if v.isdigit() else 0
    if has('routers_main'):
        st['form_org'] = num('form_org')
        st['docs_org'] = num('docs_org')
        st['committee_org'] = num('committee_org')
        mains = [int(x) for x in re.findall(r'\d+', _param('s_routers_main'))]
        backups = [int(x) for x in re.findall(r'\d+', _param('s_routers_backup')) if int(x) not in mains]
        st['routers'] = [{'pid': x, 'main': True} for x in mains] + [{'pid': x, 'main': False} for x in backups]
        st['merger_pid'] = num('merger_pid')
        ar = _param('s_access_role').strip()[:50]
        if ar and ar != (st.get('access_role') or 'Benevolence'):
            if not q.QuerySqlTop1("SELECT 1 X FROM Roles WHERE RoleName = N'%s'" % sql_str(ar)):
                raise Exception('There is no %s role in TouchPoint. Create it first (Admin, Roles).' % ar)
            st['access_role'] = ar
        gl = _param('s_go_live')
        st['go_live'] = gl if re.match(r'^\d{4}-\d{2}-\d{2}$', gl) else ''
    for k in ('affirm_on', 'affirm_required', 'handoff_on', 'email_on', 'email_router', 'email_committee', 'email_secure',
              'notify_assigned', 'notify_shared', 'notify_decision', 'notify_handoff', 'notify_docs', 'notify_paid', 'notify_nudge', 'notify_resent',
              'req_why_help', 'req_talked_help', 'req_pickup', 'req_why_decline', 'earlier_on'):
        if has(k):
            st[k] = _param('s_' + k) == '1'
    if has('router_decide'):
        st['router_decide'] = _param('s_router_decide') if _param('s_router_decide') in ('none', 'decline', 'both') else 'decline'
    if has('stuck_days'):
        st['stuck_days'] = max(1, min(60, int(to_number(_param('s_stuck_days')) or 7)))
    if has('email_secure_when'):
        st['email_secure_when'] = 'now' if _param('s_email_secure_when') == 'now' else 'later'
    if _param('s_section') == 'emails':
        st['notify_when'] = dict((k, 'now') for k, _, _, _ in NOTICES if _param('s_nw_' + k) == 'now')
    if has('affirm_over'):
        st['affirm_over'] = max(0, to_number(_param('s_affirm_over')) or 0)
    if has('text_group'):
        g = _param('s_text_group')
        st['text_group'] = int(g) if g.isdigit() else 0
        if st['text_group'] and not q.QuerySqlTop1("SELECT 1 X FROM SMSGroups WHERE Id = %d" % st['text_group']):
            raise Exception('That texting group was not found.')
        st['text_church'] = _param('s_text_church').strip()[:40]
        st['text_phone'] = _param('s_text_phone').strip()[:30]
        st['text_autosignin'] = _param('s_text_autosignin') == '1'
        for k in ('text_secure', 'text_docs'):
            t = ' '.join(_param('s_' + k).split())[:400] or DEFAULT_SETTINGS[k]
            if '{link}' not in t:
                raise Exception('Each text needs {link} in it, where the link goes.')
            st[k] = t
    if has('email_time'):
        tm = _param('s_email_time')
        st['email_time'] = tm if re.match(r'^\d{2}:\d{2}$', tm) else '07:00'
        st['email_days'] = 'all' if _param('s_email_days') == 'all' else 'weekdays'
        bd = _param('s_email_backup_days')
        st['email_backup_days'] = max(1, min(30, int(bd))) if bd.isdigit() else 3
        em = _param('s_email_from')
        if em and not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', em):
            raise Exception('The from email does not look like an email address.')
        st['email_from'] = em
        st['email_from_name'] = _param('s_email_from_name')[:60] or 'Benevolence'
        st['email_tagline'] = _param('s_email_tagline').strip()[:300]
        # Keep our line in ScheduledTasks in step with the switch.
        if st['email_on'] and not sched_install():
            raise Exception('Saved, but the line could not be added to ScheduledTasks.')
        if not st['email_on']:
            sched_uninstall()
    words = dict(st.get('form_words') or {})
    for k, _ in FORM_WORDS:
        if has('fw_' + k):
            words[k] = _param('s_fw_' + k)[:300]
    st['form_words'] = words
    # What each question is: kept only where it differs from what the tool
    # works out, so better guesses (or changed words) still reach the rest.
    if _param('s_section') == 'form' and int(st.get('form_org') or 0):
        st['form_map'], st['form_seen'] = posted_form_map(st, words)
    if _param('s_section') == 'docs' and int(st.get('docs_org') or 0):
        dmap = {}
        for x in form_questions(st['docs_org']):
            if x.T in (4, 5):
                continue
            qid = x.Qid.upper()
            v = _param('s_dm_' + qid.replace('-', ''), None)
            old = dict((k.upper(), w) for k, w in (st.get('docs_map') or {}).items()).get(qid)
            v = v if v is not None else old
            if v in ('show', 'private', 'hide') and v != docs_auto_role(x.T, strip_tags(x.Label), st):
                dmap[qid] = v
        st['docs_map'] = dmap
    if has('call_questions'):
        text = _param('s_call_questions')[:20000]
        if not parse_call_questions(text):
            raise Exception('Add at least one question.')
        st['call_questions'] = text
    for k in ('decline_reasons', 'referrals'):
        if has(k):
            st[k] = unique_lines(_param('s_' + k), 80)
    if has('pickup_places'):
        places = unique_lines(_param('s_pickup_places'), 60)
        if not places:
            raise Exception('Keep at least one pickup place.')
        st['pickup_places'] = places
    if has('help_kinds'):
        import json
        kinds, seen = [], set()
        try:
            raw = json.loads(_param('s_help_kinds') or '[]')
        except Exception:
            raise Exception('The kinds of help could not be read. Reload the page.')
        for k in raw[:30]:
            name = unicode(k.get('name') or '').strip()[:60]
            if not name or name.lower() in seen:
                continue
            seen.add(name.lower())
            reqs = [unicode(x).strip()[:80] for x in k.get('reqs') or [] if unicode(x).strip()][:10]
            kinds.append({'name': name, 'pickup': bool(k.get('pickup')), 'reqs': reqs})
        if not kinds:
            raise Exception('Keep at least one kind of help.')
        st['help_kinds'] = kinds
    save_settings(st)
    return {'success': True}


def unique_lines(text, size):
    out = []
    for x in (text or '').split('\n'):
        x = x.strip()[:size]
        if x and x not in out:
            out.append(x)
    return out


# ---------------------------------------------------------------- actions
def act_assign_note(st, role):
    """The router's note for the committee member, saved as they type."""
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].setdefault(s.Rid, {})
    if r.get('status') not in (None, 'new'):
        return {'success': True}          # already given out: the note went with it
    r['assign_note'] = clean_html(from_b64(_param('html')))
    save_record(s.Pid, rec)
    return {'success': True}


def act_assign(st, role):
    rid = _param('rid')
    s = find_submission(st, rid)
    if not s:
        raise Exception('Request not found.')
    who = int(_param('who', '0') or 0)
    members = dict((m['pid'], m['name']) for m in committee_members(st))
    if who not in members:
        raise Exception('Pick someone on the committee.')
    records = all_records()
    rec = load_record(s.Pid)
    r = rec['requests'].setdefault(rid, {})
    if _param('reassign') == '1':
        if r.get('status') not in ('committee', 'requester'):
            raise Exception('This request is not with the committee.')
        was = int(r.get('owner') or 0)
        if who == was:
            raise Exception('It is already with them. Pick someone else.')
        old = names_for([was]).get(was, 'someone')
        r['owner'] = who
        r.pop('handoff', None)
        for p in shared_pids(r):
            share_end(r, p, 'handed to %s' % members[who])
        r.pop('shared', None)
        r.pop('affirm', None)
        r.setdefault('assigned', []).append({'pid': who, 't': now_str(), 'by': int(ME), 'from': was})
        add_history(r, 'Handed to %s (was with %s)' % (members[who], old), clean_html(from_b64(_param('html'))))
        notify(r, 'assigned', [who], rid=rid)
        save_record(s.Pid, rec)
        return {'success': True, 'num': r.get('num'), 'next': '?view=home&given=%s' % rid}
    if r.get('status') not in (None, 'new'):
        raise Exception('This request was already given out. Reload the page.')
    if not r.get('num'):
        r['num'] = next_number(records)
    r.pop('assign_note', None)
    r['status'] = 'committee'
    r['owner'] = who
    r['came'] = s.T
    r.setdefault('assigned', []).append({'pid': who, 't': now_str(), 'by': int(ME)})
    add_history(r, 'Given to %s' % members[who], clean_html(from_b64(_param('html'))))
    notify(r, 'assigned', [who], rid=rid)
    save_record(s.Pid, rec)
    return {'success': True, 'num': r['num'], 'next': '?view=home&given=%s' % rid}


# ---------------------------------------------------------------- the committee member's side
def my_requests(st):
    """(given to me, shared with me for affirmation): [(submission, record)]."""
    records = all_records()
    mine, shared = [], []
    want = []
    for pid, reqs in records.items():
        for rid, r in reqs.items():
            if r.get('status') not in ('committee', 'requester'):
                continue
            if int(r.get('owner') or 0) == int(ME):
                want.append((rid, 'mine'))
            elif int(ME) in [int(x) for x in r.get('shared') or []]:
                want.append((rid, 'shared'))
    if not want:
        return [], []
    subs = dict((x.Rid.upper(), x) for x in submissions(st, "AND rp.RegPeopleId IN (%s)" % ','.join("'%s'" % w for w, _ in want), len(want)))
    for rid, kind in want:
        s = subs.get(rid.upper())
        if s:
            (mine if kind == 'mine' else shared).append((s, (records.get(s.Pid) or {}).get(rid)))
    return mine, shared


def stage_of(r):
    d = r.get('decision') or {}
    if r.get('handoff'):
        return ('Asked to hand off', 'p-warn', 'stays with you until the router passes it on')
    if (r.get('docs_in') or '') > (r.get('docs_looked') or ''):
        return ('Documents came in: take a look', 'p-warn', 'they sent them %s' % days_ago(r['docs_in']))
    if r.get('status') == 'requester':
        return ('Waiting on them for documents', 'p-grey', 'asked %s' % days_ago(([x.get('t') for x in r.get('texts') or [] if x.get('kind') == 'docs'] or [''])[-1]))
    if d.get('saved'):
        return ('Your turn: finish the decision', 'p-new', 'saved %s' % days_ago(d['saved']))
    if call_started(r) or reached(r):
        return ('Your turn: decide', 'p-new', 'talked with them' if not call_started(r) else 'call saved %s' % days_ago((r.get('call') or {}).get('saved') or ''))
    if r.get('contacts') and not reached(r):
        return ('Try again', 'p-warn', contacts_line(r))
    return ('Your turn: call them', 'p-new', 'given to you %s' % days_ago((r.get('assigned') or [{}])[-1].get('t', '')))


def next_button(r, s):
    """The one obvious next step for a request on a member's list."""
    if (r.get('decision') or {}).get('saved') or call_started(r) or reached(r):
        return '<a class="btn" href="?view=req&amp;rid=%s#dec">Finish the decision</a>' % s.Rid
    return '<a class="btn" href="?view=call&amp;rid=%s">%s</a>' % (s.Rid, 'Try again' if r.get('contacts') and not reached(r) else 'Start the call')


def my_recent(st, days=120):
    """The member's own cases decided in the last few months, newest first:
    [(submission, record, last history entry)]."""
    records = all_records()
    cut = (datetime.date.today() - datetime.timedelta(days=days)).strftime('%Y-%m-%d')
    want = []
    for pid, reqs in records.items():
        for rid, r in reqs.items():
            if r.get('status') in ('approved', 'pickup', 'done', 'declined') and int(r.get('owner') or 0) == int(ME or 0):
                hist = sorted(r.get('history') or [], key=lambda x: x.get('t') or '')
                last = hist[-1] if hist else {}
                if (last.get('t') or '')[:10] >= cut:
                    want.append((last.get('t') or '', rid, r, last))
    want.sort(reverse=True)
    want = want[:15]
    if not want:
        return []
    subs = dict((x.Rid.upper(), x) for x in submissions(st, "AND rp.RegPeopleId IN (%s)" % ','.join("'%s'" % w[1] for w in want), len(want)))
    return [(subs[rid.upper()], r, last) for t, rid, r, last in want if rid.upper() in subs]


def recent_state(st, r):
    """Where a decided case stands, in the words a member would say on the phone."""
    d = r.get('decision') or {}
    stt = r.get('status')
    total = sum(to_number(x.get('amount')) or 0 for x in d.get('items') or [] if x.get('help'))
    if stt == 'declined':
        return ('Declined', 'p-bad', d.get('decline_reason') or '')
    if stt == 'approved':
        return ('Approved %s' % money(total), 'p-new', 'the office is getting it paid')
    if stt == 'pickup':
        when = pickup_when(d)
        return ('Ready for pickup', 'p-warn', '%s%s' % (d.get('pickup') or 'place not set', (', ' + when) if when else ''))
    return ('Done', 'p-ok', outcome(r)[1])


def page_mine(st, role):
    docs_sweep(st)
    mine, shared = my_requests(st)
    h = [announcements_box(), "<h1>People I'm helping</h1>"]
    if not mine:
        h.append('<div class="bvc"><b>No one right now.</b><span class="muted">When the router gives you a request, it shows here.</span></div>')
    cache = answers_for(st, [s.Rid for s, r in mine + shared])
    for s, r in mine:
        need, total = need_summary(st, s.Rid, cache)
        label, cls, when = stage_of(r)
        if req_unread(s.Rid, r):
            when += ' · new message'
        h.append('<div class="bvc"><div class="bvr"><div><b style="font-size:calc(19px*var(--z,1))">%s · #%s</b><div>%s · %s asked</div>'
                 '<div class="muted">%s</div></div><span class="pill %s">%s</span></div>'
                 '<div class="bvr" style="justify-content:flex-start">%s<a class="btn ghost" href="?view=req&amp;rid=%s">Open their request</a></div>%s</div>' % (
                     esc(s.Name), r.get('num'), esc(need), money(total) or 'amount not given', esc(when), cls, esc(label),
                     next_button(r, s), s.Rid,
                     handoff_member_card(st, s, r, True) if r.get('status') in ('committee', 'requester') else ''))
    if shared:
        h.append('<h2>Cases you were brought in on</h2>')
        owners = names_for([r.get('owner') for s, r in shared])
        for s, r in shared:
            need, total = need_summary(st, s.Rid, cache)
            aff = (r.get('affirm') or {})
            ask = 'asked you to affirm' if int(aff.get('pid') or 0) == int(ME) and not aff.get('note') else 'brought you in on'
            new = len(req_unread(s.Rid, r))
            h.append('<div class="bvc"><div class="bvr"><div><b>%s %s #%s · %s</b><div class="muted">%s · %s%s</div></div>'
                     '<a class="btn ghost" href="?view=req&amp;rid=%s">Open the case</a></div></div>' % (
                         esc(owners.get(int(r.get('owner') or 0), '')), ask, r.get('num'), esc(s.Name), esc(need), money(total),
                         (' · %d new message%s' % (new, '' if new == 1 else 's')) if new else '', s.Rid))
    recent = my_recent(st)
    if recent:
        h.append('<h2>Your recent cases</h2><div class="muted">What happened after you sent your decision, in case they call you. '
                 'Open one to see the whole case.</div>')
        for s, r, last in recent:
            lab, cls, more = recent_state(st, r)
            h.append('<div class="bvc"><div class="bvr"><div><b>%s · #%s</b><div>%s</div><div class="muted">Last: %s %s</div></div>'
                     '<div style="text-align:right"><span class="pill %s">%s</span><div><a href="?view=req&amp;rid=%s">Open</a></div></div></div></div>' % (
                         esc(s.Name), r.get('num'), esc(more), esc(last.get('what') or ''), days_ago(last.get('t') or ''), cls, esc(lab), s.Rid))
    h.append('<div class="muted">Your list is the people the router gave you, plus any cases a teammate brought you in on. '
             'Those are theirs to decide; you can read them, write in the conversation and add notes.</div>')
    return ''.join(h)


def help_kind(st, name):
    for k in st.get('help_kinds') or []:
        if k['name'] == name:
            return k
    return {}


def needs_pickup(st, name):
    return bool(help_kind(st, name).get('pickup'))


def help_kind_options(st, chosen):
    kinds = [k['name'] for k in st.get('help_kinds') or []]
    if chosen and chosen not in kinds:
        kinds.append(chosen)
    return ''.join('<option%s>%s</option>' % (' selected' if k == chosen else '', esc(k)) for k in kinds)


def stipulation(st, how):
    return ', '.join(help_kind(st, how).get('reqs') or [])


def first_paid(st):
    """The first help type paid to the payee, the default for a bill."""
    for k in st.get('help_kinds') or []:
        if not k.get('pickup'):
            return k['name']
    return ''


def item_deliv(st, d, x):
    """'pickup' when they come to get this one, else 'mail' (the church pays
    it or mails it). A kind that is always picked up (a gas card) wins."""
    if needs_pickup(st, x.get('how')):
        return 'pickup'
    if x.get('deliv') in ('mail', 'pickup'):
        return x['deliv']
    return 'pickup' if d.get('delivery') == 'pickup' else 'mail'   # saved before each line had its own


def form_need_names(st):
    """The needs the request form asks about, in form order: the categories
    for help a committee member adds."""
    try:
        org = int(st.get('form_org') or 0)
    except Exception:
        return []
    if not org:
        return []
    if org not in _QUESTIONS:
        _QUESTIONS[org] = form_questions(org)
    import json
    words, out = st.get('form_words'), []
    for x in form_layout([(qq.Qid, qq.T, qq.Label, qq.Opt) for qq in _QUESTIONS[org]], words, st.get('form_map')):
        if x['role'] == 'need':
            out.append(need_name(x['label'], words))
        elif x['role'] == 'needs':
            try:
                out += [(o.get('Text') or o.get('Value') or '').strip()[:60] for o in json.loads(x['opts'] or '[]') or []]
            except Exception:
                pass
    seen = []
    for n in out:
        if n and n.lower() not in [z.lower() for z in seen]:
            seen.append(n)
    return seen


def added_need(cat, detail):
    """'Electric (past due notice)' from the category picked and the words typed."""
    cat, detail = (cat or '').strip(), (detail or '').strip()
    if cat and cat != 'Other':
        return '%s (%s)' % (cat, detail) if detail else cat
    return detail


def decision_lines(st, f, d):
    """The request's own needs, each with what the member chose, then any
    help they added: [{src, need, payee, address, account, asked, help,
    amount, how, deliv}]. form_*: what the form said, when it was changed."""
    saved = dict((x.get('src'), x) for x in d.get('items') or [] if x.get('src') is not None)
    lines = []
    for i, n in enumerate(f['needs']):
        x = saved.get(str(i)) or {}
        how = x.get('how') or (first_paid(st) if n['payee'] else '')
        lines.append({'src': str(i), 'need': n['need'], 'asked': n['amount'], 'help': bool(x.get('help')), 'amount': x.get('amount', ''),
                      'how': how, 'deliv': item_deliv(st, d, dict(x, how=how)),
                      'payee': x.get('payee') or n['payee'], 'address': x.get('address') or n.get('address') or '',
                      'account': x.get('account') or n.get('account') or '',
                      'form_payee': n['payee'], 'form_address': n.get('address') or '', 'form_account': n.get('account') or ''})
    for x in d.get('items') or []:
        if x.get('src') is None:
            cat, detail = x.get('cat'), x.get('detail')
            if cat is None and detail is None:
                cat, detail = 'Other', x.get('need') or ''      # added before categories
            lines.append({'src': None, 'need': x.get('need') or '', 'cat': cat or '', 'detail': detail or '', 'asked': None,
                          'help': True, 'amount': x.get('amount', ''), 'how': x.get('how') or '', 'deliv': item_deliv(st, d, x),
                          'payee': x.get('payee') or '', 'address': x.get('address') or '', 'account': x.get('account') or ''})
    return lines


def deliv_options(deliv):
    return ''.join('<option value="%s"%s>%s</option>' % (v, ' selected' if v == deliv else '', t)
                   for v, t in (('mail', 'Benevolence pays it'), ('pickup', 'They pick it up')))


def cat_options(st, chosen):
    cats = form_need_names(st)
    if chosen and chosen not in cats and chosen != 'Other':
        cats.append(chosen)
    return ('<option value="">Pick one</option>' + ''.join('<option%s>%s</option>' % (' selected' if c == chosen else '', esc(c)) for c in cats) +
            '<option value="Other"%s>Other</option>' % (' selected' if chosen == 'Other' else ''))


def payto_html(x, show_acct):
    """Who the line is paid to, from the form, with boxes to add or correct
    it. The form's account number is routers only: anyone else gets a blank
    box (blank keeps what is on file) and the last four digits."""
    acct = x.get('account') or ''
    hidden = acct and not show_acct and x['src'] is not None
    said = []
    if x['src'] is not None:
        if x.get('form_payee') and x['form_payee'] != x['payee']:
            said.append('payee ' + esc(x['form_payee']))
        if x.get('form_address') and x['form_address'] != x['address']:
            said.append(esc(x['form_address']))
        if show_acct and x.get('form_account') and x['form_account'] != acct:
            said.append('account ' + esc(x['form_account']))
    return ('<details class="d-pt"%s><summary class="muted">Pay to: <span class="d-pts"></span></summary>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap;gap:6px">'
            '<input type="text" class="d-payee" value="%s" placeholder="Payee" size="18">'
            '<input type="text" class="d-addr" value="%s" placeholder="Payee address" size="28">'
            '<input type="text" class="d-acct" value="%s" placeholder="%s" size="16" data-mask="%s"></div>%s</details>') % (
        ' open' if x['src'] is None and not x.get('payee') else '', esc(x.get('payee') or ''), esc(x.get('address') or ''),
        '' if hidden else esc(acct), ('Change account (on file, ending %s)' % esc(acct[-4:])) if hidden else 'Account or loan number',
        esc(('ending ' + acct[-4:]) if hidden else ''),
        ('<div class="muted">The form said: %s</div>' % ' · '.join(said)) if said else '')


def decision_card(st, s, r, f, show_acct=False):
    """The committee member's decision: which of the needs to help with, how
    much and how, who it is paid to or that they pick it up, any other help,
    the pickup place and the reason. Saved as they go. show_acct: the viewer
    is a router, who sees the form's account numbers."""
    d = r.get('decision') or {}
    choice = 'decline' if d.get('choice') == 'decline' else 'help'
    lines = decision_lines(st, f, d)
    rows = []
    blank = {'src': None, 'need': '', 'cat': '', 'detail': '', 'asked': None, 'help': True, 'amount': '', 'how': '', 'deliv': 'mail',
             'payee': '', 'address': '', 'account': ''}
    for i, x in enumerate(lines + [blank]):
        stip = stipulation(st, x['how'])
        what = (esc(x['need']) if x['src'] is not None else
                '<select class="d-cat">%s</select> <input type="text" class="d-need" value="%s" placeholder="What (more detail)" size="16">' % (
                    cat_options(st, x.get('cat')), esc(x.get('detail') or '')))
        rows.append(
            '<tr data-src="%s" data-need="%s"><td data-l="Help"><label class="pick"><input type="checkbox" class="d-help"%s>'
            '<span><b>%s</b></span></label>%s</td>'
            '<td data-l="Asked">%s</td>'
            '<td data-l="Give"><input type="text" class="d-amt" size="8" value="%s" placeholder="$"></td>'
            '<td data-l="How"><select class="d-how">%s</select><div class="muted d-stip">%s</div></td>'
            '<td data-l="Getting it to them"><select class="d-deliv">%s</select></td>%s</tr>' % (
                '' if x['src'] is None else esc(x['src']), esc(x['need']) if x['src'] is not None else '', ' checked' if x['help'] else '', what, payto_html(x, show_acct),
                money(x['asked']) or ('' if x['src'] is None else 'not given'), esc(x['amount']),
                help_kind_options(st, x['how']), ('They bring: ' + esc(stip)) if stip else '', deliv_options(x['deliv']),
                '' if x['src'] is not None else '<td><button type="button" class="btn red d-del" title="Remove">&times;</button></td>'))
    template = rows.pop()                 # the blank line: what + Add other help puts in
    plist = list(st.get('pickup_places') or [])
    if d.get('pickup') and d['pickup'] not in plist:
        plist.append(d['pickup'])
    places = '<option value="">Pick a place</option>' + ''.join('<option%s>%s</option>' % (' selected' if p == d.get('pickup') else '', esc(p)) for p in plist)
    reasons = ''.join('<option%s>%s</option>' % (' selected' if x == d.get('decline_reason') else '', esc(x)) for x in st.get('decline_reasons') or [])
    refs = '<option value="">(none)</option>' + ''.join('<option%s>%s</option>' % (' selected' if x == d.get('referral') else '', esc(x)) for x in st.get('referrals') or [])
    kinds_js = safe_json(dict((k['name'], {'n': ', '.join(k.get('reqs') or []), 'p': bool(k.get('pickup'))}) for k in st.get('help_kinds') or []))
    return ('<div class="bvc bvblue" id="dec" data-rid="%s"><div class="bvr"><h2>Step 2: Your decision</h2><span class="muted" id="decst">%s</span></div>'
            '<div class="bvr" style="justify-content:flex-start">'
            '<label class="pick"><input type="radio" name="dchoice" value="help"%s> Help</label>'
            '<label class="pick"><input type="radio" name="dchoice" value="decline"%s> Decline</label></div>'
            '<div id="dhelp"><b>What we will give</b><div class="muted">Tick each thing you will help with and how much. '
            'Add anything outside what they asked for.</div>'
            '<table><thead><tr><th>Their request</th><th>Asked</th><th>Give</th><th>How</th><th>Getting it to them</th></tr></thead><tbody id="dlines">%s</tbody></table>'
            '<div class="bvr"><button type="button" class="btn ghost" id="dadd">+ Add other help</button><b id="dtotal"></b></div>'
            '%s'
            '<div id="dpick" class="shform"><b>Pickup</b> <span class="muted">%s</span>'
            '<div class="muted" id="dpforced"></div>'
            '<div id="dpwhen"><div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">'
            '<label>Where <select id="dpickup">%s</select></label>'
            '<label>Day <input type="date" id="dpday" value="%s"></label><label>Time <input type="time" id="dptime" value="%s"></label></div>'
            '<div class="muted" id="dpbring"></div></div></div></div>'
            '<div id="ddecline"><label>Why <select id="dreason">%s</select></label> '
            '<label>Send them to <select id="dref">%s</select></label>'
            '<div class="muted">The reason goes to the router and the timeline, never to them by email.</div></div>'
            '<div><b>Why (the router sees this; it stays in their timeline)</b>%s</div>'
            '<div class="shform" id="dchk"><b>Before you can send</b><ul id="dchkl" style="margin:0;padding-left:4px;list-style:none"></ul></div>'
            '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn green" id="dsend">Send to the router</button>'
            '<span class="muted">Not ready? Just leave. Everything is saved and stays on your list.</span></div>%s</div>'
            '<script>(function(){var KINDS=%s, RULES=%s, card=document.getElementById("dec"), rid=card.getAttribute("data-rid"), t=null;'
            'function checks(){var v=val(), out=[], why=(card.querySelector(".rt").textContent||"").trim().length>0;'
            ' if(v.choice=="help"){var g=v.items.filter(function(x){return x.help;}), tot=0;'
            '  out.push([g.length>0,"Something ticked to help with"]);'
            '  var amt=g.every(function(x){var n=parseFloat((x.amount||"").replace(/[^0-9.]/g,"")); if(n) tot+=n; return n>0&&x.how;});'
            '  if(g.some(function(x){return x.src===null;})) out.push([g.every(function(x){return x.src!==null||x.need;}),"Each added line says what it is"]);'
            '  var noamt=g.filter(function(x){return !(parseFloat((x.amount||"").replace(/[^0-9.]/g,""))>0);}).map(function(x){return x.need||"an added line";});'
            '  if(g.length) out.push([amt,noamt.length?"Type an amount for "+noamt.join(", "):"Every ticked line has an amount and how"]);'
            '  if(RULES.pickup_req&&g.some(function(x){return x.deliv=="pickup";})) out.push([!!(v.pickup&&v.pickup_day&&v.pickup_time),"Pickup place, day and time"]);'
            '  if(RULES.why_help) out.push([why,"The Why is written"]);'
            '  if(RULES.talked_req) out.push([RULES.talked,"You talked with them (call notes or a try logged as talked with them)"]);'
            '  if(RULES.aff_req||(RULES.aff_over&&tot>RULES.aff_over)) out.push([RULES.aff_done,"A teammate is affirming it"+(RULES.aff_req?"":" (help over $"+RULES.aff_over.toLocaleString()+")")]);}'
            ' else {out.push([!!v.decline_reason,"A reason for declining is picked"]);'
            '  if(RULES.why_decline) out.push([why,"The Why is written"]);'
            '  if(RULES.aff_req) out.push([RULES.aff_done,"A teammate is affirming it"]);}'
            ' document.getElementById("dchkl").innerHTML=out.map(function(c){return \'<li style="padding:2px 0">\'+(c[0]?\'<b style="color:#1E6B45">&#10003;</b> \':\'<b style="color:#A12A22">&#10007;</b> \')+c[1]+"</li>";}).join("");'
            ' var ok=out.every(function(c){return c[0];}); var b=document.getElementById("dsend"); b.disabled=!ok;'
            ' b.textContent=ok?"Send to the router":"Finish the list above to send"; document.getElementById("dchk").style.display=ok?"none":"";}'
            'function val(){var c=card.querySelector("input[name=dchoice]:checked"); var items=[];'
            ' Array.prototype.forEach.call(card.querySelectorAll("#dlines tr"),function(tr){var src=tr.getAttribute("data-src");'
            '  var it={src:src===""?null:src, need:tr.getAttribute("data-need")||null, help:tr.querySelector(".d-help").checked, amount:tr.querySelector(".d-amt").value,'
            '   how:tr.querySelector(".d-how").value, deliv:tr.querySelector(".d-deliv").value, payee:tr.querySelector(".d-payee").value,'
            '   address:tr.querySelector(".d-addr").value, account:tr.querySelector(".d-acct").value};'
            '  if(it.src===null){it.cat=tr.querySelector(".d-cat").value; it.detail=tr.querySelector(".d-need").value;'
            '   it.need=it.cat&&it.cat!="Other"?it.cat+(it.detail.trim()?" ("+it.detail.trim()+")":""):it.detail.trim();}'
            '  items.push(it);});'
            ' return {choice:c?c.value:"help", items:items, pickup:document.getElementById("dpickup").value,'
            '  pickup_day:document.getElementById("dpday").value, pickup_time:document.getElementById("dptime").value,'
            '  decline_reason:document.getElementById("dreason").value, referral:document.getElementById("dref").value,'
            '  reason:card.querySelector(".rt").innerHTML};}'
            'function show(){var c=val().choice; document.getElementById("dhelp").style.display=c=="help"?"":"none";'
            ' document.getElementById("ddecline").style.display=c=="decline"?"":"none";'
            ' var tot=0; val().items.forEach(function(x){ if(x.help){var n=parseFloat((x.amount||"").replace(/[^0-9.]/g,"")); if(n) tot+=n;} });'
            ' document.getElementById("dtotal").textContent="Total $"+tot.toLocaleString(undefined,{maximumFractionDigits:2});'
            ' Array.prototype.forEach.call(card.querySelectorAll("#dlines tr"),function(tr){var s=(KINDS[tr.querySelector(".d-how").value]||{}).n;'
            '  tr.querySelector(".d-stip").textContent=s?"They bring: "+s:"";});'
            ' Array.prototype.forEach.call(card.querySelectorAll("#dlines tr"),function(tr){var dl=tr.querySelector(".d-deliv"), k=KINDS[tr.querySelector(".d-how").value]||{};'
            '  var am=tr.querySelector(".d-amt"), miss=tr.querySelector(".d-help").checked&&!(parseFloat((am.value||"").replace(/[^0-9.]/g,""))>0);'
            '  am.style.borderColor=miss?"#A12A22":""; am.style.background=miss?"#FBEBEB":""; am.placeholder=miss?"Amount?":"$";'
            '  am.title=miss?"Ticked, but no amount yet":""; tr.classList.toggle("d-miss",miss);'
            '  if(k.p){dl.value="pickup"; dl.disabled=true; dl.title=tr.querySelector(".d-how").value+" is always picked up";} else {dl.disabled=false; dl.title="";}'
            '  var pt=tr.querySelector(".d-pt"); pt.style.display=dl.value=="pickup"?"none":"";'
            '  var a=tr.querySelector(".d-acct"), bits=[tr.querySelector(".d-payee").value.trim()||"not given yet", tr.querySelector(".d-addr").value.trim()];'
            '  bits.push(a.value.trim()?"account "+a.value.trim():(a.getAttribute("data-mask")?"account "+a.getAttribute("data-mask"):""));'
            '  tr.querySelector(".d-pts").textContent=bits.filter(function(b){return b;}).join(" \u00b7 ");'
            '  var cs=tr.querySelector(".d-cat"); if(cs) tr.querySelector(".d-need").placeholder=cs.value=="Other"||!cs.value?"What":"More detail (optional)";});'
            ' var pk=val().items.filter(function(x){return x.help&&x.deliv=="pickup";});'
            ' document.getElementById("dpick").style.display=pk.length?"":"none";'
            ' document.getElementById("dpforced").textContent=pk.length?"Picked up: "+pk.map(function(x){return x.need||x.how;}).filter(function(n,i,a){return n&&a.indexOf(n)==i;}).join(", ")+".":"";'
            ' var br=pk.map(function(x){return (KINDS[x.how]||{}).n;}).filter(function(n,i,a){return n&&a.indexOf(n)==i;});'
            ' document.getElementById("dpbring").textContent=br.length?"Tell them to bring: "+br.join("; ")+".":""; checks();}'
            'function save(now){var st=document.getElementById("decst"); st.textContent="Saving..."; clearTimeout(t);'
            ' t=setTimeout(function(){t=null; bvPost({action:"dec_save",rid:rid,data:bvB64(JSON.stringify(val()))},function(r){'
            '  st.textContent=r.success?"Saved":"Not saved yet: sign in again";});}, now?0:800);}'
            'window.addEventListener("pagehide",function(){ if(t){clearTimeout(t); t=null; bvBeacon({action:"dec_save",rid:rid,data:bvB64(JSON.stringify(val()))});}});'
            'card.addEventListener("input",function(){show();save();}); card.addEventListener("change",function(){show();save();});'
            'document.getElementById("dadd").onclick=function(){var tr=document.createElement("tr"); tr.setAttribute("data-src","");'
            ' tr.innerHTML=%s; document.getElementById("dlines").appendChild(tr); wire(tr); show();'
            ' if(tr.scrollIntoView) tr.scrollIntoView({behavior:"smooth",block:"center"}); setTimeout(function(){tr.querySelector(".d-cat").focus({preventScroll:true});},350);};'
            'function wire(tr){var d=tr.querySelector(".d-del"); if(d) d.onclick=function(){tr.parentNode.removeChild(tr); show(); save();};'
            ' var a=tr.querySelector(".d-amt"), h=tr.querySelector(".d-help"); a.addEventListener("input",function(){ if(a.value) h.checked=true; });}'
            'Array.prototype.forEach.call(card.querySelectorAll("#dlines tr"),wire);'
            'document.getElementById("dsend").onclick=function(){var b=this; b.disabled=true; document.getElementById("decst").textContent="Sending...";'
            ' bvPost({action:"dec_send",rid:rid,data:bvB64(JSON.stringify(val()))},function(r){ if(r.success) location.href="?view=mine";'
            '  else {b.disabled=false; document.getElementById("decst").textContent=r.message;}});};'
            'show();})();</script>') % (
        s.Rid, ('Saved ' + days_ago(d['saved'])) if d.get('saved') else 'Saved as you go',
        ' checked' if choice == 'help' else '', ' checked' if choice == 'decline' else '',
        ''.join(rows), pickup_hint(st),
        'Set where and when you agreed with them. The router confirms it when the payouts go out.' if st.get('req_pickup') else
        'If you agreed on where and when, set it here (optional). The router confirms it, or sets it, with the payouts.', places, esc(d.get('pickup_day') or ''), esc(d.get('pickup_time') or ''), reasons, refs, rich_box(d.get('reason') or '', '', 100),
        ('<div class="muted">%s run it by another committee member: <a href="#people">Case management</a>, Bring in a teammate, To affirm my decision.</div>' % (
            'Before you send it,' if st.get('affirm_required') else 'If you like,')) if st.get('affirm_on') and not (r.get('affirm') or {}).get('pid') else '',
        kinds_js, safe_json(decision_rules(st, r)),
        safe_json(re.sub(r'^<tr[^>]*>|</tr>$', '', template)))


ROUTER_DECIDE = [('none', 'Nothing: always give it to a committee member'), ('decline', 'Decline only'),
                 ('both', 'Decline or help, the same decision a committee member makes')]


def router_rules(st):
    """A router's decision follows the church's Why rules; talking with them
    and a teammate's affirmation are for committee decisions (the router has
    not made the call, and the router is the church's check)."""
    return dict(st, req_talked_help=False, affirm_on=False, affirm_required=False, affirm_over=0)


def router_choice_card(st, s, r, f, rid, records):
    """A new request, for the router: one card, the same shape a committee
    member's decision has. Decline, Assign (give it to a committee member) or
    Help, as far as Settings, Decisions allows."""
    mode = st.get('router_decide', 'decline')
    assign = assign_card(st, rid, records, draft=r.get('assign_note') or '')
    if mode not in ('decline', 'both'):
        return assign
    assign = (assign.replace('<div class="bvc bvblue" id="assign"><h2>Give it to a committee member</h2>', '<div id="assign">', 1)
              .replace('<div class="bvc bvblue"><h2>Give it to a committee member</h2>', '<div id="assign">', 1))
    rr = dict(r)
    if mode == 'decline':
        rr['decision'] = dict(r.get('decision') or {}, choice='decline')
    dec = decision_card(router_rules(st), s, rr, f, True)
    dec = (dec.replace('action:"dec_save"', 'action:"rdec_save"').replace('action:"dec_send"', 'action:"rdec_send"')
           .replace('location.href="?view=mine"', 'location.href="?view=home"')
           .replace('Send to the router', 'Send the decision')
           .replace('<div class="bvc bvblue" id="dec"', '<div id="dec"', 1)
           .replace('<h2>Step 2: Your decision</h2>', '', 1))
    # The card's own Help / Decline radios are driven by the choice above.
    dec = re.sub(r'<div class="bvr" style="justify-content:flex-start"><label class="pick"><input type="radio" name="dchoice" value="help"([^>]*)> Help</label>'
                 r'<label class="pick"><input type="radio" name="dchoice" value="decline"([^>]*)> Decline</label></div>',
                 r'<input type="radio" name="dchoice" value="help"\1 style="display:none"><input type="radio" name="dchoice" value="decline"\2 style="display:none">', dec, 1)
    saved = (r.get('decision') or {}).get('choice') if (r.get('decision') or {}).get('saved') else ''
    pick = saved if saved in ('help', 'decline') and (mode == 'both' or saved == 'decline') else 'assign'
    opts = [('decline', 'Decline'), ('assign', 'Assign')] + ([('help', 'Help')] if mode == 'both' else [])
    radios = ''.join('<label class="pick"><input type="radio" name="rchoice" value="%s"%s> %s</label>' % (v, ' checked' if v == pick else '', t) for v, t in opts)
    return ('<div class="bvc bvblue" id="rchoose"><h2>Your decision</h2>'
            '<div class="bvr" style="justify-content:flex-start">%s</div>'
            '<div class="muted" id="rchelp"></div>%s%s</div>'
            '<script>(function(){var c=document.getElementById("rchoose"), dec=document.getElementById("dec"), as=document.getElementById("assign"),'
            ' TIPS={decline:"Decline it now: pick a reason, somewhere to send them, and write the Why. It closes as declined.",'
            ' assign:"Give it to a committee member to call them and decide it. The next in turn is picked; choose anyone.",'
            ' help:"Help them yourself, now: tick what we will give and how, and write the Why. No committee member is needed."};'
            'function go(first){var v=(c.querySelector("input[name=rchoice]:checked")||{}).value||"assign";'
            ' as.style.display=v=="assign"?"":"none"; if(dec) dec.style.display=v=="assign"?"none":"";'
            ' document.getElementById("rchelp").textContent=TIPS[v]||"";'
            ' if(dec&&v!="assign"){var i=dec.querySelector("input[name=dchoice][value="+v+"]"); if(i&&!i.checked){i.checked=true;'
            '  if(!first) i.dispatchEvent(new Event("change",{bubbles:true}));}}}'
            'Array.prototype.forEach.call(c.querySelectorAll("input[name=rchoice]"),function(i){i.addEventListener("change",function(e){e.stopPropagation(); go(false);});});'
            'go(true);})();</script>') % (radios, assign, dec)


def new_for_router(st, rid):
    """(submission, record, request) for a router decision on a new request."""
    if st.get('router_decide', 'decline') not in ('decline', 'both'):
        raise Exception('Routers give requests to a committee member here; they do not decide them (Settings, Decisions).')
    s = find_submission(st, rid)
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].setdefault(s.Rid, {})
    if r.get('status') not in (None, 'new'):
        raise Exception('This request was already given out or decided. Reload the page.')
    return s, rec, r


def act_rdec(st, role, send):
    if role != 'router':
        raise Exception('Only routers can do that.')
    s, rec, r = new_for_router(st, _param('rid'))
    d = fill_needs(st, s, clean_decision(st, _param('data')), r.get('decision'))
    if st.get('router_decide', 'decline') == 'decline' and d['choice'] != 'decline':
        raise Exception('Routers can only decline here; to help, give it to a committee member (Settings, Decisions).')
    if not send:
        d['saved'] = now_str()
        r['decision'] = d
        save_record(s.Pid, rec)
        return {'success': True}
    miss = decision_missing(router_rules(st), r, d)
    if miss:
        raise Exception(miss[0] if len(miss) == 1 else 'Before sending: ' + ' '.join(miss))
    if not r.get('num'):
        r['num'] = next_number(all_records())
    r['came'] = s.T
    d.update({'saved': now_str(), 'sent': now_str(), 'by': int(ME), 'router': True})
    r['decision'] = d
    if d['choice'] == 'help':
        r['status'] = 'approved'
        total = sum(to_number(x['amount']) or 0 for x in d['items'] if x['help'])
        add_history(r, 'The router approved %s without giving it out' % money(total), d['reason'], given_lines(st, d))
    else:
        r['status'] = 'declined'
        r['closed'] = now_str()
        add_history(r, 'The router declined it without giving it out: %s' % d['decline_reason'], d['reason'])
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- keeping it moving: nudge, take it over
def stuck_days(st):
    try:
        return max(1, min(60, int(st.get('stuck_days') or 7)))
    except Exception:
        return 7


def is_stuck(st, s, r):
    if (r or {}).get('status') not in ('committee', 'requester'):
        return 0
    try:
        days = (datetime.datetime.now() - datetime.datetime.strptime(last_touch(r, s)[:10], '%Y-%m-%d')).days
    except Exception:
        return 0
    return days if days >= stuck_days(st) else 0


def moving_card(st, s, r):
    """For a router on a request with the committee: nudge the case worker, or
    take it over and finish it (the case worker stays on it as a teammate)."""
    owner = int(r.get('owner') or 0)
    nm = names_for([owner]).get(owner, 'the case worker')
    days = is_stuck(st, s, r)
    head = ('<b style="color:#8A5A00">No change in %d days.</b> ' % days) if days else ''
    return ('<div class="bvc %s" id="moving" data-rid="%s"><h2>Keeping it moving</h2>'
            '<div>%sLast change %s. Check on %s, take it over to finish it yourself, or give it to someone else below.</div>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">'
            '<input type="text" id="nudgetext" size="40" placeholder="A note for %s (optional)">'
            '<button type="button" class="btn ghost" id="nudgego">Check on it</button>'
            '<button type="button" class="btn ghost" id="takego">Take it over</button><span id="movst" class="muted"></span></div>'
            '<div class="muted"><b>Check on it</b> writes to %s in the team chat and lets them know. <b>Take it over</b> makes you the case worker: '
            'the call, the documents and the decision become yours, and %s stays on it as a teammate, so nothing they learned is lost.</div></div>'
            '<script>(function(){var c=document.getElementById("moving"), st=document.getElementById("movst"), rid=c.getAttribute("data-rid");'
            'function go(d){st.textContent="Saving..."; bvPost(d,function(r){ if(r.success) location.reload(); else st.textContent=r.message;});}'
            'document.getElementById("nudgego").onclick=function(){go({action:"nudge",rid:rid,text:bvB64(document.getElementById("nudgetext").value)});};'
            'document.getElementById("takego").onclick=function(){go({action:"take_over",rid:rid});};})();</script>') % (
        'bvwarn' if days else '', s.Rid, head, days_ago(last_touch(r, s)), esc(nm), esc(nm), esc(nm), esc(nm))


def open_with_committee(st, rid):
    s = find_submission(st, rid)
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid)
    if not r or r.get('status') not in ('committee', 'requester'):
        raise Exception('This request is not open with the committee.')
    return s, rec, r


def act_nudge(st, role):
    s, rec, r = open_with_committee(st, _param('rid'))
    text = from_b64(_param('text')).strip()[:2000] or 'Checking in on this one. What is the next step, and can I help?'
    r.setdefault('messages', []).append({'t': now_str(), 'by': int(ME), 'text': text})
    add_history(r, 'The router checked on it')
    notify(r, 'nudge', [r.get('owner')], rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def act_take_over(st, role):
    s, rec, r = open_with_committee(st, _param('rid'))
    was = int(r.get('owner') or 0)
    if was == int(ME):
        raise Exception('It is already yours.')
    r['owner'] = int(ME)
    r.pop('handoff', None)
    if was:
        share_add(r, was, 'Was the case worker')
    r.setdefault('assigned', []).append({'pid': int(ME), 't': now_str(), 'by': int(ME), 'from': was})
    nm = names_for([int(ME), was])
    add_history(r, '%s took it over%s' % (nm.get(int(ME), 'The router'), (' (%s stays on it as a teammate)' % nm.get(was, '')) if was else ''))
    notify(r, 'shared', [was], 'took it over; you stay on it as a teammate', rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def pickup_hint(st):
    """Under the decision table: which kinds of help are picked up, and that
    choosing one opens a place to set where and when."""
    kinds = [k['name'] for k in st.get('help_kinds') or [] if k.get('pickup')]
    if not kinds:
        return ''
    names = ', '.join(kinds[:-1]) + ' and ' + kinds[-1] if len(kinds) > 1 else kinds[0]
    return ('<div class="muted" style="margin:4px 0">Each line says how it gets to them: <b>Benevolence pays it</b> (a check to the payee) '
            'or <b>They pick it up</b>. %s %s always picked up. When anything is picked up, set where and when below.</div>') % (
        esc(names), 'are' if len(kinds) > 1 else 'is')


def pickup_when(d):
    """'Fri Nov 14 at 2:00 PM' from the saved day and time."""
    try:
        dt = datetime.datetime.strptime('%s %s' % (d.get('pickup_day'), d.get('pickup_time') or '00:00'), '%Y-%m-%d %H:%M')
    except Exception:
        return ''
    t = dt.strftime('%I:%M %p').lstrip('0') if d.get('pickup_time') else ''
    return dt.strftime('%a %b ') + str(dt.day) + ((' at ' + t) if t else '')


def pickup_line(st, d):
    if not any(x.get('help') and item_deliv(st, d, x) == 'pickup' for x in d.get('items') or []):
        return ''
    when = pickup_when(d)
    return '<div><b>Pickup:</b> %s%s</div>' % (esc(d.get('pickup') or 'place not set'), (', ' + esc(when)) if when else '')


def decision_summary(st, r):
    """The decision as the router (and timeline) sees it."""
    d = r.get('decision') or {}
    who = names_for([d.get('by')]).get(int(d.get('by') or 0), '')
    if d.get('choice') == 'decline':
        return ('<div class="bvc bvwarn"><b>%s declined it%s</b><div>Why: %s%s</div>%s</div>' % (
            esc(who), (' on ' + nice_date(d.get('sent'))) if d.get('sent') else '', esc(d.get('decline_reason') or ''),
            (' · sent them to ' + esc(d['referral'])) if d.get('referral') else '',
            ('<div>%s</div>' % clean_html(d['reason'])) if d.get('reason') else ''))
    rows = ''.join('<tr><td><b>%s</b>%s</td><td data-l="Give">%s</td><td data-l="How">%s%s</td></tr>' % (
        esc(x.get('need') or ''), (' · ' + esc(x['payee'])) if x.get('payee') else '', money(to_number(x.get('amount'))),
        esc(x.get('how') or ''), ('<div class="muted">At pickup: %s</div>' % esc(stipulation(st, x.get('how')))) if stipulation(st, x.get('how')) else '')
        for x in d.get('items') or [] if x.get('help'))
    total = sum(to_number(x.get('amount')) or 0 for x in d.get('items') or [] if x.get('help'))
    aff = r.get('affirm') or {}
    return ('<div class="bvc"><div class="bvr"><h2>%s approved %s</h2><span class="muted">%s</span></div>'
            '<table><thead><tr><th>Help</th><th>Give</th><th>How</th></tr></thead><tbody>%s</tbody></table>'
            '%s%s%s</div>') % (
        esc(who), money(total), ('sent ' + nice_date(d.get('sent'))) if d.get('sent') else '', rows, pickup_line(st, d),
        ('<div>%s</div>' % clean_html(d['reason'])) if d.get('reason') else '',
        ('<div class="muted">Affirmed by %s (%s)%s</div>' % (esc(names_for([aff.get('pid')]).get(int(aff.get('pid') or 0), '')),
                                                           esc(aff.get('how') or ''), (': ' + esc(aff['note'])) if aff.get('note') else '')) if aff.get('pid') else '')


REPEAT_DAYS = 14   # the same person sending the form again this soon, while the first is still open


_COMBINED = []


def combined_map():
    """{combined request RID: the RID it was combined into}, once per run."""
    if _COMBINED:
        return _COMBINED[0]
    out = {}
    for row in q.QuerySql("""
            SELECT CAST(Data AS nvarchar(max)) Data FROM PeopleExtra WITH (NOLOCK)
            WHERE Field = '%s' AND Data LIKE '%%"into"%%'""" % EV_FIELD):
        for rid, req in (load_json(row.Data or '', {}).get('requests') or {}).items():
            if req.get('status') == 'combined' and req.get('into'):
                out[rid.upper()] = req['into'].upper()
    _COMBINED.append(out)
    return out


def days_apart(a, b):
    try:
        d = lambda t: datetime.datetime.strptime(t[:10], '%Y-%m-%d')
        return abs((d(a) - d(b)).days)
    except Exception:
        return 999


def repeat_pairs(st, entries):
    """{RID: (other submission, other record, 'newer'|'older')} for a person who
    sent the form twice within REPEAT_DAYS while both are still open (new, with
    the committee, or waiting on the requester), unless a router said they are
    separate requests or combined them."""
    cut = cutoff(st)
    by = {}
    for x, r in entries:
        stt = (r or {}).get('status') or 'new'
        if stt in ('new', 'committee', 'requester') and (r or x.T[:10] >= cut):
            by.setdefault(x.Pid, []).append((x, r))
    out = {}
    for lst in by.values():
        lst.sort(key=lambda e: e[0].T)
        for (a, ra), (b, rb) in zip(lst, lst[1:]):
            apart = [x.upper() for x in ((ra or {}).get('not_dup') or []) + ((rb or {}).get('not_dup') or [])]
            if a.Rid.upper() in apart or b.Rid.upper() in apart or days_apart(a.T, b.T) > REPEAT_DAYS:
                continue
            out[a.Rid.upper()] = (b, rb, 'newer')
            out[b.Rid.upper()] = (a, ra, 'older')
    return out


def merge_needs(a, b):
    parts = [x for x in (a or '').split(', ') if x]
    parts += [x for x in (b or '').split(', ') if x and x not in parts]
    return ', '.join(parts)


def same_answer(a, b):
    """The same answer, ignoring case, spacing and how a number was typed ($334 and 334)."""
    a, b = (a or '').strip().lower(), (b or '').strip().lower()
    if a == b:
        return True
    na, nb = to_number(a), to_number(b)
    return na is not None and nb is not None and abs(na - nb) < 0.005 and bool(re.match(r'^[$\s\d,.]+$', a)) and bool(re.match(r'^[$\s\d,.]+$', b))


def answers_compare(st, role, x, y, cid):
    """Every answer on two forms side by side (x the earlier, y the later),
    changed ones highlighted, with a switch to show only what changed."""
    fx, fy = form_answers(st, x.Rid), form_answers(st, y.Rid)
    priv = role == 'router'
    cell = lambda v, p: '<i>private</i>' if p and not priv else (esc(v) if v else '<span class="muted">blank</span>')
    needs = lambda f: '<br>'.join('%s %s' % (esc(n['need']), money(n['amount'])) for n in f['needs']) or 'None listed'
    else_ = lambda f: esc(f['elsewhere'] + (': ' + f['elsewhere_who'] if f['elsewhere_who'] else ''))
    head = '<thead><tr><th></th><th>Sent %s</th><th>Sent %s</th></tr></thead>' % (nice_date(x.T), nice_date(y.T))
    cols = '<colgroup><col style="width:30%"><col style="width:35%"><col style="width:35%"></colgroup>'
    changed = 0
    top = []
    for k, a, b in [('Needs', needs(fx), needs(fy)), ('Total requested', money(fx['total']) or 'Not given', money(fy['total']) or 'Not given'),
                    ('People in the home', esc(fx['people']), esc(fy['people'])), ('Help asked for elsewhere', else_(fx), else_(fy))]:
        d = a != b
        changed += 1 if d else 0
        top.append('<tr class="%s"><td><b>%s</b></td><td data-l="Earlier">%s</td><td data-l="Later">%s</td></tr>' % ('hi' if d else 'same', k, a or '<span class="muted">blank</span>', b or '<span class="muted">blank</span>'))
    rx, ry = dict(compare_rows(fx)), dict(compare_rows(fy))
    order = [n for n, _ in compare_rows(fx)] + [n for n, _ in compare_rows(fy) if n not in rx]
    secs = []
    for name in order:
        ix = dict((k, (lab, v, pv)) for k, lab, v, pv in rx.get(name, []))
        iy = dict((k, (lab, v, pv)) for k, lab, v, pv in ry.get(name, []))
        keys = [k for k, _, _, _ in rx.get(name, [])] + [k for k, _, _, _ in ry.get(name, []) if k not in ix]
        trs, nd = [], 0
        for k in keys:
            lab = ix.get(k) or iy.get(k)
            va, vb = (ix.get(k) or ('', '', False))[1], (iy.get(k) or ('', '', False))[1]
            d = not same_answer(va, vb)
            nd += 1 if d else 0
            trs.append('<tr class="%s"><td>%s</td><td data-l="Earlier">%s</td><td data-l="Later">%s</td></tr>' % ('hi' if d else 'same', esc(lab[0]), cell(va, lab[2]), cell(vb, lab[2])))
        changed += nd
        if trs:
            secs.append('<details open class="%s"><summary>%s%s</summary><table class="cmp">%s%s<tbody>%s</tbody></table></details>' % (
                'hasdiff' if nd else 'nodiff', esc(name), (' <span class="pill p-warn">%d changed</span>' % nd) if nd else '', cols, head, ''.join(trs)))
    return ('<div class="cmpbox" id="%s"><style>#%s.only tr.same,#%s.only details.nodiff{display:none} #%s table{background:#fff} #%s tr.hi td{background:#FCE3A8}</style>'
            '<div class="bvr"><b>%d answer%s differ%s</b><label class="pick"><input type="checkbox" onchange="document.getElementById(\'%s\').classList.toggle(\'only\',this.checked)"> '
            'Show only what changed</label></div><table class="cmp">%s%s<tbody>%s</tbody></table>%s</div>') % (
        cid, cid, cid, cid, cid, changed, '' if changed == 1 else 's', '' if changed != 1 else 's', cid, cols, head, ''.join(top), ''.join(secs))


def repeat_card(st, s, r, rp):
    """For a router: this person sent the form again. Same request (combine,
    keeping one) or separate requests (keep both)."""
    o, ro, which = rp
    me_new = ((r or {}).get('status') or 'new') == 'new'
    o_new = ((ro or {}).get('status') or 'new') == 'new'
    cache = answers_for(st, [s.Rid, o.Rid])
    mine, theirs = need_summary(st, s.Rid, cache)[0], need_summary(st, o.Rid, cache)[0]
    owners = names_for([int((r or {}).get('owner') or 0), int((ro or {}).get('owner') or 0)])

    def where(rr):
        stt = (rr or {}).get('status') or 'new'
        own = owners.get(int((rr or {}).get('owner') or 0))
        if stt == 'new' or not own:
            return STATUS_LABEL.get(stt, '')
        return 'With %s%s' % (own, ', waiting on the requester' if stt == 'requester' else '')
    rows = '<table><tbody>%s</tbody></table>' % ''.join(
        '<tr><td><b>%s</b></td><td data-l="Sent">%s</td><td data-l="Asking for">%s</td><td data-l="Where">%s</td></tr>' % (lab, nice_date(x.T), esc(n or 'Not given'), esc(w))
        for lab, x, n, w in sorted([('This one', s, mine, where(r)), ('The other one', o, theirs, where(ro))], key=lambda e: e[1].T))
    owner = int(((ro if not o_new else r) or {}).get('owner') or 0)
    who = owners.get(owner, 'the case worker')
    if not me_new and not o_new:
        a_who, b_who = owners.get(int((r or {}).get('owner') or 0), 'a member'), owners.get(int((ro or {}).get('owner') or 0), 'a member')
        same_person = a_who == b_who
        keep_pick = ('<div>%s <b>Keep</b> <label class="pick"><input type="radio" name="rpkeep" value="%s"%s> %s\'s case (sent %s)</label>'
                     '<label class="pick"><input type="radio" name="rpkeep" value="%s"%s> %s\'s case (sent %s)</label> '
                     '<span class="muted">The other closes as Combined: its form joins the case you keep, and its notes and team chat stay on the closed one. %s</span></div>') % (
            ('Both have been given out, to <b>%s</b>.' % esc(a_who)) if same_person else 'Both have been given out: one to <b>%s</b>, one to <b>%s</b>.' % (esc(a_who), esc(b_who)),
            s.Rid, ' checked' if which == 'newer' else '', esc(a_who), nice_date(s.T), o.Rid, ' checked' if which == 'older' else '', esc(b_who), nice_date(o.T),
            '' if same_person else 'Both are told.')
        same = 'Same request: combine the cases'
        keep_js = 'var k=document.querySelector("#repeat input[name=rpkeep]:checked").value;'
    elif me_new and o_new:
        keep_pick = ('<div><b>Keep</b> <label class="pick"><input type="radio" name="rpkeep" value="%s"%s> the one sent %s</label>'
                     '<label class="pick"><input type="radio" name="rpkeep" value="%s"%s> the one sent %s</label> '
                     '<span class="muted">(the newer has their latest answers)</span></div>') % (
            s.Rid, ' checked' if which == 'older' else '', nice_date(s.T), o.Rid, ' checked' if which == 'newer' else '', nice_date(o.T))
        same = 'Same request: combine them'
        keep_js = 'var k=document.querySelector("#repeat input[name=rpkeep]:checked").value;'
    else:
        keep, drop = (o.Rid, s.Rid) if me_new else (s.Rid, o.Rid)
        keep_pick = '<div>The other one is already with <b>%s</b>. Combining adds this form to that case and lets them know.</div>' % esc(who) if me_new else \
            '<div>The newer one has not been given out. Combining adds it to this case and lets %s know.</div>' % esc(who)
        same = 'Same request: add it to %s\'s case' % who
        keep_js = 'var k="%s";' % keep
    return ('<div class="bvc bvwarn" id="repeat" data-a="%s" data-b="%s"><h2>They sent the form twice, %d days apart</h2>'
            '<div class="muted">Often they did not hear back, or remembered something. If it is the same need, combine them: one case, '
            'and the other closes as Combined (not declined, not counted twice). If it is a different need, keep both.</div>%s'
            '<details open><summary><b>Every answer, side by side</b></summary>%s</details>%s'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><button type="button" class="btn" id="rpsame">%s</button>'
            '<button type="button" class="btn ghost" id="rpsep">Separate requests: keep both</button><span id="rpst" class="muted"></span></div></div>'
            '<script>(function(){var c=document.getElementById("repeat"), a=c.getAttribute("data-a"), b=c.getAttribute("data-b"), st=document.getElementById("rpst");'
            'document.getElementById("rpsame").onclick=function(){%s var d=(k==a)?b:a; st.textContent="Saving...";'
            ' bvPost({action:"combine",keep:k,drop:d},function(r){ if(r.success) location.href="?view=req&rid="+k; else st.textContent=r.message;});};'
            'document.getElementById("rpsep").onclick=function(){st.textContent="Saving...";'
            ' bvPost({action:"separate",a:a,b:b},function(r){ if(r.success) location.reload(); else st.textContent=r.message;});};})();</script>') % (
        s.Rid, o.Rid, days_apart(s.T, o.T), rows, answers_compare(st, 'router', *(sorted([s, o], key=lambda e: e.T) + ['rpcmp'])), keep_pick, esc(same), keep_js)


def resent_note(s, rp):
    """For the committee member on a case: the same person sent the form
    again (or had sent it before), and a router is deciding."""
    o, ro, which = rp
    stt = (ro or {}).get('status') or 'new'
    own = int((ro or {}).get('owner') or 0)
    whom = names_for([own]).get(own) if own else ''
    where = ('a router has not given it out yet' if stt == 'new' else
             'you have that one too' if own == int(ME or 0) else
             ('it is with %s' % whom) if whom else 'it is with the committee')
    return ('<div class="bvc bvwarn" id="resent"><h2>%s</h2><div>%s, %d days apart (%s). A router is deciding whether it is the same request. '
            'If it is, the two become one case and you will be told which one to work.</div></div>') % (
        'They sent the form again' if which == 'newer' else 'They had already sent the form',
        ('They sent it again on %s' if which == 'newer' else 'They also sent one on %s') % nice_date(o.T), days_apart(s.T, o.T), esc(where))


def combined_card(st, s, r, role):
    """On a request that was combined into another."""
    into = r.get('into') or ''
    k = find_submission(st, into)
    kr = (load_record(k.Pid)['requests'].get(k.Rid) or {}) if k else {}
    return ('<div class="bvc bvblue" id="combined"><h2>Sent again, combined</h2><div>This is the same request as '
            '<a href="?view=req&amp;rid=%s">%s</a>, so it was combined into that one. Its answers and documents are there.</div>%s</div>') % (
        esc(into), ('#%s' % kr['num']) if kr.get('num') else 'the other request',
        ('<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn ghost" id="uncomb" data-rid="%s">Undo: make it its own request again</button>'
         '<span id="uncst" class="muted"></span></div><script>document.getElementById("uncomb").onclick=function(){var st=document.getElementById("uncst");'
         'st.textContent="Saving..."; bvPost({action:"uncombine",rid:this.getAttribute("data-rid")},function(r){ if(r.success) location.reload(); else st.textContent=r.message;});};</script>'
         % s.Rid) if role == 'router' else '')


def also_sent_card(st, s, r, role):
    """On the request kept: the forms combined into it."""
    rows = []
    for rid in r.get('combined') or []:
        o = find_submission(st, rid)
        if o:
            rows.append('<details%s><summary><b>Also sent the form %s</b>, asking for %s</summary>%s</details>' % (
                ' open' if len(r.get('combined') or []) == 1 else '', nice_date(o.T), esc(need_summary(st, o.Rid, answers_for(st, [o.Rid]))[0] or 'nothing listed'),
                answers_compare(st, role, *(sorted([s, o], key=lambda e: e.T) + ['ascmp' + o.Rid[:8]]))))
    return ('<div class="bvc"><h2>They sent the form again</h2>%s<div class="muted">Combined into this request by the router. '
            'Check what changed between the two.</div></div>' % ''.join(rows)) if rows else ''


def act_combine(st, role):
    keep, drop = find_submission(st, _param('keep')), find_submission(st, _param('drop'))
    if not keep or not drop or keep.Rid.upper() == drop.Rid.upper():
        raise Exception('Request not found.')
    if int(keep.Pid) != int(drop.Pid):
        raise Exception('Only requests from the same person can be combined.')
    rec = load_record(keep.Pid)
    kr = rec['requests'].setdefault(keep.Rid, {})
    dr = rec['requests'].setdefault(drop.Rid, {})
    if (dr.get('status') or 'new') not in ('new', 'committee', 'requester'):
        raise Exception('That request is already decided, so it cannot be combined. Reload the page.')
    if (kr.get('status') or 'new') not in ('new', 'committee', 'requester'):
        raise Exception('That request is already decided. Reload the page.')
    dr.update({'was': dr.get('status') or 'new', 'status': 'combined', 'into': keep.Rid, 'closed': now_str(), 'came': drop.T})
    add_history(dr, 'Sent again: combined into the request sent %s' % nice_date(keep.T))
    kr['combined'] = [x for x in kr.get('combined') or [] if x.upper() != drop.Rid.upper()] + [drop.Rid]
    add_history(kr, 'They sent the form again on %s; the router combined it into this one' % nice_date(drop.T))
    if kr.get('owner'):
        notify(kr, 'resent', [kr['owner']], rid=keep.Rid)
    if dr.get('owner') and int(dr['owner']) != int(kr.get('owner') or 0):
        own = names_for([int(kr.get('owner') or 0)]).get(int(kr.get('owner') or 0), 'another member')
        notify(dr, 'combined_away', [dr['owner']], what='#%s with %s' % (kr.get('num') or '?', own), rid=drop.Rid)
    save_record(keep.Pid, rec)
    _COMBINED[:] = []
    return {'success': True}


def act_separate(st, role):
    a, b = find_submission(st, _param('a')), find_submission(st, _param('b'))
    if not a or not b or int(a.Pid) != int(b.Pid):
        raise Exception('Request not found.')
    rec = load_record(a.Pid)
    for x, y in ((a, b), (b, a)):
        rr = rec['requests'].setdefault(x.Rid, {})
        rr['not_dup'] = list(set((rr.get('not_dup') or []) + [y.Rid.upper()]))
    add_history(rec['requests'][max(a, b, key=lambda e: e.T).Rid], 'The router marked it a separate request from the one sent %s' % nice_date(min(a.T, b.T)))
    save_record(a.Pid, rec)
    return {'success': True}


def act_uncombine(st, role):
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid) or {}
    if r.get('status') != 'combined':
        raise Exception('It is not combined.')
    into = r.pop('into', '')
    r['status'] = r.pop('was', 'new')
    r.pop('closed', None)
    r['not_dup'] = list(set((r.get('not_dup') or []) + [into.upper()]))
    add_history(r, 'The router made it its own request again')
    kr = rec['requests'].get(into) or {}
    kr['combined'] = [x for x in kr.get('combined') or [] if x.upper() != s.Rid.upper()]
    kr['not_dup'] = list(set((kr.get('not_dup') or []) + [s.Rid.upper()]))
    save_record(s.Pid, rec)
    _COMBINED[:] = []
    return {'success': True}


def earlier_help_card(st, s, r, role, fam_reqs, records):
    """Earlier help to this family, for whoever is working this case: each
    earlier request's outcome, what was given or why it was declined, the
    Why, and the call. Read only; the earlier team chat stays with its case,
    and this opens nothing else (Settings, Decisions can turn it off)."""
    me = int(ME or 0)
    if not st.get('earlier_on', True):
        return ''
    if not (role == 'router' or me == int(r.get('owner') or 0) or me in shared_pids(r)):
        return ''
    earlier = [fr for fr in fam_reqs if fr.Rid.upper() != s.Rid.upper() and fr.T < s.T]
    if not earlier:
        return ''
    names = names_for(list(set(fr.Pid for fr in earlier)))
    blocks = []
    for fr in reversed(earlier):
        er = (records.get(fr.Pid) or {}).get(fr.Rid) or (records.get(fr.Pid) or {}).get(fr.Rid.upper())
        kind, line = outcome(er)
        owner = int((er or {}).get('owner') or 0)
        head = '%s%s · %s%s · <span class="%s">%s</span>' % (
            ('#%s · ' % er['num']) if er and er.get('num') else '', nice_date(fr.T), esc(names.get(int(fr.Pid), '')),
            (' · with %s' % esc(names_for([owner]).get(owner, ''))) if owner else '', 'pill p-bad' if kind == 'declined' else '', esc(line))
        if not er or er.get('older'):
            body = '<div class="muted">From before this tool: only the outcome is known.</div>'
        else:
            body = ((decision_summary(st, er) if (er.get('decision') or {}).get('sent') else '<div class="muted">No decision was sent.</div>')
                    + (call_summary(st, er) or '<div class="muted">No call notes.</div>'))
        blocks.append('<details class="post"><summary><b>%s</b></summary>%s</details>' % (head, body))
    return ('<div class="bvc" id="earlier"><div class="bvr"><h2>Earlier help to this family</h2><span class="muted">%d earlier request%s</span></div>'
            '<div class="muted">What was decided before, and why, with the call notes. Read only; each earlier case\'s team chat stays with it.</div>%s</div>') % (
        len(earlier), '' if len(earlier) == 1 else 's', ''.join(blocks))


def owned_request(st, rid, need_owner=True):
    """(submission, record, request) for an action, checking it is this
    member's (or shared with them when need_owner is False)."""
    s = find_submission(st, rid)
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(rid)
    if not r or r.get('status') not in ('committee', 'requester'):
        raise Exception('This request is not open with the committee.')
    mine = int(r.get('owner') or 0) == int(ME)
    if need_owner and not mine:
        raise Exception('This request was not given to you.')
    if not need_owner and not mine and int(ME) not in [int(x) for x in r.get('shared') or []]:
        raise Exception('This request was not shared with you.')
    return s, rec, r


def clean_decision(st, raw):
    import json
    try:
        d = json.loads(from_b64(raw) or '{}')
    except Exception:
        raise Exception('The decision could not be read. Reload the page.')
    choice = d.get('choice') if d.get('choice') in ('help', 'decline') else 'help'
    items = []
    old = d.get('delivery') if d.get('delivery') in ('mail', 'pickup') else ''
    for x in (d.get('items') or [])[:30]:
        src = x.get('src')
        added = src in (None, '')
        it = {'src': unicode(src) if not added else None, 'help': bool(x.get('help')),
              'amount': unicode(x.get('amount') or '')[:20], 'how': unicode(x.get('how') or '')[:60],
              'payee': unicode(x.get('payee') or '').strip()[:80], 'address': unicode(x.get('address') or '').strip()[:160],
              'account': unicode(x.get('account') or '').strip()[:40]}
        if added:
            if x.get('cat') is not None or x.get('detail') is not None:
                it['cat'], it['detail'] = unicode(x.get('cat') or '')[:60], unicode(x.get('detail') or '').strip()[:80]
                it['need'] = added_need(it['cat'], it['detail'])[:80]
            else:
                it['need'] = unicode(x.get('need') or '')[:80]
        else:
            it['need'] = None
        it['deliv'] = item_deliv(st, {'delivery': old}, dict(x, how=it['how']))
        items.append(it)
    day = unicode(d.get('pickup_day') or '')
    tm = unicode(d.get('pickup_time') or '')
    # The request is picked up when anything given is; otherwise the church pays it.
    delivery = 'pickup' if any(x['help'] and x['deliv'] == 'pickup' for x in items) else ('mail' if items or old else '')
    return {'choice': choice, 'items': items, 'delivery': delivery, 'pickup': unicode(d.get('pickup') or '')[:60],
            'pickup_day': day if re.match(r'^\d{4}-\d{2}-\d{2}$', day) else '',
            'pickup_time': tm if re.match(r'^\d{2}:\d{2}$', tm) else '',
            'decline_reason': unicode(d.get('decline_reason') or '')[:80], 'referral': unicode(d.get('referral') or '')[:80],
            'reason': clean_html(d.get('reason') or '')}


def fill_needs(st, s, d, prev=None):
    """Copy each need's name, payee, address and account onto the saved item,
    so the router's view and reports don't depend on re-reading the form. A
    payee, address or account typed on the decision wins over the form's.
    prev: the decision saved before. The account box is blank for anyone who
    is not a router (the form's account is routers only), so blank keeps
    what was there."""
    f = form_answers(st, s.Rid)
    before = dict((x.get('src'), x) for x in (prev or {}).get('items') or [] if x.get('src') is not None)
    for x in d['items']:
        if x['src'] is not None and x['src'].isdigit() and int(x['src']) < len(f['needs']):
            n = f['needs'][int(x['src'])]
            x['need'] = n['need']
            x['payee'] = x.get('payee') or n['payee']
            x['address'] = x.get('address') or n.get('address') or ''
            x['account'] = x.get('account') or (before.get(x['src']) or {}).get('account') or n.get('account') or ''
    return d


def decision_rules(st, r):
    """The checks that apply to this request, for the server and the page's
    live checklist. Help and Decline each have their own: picking Decline
    never trips on help lines, and the other way round."""
    return {'why_help': bool(st.get('req_why_help', True)), 'why_decline': bool(st.get('req_why_decline', True)),
            'talked_req': bool(st.get('req_talked_help', True)), 'talked': bool(call_started(r) or reached(r)),
            'aff_on': bool(st.get('affirm_on')), 'aff_req': bool(st.get('affirm_on') and st.get('affirm_required')),
            'aff_over': float(st.get('affirm_over') or 0) if st.get('affirm_on') else 0.0,
            'aff_done': bool((r.get('affirm') or {}).get('pid')),
            'pickups': [k['name'] for k in st.get('help_kinds') or [] if k.get('pickup')],
            'pickup_req': bool(st.get('req_pickup'))}


def picked_up(k, d, given):
    """Do they come to get it? Yes when the committee member chose pickup, or
    when anything ticked is a kind that is always picked up (a gas card)."""
    return d.get('delivery') == 'pickup' or any(x['how'] in k['pickups'] or x.get('deliv') == 'pickup' for x in given)


def decision_missing(st, r, d):
    """[what is still missing, in plain words] for the choice made."""
    k = decision_rules(st, r)
    why = bool(strip_tags(d.get('reason') or ''))
    out = []
    if d['choice'] == 'help':
        given = [x for x in d['items'] if x['help']]
        if not given:
            out.append('Tick at least one thing to help with, or pick Decline.')
        for x in given:
            if not to_number(x['amount']):
                out.append('Type an amount for %s.' % (x['need'] or 'each thing you will help with'))
            if not x['how']:
                out.append('Pick how for %s.' % (x['need'] or 'each thing'))
            if x['src'] is None and not x['need']:
                out.append('Say what the added help is.')
        # Pickup place, day and time: optional unless the church requires them
        # (Settings, Decisions); otherwise the router sets them with the payouts.
        if k['pickup_req'] and picked_up(k, d, given):
            if not d['pickup']:
                out.append('Pick where they will pick it up.')
            if not d['pickup_day'] or not d['pickup_time']:
                out.append('Add the day and time agreed for pickup.')
        if k['why_help'] and not why:
            out.append('Write the Why: what you learned and why we are helping.')
        if k['talked_req'] and not k['talked']:
            out.append('Talk with them first: save call notes, or log a try as talked with them.')
        total = sum(to_number(x['amount']) or 0 for x in given)
        if not k['aff_done'] and (k['aff_req'] or (k['aff_over'] and total > k['aff_over'])):
            out.append('Run it by another committee member first (Affirmation)%s.' % (
                '' if k['aff_req'] else ': help over %s needs one' % money(k['aff_over'])))
    else:
        if not d['decline_reason']:
            out.append('Pick why it is declined.')
        if k['why_decline'] and not why:
            out.append('Write the Why: what you learned and why we are declining.')
        if k['aff_req'] and not k['aff_done']:
            out.append('Run it by another committee member first (Affirmation).')
    return out


def act_dec_save(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    d = fill_needs(st, s, clean_decision(st, _param('data')), r.get('decision'))
    d['saved'] = now_str()
    r['decision'] = d
    save_record(s.Pid, rec)
    return {'success': True}


def act_dec_send(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    d = fill_needs(st, s, clean_decision(st, _param('data')), r.get('decision'))
    miss = decision_missing(st, r, d)
    if miss:
        raise Exception(miss[0] if len(miss) == 1 else 'Before sending: ' + ' '.join(miss))
    d.update({'saved': now_str(), 'sent': now_str(), 'by': int(ME)})
    r['decision'] = d
    notify(r, 'decision', main_routers(st), rid=s.Rid)
    if d['choice'] == 'help':
        r['status'] = 'approved'
        total = sum(to_number(x['amount']) or 0 for x in d['items'] if x['help'])
        add_history(r, 'Approved %s' % money(total), d['reason'], given_lines(st, d))
    else:
        r['status'] = 'declined'
        r['closed'] = now_str()
        add_history(r, 'Declined: %s' % d['decline_reason'], d['reason'])
    save_record(s.Pid, rec)
    return {'success': True}


def act_affirm_set(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    who = int(_param('who', '0') or 0)
    if who not in [m['pid'] for m in committee_members(st)] or who == int(ME):
        raise Exception('Pick another committee member.')
    # Shared: they read the case here. Not shared: only their name and how
    # it was talked over (a phone call) go on the case.
    share = _param('share', '1') != '0'
    how = 'They read it here' if share else (_param('how') if _param('how') in ('By phone', 'In person') else 'By phone')
    aff = r.get('affirm') or {}
    if int(aff.get('pid') or 0) != who:
        aff = {'pid': who}
    aff['how'] = how
    r['affirm'] = aff
    if share:
        share_add(r, who, 'Affirmation')
    add_history(r, 'Asked %s to affirm' % names_for([who]).get(who, ''), '' if share else esc(how))
    notify(r, 'shared', [who], 'asked you to affirm their decision', rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def act_affirm_note(st, role):
    s, rec, r = owned_request(st, _param('rid'), need_owner=False)
    if int(ME) not in [int(x) for x in r.get('shared') or []]:
        raise Exception('This request was not shared with you.')
    aff = r.get('affirm') or {'pid': int(ME)}
    aff.update({'pid': int(ME), 'note': _param('note')[:2000], 'at': now_str()})
    r['affirm'] = aff
    add_history(r, 'Affirmation note from %s' % names_for([ME]).get(int(ME), ''), esc(aff['note']))
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- messages and announcements
# Team posts live on each poster's own record (TEAM_FIELD, hidden like the
# requests), so nobody overwrites anyone else's. The same value keeps what each
# person has seen. Messages about one request live on that request.
_TEAM = {}


def my_team():
    """This person's own team value: {'posts': [...], 'seen': {key: time}}."""
    if 'me' not in _TEAM:
        d = load_json(model.ExtraValueText(int(ME), TEAM_FIELD) or '', {}) if ME else {}
        d.setdefault('posts', [])
        d.setdefault('seen', {})
        _TEAM['me'] = d
    return _TEAM['me']


def save_my_team(d):
    if not TEAM_PROTECTED:
        raise Exception('Team messages are not protected yet, so nothing was saved. An admin can fix it in Settings.')
    d['posts'] = d['posts'][-300:]
    model.AddExtraValueText(int(ME), TEAM_FIELD, safe_json(d))
    _TEAM['me'] = d


def team_posts():
    """Every team post, oldest first, each with 'by'."""
    if 'all' not in _TEAM:
        out = []
        for r in q.QuerySql("""
                SELECT PeopleId, Field, CAST(Data AS nvarchar(max)) Data FROM PeopleExtra WITH (NOLOCK)
                WHERE Field IN (%s) AND Data IS NOT NULL""" % ','.join("'%s'" % f for f in TEAM_NAMES)):
            for p in load_json(r.Data or '', {}).get('posts') or []:
                out.append(dict(p, by=int(r.PeopleId)))
        out.sort(key=lambda p: p.get('t') or '')
        _TEAM['all'] = out
    return _TEAM['all']


def mark_seen(key, latest):
    """Remember this person has seen everything up to now on key."""
    if not ME or not TEAM_PROTECTED or not latest:
        return
    d = my_team()
    if (d['seen'].get(key) or '') >= latest:
        return
    d['seen'][key] = now_str()
    try:
        save_my_team(d)
    except Exception:
        pass


HERE_EVERY = 30      # seconds between presence saves (a chat open checks in every 3 to 15)
HERE_FOR = 75        # someone counts as "here" in a chat this long after their last check-in


def note_here(key):
    """Remember this person has this chat open (and is in the tool), at most
    every HERE_EVERY seconds, so "is here" and the email pause can use it."""
    if not ME or not TEAM_PROTECTED:
        return
    d = my_team()
    here = d.setdefault('here', {})
    now = datetime.datetime.now()
    try:
        last = datetime.datetime.strptime((here.get(key) or '')[:19], TS)
        if (now - last).total_seconds() < HERE_EVERY:
            return
    except Exception:
        pass
    stamp = now.strftime(TS)
    here[key] = stamp
    here['*'] = stamp
    for k in [k for k, v in here.items() if k != '*' and v < (now - datetime.timedelta(days=2)).strftime(TS)]:
        here.pop(k, None)
    try:
        save_my_team(d)
    except Exception:
        pass


def here_now(key, pids):
    """[names] of these people (not me) who have this chat open right now."""
    cut = (datetime.datetime.now() - datetime.timedelta(seconds=HERE_FOR)).strftime(TS)
    others = [int(x) for x in pids if x and int(x) != int(ME or 0)]
    on = [x for x in others if (team_of(x).get('here') or {}).get(key, '') >= cut]
    names = names_for(on)
    return [names.get(x, '') for x in on if names.get(x)]


def chat_live(key, posts_times, pids):
    """(here names, live): live while someone else is here or a message came in
    the last 3 minutes, so the chat checks every few seconds."""
    names = here_now(key, pids)
    recent = (datetime.datetime.now() - datetime.timedelta(minutes=3)).strftime(TS)
    return names, bool(names) or any((t or '') >= recent for t in posts_times[-3:])


def req_unread(rid, r):
    """Messages on a request from others that this person has not seen."""
    seen = my_team()['seen'].get(rid.upper()) or ''
    return [m for m in (r or {}).get('messages') or [] if int(m.get('by') or 0) != int(ME or 0) and (m.get('t') or '') > seen and msg_visible(r, m, ME)]


def text_html(t):
    return esc(t).replace('\n', '<br>')


def announcements_box(can_post=False, days=14):
    """Announcements, where people land: above People I'm helping and at the
    top of Waiting on you. Routers post them from Waiting on you."""
    since = (datetime.datetime.now() - datetime.timedelta(days=days)).strftime('%Y-%m-%d')
    ann = [p for p in reversed(team_posts()) if p.get('kind') == 'announce' and (p.get('t') or '') >= since][:4]
    me = int(ME or 0)
    names = names_for([p['by'] for p in ann])
    h = ''.join('<div class="bvc bvblue" style="padding:10px 16px"><div class="bvr"><b>&#128227; Announcement from %s, %s</b>%s</div><div>%s</div></div>' % (
        esc(names.get(p['by'], '')), nice_date(p.get('t')),
        ('<a href="#" class="rmann muted" data-id="%s">Remove</a>' % esc(p.get('id'))) if p['by'] == me else '', text_html(p.get('text'))) for p in ann)
    if can_post:
        h += ('<details class="annnew"><summary>&#128227; New announcement</summary><div class="bvc" style="margin-top:6px">'
              '<textarea id="anntext" rows="2" style="width:100%" placeholder="For the whole team; shows here and above People I\'m helping for two weeks"></textarea>'
              '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn" id="anngo">Post announcement</button>'
              '<span id="annst" class="muted"></span></div></div></details>')
    if h:
        h += ('<script>(function(){var g=document.getElementById("anngo"); if(g) g.onclick=function(){var t=document.getElementById("anntext").value,'
              'st=document.getElementById("annst"); if(!t.trim()) return; st.textContent="Posting...";'
              'bvPost({action:"team_post",kind:"announce",text:bvB64(t)},function(r){if(r.success) location.reload(); else st.textContent=r.message;});};'
              'Array.prototype.forEach.call(document.querySelectorAll(".rmann"),function(a){a.onclick=function(e){e.preventDefault();'
              'bvPost({action:"team_remove",id:a.getAttribute("data-id")},function(r){if(r.success) location.reload();});};});})();</script>')
    return h

def team_people(st):
    """{PeopleId: name} for everyone on the team: routers and the committee."""
    out = dict((m['pid'], m['name']) for m in committee_members(st))
    rp = [int(r['pid']) for r in st.get('routers') or []]
    out.update(names_for(rp))
    return out


def chat_id(pids):
    return 'c-' + '-'.join(str(p) for p in sorted(set(int(x) for x in pids)))


def chat_people(cid):
    """The PeopleIds in a private chat id, or [] for a bad one."""
    if not re.match(r'^c(-\d+){2,30}$', cid or ''):
        return []
    return sorted(set(int(x) for x in cid.split('-')[1:]))


def chat_posts(cid):
    """A chat's messages, oldest first."""
    if cid == 'team':
        return [p for p in team_posts() if p.get('kind') == 'msg']
    return [p for p in team_posts() if p.get('kind') == 'dm' and p.get('chat') == cid]


def my_chats():
    """{chat id: [posts]} for the private chats this person is in."""
    me = int(ME or 0)
    out = {}
    for p in team_posts():
        if p.get('kind') == 'dm' and me in chat_people(p.get('chat')):
            out.setdefault(p['chat'], []).append(p)
    return out


def chat_unread(cid, posts):
    seen = my_team()['seen'].get('team' if cid == 'team' else cid) or ''
    me = int(ME or 0)
    return sum(1 for p in posts if p['by'] != me and (p.get('t') or '') > seen)


def team_unread():
    n = chat_unread('team', chat_posts('team'))
    for cid, posts in my_chats().items():
        n += chat_unread(cid, posts)
    return n


def chat_title(cid, names):
    if cid == 'team':
        return 'Everyone'
    me = int(ME or 0)
    return ', '.join(names.get(p, 'Someone') for p in chat_people(cid) if p != me) or 'Just you'


def thread_html(cid, posts, names):
    """The chat bubbles: mine on the right, everyone else's on the left."""
    me = int(ME or 0)
    seen = my_team()['seen'].get('team' if cid == 'team' else cid) or ''
    out = []
    for p in posts[-200:]:
        mine = p['by'] == me
        out.append('<div class="bub%s%s%s"><div class="bubh">%s<span>%s %s</span>%s</div><div>%s</div></div>' % (
            ' me' if mine else '', ' ann' if p.get('kind') == 'announce' else '',
            ' new' if not mine and (p.get('t') or '') > seen else '',
            ('<b>Announcement</b> ' if p.get('kind') == 'announce' else '') + ('' if mine else '<b>%s</b> ' % esc(names.get(p['by'], ''))),
            nice_date(p.get('t')), esc((p.get('t') or '')[11:16]),
            (' <a href="#" class="rmpost" data-id="%s">Remove</a>' % esc(p.get('id'))) if mine else '', text_html(p.get('text'))))
    return ''.join(out) or '<div class="muted" style="text-align:center">No messages yet. Say hello.</div>'


def page_team(st, role):
    cid = _param('c')
    if cid:
        return page_chat(st, role, cid)
    people = team_people(st)
    me = int(ME or 0)
    h = ['<h1>Messages</h1>']
    if not TEAM_PROTECTED:
        h.append('<div class="bvc bvbad"><b>Team messages are not protected yet, so nothing can be sent.</b>'
                 '<span>Ask an admin to open Benevolence once; it protects them by itself.</span></div>')
    rows = []
    chats = [('team', chat_posts('team'))] + sorted(my_chats().items(), key=lambda x: x[1][-1].get('t') or '', reverse=True)
    names = names_for(list(set([p['by'] for c, ps in chats for p in ps] + [x for c, ps in chats for x in chat_people(c)])))
    names.update(people)
    for c, posts in chats:
        last = posts[-1] if posts else {}
        n = chat_unread(c, posts)
        rows.append('<a class="chatrow%s" href="?view=team&amp;c=%s"><div class="bvr"><b>%s%s</b><span class="muted">%s</span></div>'
                    '<div class="muted">%s</div></a>' % (
                        ' unread' if n else '', esc(c), '&#128274; ' if c != 'team' else '', esc(chat_title(c, names)),
                        nice_date(last.get('t')) if last else '',
                        (('<b>%d new</b> · ' % n) if n else '') + (esc((names.get(last['by'], '') + ': ') if last.get('by') != me else 'You: ') +
                                                                  esc((last.get('text') or '')[:80]) if last else
                                                                  ('The whole team: routers and the committee.' if c == 'team' else ''))))
    pick = ''.join('<label class="pick"><input type="checkbox" class="dmto" value="%d"> %s</label>' % (p, esc(n))
                   for p, n in sorted(people.items(), key=lambda x: x[1]) if p != me)
    h.append('<div class="bvc"><details><summary><b>+ New private chat</b> <span class="muted">with one or more teammates</span></summary>'
             '<div class="pickgrid">%s</div><button type="button" class="btn" id="dmgo">Open the chat</button></details></div>' % (
                 pick or '<div class="muted">No one else is on the team yet.</div>'))
    h.append('<div class="chatlist">%s</div>' % ''.join(rows))
    h.append('<div class="muted">&#128274; Private chats show only to the people in them. Email will only say you have a secure message, '
             'with a link back here; never the message itself.</div>')
    h.append('<script>document.getElementById("dmgo").onclick=function(){var ids=Array.prototype.filter.call(document.querySelectorAll(".dmto"),'
             'function(x){return x.checked;}).map(function(x){return x.value;}); if(!ids.length) return; ids.push("%d");'
             'ids.sort(function(a,b){return a-b;}); location.href="?view=team&c=c-"+ids.filter(function(v,i){return ids.indexOf(v)==i;}).join("-");};</script>' % me)
    return ''.join(h)


def page_chat(st, role, cid):
    people = team_people(st)
    me = int(ME or 0)
    if cid != 'team':
        who = chat_people(cid)
        if me not in who or any(p not in people for p in who):
            return '<div class="bvc"><b>This chat is not available.</b><span>Private chats open only for the people in them.</span></div>'
    posts = chat_posts(cid)
    names = names_for(list(set(p['by'] for p in posts)))
    names.update(people)
    title = chat_title(cid, names)
    body = ('<div class="bvr"><div><a href="?view=team">&larr; Messages</a><h1 style="margin:4px 0 0">%s%s</h1>'
            '<div class="muted">%s</div></div></div>'
            '<div class="thread" id="thread">%s</div>'
            '<div class="bvc chatbox"><textarea id="ctext" rows="2" placeholder="Write a message"></textarea>'
            '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn" data-k="%s">Send</button>%s'
            '<span id="cst" class="muted"></span></div></div>') % (
        '&#128274; ' if cid != 'team' else '', esc(title),
        ('Private: only the people in this chat see it.' if cid != 'team' else 'Everyone on the team: the routers and the committee.')
        + ' <span id="chere" style="color:#1E6B45;font-weight:700"></span>',
        thread_html(cid, posts, names), 'dm' if cid != 'team' else 'msg', '')
    body += ('<script>(function(){var cid=%s, st=document.getElementById("cst"), th=document.getElementById("thread"), last=%s;'
             'function bottom(){window.scrollTo(0, document.body.scrollHeight);}'
             'function wire(){Array.prototype.forEach.call(th.querySelectorAll(".rmpost"),function(a){a.onclick=function(e){e.preventDefault();'
             ' bvPost({action:"team_remove",id:a.getAttribute("data-id")},function(r){if(r.success) poll(true);});};});}'
             'var live=false, tm=null;'
             'function here(r){var h=document.getElementById("chere"); if(h) h.textContent=(r.here&&r.here.length)?" \u25CF "+r.here.join(", ")+(r.here.length==1?" is":" are")+" here":"";}'
             'function next(){clearTimeout(tm); tm=setTimeout(function(){ if(!document.hidden) poll(false); else next(); }, live?3000:15000);}'
             'function poll(force){bvPost({action:"chat_poll",c:cid,last:force?"":last},function(r){ if(r.success){live=!!r.live; here(r);'
             ' if(r.html!=null){th.innerHTML=r.html; last=r.last; wire(); bottom();}} next();});}'
             'document.addEventListener("visibilitychange",function(){ if(!document.hidden) poll(false); });'
             'Array.prototype.forEach.call(document.querySelectorAll(".chatbox button[data-k]"),function(b){b.onclick=function(){'
             ' var t=document.getElementById("ctext").value; if(!t.trim()) return; st.textContent="Sending...";'
             ' bvPost({action:"team_post",kind:b.getAttribute("data-k"),c:cid,text:bvB64(t)},function(r){'
             '  if(r.success){document.getElementById("ctext").value=""; st.textContent=""; poll(true);} else st.textContent=r.message;});};});'
             'document.getElementById("ctext").addEventListener("keydown",function(e){if(e.key=="Enter"&&(e.ctrlKey||e.metaKey)){'
             ' document.querySelector(".chatbox button[data-k]").click();}});'
             'wire(); bottom(); poll(false);})();</script>') % (
        safe_json(cid), safe_json(posts[-1].get('t') if posts else ''))
    mark_seen('team' if cid == 'team' else cid, posts[-1].get('t') if posts else '')
    return body


def act_chat_poll(st, role):
    """New messages for an open chat (every 15 seconds while it is open)."""
    cid = _param('c')
    people = team_people(st)
    if cid != 'team' and (int(ME or 0) not in chat_people(cid) or any(p not in people for p in chat_people(cid))):
        raise Exception('This chat is not available.')
    posts = chat_posts(cid)
    last = posts[-1].get('t') if posts else ''

    key = 'team' if cid == 'team' else cid
    note_here(key)
    who, live = chat_live(key, [p.get('t') for p in posts], [] if cid == 'team' else chat_people(cid))
    if _param('last') and _param('last') == last:
        return {'success': True, 'html': None, 'here': who, 'live': live}
    names = names_for(list(set(p['by'] for p in posts)))
    names.update(people)
    html = thread_html(cid, posts, names)
    mark_seen(key, last)
    return {'success': True, 'html': html, 'last': last, 'here': who, 'live': live}


def act_team_post(st, role):
    kind = _param('kind')
    if kind == 'announce' and role != 'router':
        raise Exception('Only routers post announcements.')
    text = from_b64(_param('text')).strip()[:2000]
    if not text:
        raise Exception('Write something first.')
    d = my_team()
    t = now_str()
    post = {'id': '%s-%s' % (t, ME), 't': t, 'kind': kind if kind in ('announce', 'dm') else 'msg', 'text': text}
    if post['kind'] == 'dm':
        cid = _param('c')
        who = chat_people(cid)
        people = team_people(st)
        if int(ME) not in who or len(who) < 2 or any(p not in people for p in who):
            raise Exception('Pick teammates for this chat.')
        post['chat'] = cid
        d['seen'][cid] = t
        sent = secure_now(st, [x for x in who if x != int(ME)], cid, [x for x in d['posts'] if x.get('chat') == cid], 'view=team&c=%s' % cid, None)
        if sent:
            post['mailed'] = sent
    else:
        d['seen']['team'] = t
    d['posts'].append(post)
    save_my_team(d)
    return {'success': True}


def act_team_remove(st, role):
    d = my_team()
    pid = _param('id')
    d['posts'] = [p for p in d['posts'] if p.get('id') != pid]
    save_my_team(d)
    return {'success': True}


def can_message(role, r):
    """Router, the member it was given to, or a teammate asked to affirm."""
    if role == 'router':
        return True
    me = int(ME or 0)
    return bool(me) and (int(r.get('owner') or 0) == me or me in [int(x) for x in r.get('shared') or []])


MSG_TO = {'all': '', 'router': 'to the router only', 'case': 'kept from the router'}


def router_ids():
    try:
        st = SETTINGS
    except NameError:
        st = load_settings()
    return [int(x['pid']) for x in (st or {}).get('routers') or []]


def msg_visible(r, m, pid):
    """A case message is for everyone on it, the router only, or the people
    on the case without the router. The writer always sees their own."""
    pid = int(pid or 0)
    to = m.get('to') or 'all'
    if to == 'all' or int(m.get('by') or 0) == pid or int(m.get('re') or 0) == pid:
        return True
    if to == 'router':
        return pid in router_ids()
    return pid == int(r.get('owner') or 0) or pid in shared_pids(r)


def chat_thread(s, r):
    """(thread html, newest time, how many are new) of the team's messages on
    this case that this person may see, oldest first. Marks nothing seen."""
    me = int(ME or 0)
    ms = [m for m in r.get('messages') or [] if msg_visible(r, m, me)]
    names = names_for(list(set([int(m.get('by') or 0) for m in ms] + [int(m.get('re') or 0) for m in ms])))
    seen = my_team()['seen'].get(s.Rid.upper()) or ''

    def to_line(m):
        to = m.get('to') or 'all'
        if to == 'all':
            return ''
        if m.get('re'):
            return 'private reply to %s' % esc(names.get(int(m['re']), ''))
        return MSG_TO[to]
    out, nnew = [], 0
    for m in ms:
        by = int(m.get('by') or 0)
        mine = by == me
        new = not mine and (m.get('t') or '') > seen
        nnew += 1 if new else 0
        out.append('<div class="bub%s%s"><div class="bubh">%s<span>%s %s</span></div><div>%s</div>'
                   '<div class="bvr" style="justify-content:space-between;gap:8px"><span class="to">%s</span>%s</div></div>' % (
                       ' me' if mine else '', ' new' if new else '',
                       '' if mine else '<b>%s</b>' % esc(names.get(by, '')), nice_date(m.get('t')), esc((m.get('t') or '')[11:16]),
                       text_html(m.get('text')), to_line(m),
                       '' if mine else '<button type="button" class="lnk rep" data-by="%d" data-to="%s" data-name="%s">Reply</button>' % (
                           by, esc(m.get('to') or 'all'), esc(names.get(by, '')))))
    html = ''.join(out) or '<div class="muted" style="text-align:center;padding:20px 0">No messages yet. Write the first one below.</div>'
    return html, (ms[-1].get('t') if ms else ''), nnew


def chat_widget(s, r):
    """Team chat: one big button that stays put at the bottom right and
    turns orange with the count when something new comes in (one gentle
    pulse, never a popup). Pressing it opens the whole thread large, with the
    reply box, and Close. It checks for new messages every 12 seconds and
    never reloads the page, so nothing typed is lost."""
    me = int(ME or 0)
    owner = int(r.get('owner') or 0)
    html, last, nnew = chat_thread(s, r)
    others = [p for p in [owner] + shared_pids(r) if p and p != me]
    if (me in router_ids() and me != owner) or not others:
        to = '<span class="muted">%s</span>' % ('Goes to everyone on the case.' if others else 'Goes to the router.')
    else:
        # Name the people, not the roles: "Only Pat Smith (router)".
        rp = [p for p in router_ids() if p != me]
        cp = [p for p in [owner] + shared_pids(r) if p and p != me and p not in rp]
        nm = names_for(rp + cp)

        def lst(ps):
            ns = [esc(nm.get(p, '')) for p in ps if nm.get(p)]
            return (', '.join(ns[:-1]) + ' and ' + ns[-1]) if len(ns) > 1 else (ns[0] if ns else '')
        # The case worker (and anyone they brought in) first; the router is added
        # only when chosen. With nobody else on the case, it goes to the router.
        if rp and cp:
            opts = ('<option value="case" selected>%s</option>' % lst(cp)
                    + '<option value="all">%s and the %s (%s)</option>' % (lst(cp), 'router' if len(rp) == 1 else 'routers', lst(rp))
                    + '<option value="router">Only the %s (%s)</option>' % ('router' if len(rp) == 1 else 'routers', lst(rp)))
        else:
            opts = '<option value="all">%s</option>' % lst(rp + cp)
        to = '<label>Send to <select id="rmto">%s</select></label>' % opts
    return ('<button type="button" class="tcbtn%s" id="tcbtn">Team chat<span id="tcn">%s</span></button>'
            '<div class="tcpanel" id="rmsg" hidden role="dialog" aria-label="Team chat">'
            '<div class="tchead"><div class="t"><b>Team chat</b> · #%s %s<div class="muted" title="Seen by %s">Only the team. %s never sees this.</div>'
            '<div id="rmhere" style="color:#1E6B45;font-weight:700;font-size:.9em"></div></div>'
            '<button type="button" class="tcclose" id="tcclose">Close</button></div>'
            '<div class="tcbody" id="cthread">%s</div>'
            '<div class="tcfoot"><div class="replying" id="rmre" hidden></div>'
            '<textarea id="rmtext" rows="3" placeholder="Write to the team. %s never sees this."></textarea>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">%s<button type="button" class="btn" id="rmgo">Send</button>'
            '<button type="button" class="btn ghost" id="tcclose2">Close</button><span id="rmst" class="muted"></span></div></div></div>'
            '<script>(function(){var rid="%s", last=%s, RE=0, RTO="", seenN=%d,'
            ' btn=document.getElementById("tcbtn"), pn=document.getElementById("rmsg"), th=document.getElementById("cthread"),'
            ' box=document.getElementById("rmtext"), o=document.getElementById("rmto"), rb=document.getElementById("rmre"), n=document.getElementById("tcn");'
            'function badge(k){n.textContent=k?" \\u00b7 "+k+" new":""; btn.className="tcbtn"+(k?" hasnew":"");'
            ' if(k>seenN){void btn.offsetWidth; btn.className+=" pulse";} seenN=k;'
            ' Array.prototype.forEach.call(document.querySelectorAll("[data-chatn]"),function(x){x.textContent=k?" ("+k+" new)":"";});}'
            'function clear(){RE=0; RTO=""; rb.hidden=true; if(o) o.disabled=false;}'
            'function wire(){Array.prototype.forEach.call(th.querySelectorAll(".rep"),function(b){b.onclick=function(){var t=b.getAttribute("data-to");'
            ' if(t=="all"){clear(); if(o) o.value="all"; rb.innerHTML="Replying to "+b.getAttribute("data-name")+", to everyone. <a href=\\"#\\" id=\\"rmx\\">Cancel</a>";}'
            ' else {RE=b.getAttribute("data-by"); RTO=t; if(o){o.value=t; o.disabled=true;}'
            '  rb.innerHTML="Replying privately to "+b.getAttribute("data-name")+". <a href=\\"#\\" id=\\"rmx\\">Cancel</a>";}'
            ' rb.hidden=false; document.getElementById("rmx").onclick=function(e){e.preventDefault(); clear();}; box.focus();};});}'
            'function bottom(){th.scrollTop=th.scrollHeight;}'
            'function poll(force){var open=!pn.hidden; bvPost({action:"req_thread",rid:rid,mark:open?"1":"0"},function(r){ if(!r.success) return;'
            ' var near=th.scrollHeight-th.scrollTop-th.clientHeight<80;'
            ' if(force||r.last!=last){th.innerHTML=r.html; last=r.last; wire(); if(near||force) bottom();}'
            ' badge(open?0:r["new"]); live=open&&!!r.live; var h=document.getElementById("rmhere");'
            ' if(h) h.textContent=(open&&r.here&&r.here.length)?"\u25CF "+r.here.join(", ")+(r.here.length==1?" is":" are")+" here":""; next();});}'
            'var live=false, tm=null;'
            'function next(){clearTimeout(tm); tm=setTimeout(function(){ if(!document.hidden) poll(false); else next(); }, live?3000:(pn.hidden?15000:8000));}'
            'document.addEventListener("visibilitychange",function(){ if(!document.hidden) poll(false); });'
            # Phones: after turning sideways and back, Safari can leave the page zoomed or shifted, and
            # "fixed to the bottom" then follows the page, not the screen. So pin both to what is
            # actually on screen (visualViewport), and redo it whenever that changes.
            'function place(){var v=window.visualViewport; if(!v) return; var m=18, sc=1/(v.scale||1);'
            ' btn.style.right="auto"; btn.style.bottom="auto"; btn.style.transformOrigin="0 0"; btn.style.transform="scale("+sc+")";'
            ' btn.style.left=(v.offsetLeft+v.width-(btn.offsetWidth*sc)-m*sc)+"px"; btn.style.top=(v.offsetTop+v.height-(btn.offsetHeight*sc)-m*sc)+"px";'
            ' if(!pn.hidden){ if(v.width<=640){pn.style.right="auto"; pn.style.bottom="auto"; pn.style.transformOrigin="0 0"; pn.style.transform="scale("+sc+")";'
            '   pn.style.left=v.offsetLeft+"px"; pn.style.top=v.offsetTop+"px"; pn.style.width=(v.width/sc)+"px"; pn.style.height=(v.height/sc)+"px";}'
            '  else {pn.style.transform=""; pn.style.left=pn.style.top=pn.style.width=pn.style.height=""; pn.style.right=""; pn.style.bottom="";}}}'
            'if(window.visualViewport){window.visualViewport.addEventListener("resize",place); window.visualViewport.addEventListener("scroll",place);}'
            'window.addEventListener("orientationchange",function(){place(); setTimeout(place,350); setTimeout(place,900);});'
            'function show(){pn.hidden=false; btn.hidden=true; place(); bottom(); poll(true); setTimeout(function(){box.focus();},50);}'
            'function hide(){pn.hidden=true; btn.hidden=false; place(); if(location.hash=="#rmsg") history.replaceState(null,"",location.pathname+location.search);}'
            'btn.onclick=show; document.getElementById("tcclose").onclick=hide; document.getElementById("tcclose2").onclick=hide;'
            'document.addEventListener("keydown",function(e){ if(e.key=="Escape"&&!pn.hidden) hide(); });'
            'Array.prototype.forEach.call(document.querySelectorAll("[data-chat=open],a[href=\\"#rmsg\\"]"),function(a){a.addEventListener("click",function(e){e.preventDefault(); show();});});'
            'if(location.hash=="#rmsg") show();'
            'function send(){var t=box.value,st=document.getElementById("rmst"); if(!t.trim()) return; st.textContent="Sending...";'
            ' bvPost({action:"req_msg",rid:rid,text:bvB64(t),to:RTO||(o?o.value:"all"),re:RE},function(r){'
            ' if(r.success){box.value=""; st.textContent=""; clear(); poll(true);} else st.textContent=r.message;});}'
            'document.getElementById("rmgo").onclick=send;'
            'box.addEventListener("keydown",function(e){ if(e.key=="Enter"&&(e.ctrlKey||e.metaKey)){e.preventDefault(); send();} });'
            'wire(); badge(seenN); seenN=%d; place(); next();})();</script>') % (
        ' hasnew' if nnew else '', (' &middot; %d new' % nnew) if nnew else '',
        r.get('num') or '', esc(s.Name), esc(case_audience(r)), esc(s.Name), html, esc(s.Name), to,
        s.Rid, safe_json(last), nnew, nnew)


def act_req_thread(st, role):
    """The thread for the chat panel; marks it seen when the panel is open."""
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    r = load_record(s.Pid)['requests'].get(s.Rid)
    if not r or not can_message(role, r):
        raise Exception('You are not part of this request.')
    html, last, nnew = chat_thread(s, r)
    who, live = [], False
    if _param('mark') == '1':
        mark_seen(s.Rid.upper(), last)
        note_here(s.Rid.upper())
        nnew = 0
        team = [r.get('owner')] + list(r.get('shared') or []) + router_ids()
        who, live = chat_live(s.Rid.upper(), [m.get('t') for m in r.get('messages') or [] if msg_visible(r, m, ME)], team)
    return {'success': True, 'html': html, 'last': last, 'new': nnew, 'here': who, 'live': live}


def act_req_msg(st, role):
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid)
    if not r or not can_message(role, r):
        raise Exception('You are not part of this request.')
    text = from_b64(_param('text')).strip()[:2000]
    if not text:
        raise Exception('Write something first.')
    to = _param('to') if _param('to') in ('router', 'case') else 'all'
    try:
        re_ = int(_param('re', '0') or 0)
    except Exception:
        re_ = 0
    team = set(router_ids() + [int(r.get('owner') or 0)] + shared_pids(r))
    if re_ not in team or re_ == int(ME) or to == 'all':
        re_ = 0
    if to == 'case' and int(ME) not in [int(r.get('owner') or 0)] + shared_pids(r):
        to = 'all'   # a router outside the case cannot write past themselves
    if int(ME) in router_ids() and int(ME) != int(r.get('owner') or 0) and not re_:
        to = 'all'
    t = now_str()
    m = {'t': t, 'by': int(ME), 'text': text}
    if to != 'all':
        m['to'] = to
    if re_:
        m['re'] = re_
    to = [x for x in set(router_ids() + [int(r.get('owner') or 0)] + shared_pids(r)) if x and x != int(ME) and msg_visible(r, m, x)]
    sent = secure_now(st, to, s.Rid.upper(), r.get('messages') or [], 'view=req&rid=%s#rmsg' % s.Rid, r.get('num'))
    if sent:
        m['mailed'] = sent
    r.setdefault('messages', []).append(m)
    save_record(s.Pid, rec)
    mark_seen(s.Rid.upper(), t)
    return {'success': True}


def case_audience(r):
    """'Pat (router), Sam (case worker) and Lee' for the conversation."""
    st = SETTINGS or {}
    routers = [int(x['pid']) for x in st.get('routers') or []]
    owner = int(r.get('owner') or 0)
    names = names_for(routers + [owner] + shared_pids(r))
    parts = ['%s (%s)' % (names[p], 'router and case worker' if p == owner else 'router') for p in routers if names.get(p)]
    if owner and names.get(owner) and owner not in routers:
        parts.append('%s (case worker)' % names[owner])
    parts += [names[p] for p in shared_pids(r) if names.get(p) and p not in routers and p != owner]
    if len(parts) > 1:
        return ', '.join(parts[:-1]) + ' and ' + parts[-1]
    return parts[0] if parts else 'the router'


def need_aff_box(st, r):
    """Affirmation is required and nobody is asked yet."""
    return bool(st.get('affirm_on') and st.get('affirm_required') and not (r.get('affirm') or {}).get('pid'))


def people_card(st, s, r, role, handoff=''):
    """Everyone on the case and their talk, in one place: the case worker,
    each teammate brought in (to help or to affirm), the affirmation note,
    who had it before, and the conversation. The case worker and the router
    bring people in; only the case worker asks for an affirmation, and can do
    it without sharing the case (a phone call)."""
    stt = r.get('status')
    owner = int(r.get('owner') or 0)
    me = int(ME or 0)
    open_ = stt in ('committee', 'requester')
    can_edit = open_ and (role == 'router' or owner == me)
    aff = r.get('affirm') or {}
    aff_pid = int(aff.get('pid') or 0)
    affirm_on = bool(st.get('affirm_on'))
    shares = r.get('shares') or []
    asg = r.get('assigned') or []
    pids = [int(a.get('pid') or 0) for a in asg] + [int(a.get('by') or 0) for a in asg]
    pids += [int(x.get('pid') or 0) for x in shares] + [int(x.get('by') or 0) for x in shares] + shared_pids(r) + [aff_pid]
    names = names_for(list(set(p for p in pids if p)))

    def note_of(p):
        if p != aff_pid or not aff.get('note'):
            return ''
        return '<div class="post"><div class="muted">Affirmation note%s</div><div>%s</div></div>' % (
            (' · ' + nice_date(aff.get('at'))) if aff.get('at') else '', esc(aff['note']))

    def off_btn(p):
        if not open_ or not (can_edit or p == me):
            return ''
        return '<button type="button" class="lnk shrm" data-who="%d">%s</button>' % (p, 'Step off' if p == me else 'Take off')

    rows = []
    if asg:
        a = asg[-1]
        rows.append('<div class="cmrow" title="Given by %s"><b>%s</b><span class="pill p-grey">case worker</span><span class="muted">since %s</span></div>' % (
            esc(names.get(int(a.get('by') or 0), 'the router')), esc(names.get(owner, '')), days_ago(a.get('t') or '')))
    current = dict((int(x['pid']), x) for x in shares if not x.get('ended'))
    for p in shared_pids(r):
        x = current.get(p) or {}
        what = ('<span class="pill p-warn">affirming</span>' if not aff.get('note') else '<span class="pill p-ok">affirmed</span>') if p == aff_pid else '<span class="pill p-grey">helping</span>'
        why = esc(x['why']) if x.get('why') and x.get('why') != 'Affirmation' else ''
        rows.append('<div class="cmrow" title="Brought in%s%s"><b>%s</b>%s<span class="muted">%s</span>%s</div>%s' % (
            (' by ' + esc(names.get(int(x.get('by') or 0), ''))) if x.get('by') else '', (' ' + days_ago(x['t'])) if x.get('t') else '',
            esc(names.get(p, '')), what, why, off_btn(p), note_of(p)))
    if aff_pid and aff_pid not in shared_pids(r):
        rows.append('<div class="cmrow"><b>%s</b>%s<span class="muted">talked it over %s, the case is not shared with them</span>%s</div>%s' % (
            esc(names.get(aff_pid, '')), '<span class="pill p-ok">affirmed</span>' if aff.get('note') else '<span class="pill p-warn">affirming</span>',
            esc((aff.get('how') or 'by phone').lower()), off_btn(aff_pid) if owner == me else '', note_of(aff_pid)))

    # The teammate asked to affirm writes their note right here.
    mine = ''
    if open_ and affirm_on and aff_pid == me and me in shared_pids(r):
        mine = ('<div class="bvc bvblue" id="affn"><b>%s asked you to affirm this</b><div class="muted">Look it over. '
                'The decision is theirs; your note goes to them and the router.</div>'
                '<textarea id="affnote" rows="3" style="width:100%%">%s</textarea>'
                '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn" id="affsave">Save my note</button>'
                '<span class="muted" id="affst"></span></div></div>'
                '<script>document.getElementById("affsave").onclick=function(){var st=document.getElementById("affst");st.textContent="Saving...";'
                'bvPost({action:"affirm_note",rid:"%s",note:document.getElementById("affnote").value},function(r){st.textContent=r.success?"Saved.":r.message;});};</script>') % (
            esc(names.get(owner, 'The case worker')), esc(aff.get('note') or ''), s.Rid)

    before = []
    for i, a in enumerate(asg[:-1]):
        nxt = asg[i + 1]
        before.append('<div class="muted">%s · case worker %s to %s</div>' % (
            esc(names.get(int(a.get('pid') or 0), '')), nice_date(a.get('t')), nice_date(nxt.get('t'))))
    for x in shares:
        if x.get('ended'):
            before.append('<div class="muted">%s · brought in %s to %s%s</div>' % (
                esc(names.get(int(x.get('pid') or 0), '')), nice_date(x.get('t')), nice_date(x['ended']),
                (' (' + esc(x['end_why']) + ')') if x.get('end_why') else ''))

    add, add_btn = '', ''
    can_affirm = affirm_on and owner == me and open_
    if can_edit:
        taken = set([owner] + shared_pids(r))
        opts = ''.join('<option value="%d">%s</option>' % (m['pid'], esc(m['name'])) for m in committee_members(st) if m['pid'] not in taken)
        aff_opts = ''.join('<option value="%d">%s</option>' % (m['pid'], esc(m['name'])) for m in committee_members(st) if m['pid'] != owner)
        need_aff = can_affirm and st.get('affirm_required') and not aff_pid
        if opts or can_affirm:
            first_aff = not aff_pid
            purpose = ('<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">'
                       '<label class="pick"><input type="radio" name="shfor" value="affirm"%s> To affirm my decision%s</label>'
                       '<label class="pick"><input type="radio" name="shfor" value="help"%s> To help with it</label></div>' % (
                           ' checked' if first_aff else '', ' (required)' if st.get('affirm_required') else '', '' if first_aff else ' checked')
                       ) if can_affirm else ''
            add = ('<div class="shform" id="shform"%s><b>%s</b>%s'
                   '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">'
                   '<select id="shwho"><option value="">Pick someone</option>%s</select>'
                   '<input type="text" id="shwhy" placeholder="Why (optional), e.g. knows the family" size="28">'
                   '<label class="pick" id="shsharel"><input type="checkbox" id="shshare" checked> Share the case so they can read it here</label>'
                   '<label id="shhowl">How you talked it over <select id="shhow"><option>By phone</option><option>In person</option></select></label>'
                   '<button type="button" class="btn" id="shgo">Save</button></div>'
                   '<div class="muted" id="shhelp"></div></div>') % (
                '' if need_aff else ' hidden', 'Add a teammate' + (' or ask one to affirm' if can_affirm else ''), purpose, opts, )
            add_btn = '<button type="button" class="lnk" id="shopen">+ Add</button>'
            add += ('<script>document.getElementById("shopen").onclick=function(){var f=document.getElementById("shform"); f.hidden=!f.hidden;'
                    ' if(!f.hidden) document.getElementById("shwho").focus();};</script>')
            add += ('<script>(function(){var OPTS=%s, AFF=%s;'
                    'function f(){var c=document.querySelector("input[name=shfor]:checked"); return c?c.value:"help";}'
                    'function show(){var a=f()=="affirm", sh=document.getElementById("shshare").checked, w=document.getElementById("shwho"), v=w.value;'
                    ' w.innerHTML=\'<option value="">Pick someone</option>\'+(a?AFF:OPTS); w.value=v;'
                    ' document.getElementById("shwhy").style.display=a?"none":"";'
                    ' document.getElementById("shsharel").style.display=a?"":"none";'
                    ' document.getElementById("shhowl").style.display=a&&!sh?"":"none";'
                    ' document.getElementById("shhelp").textContent=!a?"They can read this case, write in its conversation and add notes. They cannot call, decide or hand it off."'
                    '  :sh?"They read the case here and leave their note. The decision stays yours."'
                    '  :"Only their name and how you talked goes on the case. They will not see it here.";}'
                    'Array.prototype.forEach.call(document.querySelectorAll("input[name=shfor],#shshare"),function(x){x.onchange=show;});'
                    'show();'
                    'document.getElementById("shgo").onclick=function(){var w=document.getElementById("shwho").value, st=document.getElementById("shst"); if(!w) return;'
                    ' st.textContent="Saving..."; var d=f()=="affirm"?{action:"affirm_set",rid:"%s",who:w,share:document.getElementById("shshare").checked?"1":"0",'
                    '  how:document.getElementById("shhow").value}:{action:"share_add",rid:"%s",who:w,why:document.getElementById("shwhy").value};'
                    ' bvPost(d,function(r){ if(r.success) location.reload(); else st.textContent=r.message;});};})();</script>') % (
                safe_json(opts), safe_json(aff_opts), s.Rid, s.Rid)

    talk = ''
    can_talk = bool(stt and stt != 'new' and can_message(role, r))
    nmsg = len([m for m in r.get('messages') or [] if msg_visible(r, m, me)])
    nnew = len(req_unread(s.Rid, r)) if can_talk else 0   # before the conversation marks them seen
    if can_talk:
        talk = ('<div class="sec"><div class="cmhead"><div class="cml">Team chat</div><span class="muted">%s</span></div>'
                '<button type="button" class="btn ghost" data-chat="open">Open team chat<span data-chatn>%s</span></button></div>') % (
                    ('%d message%s' % (nmsg, '' if nmsg == 1 else 's')) if nmsg else 'no messages yet', (' (%d new)' % nnew) if nnew else '')
    others = [names.get(p, '') for p in shared_pids(r)] + ([names.get(aff_pid, '')] if aff_pid and aff_pid not in shared_pids(r) else [])
    glance = [('with ' + ', '.join(x for x in others if x)) if others else ('just ' + ('you' if owner == me else esc(names.get(owner, 'the case worker'))))]
    if nmsg:
        glance.append('%d message%s%s' % (nmsg, '' if nmsg == 1 else 's', (' <b>(%d new)</b>' % nnew) if nnew else ''))
    if r.get('handoff'):
        glance.append('<b>you asked to hand it off</b>' if owner == me else '<b>hand-off asked</b>')
    opened = bool(nnew or r.get('handoff') or (me != owner and me in shared_pids(r)) or (need_aff_box(st, r) and owner == me and open_))
    return ('<details class="bvc cm" id="people"%s><summary><b>Case management</b><span class="muted" style="font-weight:normal">%s</span></summary>'
            '%s<div class="sec"><div class="cmhead"><div class="cml">Team</div>%s</div>%s%s%s%s<div class="muted" id="shst"></div></div>%s</details>'
            '<script>(function(){var d=document.getElementById("people");'
            'function o(){if(location.hash=="#people") d.open=true;} o(); window.addEventListener("hashchange",o);'
            'Array.prototype.forEach.call(document.querySelectorAll("a[href=\'#people\']"),function(a){a.addEventListener("click",function(){d.open=true;});});})();</script>'
            '<script>(function(){var st=document.getElementById("shst");'
            'Array.prototype.forEach.call(document.querySelectorAll("#people .shrm"),function(x){x.onclick=function(){st.textContent="Saving...";'
            ' bvPost({action:"share_remove",rid:"%s",who:x.getAttribute("data-who")},function(r){ if(r.success) location.reload(); else st.textContent=r.message;});};});})();</script>') % (
        ' open' if opened else '', ' · '.join(glance), ('<div class="sec">%s</div>' % handoff) if handoff else '', add_btn,
        ''.join(rows) or '<div class="muted">Not given to anyone yet.</div>', mine,
        ('<details class="flat"><summary>Before (%d)</summary>%s</details>' % (len(before), ''.join(before))) if before else '',
        add, talk, s.Rid)


def plan_so_far(st, r):
    """The case worker's decision draft, read only, for a teammate and the router."""
    d = r.get('decision') or {}
    if not d.get('saved') or d.get('choice') == 'notyet' and not any(x.get('help') for x in d.get('items') or []):
        return ''
    who = names_for([r.get('owner')]).get(int(r.get('owner') or 0), 'The case worker')
    if d.get('choice') == 'decline':
        body = '<div>Leaning toward declining%s.</div>' % ((': ' + esc(d['decline_reason'])) if d.get('decline_reason') else '')
    else:
        lines = [x for x in d.get('items') or [] if x.get('help')]
        body = ''.join('<div class="bvr"><span><b>%s</b>%s</span><span>%s · %s</span></div>' % (
            esc(x.get('need') or ''), (' · ' + esc(x['payee'])) if x.get('payee') else '', money(to_number(x.get('amount'))) or '$?',
            esc(x.get('how') or '')) for x in lines) or '<div class="muted">No help ticked yet.</div>'
        body += pickup_line(st, d)
    return ('<div class="bvc" style="background:#F7F5F0"><div class="bvr"><h2>%s\'s plan so far</h2><span class="muted">not sent · saved %s</span></div>%s%s</div>') % (
        esc(who), days_ago(d['saved']), body, ('<div>%s</div>' % clean_html(d['reason'])) if d.get('reason') else '')


def act_share_add(st, role):
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid)
    if not r or r.get('status') not in ('committee', 'requester'):
        raise Exception('This request is not open with the committee.')
    if role != 'router' and int(r.get('owner') or 0) != int(ME):
        raise Exception('Only the case worker or the router can bring someone in.')
    who = int(_param('who', '0') or 0)
    members = dict((m['pid'], m['name']) for m in committee_members(st))
    if who not in members or who == int(r.get('owner') or 0):
        raise Exception('Pick another committee member.')
    if not share_add(r, who, _param('why')):
        raise Exception('They are already on this case.')
    add_history(r, 'Brought %s in on the case' % members[who], esc(_param('why')[:200]))
    notify(r, 'shared', [who], 'brought you in on it', rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def act_share_remove(st, role):
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid)
    who = int(_param('who', '0') or 0)
    aff_only = r and who not in shared_pids(r) and int((r.get('affirm') or {}).get('pid') or 0) == who
    if not r or (who not in shared_pids(r) and not aff_only):
        raise Exception('They are not on this case.')
    if who != int(ME) and role != 'router' and int(r.get('owner') or 0) != int(ME):
        raise Exception('Only the case worker or the router can take someone off.')
    nm = names_for([who]).get(who, '')
    if aff_only:
        r.pop('affirm', None)
    else:
        share_end(r, who, 'stepped off' if who == int(ME) else 'taken off by %s' % names_for([ME]).get(int(ME), ''))
    if int((r.get('affirm') or {}).get('pid') or 0) == who and not (r.get('affirm') or {}).get('note'):
        r.pop('affirm', None)
    add_history(r, '%s is no longer on the case' % nm)
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- texting the person who asked
# Send only, through a TouchPoint texting group picked in Settings. The text is
# a fixed line with a link; anything personal stays in this tool, on the case
# (r['texts']), and the person reads it on their own page (view=letter), which
# shows only what was sent to them. TouchPoint keeps texts in its own history,
# so a text never carries details, and only carries a sign-in link when an
# admin turned that on.
TEXT_STATUSES = ('committee', 'requester', 'approved', 'pickup')


def church_name(st):
    if (st.get('text_church') or '').strip():
        return st['text_church'].strip()
    try:
        return model.Setting('NameOfChurch', '') or 'Your church'
    except Exception:
        return 'Your church'


def requester_reach(pid):
    """(cell phone, has a TouchPoint login)."""
    c = reach_ways(pid)
    return c['cell'], c['login']


def reach_ways(pid):
    """{'cell', 'email', 'login'}: how the person who asked can be reached."""
    p = q.QuerySqlTop1("""SELECT ISNULL(p.CellPhone, '') Cell, ISNULL(p.EmailAddress, '') Email,
        (SELECT COUNT(*) FROM Users u WHERE u.PeopleId = p.PeopleId) U FROM People p WHERE p.PeopleId = %d""" % int(pid))
    return {'cell': (p.Cell or '') if p else '', 'email': (p.Email or '').strip() if p else '', 'login': bool(p and p.U)}


def send_ways(st, pid):
    """The ways this case can message them: ['text', 'email'], whatever is set up and on file."""
    c = reach_ways(pid)
    out = []
    if int(st.get('text_group') or 0) and c['cell']:
        out.append('text')
    if c['email'] and '@' in c['email']:
        out.append('email')
    return out, c


def text_link(st, pid, url):
    """The link for a text. Signs them in only when an admin turned that on,
    and never in Testing (it would let a tester in as them)."""
    if st.get('text_autosignin') and st.get('mode') == 'live':
        try:
            return model.GetAuthenticatedUrl(int(pid), url, True)
        except Exception:
            return url
    return url


EMAIL_SUBJECTS = {'secure': '{church}: a message about your request for help',
                  'docs': '{church}: documents for your request for help'}


def send_to_them(st, s, r, kind, url, ways):
    """Text, email or both. Each carries only the fixed line and the link,
    never what was written. Returns {'text': sent?, 'email': sent?}; in
    Testing nothing goes out."""
    out = {}
    if 'text' in ways:
        out['text'] = send_text(st, s, r, kind, url)
    if 'email' in ways:
        link = text_link(st, s.Pid, url)
        line = esc(st.get('text_secure' if kind == 'secure' else 'text_docs') or DEFAULT_SETTINGS['text_' + kind]).replace('{church}', esc(church_name(st)))
        line = line.replace('{link}', '<a href="%s">%s</a>' % (esc(link), esc(link)))
        line = line.replace('Replies to this number are not monitored.', '')
        if st.get('mode') == 'live':
            send_email(st, s.Pid, EMAIL_SUBJECTS[kind].replace('{church}', church_name(st)), '<div><p>%s</p></div>' % line.strip())
            out['email'] = True
        else:
            out['email'] = False
    return out


def by_words(ways):
    return ' and '.join({'text': 'text', 'email': 'email'}[w] for w in ways)


def send_text(st, s, r, kind, url):
    """Text them. In Testing nothing is sent; the entry says so."""
    msg = (st.get('text_secure' if kind == 'secure' else 'text_docs') or DEFAULT_SETTINGS['text_' + kind])
    msg = msg.replace('{church}', church_name(st)).replace('{link}', text_link(st, s.Pid, url))
    if st.get('mode') != 'live':
        return False
    model.SendSms(str(int(s.Pid)), int(st['text_group']), 'Benevolence #%s' % (r.get('num') or ''), msg[:1600])
    return True


def can_text(st, role, r):
    return (r.get('status') in TEXT_STATUSES and (role == 'router' or int(r.get('owner') or 0) == int(ME or 0)))


def texts_card(st, s, r, role, inner=False):
    """Texting them. inner: folded inside Step 1, closed until opened."""
    out = texts_body(st, s, r, role, inner)
    return (out + '</details>') if inner and out else out


def texts_body(st, s, r, role, inner):
    if not can_text(st, role, r):
        return ''
    ways, c = send_ways(st, s.Pid)
    sent = r.get('texts') or []
    names = names_for([x.get('by') for x in sent])
    rows = ''.join('<div class="post"><div class="bvr"><b>%s</b><span class="muted">%s · %s · by %s%s</span></div>%s</div>' % (
        'Secure message' if x.get('kind') == 'secure' else 'Documents link', esc(names.get(int(x.get('by') or 0), '')), nice_date(x.get('t')),
        by_words(x.get('via') or ['text']),
        (' · read %s' % nice_date(x['read'])) if x.get('read') else (' · <b>Testing: not sent</b>' if x.get('test') else ' · not opened yet'),
        ('<div>%s</div>' % text_html(x['text'])) if x.get('text') else '') for x in reversed(sent))
    who = esc(s.Name)
    head = '<div class="bvc" id="texts"><h2>Message %s</h2>' % who
    if inner:
        head = ('<details id="texts" class="sec"><summary><b>Message %s</b><span class="muted" style="font-weight:normal">by text or email: a secure message or the documents link%s</span></summary><div>' % (
            who, (' · %d sent' % len(sent)) if sent else ''))
    if not ways:
        why = []
        if not c['cell']:
            why.append('no cell phone')
        elif not int(st.get('text_group') or 0):
            why.append('texting is not set up (Settings, Texting them)')
        if not c['email']:
            why.append('no email address')
        return head + '<div class="muted">They cannot be messaged here: %s on their record. Call them instead.</div>%s</div>' % (' and '.join(why), rows)
    labels = {'text': 'Text to %s' % esc(c['cell']), 'email': 'Email to %s' % esc(c['email'])}
    pick = ''.join('<label class="pick"><input type="radio" name="txby" value="%s"%s> %s</label>' % (w, ' checked' if w == ways[0] else '', labels[w]) for w in ways)
    if len(ways) == 2:
        pick += '<label class="pick"><input type="radio" name="txby" value="both"> Both</label>'
    warn = '' if c['login'] else ('<div class="bvc bvwarn" style="padding:8px 12px">They have no TouchPoint login, so they could not open a secure message. '
                                  'Call them instead.%s</div>' % (' The documents link still works: the form lets them fill it out without one.' if st.get('docs_org') else ''))
    docs = (('<button type="button" class="btn ghost" id="txdocs">Send them the documents link</button>')
            if int(st.get('docs_org') or 0) else '<span class="muted">The documents form is not set up yet.</span>')
    return (head + '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><b>Send by</b>%s</div>'
            '<div class="muted">Send only: they cannot reply%s. What you write stays here and on their timeline; '
            'the text or email only says they have a message, with a link.</div>%s'
            '%s<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">%s%s<span class="muted" id="txst"></span></div>%s</div>'
            '<script>(function(){var st=document.getElementById("txst");'
            'function by(){var c=document.querySelector("input[name=txby]:checked"); return c?c.value:"";}'
            'function go(d){d.by=by(); st.textContent="Sending..."; bvPost(d,function(r){ if(r.success) location.reload(); else st.textContent=r.message;});}'
            'var b=document.getElementById("txgo"), tx=document.getElementById("txtext");'
            'function ready(){ if(b&&tx){var e=!tx.value.trim(); b.disabled=e; b.title=e?"Write the message first":"";}}'
            'if(tx) tx.addEventListener("input",ready); ready();'
            'if(b) b.onclick=function(){var t=tx.value; if(!t.trim()) return;'
            ' go({action:"text_secure",rid:"%s",text:bvB64(t)});};'
            'var dd=document.getElementById("txdocs"); if(dd) dd.onclick=function(){ go({action:"text_docs",rid:"%s"}); };})();</script>') % (
        pick, (' (questions go to %s)' % esc(st['text_phone'])) if st.get('text_phone') else '', warn,
        ('<textarea id="txtext" rows="3" style="width:100%%;border-color:#1E6B45" placeholder="Write to %s. This is exactly what they will read."></textarea>' % who) if c['login'] else '',
        ('<button type="button" class="btn" id="txgo" style="background:#1E6B45">Send to %s</button>' % who) if c['login'] else '', docs,
        rows, s.Rid, s.Rid)


def text_target(st, role):
    """(submission, record, request, has login, ways chosen) for a send."""
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid)
    if not r or not can_text(st, role, r):
        raise Exception('Only the case worker or the router can message them.')
    ways, c = send_ways(st, s.Pid)
    if not ways:
        raise Exception('They have no cell phone or email we can use. Call them instead.')
    want = _param('by') or ways[0]
    chosen = ways if want == 'both' else [w for w in ways if w == want]
    if not chosen:
        raise Exception('Pick how to send it.' if want not in ('text', 'email') else
                        ('They cannot be texted.' if want == 'text' else 'There is no email address on their record.'))
    return s, rec, r, c['login'], chosen


def act_text_secure(st, role):
    s, rec, r, login, ways = text_target(st, role)
    if not login:
        raise Exception('They have no TouchPoint login, so they could not read it. Call them instead.')
    text = from_b64(_param('text')).strip()[:2000]
    if not text:
        raise Exception('Write the message first.')
    res = send_to_them(st, s, r, 'secure', site_link('view=letter&rid=%s' % s.Rid), ways)
    test = not any(res.values())
    r.setdefault('texts', []).append({'t': now_str(), 'by': int(ME), 'kind': 'secure', 'text': text, 'test': test, 'via': ways})
    add_history(r, 'Sent them a secure message by %s%s' % (by_words(ways), ' (Testing: not sent)' if test else ''), text_html(text))
    save_record(s.Pid, rec)
    return {'success': True}


def act_text_docs(st, role):
    s, rec, r, login, ways = text_target(st, role)
    org = int(st.get('docs_org') or 0)
    if not org:
        raise Exception('The documents form is not set up yet.')
    res = send_to_them(st, s, r, 'docs', '%s/OnlineReg/%d' % ((model.CmsHost or '').rstrip('/'), org), ways)
    test = not any(res.values())
    r.setdefault('texts', []).append({'t': now_str(), 'by': int(ME), 'kind': 'docs', 'test': test, 'via': ways})
    if r.get('status') == 'committee':
        r['status'] = 'requester'
    add_history(r, 'Asked for documents by %s%s' % (by_words(ways), ' (Testing: not sent)' if test else ''))
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- documents they upload
# The documents form (Settings, docs_org) is a TouchPoint form with file
# uploads. Each upload belongs to the newest request from anyone in that
# family made before it, up to the family's next request. When documents
# arrive for a case waiting on them, it goes back to the committee member,
# who sees "Documents came in" on their list, in their daily email and on
# the case. Files stay in TouchPoint: the tool only links to them.
def doc_uploads(st, fams, rids=None):
    """Every answer on the documents form by these families, in each
    submission's form order, oldest submission first. Works for any layout:
    a file question is {'kind': 'file', 'name', 'image', 'doc'} (doc is None
    when TouchPoint no longer has the file, e.g. the person was dropped);
    anything else they typed is {'kind': 'text', 'text'}, like a note about
    an upload. Headings and the name/contact block are left out, the same as
    the request form. Each item also has 'rid', 'pid', 'fam', 't', 'label', 'order'."""
    org = int(st.get('docs_org') or 0)
    fams = [int(f) for f in fams if f]
    rids = [x.upper() for x in rids or [] if GUID.match(x or '')]
    if not org or not (fams or rids):
        return []
    who = ' OR '.join(([('p.FamilyId IN (%s)' % ','.join(str(f) for f in set(fams)))] if fams else []) +
                      ([("rp.RegPeopleId IN (%s)" % ','.join("'%s'" % x for x in set(rids)))] if rids else []))
    import json
    out = []
    for a in q.QuerySql("""
            SELECT CAST(rp.RegPeopleId AS varchar(36)) Rid, rp.PeopleId Pid, p.FamilyId Fam, a.AnswerValue A,
                   q.QuestionTypeId Qt, q.[Order] O, q.Label, CAST(q.RegQuestionId AS varchar(36)) Qid,
                   CONVERT(varchar(19), ISNULL(rp.CompletedDate, r.CompletedDate), 120) T
            FROM RegAnswer a JOIN RegQuestion q ON q.RegQuestionId = a.RegQuestionId
            JOIN RegPeople rp ON rp.RegPeopleId = a.RegPeopleId JOIN Registration r ON r.RegistrationId = rp.RegistrationId
            JOIN People p ON p.PeopleId = rp.PeopleId
            WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL AND q.QuestionTypeId NOT IN (4, 5)
              AND a.AnswerValue IS NOT NULL AND (%s)
            ORDER BY ISNULL(rp.CompletedDate, r.CompletedDate), q.[Order]""" % (org, who)):
        base = {'rid': a.Rid.upper(), 'pid': a.Pid, 'fam': a.Fam, 't': a.T, 'label': strip_tags(a.Label), 'order': a.O or 0,
                'qid': (a.Qid or '').upper()}
        if a.Qt == 7:
            try:
                v = json.loads(a.A)
            except Exception:
                v = a.A
            for one in (v if isinstance(v, list) else [v]):
                one = unicode(one or '')
                img, _, name = one.partition(':')
                if img.strip().isdigit():
                    out.append(dict(base, kind='file', name=name.strip(), image=int(img), doc=None))
        else:
            val = answer_value(a.A).strip()
            if val:
                out.append(dict(base, kind='text', text=val))
    files = [x for x in out if x['kind'] == 'file']
    if files:
        ids = dict((x.ImageId, x.DocumentId) for x in q.QuerySql(
            "SELECT ImageId, DocumentId FROM OrgMemberDocuments WHERE OrganizationId = %d AND ImageId IN (%s)" % (
                org, ','.join(str(x['image']) for x in files))))
        for x in files:
            x['doc'] = ids.get(x['image'])
    return out


def doc_links():
    """What the routers decided about documents the tool could not place on
    its own: ({documents RegPeopleId: request RegPeopleId}, {set aside}).
    Kept in the protected records: a link on the request it was attached to
    (docs_linked), a set-aside on the record of whoever sent it."""
    links, aside = {}, set()
    for r in q.QuerySql("""
            SELECT PeopleId, CAST(Data AS nvarchar(max)) Data FROM PeopleExtra WITH (NOLOCK)
            WHERE Field = '%s' AND Data IS NOT NULL AND (Data LIKE '%%docs_linked%%' OR Data LIKE '%%docs_aside%%')""" % EV_FIELD):
        rec = load_json(r.Data or '', {})
        for rid, req in (rec.get('requests') or {}).items():
            for d in req.get('docs_linked') or []:
                links[d.upper()] = rid.upper()
        aside.update(d.upper() for d in rec.get('docs_aside') or [])
    return links, aside


def docs_by_request(st, subs):
    """{RID: [uploads]} for these requests: each upload goes to the family's
    newest request made before it, unless a router attached it to another
    request (or set it aside)."""
    subs = [x for x in subs if x]
    fams = list(set(x.Fam for x in subs))
    links, aside = doc_links()
    mine = set(x.Rid.upper() for x in subs)
    linked_here = [d for d, to in links.items() if to in mine]
    ups = doc_uploads(st, fams, linked_here)
    if not ups:
        return {}
    starts = {}
    for f in q.QuerySql("""
            SELECT CAST(rp.RegPeopleId AS varchar(36)) Rid, p.FamilyId Fam,
                   CONVERT(varchar(19), ISNULL(rp.CompletedDate, r.CompletedDate), 120) T
            FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId JOIN People p ON p.PeopleId = rp.PeopleId
            WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL AND p.FamilyId IN (%s)""" % (
                int(st['form_org']), ','.join(str(int(f)) for f in fams if f) or '0')):
        starts.setdefault(f.Fam, []).append((f.T, f.Rid.upper()))
    out = {}
    for u in ups:
        if u['rid'] in links:
            out.setdefault(links[u['rid']], []).append(dict(u, linked=True))
            continue
        if u['rid'] in aside:
            continue
        before = sorted(x for x in starts.get(u['fam'], []) if x[0] <= u['t'])
        if before:
            out.setdefault(before[-1][1], []).append(u)
    for k, to in combined_map().items():
        if k in out:
            out.setdefault(to, []).extend(out.pop(k))
    return out


def unsorted_docs(st, days=180):
    """Documents the tool could not place: no request from that family before
    they were sent, and no router has attached them or set them aside.
    [(submission row, [uploads])], newest first; and the set-aside ones."""
    org = int(st.get('docs_org') or 0)
    if not org or not int(st.get('form_org') or 0):
        return [], []
    cut = (datetime.date.today() - datetime.timedelta(days=days)).strftime('%Y-%m-%d')
    sent = list(q.QuerySql("""
        SELECT CAST(rp.RegPeopleId AS varchar(36)) Rid, rp.PeopleId Pid, p.FamilyId Fam,
               CONVERT(varchar(19), ISNULL(rp.CompletedDate, r.CompletedDate), 120) T,
               ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName Name
        FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId JOIN People p ON p.PeopleId = rp.PeopleId
        WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL AND ISNULL(rp.CompletedDate, r.CompletedDate) >= '%s'
        ORDER BY ISNULL(rp.CompletedDate, r.CompletedDate) DESC""" % (org, cut)))
    if not sent:
        return [], []
    links, aside = doc_links()
    fams = list(set(x.Fam for x in sent if x.Fam))
    first = {}
    for f in q.QuerySql("""
            SELECT p.FamilyId Fam, MIN(CONVERT(varchar(19), ISNULL(rp.CompletedDate, r.CompletedDate), 120)) T
            FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId JOIN People p ON p.PeopleId = rp.PeopleId
            WHERE r.OrganizationId = %d AND r.CompletedDate IS NOT NULL AND p.FamilyId IN (%s) GROUP BY p.FamilyId""" % (
                int(st['form_org']), ','.join(str(int(f)) for f in fams) or '0')):
        first[f.Fam] = f.T
    loose = [x for x in sent if x.Rid.upper() not in links and not (first.get(x.Fam) and first[x.Fam] <= x.T)]
    ups = doc_uploads(st, [], [x.Rid for x in loose]) if loose else []
    by = {}
    for u in ups:
        by.setdefault(u['rid'], []).append(u)
    out = [(x, by.get(x.Rid.upper(), [])) for x in loose if x.Rid.upper() not in aside and by.get(x.Rid.upper())]
    put_aside = [(x, by.get(x.Rid.upper(), [])) for x in loose if x.Rid.upper() in aside and by.get(x.Rid.upper())]
    return out, put_aside


def docs_role(st, u):
    """show, private or hide for one answer on the documents form."""
    fmap = dict((k.upper(), v) for k, v in (st.get('docs_map') or {}).items())
    if u.get('qid') in fmap:
        return fmap[u['qid']]
    return 'private' if u['kind'] != 'file' and word_test(st.get('form_words') or DEFAULT_SETTINGS['form_words'], 'private')(u['label']) else 'show'


def docs_rows(st, items, role):
    """One submission's answers as label/value rows."""
    rows = []
    for u in sorted(items, key=lambda x: x['order']):
        how = docs_role(st, u)
        if how == 'hide':
            continue
        if u['kind'] == 'file':
            val = ('<b>%s</b> <a class="btn ghost" target="_blank" rel="noopener" href="/OrgMemberDialog/DocumentDownload/%d">Open</a>' % (
                esc(u['name'] or 'File'), u['doc'])) if u['doc'] else '<b>%s</b> <span class="muted">removed from TouchPoint</span>' % esc(u['name'] or 'File')
        elif how == 'private' and role != 'router':
            val = '<i>private</i>'
        else:
            val = text_html(u['text'])
        rows.append('<span>%s</span><span class="a">%s</span>' % (esc(u['label'][:120]), val))
    return ''.join(rows)


def docs_place_select(options, current=''):
    return ''.join('<option value="%s"%s>%s</option>' % (esc(v), ' selected' if v == current else '', esc(l)) for v, l in options)


def request_choices(st, pids_first=None, skip=''):
    """Requests a router can attach documents to: the sender's family first,
    then the newest others. [(rid, label)]."""
    records = all_records()
    seen, out = set(), []
    rows = []
    for pid in pids_first or []:
        rows += [(x, True) for x in reversed(family_requests(st, pid))]
    rows += [(x, False) for x in submissions(st, '', 60)]
    names = {}
    for x, fam in rows:
        rid = x.Rid.upper()
        if rid in seen or rid == skip.upper():
            continue
        seen.add(rid)
        r = (records.get(x.Pid) or {}).get(x.Rid) or (records.get(x.Pid) or {}).get(rid) or {}
        nm = getattr(x, 'Name', None) or names.setdefault(x.Pid, names_for([x.Pid]).get(int(x.Pid), ''))
        out.append((rid, '%s#%s %s, %s' % ('Same family: ' if fam else '', r.get('num') or '?', nm, nice_date(x.T))))
    return out


def unsorted_docs_card(st):
    """Router's home: documents that match no request, to attach or set aside."""
    loose, aside = unsorted_docs(st)
    if not loose and not aside:
        return ''
    blocks = []
    for x, ups in loose:
        opts = [('', 'Attach to...')] + request_choices(st, [x.Pid])
        blocks.append('<div class="post" data-doc="%s"><div class="bvr"><b>%s sent documents %s %s</b></div>'
                      '<div class="qa" style="padding-left:0">%s</div>'
                      '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><select class="docto">%s</select>'
                      '<button type="button" class="btn ghost docgo">Attach</button>'
                      '<button type="button" class="btn ghost docaside">Not for a request</button><span class="docst"></span></div></div>' % (
                          x.Rid, esc(x.Name), nice_date(x.T), esc(x.T[11:16]), docs_rows(st, ups, 'router'), docs_place_select(opts)))
    aside_html = ''
    if aside:
        aside_html = ('<details><summary>Set aside (%d)</summary>%s</details>') % (len(aside), ''.join(
            '<div class="post" data-doc="%s"><div class="bvr"><b>%s, %s</b> <button type="button" class="btn ghost docback">Put back</button>'
            '<span class="docst"></span></div></div>' % (x.Rid, esc(x.Name), nice_date(x.T)) for x, ups in aside))
    return ('<details class="bvc" id="unsorted"%s><summary><b>Documents that match no request (%d)</b> '
            '<span class="muted">Sent on the documents form by someone with no request from their family before it. Attach each to the right request.</span></summary>'
            '%s%s</details>'
            '<script>(function(){var c=document.getElementById("unsorted");'
            'function go(p,d){var st=p.querySelector(".docst"); st.textContent="Saving..."; d.action="docs_place"; d.doc=p.getAttribute("data-doc");'
            ' bvPost(d,function(r){if(r.success) location.reload(); else st.textContent=r.message;});}'
            'Array.prototype.forEach.call(c.querySelectorAll(".post"),function(p){'
            ' var g=p.querySelector(".docgo"); if(g) g.onclick=function(){var to=p.querySelector(".docto").value; if(!to){p.querySelector(".docst").textContent="Pick a request first.";return;} go(p,{to:to});};'
            ' var a=p.querySelector(".docaside"); if(a) a.onclick=function(){go(p,{to:"aside"});};'
            ' var b=p.querySelector(".docback"); if(b) b.onclick=function(){go(p,{to:"auto"});};});})();</script>') % (
        ' open' if loose else '', len(loose), ''.join(blocks) or '<div class="muted">Nothing waiting.</div>', aside_html)


def act_docs_place(st, role):
    """A router attaches one documents submission to a request, sets it aside
    (to: aside), or hands it back to the tool's own matching (to: auto)."""
    if role != 'router':
        raise Exception('Only routers can move documents.')
    doc = _param('doc').upper()
    to = _param('to')
    if not GUID.match(doc):
        raise Exception('Those documents were not found.')
    sender = q.QuerySqlTop1("""
        SELECT rp.PeopleId Pid FROM RegPeople rp JOIN Registration r ON r.RegistrationId = rp.RegistrationId
        WHERE r.OrganizationId = %d AND rp.RegPeopleId = '%s'""" % (int(st.get('docs_org') or 0), sql_str(doc)))
    if not sender:
        raise Exception('Those documents were not found on the documents form.')
    target = None
    if to not in ('aside', 'auto'):
        target = find_submission(st, to)
        if not target:
            raise Exception('That request was not found.')
    # Take it off wherever it was placed before.
    for r0 in q.QuerySql("""
            SELECT PeopleId FROM PeopleExtra WITH (NOLOCK)
            WHERE Field = '%s' AND CAST(Data AS nvarchar(max)) LIKE '%%%s%%'""" % (EV_FIELD, sql_str(doc))):
        rec = load_record(r0.PeopleId)
        changed = False
        for req in rec['requests'].values():
            keep = [d for d in req.get('docs_linked') or [] if d.upper() != doc]
            if keep != (req.get('docs_linked') or []):
                req['docs_linked'] = keep
                changed = True
        keep = [d for d in rec.get('docs_aside') or [] if d.upper() != doc]
        if keep != (rec.get('docs_aside') or []):
            rec['docs_aside'] = keep
            changed = True
        if changed:
            save_record(r0.PeopleId, rec)
    if to == 'aside':
        rec = load_record(sender.Pid)
        rec.setdefault('docs_aside', []).append(doc)
        save_record(sender.Pid, rec)
    elif target:
        rec = load_record(target.Pid)
        r = rec['requests'].setdefault(target.Rid, {})
        r.setdefault('docs_linked', []).append(doc)
        add_history(r, 'Documents attached by the router')
        save_record(target.Pid, rec)
    return {'success': True}


def docs_note(r, ups):
    """Record what came in on the request; send it back to the committee when
    it was waiting on them. True when something changed."""
    if not ups:
        return False
    newest = max(u['t'] for u in ups)
    if newest <= (r.get('docs_in') or ''):
        return False
    fresh = [u for u in ups if u['t'] > (r.get('docs_in') or '') and u['kind'] == 'file']
    r['docs_in'] = newest
    r.setdefault('history', []).append({'t': now_str(), 'by': 0, 'what': 'Documents came in (%s)' % (
        ('%d %s' % (len(fresh), 'file' if len(fresh) == 1 else 'files')) if fresh else 'a note, no files'), 'note': ''})
    if r.get('status') == 'requester':
        r['status'] = 'committee'
    if r.get('status') in ('committee', 'requester'):
        notify(r, 'docs', [r.get('owner')])
    return True


def docs_sweep(st):
    """Bring in documents for every open case. Run on the router's and the
    committee's lists and every 15 minutes by the email tick."""
    if not int(st.get('docs_org') or 0):
        return
    records = all_records()
    want = [(pid, rid) for pid, reqs in records.items() for rid, r in reqs.items()
            if r.get('status') in ('committee', 'requester') and GUID.match(rid or '')]
    if not want:
        return
    subs = submissions(st, "AND rp.RegPeopleId IN (%s)" % ','.join("'%s'" % rid for _, rid in want), len(want))
    found = docs_by_request(st, subs)
    for x in subs:
        ups = found.get(x.Rid.upper())
        if not ups:
            continue
        rec = load_record(x.Pid)
        r = rec['requests'].get(x.Rid) or rec['requests'].get(x.Rid.upper())
        if r and docs_note(r, ups):
            save_record(x.Pid, rec)


def docs_card(st, s, r, ups, role='committee'):
    """What they sent on the documents form, newest first. Each time they sent
    it is one block, its answers in the form's order: a file with Open, a
    note as text. New until this person has seen it. Answers the request form
    marks private (Settings, Reading your form) show only to routers."""
    if not int(st.get('docs_org') or 0):
        return ''
    seen = my_team()['seen'].get('DOCS:' + s.Rid.upper()) or ''
    asked = [x for x in r.get('texts') or [] if x.get('kind') == 'docs']
    subs = {}
    for u in ups:
        subs.setdefault((u['t'], u['rid']), []).append(u)
    blocks = []
    choices = request_choices(st, [s.Pid], skip=s.Rid) if role == 'router' and subs else []
    for (t, rid), items in sorted(subs.items(), reverse=True):
        move = ''
        if role == 'router':
            # The tool placed it by family and time; a router can put it right.
            move = ('<details class="muted"><summary>%s</summary><div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">'
                    '<select class="docto">%s</select><button type="button" class="btn ghost docgo">Move</button>'
                    '<button type="button" class="btn ghost docaside">Not for a request</button>%s<span class="docst"></span></div></details>') % (
                'Attached by a router. Move it?' if any(u.get('linked') for u in items) else 'Belongs to a different request?',
                docs_place_select([('', 'Move to...')] + choices),
                '<button type="button" class="btn ghost docback">Let the tool place it</button>' if any(u.get('linked') for u in items) else '')
        blocks.append('<div class="post" data-doc="%s"><div class="bvr"><b>Sent %s %s</b>%s</div><div class="qa" style="padding-left:0">%s</div>%s</div>' % (
            rid, nice_date(t), esc(t[11:16]), ' <span class="pill p-warn">New</span>' if t > seen else '', docs_rows(st, items, role), move))
    if ups:
        mark_seen('DOCS:' + s.Rid.upper(), max(u['t'] for u in ups))
    head = ('Asked for them %s by %s.' % (nice_date(asked[-1].get('t')), by_words(asked[-1].get('via') or ['text']))) if asked else ''
    body = ''.join(blocks) or '<div class="muted">None yet.%s</div>' % (
        '' if asked else ' To ask for them, open <b>Message %s</b> in Step 1 and press <b>Send them the documents link</b>.' % esc(s.Name))
    wire = ('<script>(function(){var c=document.getElementById("docs");'
            'function go(p,d){var st=p.querySelector(".docst"); st.textContent="Saving..."; d.action="docs_place"; d.doc=p.getAttribute("data-doc");'
            ' bvPost(d,function(r){if(r.success) location.reload(); else st.textContent=r.message;});}'
            'Array.prototype.forEach.call(c.querySelectorAll(".post"),function(p){'
            ' var g=p.querySelector(".docgo"); if(g) g.onclick=function(){var to=p.querySelector(".docto").value; if(!to){p.querySelector(".docst").textContent="Pick a request first.";return;} go(p,{to:to});};'
            ' var a=p.querySelector(".docaside"); if(a) a.onclick=function(){go(p,{to:"aside"});};'
            ' var b=p.querySelector(".docback"); if(b) b.onclick=function(){go(p,{to:"auto"});};});})();</script>') if role == 'router' else ''
    return ('<div class="bvc" id="docs"><div class="bvr"><h2>Documents they sent</h2><span class="muted">%s</span></div>%s'
            '<div class="muted">The files stay in TouchPoint. Open each one here, and please don\'t forward them.</div></div>%s') % (esc(head), body, wire)


def page_letter(st):
    """The person who asked reads what was sent to them, and nothing else."""
    s = find_submission(st, _param('rid'))
    if not s or not ME or int(s.Pid) != int(ME):
        return page_denied()
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid) or {}
    mine = [x for x in r.get('texts') or [] if x.get('kind') == 'secure']
    changed = False
    for x in mine:
        if not x.get('read'):
            x['read'] = now_str()
            changed = True
    if changed and PROTECTED:
        save_record(s.Pid, rec)
    body = ''.join('<div class="bvc"><div class="muted">%s</div><div>%s</div></div>' % (nice_date(x.get('t')), text_html(x.get('text')))
                   for x in reversed(mine)) or '<div class="bvc">There are no messages for you here.</div>'
    return CSS + ('<div id="bv"><main><h1>A message from %s</h1>%s<div class="bvc"><div>This page cannot take replies.%s</div></div>'
                  '</main></div>') % (esc(church_name(st)), body,
                                       (' If you have a question, please call %s.' % esc(st['text_phone'])) if st.get('text_phone') else '')


# ---------------------------------------------------------------- handoff
HANDOFF_REASONS = ['Too much on my plate right now', 'I know them personally', "I'll be away", 'Other']


def handoff_member_card(st, s, r, small=False):
    """For the member it was given to: ask the router to pass it on. At the
    top of the case, and (small) on each card of People I'm helping."""
    if not st.get('handoff_on'):
        return ''
    k = re.sub(r'[^0-9A-Za-z]', '', s.Rid)[:12]
    ho = r.get('handoff')
    wire = ('<script>(function(){var c=document.getElementById("ho%s"), st=c.querySelector(".host");'
            'function go(d){st.textContent="Sending..."; bvPost(d,function(x){if(x.success) location.reload(); else st.textContent=x.message;});}'
            'var x=c.querySelector(".hocancel"); if(x) x.onclick=function(){go({action:"handoff_cancel",rid:"%s"});};'
            'var g=c.querySelector(".hogo"); if(g) g.onclick=function(){go({action:"handoff_ask",rid:"%s",'
            'reason:c.querySelector(".horeason").value,note:c.querySelector(".honote").value});};})();</script>') % (k, s.Rid, s.Rid)
    if ho:
        return ('<div class="%s" id="ho%s"><b>You asked the router to give this to someone else</b>'
                '<div>%s%s · %s. It stays on your list until the router passes it on.</div>'
                '<div><button type="button" class="btn ghost hocancel">I will keep it after all</button> <span class="host"></span></div></div>%s') % (
            'muted' if small else 'bvc bvwarn', k, esc(ho.get('reason') or ''), (': ' + esc(ho['note'])) if ho.get('note') else '',
            days_ago(ho.get('t') or ''), wire)
    return ('<details class="%s" id="ho%s"><summary><b>Can\'t take this one?</b>%s</summary>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><select class="horeason">%s</select>'
            '<input type="text" class="honote" placeholder="Note for the router (optional)" size="34">'
            '<button type="button" class="btn ghost hogo">Ask to hand it off</button><span class="host"></span></div>'
            '<div class="muted">Anything you have done so far goes with it.</div></details>%s') % (
        'shadd flat' if small else 'bvc', k, '' if small else ' Ask the router to give it to someone else',
        ''.join('<option>%s</option>' % esc(x) for x in HANDOFF_REASONS), wire)


def handoff_router_card(st, s, r, records):
    """For the router: an asked-for handoff, or reassigning any time."""
    owner = int(r.get('owner') or 0)
    nm = names_for([owner]).get(owner, '')
    ho = r.get('handoff')
    if ho:
        return ('<div class="bvc bvwarn"><h2>%s asked to hand this off</h2><div>%s%s · %s</div>'
                '<div class="bvr" style="justify-content:flex-start"><input type="text" id="hokeepnote" placeholder="Note back to them (optional)" size="34">'
                '<button type="button" class="btn ghost" id="hokeep">Keep it with %s</button><span id="hokst"></span></div></div>%s'
                '<script>document.getElementById("hokeep").onclick=function(){var st=document.getElementById("hokst"); st.textContent="Saving...";'
                'bvPost({action:"handoff_keep",rid:"%s",note:document.getElementById("hokeepnote").value},function(r){if(r.success) location.reload(); else st.textContent=r.message;});};</script>') % (
            esc(nm), esc(ho.get('reason') or ''), (': ' + esc(ho['note'])) if ho.get('note') else '', days_ago(ho.get('t') or ''),
            esc(nm.split(' ')[0] if nm else 'them'), assign_card(st, s.Rid, records, owner), s.Rid)
    return '<details class="bvc"><summary><b>Give it to someone else</b> <span class="muted">if %s cannot get to it</span></summary>%s</details>' % (
        esc(nm or 'they'), assign_card(st, s.Rid, records, owner))


def act_handoff_ask(st, role):
    if not st.get('handoff_on'):
        raise Exception('Handoff is turned off. Talk with the router.')
    s, rec, r = owned_request(st, _param('rid'))
    reason = _param('reason')
    if reason not in HANDOFF_REASONS:
        raise Exception('Pick a reason.')
    r['handoff'] = {'t': now_str(), 'by': int(ME), 'reason': reason, 'note': _param('note')[:300]}
    add_history(r, 'Asked to hand it off: %s' % reason, esc(r['handoff']['note']))
    notify(r, 'handoff', main_routers(st), rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def act_handoff_cancel(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    if r.pop('handoff', None) is not None:
        add_history(r, 'Will keep it after all')
        save_record(s.Pid, rec)
    return {'success': True}


def act_handoff_keep(st, role):
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid) or {}
    if not r.pop('handoff', None):
        raise Exception('There is no handoff to answer. Reload the page.')
    note = _param('note')[:300]
    r['handoff_kept'] = {'t': now_str(), 'by': int(ME), 'note': note}
    add_history(r, 'Router kept it with them', esc(note))
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- tries to reach them
CONTACT_KINDS = ['Called: talked with them', 'Called: no answer', 'Called: left a voicemail', 'Texted', 'Emailed',
                 'Number not working', 'In person']
REACHED = ('Called: talked with them', 'In person')


def contacts_line(r):
    """'3 tries, last: no answer Tue' or ''."""
    cs = r.get('contacts') or []
    if not cs:
        return ''
    last = cs[-1]
    return '%d %s, last: %s %s' % (len(cs), 'try' if len(cs) == 1 else 'tries', last.get('how', '').replace('Called: ', '').lower(),
                                   days_ago(last.get('t') or ''))


def reached(r):
    return any(c.get('how') in REACHED for c in r.get('contacts') or [])


def contacts_list(r):
    cs = r.get('contacts') or []
    if not cs:
        return '<div class="muted">No tries logged yet.</div>'
    who = names_for([c.get('by') for c in cs])
    return '<table><tbody>%s</tbody></table>' % ''.join(
        '<tr><td data-l="When">%s</td><td data-l="What"><b>%s</b>%s</td><td data-l="By">%s</td></tr>' % (
            nice_date(c.get('t')) + ' ' + esc((c.get('t') or '')[11:16]), esc(c.get('how') or ''),
            (' · ' + esc(c['note'])) if c.get('note') else '', esc(who.get(int(c.get('by') or 0), ''))) for c in reversed(cs))


def contacts_card(s, r, can_log, inner=False):
    cs = r.get('contacts') or []
    tries = [c for c in cs if c.get('how') not in REACHED]
    hint = ''
    if len(tries) >= 3 and not reached(r):
        hint = ('<div class="bvc bvwarn" style="padding:8px 12px">%d tries without reaching them. Keep trying, or on the decision '
                'pick <b>Decline</b> with <b>Could not reach them</b>.</div>' % len(tries))
    form = ''
    if can_log:
        form = ('<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><select id="ckhow">%s</select>'
                '<input type="text" id="cknote" placeholder="Note (optional)" size="28"><button type="button" class="btn ghost" id="ckgo">Log it</button>'
                '<span id="ckst" class="muted"></span></div>'
                '<script>document.getElementById("ckgo").onclick=function(){var st=document.getElementById("ckst"); st.textContent="Saving...";'
                'bvPost({action:"contact_log",rid:"%s",how:document.getElementById("ckhow").value,note:document.getElementById("cknote").value},'
                'function(r){if(r.success) location.reload(); else st.textContent=r.message;});};</script>') % (
            ''.join('<option>%s</option>' % esc(k) for k in CONTACT_KINDS), s.Rid)
    if inner:
        return '<div class="sec" id="tries"><b>Tries to reach them</b>%s%s%s</div>' % (
            form, hint, contacts_list(r) if cs else '')
    return '<div class="bvc" id="tries"><div class="bvr"><h2>Tries to reach them</h2><span class="muted">%s</span></div>%s%s%s</div>' % (
        esc(contacts_line(r)), form, hint, contacts_list(r))


def call_answers(st, r):
    """What was logged on the call: answers by topic, then the call notes."""
    ans = (r.get('call') or {}).get('answers') or {}
    rows = []
    known = set()
    for t in parse_call_questions(st.get('call_questions')):
        items = ''.join('<span>%s</span><span class="a">%s</span>' % (esc(q['q']), esc(ans.get(q['q'])))
                        for q in t['qs'] if (ans.get(q['q']) or '').strip())
        known.update(q['q'] for q in t['qs'])
        if items:
            rows.append('<details open><summary>%s</summary><div class="qa">%s</div></details>' % (esc(t['name']), items))
    extra = ''.join('<span>%s</span><span class="a">%s</span>' % (esc(k), esc(v)) for k, v in sorted(ans.items()) if k not in known and (v or '').strip())
    if extra:
        rows.append('<details><summary>Earlier questions</summary><div class="qa">%s</div></details>' % extra)
    notes = ((r.get('call') or {}).get('notes') or '').strip()
    if notes:
        rows.append('<div><b>Call notes</b><div>%s</div></div>' % clean_html(notes))
    return ''.join(rows)


def act_contact_log(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    how = _param('how')
    if how not in CONTACT_KINDS:
        raise Exception('Pick what happened.')
    c = {'t': now_str(), 'by': int(ME), 'how': how, 'note': _param('note')[:300]}
    r.setdefault('contacts', []).append(c)
    add_history(r, how, esc(c['note']))
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- the call
def parse_call_questions(text):
    """[{'name': topic, 'qs': [{'q', 'yn'}]}] from the Settings text."""
    topics, cur = [], None
    for line in (text or '').split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('#'):
            cur = {'name': line.lstrip('#').strip()[:60] or 'More', 'qs': []}
            topics.append(cur)
            continue
        yn = bool(re.search(r'(?i)\[\s*yes\s*/\s*no\s*\]\s*$', line))
        q = re.sub(r'(?i)\s*\[\s*yes\s*/\s*no\s*\]\s*$', '', line)[:300]
        if cur is None:
            cur = {'name': 'Questions', 'qs': []}
            topics.append(cur)
        cur['qs'].append({'q': q, 'yn': yn})
    return [t for t in topics if t['qs']]


def call_progress(st, r):
    """(answered, total) for this request's call."""
    qs = [q['q'] for t in parse_call_questions(st.get('call_questions')) for q in t['qs']]
    ans = ((r or {}).get('call') or {}).get('answers') or {}
    return (sum(1 for q in qs if (ans.get(q) or '').strip()), len(qs))


def call_started(r):
    c = (r or {}).get('call') or {}
    return bool(c.get('saved') or (c.get('notes') or '').strip())


def call_questions_card(st):
    topics = parse_call_questions(st.get('call_questions'))
    return ('<h2>The call</h2><div class="muted">The questions a committee member walks through on the phone. One per line. '
            'Start a topic with a line beginning with <b>#</b>. End a question with <b>[yes/no]</b> to give Yes, No and '
            "Didn't say buttons. Changing a question keeps answers already given; they show under Earlier questions.</div>"
            '<textarea name="call_questions" rows="18" style="width:100%%;font-family:monospace">%s</textarea>'
            '<div class="muted">Now: %d topics, %d questions (%s).</div>') % (
        esc(st.get('call_questions') or ''), len(topics), sum(len(t['qs']) for t in topics),
        esc(', '.join(t['name'] for t in topics)))


def call_summary(st, r):
    """The call and the tries, read-only, for the router and a teammate."""
    c = r.get('call') or {}
    tries = ('<div><b>Tries to reach them</b> <span class="muted">%s</span>%s</div>' % (esc(contacts_line(r)), contacts_list(r))) if r.get('contacts') else ''
    if not call_started(r):
        return ('<div class="bvc"><h2>The call</h2><div class="muted">Not started yet.</div>%s</div>' % tries) if tries else ''
    who = names_for([c.get('by')]).get(int(c.get('by') or 0), '')
    a, n = call_progress(st, r)
    return ('<div class="bvc"><div class="bvr"><h2>The call</h2><span class="muted">%s · %d of %d answered · saved %s</span></div>%s%s</div>') % (
        esc(who), a, n, days_ago(c.get('saved') or ''), tries, call_answers(st, r))


def call_card(st, s, r):
    """Step 1 for the member it was given to: how to reach them, the call,
    each try, texting (folded) and what the call has so far, each said once."""
    a, n = call_progress(st, r)
    if call_started(r):
        msg, btn = '%d of %d questions answered, saved %s.' % (a, n, days_ago((r.get('call') or {}).get('saved') or '')), 'Pick up the call where you left off'
    elif r.get('contacts') and not reached(r):
        msg, btn = 'Not reached yet. Log each try below.', 'Call again'
    else:
        msg, btn = 'The call opens on its own page with your notes and the questions. Everything saves as you type.', 'Start the call'
    got = call_answers(st, r)
    reach = contact_card(s).replace('<div class="bvc"><h2>How to reach them</h2>', '<div>', 1)
    return ('<div class="bvc" id="contact"><div class="bvr"><h2>Step 1: Contact them</h2>%s</div>%s'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><a class="btn" href="?view=call&amp;rid=%s">%s</a><span class="muted">%s</span></div>'
            '%s%s%s</div>') % (
        '<span class="pill p-ok">Started</span>' if call_started(r) else '', reach, s.Rid, btn, esc(msg),
        contacts_card(s, r, True, True),
        texts_card(st, s, r, 'committee', True),
        ('<div class="sec"><b>What you have from the call</b>%s</div>' % got) if got else '')


def page_call(st, role):
    rid = _param('rid')
    s = find_submission(st, rid)
    rec = load_record(s.Pid) if s else None
    r = (rec or {}).get('requests', {}).get(s.Rid) if s else None
    if not s or not r or int(r.get('owner') or 0) != int(ME or 0):
        return '<div class="bvc"><b>This call is not yours to make.</b><span>Only the committee member a request was given to can make its call.</span></div>'
    if r.get('status') not in ('committee', 'requester'):
        return '<div class="bvc"><b>This request has already been decided.</b> <a href="?view=req&amp;rid=%s">Open it</a></div>' % s.Rid
    c = r.get('call') or {}
    ans = c.get('answers') or {}
    topics = parse_call_questions(st.get('call_questions'))
    f = form_answers(st, s.Rid)
    need = ', '.join(n['need'] for n in f['needs'])
    mode = c.get('mode') or ''
    secs, chips, guide = [], [], []
    known = set()
    for i, t in enumerate(topics):
        qs = []
        for q in t['qs']:
            known.add(q['q'])
            v = ans.get(q['q']) or ''
            if q['yn']:
                field = ''.join('<label class="pick"><input type="radio" name="yn%d_%d" value="%s" data-q="%s"%s> %s</label>' % (
                    i, len(qs), o, esc(q['q']), ' checked' if v == o else '', o) for o in ('Yes', 'No', "Didn't say"))
                field = '<div class="bvr" style="justify-content:flex-start">%s</div>' % field
            else:
                field = '<textarea rows="2" data-q="%s">%s</textarea>' % (esc(q['q']), esc(v))
            qs.append('<div class="cq"><b>%s</b>%s</div>' % (esc(q['q']), field))
        secs.append('<section class="ct" id="t%d"><h2>%s</h2>%s</section>' % (i, esc(t['name']), ''.join(qs)))
        chips.append('<a href="#t%d" class="chip" data-t="%d">%s</a>' % (i, i, esc(t['name'])))
        guide.append('<label class="pick"><input type="checkbox" class="cov" value="%s"%s> %s</label>' % (
            esc(t['name']), ' checked' if t['name'] in (c.get('covered') or []) else '', esc(t['name'])))
    earlier = ''.join('<span>%s</span><span class="a">%s</span>' % (esc(k), esc(v)) for k, v in sorted(ans.items()) if k not in known and (v or '').strip())
    return ('<div class="callgrid" id="cgrid">'
            '<div class="callmain">'
            '<div class="backbar"><a class="btn ghost" href="?view=req&amp;rid=%s">&larr; Back to the case</a>'
            '<span class="muted">Everything here is saved as you type.</span></div>'
            '<div class="bvr"><div><h1 style="margin:0">Call with %s · #%s</h1><div class="muted">%s · %s asked</div></div>'
            '<span class="pill p-ok" id="callst">%s</span></div>'
            '%s'
            '<details class="bvc" id="cguide"><summary><b>Guide</b> <span class="muted">topics to cover, or walk through the questions</span></summary>'
            '<div class="seg" id="cmode"><button type="button" data-m="notes">My own notes, guide beside me</button>'
            '<button type="button" data-m="walk">Walk through the questions</button></div>'
            '<div class="muted">Either way works. Switch any time; nothing typed is lost.</div>'
            '<div id="walk"><div class="bvr"><b id="cprog"></b><span class="muted">Skip any; come back to it later.</span></div>'
            '<div class="bar"><div id="cbar"></div></div><div class="chips">%s</div><div class="bvc" id="cqs">%s</div>%s</div>'
            '<div id="notesguide"><div class="muted">Tick topics off as you cover them.</div>%s</div></details>'
            '<script>(function(){var g=document.getElementById("cguide"); try{if(localStorage.getItem("bv.guideopen")=="1") g.open=true;}catch(e){}'
            'g.addEventListener("toggle",function(){try{localStorage.setItem("bv.guideopen",g.open?"1":"0");}catch(e){}});})();</script>'
            '<div class="bvc"><b>Call notes</b>%s</div>'
            '<div class="bvc"><h2>When the call is done</h2><div class="bvr" style="justify-content:flex-start">'
            '<a class="btn" href="?view=req&amp;rid=%s#dec">Go to the decision</a>'
            '<a class="btn ghost" href="?view=mine">Back to People I\'m helping</a></div>'
            '<div class="muted">Everything is saved to TouchPoint as you type. Come back any time; it opens where you stopped.</div></div>'
            '</div><aside class="callform bvc" id="cform"><h2>What they wrote on the form</h2>'
            '<div class="muted">Open the parts you want to check.</div>%s</aside></div>'
            '<script>(function(){var rid="%s", KEY="bv.call."+rid, t=null, st=document.getElementById("callst");'
            'function val(){var a={}; Array.prototype.forEach.call(document.querySelectorAll("#cqs textarea[data-q]"),function(x){a[x.getAttribute("data-q")]=x.value;});'
            ' Array.prototype.forEach.call(document.querySelectorAll("#cqs input[type=radio][data-q]"),function(x){var k=x.getAttribute("data-q"); if(!(k in a)) a[k]=""; if(x.checked) a[k]=x.value;});'
            ' return {mode:document.getElementById("cmode").getAttribute("data-on")||"notes", answers:a,'
            '  covered:Array.prototype.filter.call(document.querySelectorAll(".cov"),function(x){return x.checked;}).map(function(x){return x.value;})};}'
            'function prog(){var v=val().answers, n=0, done=0; for(var k in v){n++; if((v[k]||"").trim()) done++;}'
            ' document.getElementById("cprog").textContent=done+" of "+n+" answered"; document.getElementById("cbar").style.width=(n?Math.round(100*done/n):0)+"%%";'
            ' Array.prototype.forEach.call(document.querySelectorAll(".chip"),function(c){var sec=document.getElementById("t"+c.getAttribute("data-t")), all=true;'
            '  Array.prototype.forEach.call(sec.querySelectorAll(".cq"),function(q){var ta=q.querySelector("textarea"), r=q.querySelector("input:checked");'
            '   if(!(ta?ta.value.trim():r)) all=false;}); c.className="chip"+(all?" done":"");});}'
            'function save(){try{localStorage.setItem(KEY,JSON.stringify(val()));}catch(e){} st.textContent="Saving..."; st.className="pill p-warn"; clearTimeout(t);'
            ' t=setTimeout(function(){bvPost({action:"call_save",rid:rid,data:bvB64(JSON.stringify(val()))},function(r){'
            '  if(r.success){try{localStorage.removeItem(KEY);}catch(e){} st.textContent="Saved"; st.className="pill p-ok";}'
            '  else {st.textContent="Not saved yet: sign in again"; st.className="pill p-bad";}});},800);}'
            'function setMode(m){var c=document.getElementById("cmode"); c.setAttribute("data-on",m);'
            ' Array.prototype.forEach.call(c.querySelectorAll("button"),function(b){b.className=b.getAttribute("data-m")==m?"on":"";});'
            ' document.getElementById("walk").style.display=m=="walk"?"":"none"; document.getElementById("notesguide").style.display=m=="notes"?"":"none";'
            ' try{localStorage.setItem("bv.callmode",m);}catch(e){}}'
            'Array.prototype.forEach.call(document.querySelectorAll("#cmode button"),function(b){b.onclick=function(){setMode(b.getAttribute("data-m")); save();};});'
            'var m0=%s; try{m0=m0||localStorage.getItem("bv.callmode");}catch(e){} setMode(m0||"notes");'
            'try{var kept=JSON.parse(localStorage.getItem(KEY)||"null"); if(kept){'
            ' Array.prototype.forEach.call(document.querySelectorAll("#cqs textarea[data-q]"),function(x){var v=kept.answers[x.getAttribute("data-q")]; if(v!=null) x.value=v;});'
            ' Array.prototype.forEach.call(document.querySelectorAll("#cqs input[type=radio][data-q]"),function(x){x.checked=kept.answers[x.getAttribute("data-q")]==x.value;});'
            ' Array.prototype.forEach.call(document.querySelectorAll(".cov"),function(x){x.checked=(kept.covered||[]).indexOf(x.value)>-1;});'
            ' setMode(kept.mode||"notes"); save();}}catch(e){}'
            'function grow(x){x.style.height="auto"; x.style.height=(x.scrollHeight+4)+"px";}'
            'Array.prototype.forEach.call(document.querySelectorAll("#cqs textarea"),function(x){grow(x); x.addEventListener("input",function(){grow(x);});});'
            'document.getElementById("cqs").addEventListener("input",function(){prog();save();});'
            'document.getElementById("cqs").addEventListener("change",function(){prog();save();});'
            'document.getElementById("notesguide").addEventListener("change",save);'
            'prog();})();</script>') % (
        s.Rid, esc(s.Name), r.get('num') or '', esc(need or 'see their request'), money(f['total']) or 'amount not given',
        'Saved' if c.get('saved') else 'Saves as you type',
        contact_card(s) + contacts_card(s, r, True), ''.join(chips), ''.join(secs),
        ('<details class="bvc"><summary>Earlier questions</summary><div class="qa">%s</div></details>' % earlier) if earlier else '',
        ''.join(guide), rich_box(c.get('notes') or '', 'call_notes:%s:notes' % s.Rid, 160), s.Rid,
        form_sections(f, role == 'router'), s.Rid, safe_json(mode)) + chat_widget(s, r)


def act_call_save(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    import json
    try:
        d = json.loads(from_b64(_param('data')) or '{}')
    except Exception:
        raise Exception('The call could not be read. Reload the page.')
    c = r.get('call') or {}
    ans = dict(c.get('answers') or {})
    for k, v in list((d.get('answers') or {}).items())[:200]:
        k, v = unicode(k)[:300], unicode(v or '')[:4000]
        if v.strip():
            ans[k] = v
        else:
            ans.pop(k, None)
    first = not call_started(r)
    c.update({'answers': ans, 'mode': d.get('mode') if d.get('mode') in ('walk', 'notes') else 'notes',
              'covered': [unicode(x)[:60] for x in d.get('covered') or []][:30], 'saved': now_str(), 'by': int(ME)})
    r['call'] = c
    if first:
        add_history(r, 'Call started')
    save_record(s.Pid, rec)
    return {'success': True}


def act_call_notes(st, role):
    s, rec, r = owned_request(st, _param('rid'))
    c = r.get('call') or {}
    first = not call_started(r)
    c.update({'notes': clean_html(from_b64(_param('html'))), 'saved': now_str(), 'by': int(ME)})
    r['call'] = c
    if first:
        add_history(r, 'Call started')
    save_record(s.Pid, rec)
    return {'success': True}


# ---------------------------------------------------------------- payouts and pickup (the router)
def given_items(r):
    """[(index, item)] for each thing the committee approved."""
    return [(i, x) for i, x in enumerate((r.get('decision') or {}).get('items') or []) if x.get('help')]


def pickup_items(st, r):
    """What is (or will be) held for pickup. Before the payout: the kinds of help
    marked "picked up" in Settings. After it: what the router chose to hold."""
    pays = r.get('payouts')
    if pays:
        held = set(p.get('i') for p in pays if p.get('pickup'))
        return [(i, x) for i, x in given_items(r) if i in held]
    d = r.get('decision') or {}
    return [(i, x) for i, x in given_items(r) if item_deliv(st, d, x) == 'pickup']


def pickup_due(r):
    """'YYYY-MM-DD' agreed for pickup, or ''."""
    return (r.get('decision') or {}).get('pickup_day') or ''


def payout_card(st, s, r):
    """The router records each approved item: paid and sent (mailed or handed
    over) or held for pickup. Anything held needs a place, day and time, and
    the request then waits for pickup until the router marks it picked up."""
    d = r.get('decision') or {}
    rows = []
    for i, x in given_items(r):
        card = needs_pickup(st, x.get('how'))
        held = item_deliv(st, d, x) == 'pickup'
        rows.append(
            '<tr data-i="%d"><td data-l="Help"><label class="pick"><input type="checkbox" class="p-done"> <span><b>%s</b>%s</span></label>%s</td>'
            '<td data-l="Amount"><input type="text" class="p-amt" size="8" value="%s"></td><td data-l="How">%s</td>'
            '<td data-l="%s"><input type="text" class="p-ref" size="14" placeholder="%s"></td>'
            '<td data-l="Then"><select class="p-hold"><option value="0"%s>Mailed or handed over</option><option value="1"%s>Held for pickup</option></select></td></tr>' % (
                i, esc(x.get('need') or ''), (' · ' + esc(x['payee'])) if x.get('payee') else '',
                ('<div class="muted">%s</div>' % ' · '.join(esc(v) for v in [x.get('address'), ('Account ' + x['account']) if x.get('account') else ''] if v))
                if (x.get('address') or x.get('account')) and not held else '', esc(x.get('amount') or ''),
                esc(x.get('how') or ''), 'Card' if card else 'Check', 'Card last 4 (optional)' if card else 'Check number',
                '' if held else ' selected', ' selected' if held else ''))
    plist = list(st.get('pickup_places') or [])
    if d.get('pickup') and d['pickup'] not in plist:
        plist.append(d['pickup'])
    places = '<option value="">Pick a place</option>' + ''.join('<option%s>%s</option>' % (' selected' if p == d.get('pickup') else '', esc(p)) for p in plist)
    return ('<div class="bvc bvblue" id="pay" data-rid="%s"><h2>Confirm payouts</h2>'
            '<div class="muted">Tick each one as you write the check or set the card aside, and say what happens to it. You can change an amount if what was given differs.</div>'
            '<div class="tw"><table><thead><tr><th>Help</th><th>Amount</th><th>How</th><th>Check or card</th><th>Then</th></tr></thead><tbody>%s</tbody></table></div>'
            '<div id="payhold" class="shform"><b>Held for pickup</b>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap"><label>Where <select id="pyplace">%s</select></label>'
            '<label>Day <input type="date" id="pyday" value="%s"></label><label>Time <input type="time" id="pytime" value="%s"></label></div>'
            '<div class="muted">%s It waits in <b>Waiting for pickup</b> until you mark it picked up.</div></div>'
            '<div class="muted" id="paydone">Once you confirm, the request is done.</div>'
            '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn green" id="paygo">Confirm payouts</button><span id="payst"></span></div></div>'
            '<script>(function(){var card=document.getElementById("pay");'
            'function show(){var any=Array.prototype.some.call(card.querySelectorAll(".p-hold"),function(x){return x.value=="1";});'
            ' document.getElementById("payhold").style.display=any?"":"none"; document.getElementById("paydone").style.display=any?"none":"";}'
            'card.addEventListener("change",show); show();'
            'document.getElementById("paygo").onclick=function(){var b=this,st=document.getElementById("payst"),lines=[];'
            'Array.prototype.forEach.call(card.querySelectorAll("tbody tr"),function(tr){lines.push({i:+tr.getAttribute("data-i"),'
            ' done:tr.querySelector(".p-done").checked, amount:tr.querySelector(".p-amt").value, ref:tr.querySelector(".p-ref").value, hold:tr.querySelector(".p-hold").value=="1"});});'
            'b.disabled=true; st.textContent="Saving...";'
            'bvPost({action:"pay_confirm",rid:"%s",data:bvB64(JSON.stringify(lines)),place:document.getElementById("pyplace").value,'
            ' day:document.getElementById("pyday").value,time:document.getElementById("pytime").value},function(r){if(r.success) location.reload(); else {b.disabled=false; st.textContent=r.message;}});};})();</script>') % (
        s.Rid, ''.join(rows), places, esc(d.get('pickup_day') or ''), esc(d.get('pickup_time') or ''),
        ('The committee member set this with them on the call; change it only if it changed.' if d.get('pickup') and d.get('pickup_day')
         else 'The committee member did not set it; agree it with them and set it here.'), s.Rid)


def pickup_card(st, s, r):
    """At pickup: check what they must bring, then mark it picked up; or move the day."""
    d = r.get('decision') or {}
    items, checks = [], []
    for i, x in pickup_items(st, r):
        items.append('<li><b>%s</b> · %s%s</li>' % (esc(x.get('need') or ''), money(to_number(paid_amount(r, i, x))), (' · ' + esc(x['how'])) if x.get('how') else ''))
        for q in help_kind(st, x.get('how')).get('reqs') or []:
            if q not in checks:
                checks.append(q)
    late = pickup_due(r) and pickup_due(r) < datetime.date.today().strftime('%Y-%m-%d')
    boxes = ''.join('<label class="pick"><input type="checkbox" class="k-req" value="%s"> %s</label>' % (esc(q), esc(q)) for q in checks)
    return ('<div class="bvc %s" id="pk" data-rid="%s"><div class="bvr"><h2>Waiting for pickup</h2><span class="pill %s">%s</span></div>'
            '<div><b>%s</b>%s</div><ul>%s</ul>'
            '<h2>When they come</h2>%s'
            '<label>Picked up by <input type="text" id="pkwho" value="%s" size="24"></label>'
            '<div class="bvr" style="justify-content:flex-start"><button type="button" class="btn green" id="pkgo">Mark picked up</button><span id="pkst"></span></div>'
            '<details><summary>Change the day or time</summary><div class="bvr" style="justify-content:flex-start">'
            '<label>Day <input type="date" id="pkday" value="%s"></label><label>Time <input type="time" id="pktime" value="%s"></label>'
            '<button type="button" class="btn ghost" id="pkmove">Save new time</button></div></details></div>'
            '<script>(function(){var rid="%s",st=document.getElementById("pkst");'
            'document.getElementById("pkgo").onclick=function(){var c=Array.prototype.filter.call(document.querySelectorAll(".k-req"),function(x){return x.checked;}).map(function(x){return x.value;});'
            ' st.textContent="Saving..."; bvPost({action:"pickup_done",rid:rid,who:document.getElementById("pkwho").value,checked:c.join("\\n")},function(r){if(r.success) location.reload(); else st.textContent=r.message;});};'
            'document.getElementById("pkmove").onclick=function(){st.textContent="Saving...";'
            ' bvPost({action:"pickup_move",rid:rid,day:document.getElementById("pkday").value,time:document.getElementById("pktime").value},function(r){if(r.success) location.reload(); else st.textContent=r.message;});};})();</script>') % (
        'bvwarn' if late else 'bvblue', s.Rid, 'p-warn' if late else 'p-new', 'Missed: call them' if late else 'Set aside',
        esc(d.get('pickup') or 'Pickup place not set'), (', ' + esc(pickup_when(d))) if pickup_when(d) else '', ''.join(items),
        ('<div class="muted">Check each before handing it over:</div>' + boxes) if boxes else '<div class="muted">Nothing to check for these.</div>',
        esc(s.Name), esc(pickup_due(r)), esc(d.get('pickup_time') or ''), s.Rid)


def paid_amount(r, i, x):
    for p in r.get('payouts') or []:
        if p.get('i') == i:
            return p.get('amount')
    return x.get('amount')


def payouts_list(r):
    """What was paid, for a finished request."""
    ps = r.get('payouts') or []
    if not ps:
        return ''
    rows = ''.join('<tr><td><b>%s</b></td><td data-l="Paid">%s</td><td data-l="How">%s%s</td></tr>' % (
        esc(p.get('need') or ''), money(p.get('amount')), esc(p.get('how') or ''), (' · ' + esc(p['ref'])) if p.get('ref') else '') for p in ps)
    pk = r.get('picked') or {}
    return ('<div class="bvc"><h2>Paid out %s</h2><table><tbody>%s</tbody></table>%s</div>') % (
        money(sum(float(p.get('amount') or 0) for p in ps)), rows,
        ('<div class="muted">Picked up by %s, %s%s</div>' % (esc(pk.get('who') or ''), nice_date(pk.get('at')),
                                                         (' · checked ' + esc(', '.join(pk['checked']))) if pk.get('checked') else '')) if pk else '')


def router_request(st, s, r):
    """Load a request a router acts on, with its record."""
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    r = rec['requests'].get(s.Rid)
    if not r:
        raise Exception('Request not found.')
    return s, rec, r


def act_pay_confirm(st, role):
    import json
    s, rec, r = router_request(st, None, None)
    if r.get('status') != 'approved':
        raise Exception('This request is not waiting for payouts.')
    try:
        lines = dict((int(x.get('i')), x) for x in json.loads(from_b64(_param('data')) or '[]'))
    except Exception:
        raise Exception('The payouts could not be read. Reload the page.')
    pays = []
    for i, x in given_items(r):
        l = lines.get(i) or {}
        if not l.get('done'):
            raise Exception('Tick %s once it is paid or set aside.' % (x.get('need') or 'each line'))
        amt = to_number(l.get('amount'))
        if not amt:
            raise Exception('Type the amount for %s.' % (x.get('need') or 'each line'))
        card = needs_pickup(st, x.get('how'))
        ref = unicode(l.get('ref') or '').strip()[:40]
        if not card and not ref:
            raise Exception('Type the check number for %s.' % (x.get('need') or 'each check'))
        hold = bool(l['hold']) if 'hold' in l else card
        pays.append({'i': i, 'need': x.get('need') or '', 'how': x.get('how') or '', 'amount': amt, 'ref': ref,
                     'pickup': hold, 'by': int(ME), 'at': now_str()})
    held = [p for p in pays if p['pickup']]
    if held:
        place, day, tm = _param('place').strip()[:80], _param('day'), _param('time')
        d = r.get('decision') or {}
        place, day, tm = place or d.get('pickup') or '', day or d.get('pickup_day') or '', tm or d.get('pickup_time') or ''
        if not place:
            raise Exception('Pick where it is held for pickup.')
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', day) or not re.match(r'^\d{2}:\d{2}$', tm):
            raise Exception('Add the day and time they will pick it up.')
        d['pickup'], d['pickup_day'], d['pickup_time'] = place, day, tm
        r['decision'] = d
    r['payouts'] = pays
    total = sum(p['amount'] for p in pays)
    paid = [{'need': p['need'], 'amount': p['amount'], 'how': p['how'], 'ref': p['ref'], 'held': p['pickup'],
             'card': needs_pickup(st, p['how'])} for p in pays]
    if held:
        r['status'] = 'pickup'
        add_history(r, 'Paid out %s; held for pickup at %s, %s' % (money(total), r['decision']['pickup'], pickup_when(r['decision'])), '', paid)
        notify(r, 'paid', [r.get('owner')], 'payouts are confirmed; it is waiting for pickup', rid=s.Rid)
    else:
        r['status'] = 'done'
        r['closed'] = now_str()
        add_history(r, 'Paid out %s' % money(total), '', paid)
        notify(r, 'paid', [r.get('owner')], 'payouts are confirmed', rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def act_pickup_done(st, role):
    s, rec, r = router_request(st, None, None)
    if r.get('status') != 'pickup':
        raise Exception('This request is not waiting for pickup.')
    need = []
    for i, x in pickup_items(st, r):
        for q in help_kind(st, x.get('how')).get('reqs') or []:
            if q not in need:
                need.append(q)
    got = [x.strip() for x in _param('checked').split('\n') if x.strip()]
    missing = [q for q in need if q not in got]
    if missing:
        raise Exception('Check %s before handing it over.' % ', '.join(missing))
    who = _param('who')[:80]
    if not who:
        raise Exception('Type who picked it up.')
    r['picked'] = {'who': who, 'at': now_str(), 'by': int(ME), 'checked': got}
    r['status'] = 'done'
    r['closed'] = now_str()
    add_history(r, 'Picked up by %s' % who)
    notify(r, 'paid', [r.get('owner')], 'it was picked up', rid=s.Rid)
    save_record(s.Pid, rec)
    return {'success': True}


def act_pickup_move(st, role):
    s, rec, r = router_request(st, None, None)
    if r.get('status') != 'pickup':
        raise Exception('This request is not waiting for pickup.')
    day, tm = _param('day'), _param('time')
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', day) or not re.match(r'^\d{2}:\d{2}$', tm):
        raise Exception('Pick a day and a time.')
    d = r.get('decision') or {}
    d['pickup_day'], d['pickup_time'] = day, tm
    r['decision'] = d
    add_history(r, 'Pickup moved to %s' % pickup_when(d))
    save_record(s.Pid, rec)
    return {'success': True}


def pickups_section(st, items):
    """Router home: every pickup, soonest first, missed ones flagged."""
    if not items:
        return ''
    today = datetime.date.today().strftime('%Y-%m-%d')
    items = sorted(items, key=lambda x: pickup_due(x[1]) or '9999')
    trs = []
    for s, r in items:
        d = r.get('decision') or {}
        what = ', '.join('%s %s' % (esc(x.get('how') or ''), money(to_number(paid_amount(r, i, x)))) for i, x in pickup_items(st, r))
        late = pickup_due(r) and pickup_due(r) < today
        trs.append('<tr class="%s"><td><a href="?view=req&amp;rid=%s"><b>#%s %s</b></a></td><td data-l="When">%s%s</td>'
                   '<td data-l="Where">%s</td><td data-l="What">%s</td></tr>' % (
                       'hi' if late else '', s.Rid, r.get('num') or '', esc(s.Name), esc(pickup_when(d) or 'not set'),
                       ' <span class="pill p-warn">Missed</span>' if late else (' <span class="pill p-new">Today</span>' if pickup_due(r) == today else ''),
                       esc(d.get('pickup') or ''), what))
    return ('<h2>Pickups (%d)</h2><div class="bvc"><table><thead><tr><th>Who</th><th>When</th><th>Where</th><th>What</th></tr></thead>'
            '<tbody>%s</tbody></table><div class="muted">Soonest first. Missed ones are highlighted: give them a call to set a new time.</div></div>') % (
        len(items), ''.join(trs))


# ---------------------------------------------------------------- emails
# TouchPoint's ScheduledTasks runs every 15 minutes and calls this script with
# Data.bvtick = 'tick' (and the slot, since CallScript drops ScheduledTime).
# Each tick: secure message notices for messages 10+ minutes old and still
# unread, and once a day at the set time the daily emails. Nothing is sent in
# Testing mode. Emails carry request numbers and how long things have waited,
# never names, needs, amounts or message text: everything else is one click
# away in the tool. The log (Special Content, readable by anyone) holds counts
# and times only.
EMAIL_LOG = 'TPxi_BenevolenceEmailLog'
SCHED_START = '# >>> Benevolence emails (TPxi_Benevolence) start (managed by the tool, do not edit) >>>'
SCHED_END = '# <<< Benevolence emails end <<<'


def sched_read():
    try:
        return model.PythonContent('ScheduledTasks') or ''
    except Exception:
        return ''


def sched_remove_block(text):
    if SCHED_START not in text:
        return text
    head, rest = text.split(SCHED_START, 1)
    tail = rest.split(SCHED_END, 1)[1] if SCHED_END in rest else ''
    return head.rstrip('\n') + '\n' + tail.lstrip('\n')


def sched_block(script_name):
    return '\n'.join([
        SCHED_START,
        'try:',
        "    Data.bvtick = 'tick'",
        '    Data.bvslot = model.ScheduledTime or ""',
        "    model.CallScript('%s')" % script_name.replace("'", ''),
        'except Exception as _bv_e:',
        "    print 'Benevolence emails: ' + str(_bv_e)",
        "Data.bvtick = ''",
        SCHED_END])


def sched_state():
    text = sched_read()
    if SCHED_START not in text:
        return 'missing'
    return 'ok' if sched_block(script_name()) in text else 'old'


def script_name():
    try:
        return model.ScriptName or 'TPxi_Benevolence'
    except Exception:
        return 'TPxi_Benevolence'


def sched_install():
    text = sched_remove_block(sched_read())
    model.WriteContentPython('ScheduledTasks', text.rstrip('\n') + '\n\n' + sched_block(script_name()) + '\n')
    return sched_state() == 'ok'


def sched_uninstall():
    text = sched_read()
    if SCHED_START in text:
        model.WriteContentPython('ScheduledTasks', sched_remove_block(text))
    return SCHED_START not in sched_read()


def email_log(line):
    try:
        old = (model.TextContent(EMAIL_LOG) or '').splitlines()[-200:]
        model.WriteContentText(EMAIL_LOG, '\n'.join(old + [now_str()[:16] + '  ' + line]), '')
    except Exception:
        pass


def team_of(pid):
    d = load_json(model.ExtraValueText(int(pid), TEAM_FIELD) or '', {}) if pid else {}
    d.setdefault('posts', [])
    d.setdefault('seen', {})
    return d


def site_link(qs):
    try:
        host = (model.CmsHost or '').rstrip('/')
    except Exception:
        host = ''
    return '%s/PyScriptForm/%s?%s' % (host, script_name(), qs)


def days_waiting(ts):
    try:
        d = (datetime.datetime.now() - datetime.datetime.strptime((ts or '')[:10], '%Y-%m-%d')).days
    except Exception:
        return ''
    return 'today' if d <= 0 else ('1 day' if d == 1 else '%d days' % d)


def email_html(greeting, sections, button_qs, button_text):
    """A plain, friendly email: a greeting, short lists, one button."""
    body = ''.join('<h3 style="margin:18px 0 6px;font-size:16px;color:#1C2433">%s</h3><ul style="margin:0;padding-left:20px">%s</ul>' % (
        esc(title), ''.join('<li style="margin:4px 0">%s</li>' % esc(x) for x in items)) for title, items in sections if items)
    return ('<div style="font-family:Verdana,Arial,sans-serif;font-size:15px;color:#1C2433;max-width:560px">'
            '<p>%s</p>%s<p style="margin:22px 0"><a href="%s" style="background:#1F4E8C;color:#fff;text-decoration:none;'
            'font-weight:bold;padding:10px 18px;border-radius:8px;display:inline-block">%s</a></p>'
            '<p style="color:#6B7280;font-size:13px">Names and details stay in the Benevolence tool, so this email only '
            'shows request numbers.</p></div>') % (esc(greeting), body, esc(site_link(button_qs)), esc(button_text))


def daily_for(st, pid, first, records, by_status, today):
    """(subject, html) of today's email for one person, or None."""
    rinfo = dict((int(r['pid']), bool(r.get('main'))) for r in st.get('routers') or [])
    seen = team_of(pid)['seen']
    names = {}
    sections = []

    def unread(rid, r):
        s_ = seen.get(rid.upper()) or ''
        return [m for m in r.get('messages') or [] if int(m.get('by') or 0) != pid and (m.get('t') or '') > s_ and msg_visible(r, m, pid)]
    open_ = by_status.get('committee', []) + by_status.get('requester', [])
    if pid in rinfo and st.get('email_router', True):
        main = rinfo[pid]
        backup_days = int(st.get('email_backup_days') or 3)
        cut = (datetime.date.today() - datetime.timedelta(days=backup_days)).strftime('%Y-%m-%d')
        items = []
        new = [s for s, r in by_status.get('new', []) if main or s.T[:10] <= cut]
        if new:
            oldest = min(s.T for s in new)
            items.append('%d new %s to give out (the oldest has waited %s)' % (
                len(new), 'request' if len(new) == 1 else 'requests', days_waiting(oldest)) if len(new) > 1
                else '1 new request to give out (waiting %s)' % days_waiting(oldest))
        for s, r in by_status.get('approved', []):
            sent = (r.get('decision') or {}).get('sent') or ''
            if main or sent[:10] <= cut:
                items.append('#%s approved: confirm payouts (waiting %s)' % (r.get('num'), days_waiting(sent)))
        for s, r in open_:
            ho = r.get('handoff')
            if ho and (main or (ho.get('t') or '')[:10] <= cut):
                items.append('#%s: a committee member asked to hand it off (%s ago)' % (r.get('num'), days_waiting(ho.get('t'))))
            elif main and unread(s.Rid, r):
                items.append('#%s: a new message for you' % r.get('num'))
        for s, r in by_status.get('pickup', []):
            due = pickup_due(r)
            if due and due <= today:
                items.append('#%s: pickup %s' % (r.get('num'), 'today' if due == today else 'missed (%s ago), give them a call' % days_waiting(due)))
        stuck = []
        if main:
            week = (datetime.date.today() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
            owners = names_for([int(r.get('owner') or 0) for s, r in open_])
            for s, r in open_:
                lt = last_touch(r, s)
                if lt[:10] <= week:
                    stuck.append('#%s with %s: no change in %s' % (r.get('num'), owners.get(int(r.get('owner') or 0), 'a member'), days_waiting(lt)))
        if items or stuck:
            sections.append(('Waiting on you' if main else 'Waiting %d days or more (you are a backup router)' % backup_days, items))
            sections.append(('No change in a week', stuck))
    if pid in [m['pid'] for m in committee_members(st)] and st.get('email_committee', True):
        mine = []
        for s, r in open_:
            if int(r.get('owner') or 0) == pid:
                label = stage_of(r)[0].replace('Your turn: ', '')
                mine.append('#%s: %s (given to you %s ago)' % (r.get('num'), label[:1].lower() + label[1:],
                                                            days_waiting((r.get('assigned') or [{}])[-1].get('t'))))
                if unread(s.Rid, r):
                    mine.append('#%s: a new message about it' % r.get('num'))
            elif pid in shared_pids(r):
                aff = r.get('affirm') or {}
                seen = team_of(pid)['seen'].get(s.Rid.upper()) or ''
                if int(aff.get('pid') or 0) == pid and not aff.get('note'):
                    mine.append('#%s: a teammate asked you to affirm their decision' % r.get('num'))
                elif any(int(x.get('pid') or 0) == pid and not x.get('ended') and (x.get('t') or '') > seen for x in r.get('shares') or []):
                    mine.append('#%s: a teammate brought you in on it' % r.get('num'))
        sections.append(('When you have a moment', mine))
    sections = [x for x in sections if x[1]]
    if not sections:
        return None
    n = sum(len(x[1]) for x in sections)
    subject = 'Benevolence: %d %s for you today' % (n, 'thing' if n == 1 else 'things')
    who = 'router' if pid in rinfo else 'committee'
    return subject, email_html('Hi %s,' % first, sections, 'view=%s' % ('home' if who == 'router' else 'mine'),
                               'Open Benevolence')


def secure_email(pid, first, senders, num, qs):
    """The "You have a secure message" email (never the message itself)."""
    subject = ('Secure message from %s about request #%s' % (senders, num)) if num else 'Secure message from %s' % senders
    html = ('<div style="font-family:Verdana,Arial,sans-serif;font-size:15px;color:#1C2433;max-width:560px">'
            '<p>Hi %s,</p><p>You have a secure message from <b>%s</b>%s.</p>'
            '<p style="margin:22px 0"><a href="%s" style="background:#1F4E8C;color:#fff;text-decoration:none;font-weight:bold;'
            'padding:10px 18px;border-radius:8px;display:inline-block">Read it</a></p>'
            '<p style="color:#6B7280;font-size:13px">For privacy the message is not in this email. Sign in to read it.</p></div>') % (
        esc(first), esc(senders), (' about request #%s' % esc(unicode(num))) if num else '', esc(site_link(qs)))
    return subject, html


def secure_now(st, to, key, earlier, qs, num):
    """Right-away secure message notices (Settings, Emails): at once to each
    person who is not in that chat or the tool just now, at most one per chat
    every 30 minutes. earlier: this chat's earlier messages (for the limit).
    Returns who was emailed, to mark on the message."""
    if not (emails_live(st) and st.get('email_secure', True) and st.get('email_secure_when') == 'now'):
        return []
    now = datetime.datetime.now()
    recent = (now - datetime.timedelta(seconds=120)).strftime(TS)
    limit = (now - datetime.timedelta(minutes=30)).strftime(TS)
    me = names_for([int(ME)]).get(int(ME), 'A teammate') if ME else 'A teammate'
    names = names_for(list(to))
    sent = []
    for pid in to:
        here = team_of(pid).get('here') or {}
        if (here.get(key) or '') >= recent or (here.get('*') or '') >= recent:
            continue        # in the chat or the tool right now: they will see it
        if any(pid in (m.get('mailed') or []) and (m.get('t') or '') >= limit for m in earlier):
            continue        # already emailed about this chat in the last 30 minutes
        subject, html = secure_email(pid, (names.get(pid) or '').split(' ')[0], me, num, qs)
        try:
            send_email(st, pid, subject, html)
            sent.append(pid)
        except Exception as ex:
            email_log('secure notice failed: %s' % unicode(ex)[:120])
    return sent


def secure_notices(st, people, since, upto):
    """[(pid, subject, html)] for private and request messages that reached
    upto (10 minutes ago) since the last run and are still unread."""
    out = []
    names = names_for(list(people))
    posts = team_posts()
    records = all_records()
    for pid in people:
        mine = team_of(pid)
        seen = mine['seen']
        # In the tool after it was sent (the chat open, or the unread badge in
        # front of them): no email. Only people who have not been back get one.
        back = (mine.get('here') or {}).get('*') or ''
        froms = {}
        for p in posts:
            t = p.get('t') or ''
            if p.get('kind') != 'dm' or p['by'] == pid or pid not in chat_people(p.get('chat')) or not (since < t <= upto):
                continue
            if t > (seen.get(p['chat']) or '') and back < t and pid not in (p.get('mailed') or []):
                froms.setdefault(p['chat'], set()).add(p['by'])
        for reqs in records.values():
            for rid, r in reqs.items():
                if r.get('status') not in ('committee', 'requester', 'approved', 'pickup'):
                    continue
                inv = int(r.get('owner') or 0) == pid or pid in [int(x) for x in r.get('shared') or []] or pid in [int(x['pid']) for x in st.get('routers') or []]
                if not inv:
                    continue
                for m in r.get('messages') or []:
                    t = m.get('t') or ''
                    if (int(m.get('by') or 0) != pid and since < t <= upto and t > (seen.get(rid.upper()) or '') and back < t
                            and pid not in (m.get('mailed') or []) and msg_visible(r, m, pid)):
                        froms.setdefault('req:%s:%s' % (rid, r.get('num')), set()).add(int(m['by']))
        for key, who in froms.items():
            senders = ', '.join(sorted(names_for(list(who)).values()))
            if key.startswith('req:'):
                _, rid, num = key.split(':')
                qs = 'view=req&rid=%s#rmsg' % rid
            else:
                num, qs = None, 'view=team&c=%s' % key
            subject, html = secure_email(pid, (names.get(pid) or '').split(' ')[0], senders, num, qs)
            out.append((pid, subject, html))
    return out


def event_notices(st, since, upto):
    """[(pid, subject, html)]: one email per person for everything that
    happened since the last run and involves them, skipping what they have
    already opened since."""
    on = set(k for k, _, _, _ in NOTICES if st.get('notify_' + k, k != 'paid'))
    lines = {}
    team = {}
    for reqs in all_records().values():
        for rid, r in reqs.items():
            for e in r.get('events') or []:
                t = e.get('t') or ''
                # [since, upto): each event lands in exactly one run, even one in the same second as a run
                if e.get('k') not in on or not (since <= t < upto):
                    continue
                for pid in e.get('to') or []:
                    if pid in (e.get('sent') or []):
                        continue          # already sent at the click
                    if pid not in team:
                        team[pid] = team_of(pid).get('here') or {}
                    if (team[pid].get('open:' + rid.upper()) or '') >= t:
                        continue          # they opened it since: nothing to tell them
                    lines.setdefault(pid, []).append((t, rid, r.get('num'), e))
    if not lines:
        return []
    names = names_for(list(set(int(e['by']) for v in lines.values() for _, _, _, e in v if e.get('by'))) + list(lines.keys()))
    out = []
    for pid, items in lines.items():
        items.sort()
        text = [notice_text(e['k'], num, names.get(int(e.get('by') or 0), ''), e.get('w')) for t, rid, num, e in items]
        subject = ('Benevolence: ' + text[0]) if len(text) == 1 else 'Benevolence: %d updates for you' % len(text)
        first = (names.get(pid) or '').split(' ')[0]
        qs = ('view=req&rid=%s' % items[0][1]) if len(items) == 1 else ('view=home' if pid in router_ids() else 'view=mine')
        out.append((pid, subject, email_html('Hi %s,' % first, [('Just now', text)], qs, 'Open it' if len(items) == 1 else 'Open Benevolence')))
    return out


def with_tagline(st, html):
    """The church's tag line (Settings, Emails) at the foot of every email,
    e.g. that the sending mailbox is not monitored."""
    t = (st.get('email_tagline') or '').strip()
    if not t or not html.endswith('</div>'):
        return html
    return html[:-6] + ('<p style="color:#6B7280;font-size:13px;border-top:1px solid #E6E1D8;padding-top:10px">%s</p></div>' % text_html(t))


def send_email(st, pid, subject, html):
    rs = [int(r['pid']) for r in st.get('routers') or [] if r.get('main')] or [int(pid)]
    from_email = st.get('email_from') or model.Setting('AdminMail', '')
    model.Email(str(int(pid)), rs[0], from_email, st.get('email_from_name') or 'Benevolence', subject, with_tagline(st, html))


def church_now(server=None):
    """The church's local date and time for this run. TouchPoint's runs are
    never exactly on time, and the server clock may be in another time zone,
    but the ScheduledTasks slot ('0715') is the church's local time."""
    server = server or datetime.datetime.now()
    raw = re.sub(r'\D', '', unicode(_param('bvslot') or ''))
    if len(raw) != 4:
        return server
    slot = int(raw[:2]) * 60 + int(raw[2:])
    here = server.hour * 60 + server.minute
    diff = (slot - here) % 1440
    if diff > 720:
        diff -= 1440                       # e.g. slot 23:45, server 00:02 next day: -17 minutes, not +23 hours
    # Moving the server clock by exactly that gap lands on the slot itself, on
    # the right church day: the time zone difference and the few minutes
    # early or late together. So a run never lands on the wrong side of the
    # send time, or of midnight.
    return (server + datetime.timedelta(minutes=diff)).replace(second=0, microsecond=0)


def email_tick(st):
    """Run by ScheduledTasks every 15 minutes."""
    now = datetime.datetime.now()
    if (st.get('email_last_tick') or '')[:13] != now_str()[:13]:
        st['email_last_tick'] = now_str()
        save_settings(st)
    try:
        docs_sweep(st)
    except Exception:
        pass
    if not st.get('email_on') or st.get('mode') != 'live':
        return
    people = set(team_people(st).keys())
    # Secure message notices: messages that turned 10 minutes old since the last run.
    upto = (now - datetime.timedelta(minutes=10)).strftime(TS)
    since = st.get('email_msg_cut') or upto
    if st.get('email_secure', True) and since < upto:
        sent = 0
        for pid, subject, html in secure_notices(st, people, since, upto):
            try:
                send_email(st, pid, subject, html)
                sent += 1
            except Exception as e:
                email_log('secure notice failed: %s' % unicode(e)[:120])
        if sent:
            email_log('secure notices sent: %d' % sent)
    st['email_msg_cut'] = upto
    # Right-away notices: whatever happened since the last run, one email each.
    ev_upto = now_str()
    ev_since = st.get('email_ev_cut') or ev_upto
    if ev_since <= ev_upto:
        sent = 0
        for pid, subject, html in event_notices(st, ev_since, ev_upto):
            try:
                send_email(st, pid, subject, html)
                sent += 1
            except Exception as e:
                email_log('notice failed: %s' % unicode(e)[:120])
        if sent:
            email_log('right-away notices sent: %d' % sent)
    st['email_ev_cut'] = ev_upto
    # The daily emails: on the first run at or after the set time, once per
    # church day. A late or skipped run just means a later run sends it; an
    # early one waits for the next.
    local = church_now(now)
    today = local.strftime('%Y-%m-%d')
    if st.get('email_last_daily') == today:
        save_settings(st)
        return
    if st.get('email_days', 'weekdays') == 'weekdays' and local.weekday() >= 5:
        save_settings(st)
        return
    try:
        hh, mm = [int(x) for x in (st.get('email_time') or '07:00').split(':')[:2]]
    except Exception:
        hh, mm = 7, 0
    if (local.hour, local.minute) < (hh, mm):
        save_settings(st)
        return
    st['email_last_daily'] = today
    save_settings(st)
    records = all_records()
    by_status = {}
    for s in current_subs(st, records, 300):
        r = (records.get(s.Pid) or {}).get(s.Rid)
        by_status.setdefault((r or {}).get('status') or 'new', []).append((s, r))
    names = team_people(st)
    sent = skipped = 0
    for pid in sorted(people):
        try:
            e = daily_for(st, pid, (names.get(pid) or '').split(' ')[0], records, by_status, today)
            if e:
                send_email(st, pid, e[0], e[1])
                sent += 1
            else:
                skipped += 1
        except Exception as ex:
            email_log('daily email failed for one person: %s' % unicode(ex)[:120])
    email_log('daily emails: %d sent, %d had nothing waiting' % (sent, skipped))


def page_texts_section(st):
    groups = list(q.QuerySql("SELECT Id, Name FROM SMSGroups WHERE ISNULL(IsDeleted, 0) = 0 ORDER BY Name"))
    cur = int(st.get('text_group') or 0)
    opts = '<option value="0">Off: no texts</option>' + ''.join('<option value="%d"%s>%s</option>' % (
        g.Id, ' selected' if g.Id == cur else '', esc(g.Name)) for g in groups)
    return ('<div class="bvc" id="setf" data-sec="texts" data-reload="1"><h2>Texting them</h2>'
            '<div class="muted">The case worker or router can text the person who asked: a secure message, or the link to the documents form. '
            'Texts are send only. The text itself never holds names, needs or amounts, only a link; what was said stays in Benevolence, '
            'where it is kept with the case.</div>'
            '<label>Send from texting group <select name="text_group">%s</select> '
            '<span class="muted">Make a group just for this (e.g. Benevolence) in TouchPoint\'s texting; replies to it are not watched.</span></label>'
            '<label>Church name in the text <input type="text" name="text_church" value="%s" size="24" placeholder="%s"></label>'
            '<label>Phone for questions <input type="text" name="text_phone" value="%s" size="16"> '
            '<span class="muted">shown on the secure message page, since they cannot reply</span></label>'
            '<label style="display:block">Secure message text <textarea name="text_secure" rows="2" style="width:100%%">%s</textarea></label>'
            '<label style="display:block">Documents text <textarea name="text_docs" rows="2" style="width:100%%">%s</textarea></label>'
            '<div class="muted">{church} becomes the church name and {link} the link. Keep them short.</div>'
            '<div class="bvc bvwarn"><label class="pick"><input type="checkbox" name="text_autosignin"%s> '
            '<span><b>Sign them in automatically from the link</b></span></label>'
            '<div>Off (safer): the link opens the page and TouchPoint asks them to sign in. '
            'On: the link signs them straight in, but it works for 7 days, more than once, into their whole TouchPoint account, '
            'and TouchPoint keeps every text in its texting history, where anyone who can read texts could use it. '
            'Either way, they need a TouchPoint login to read a secure message.</div></div>') % (
        opts, esc(st.get('text_church') or ''), esc(church_name(dict(st, text_church=''))), esc(st.get('text_phone') or ''),
        esc(st.get('text_secure') or DEFAULT_SETTINGS['text_secure']), esc(st.get('text_docs') or DEFAULT_SETTINGS['text_docs']),
        ' checked' if st.get('text_autosignin') else '')


def when_select(name, cur, opts):
    return '<select name="%s" style="width:auto;margin-left:6px">%s</select>' % (
        name, ''.join('<option value="%s"%s>%s</option>' % (v, ' selected' if v == cur else '', esc(l)) for v, l in opts))


NOTICE_WHEN = [('run', 'At the next 15-minute run'), ('now', 'Right away')]


def notices_block(st, chk):
    """Right-away notices: a switch per event, saying who gets it, and when."""
    when = st.get('notify_when') or {}
    return ('<div class="shform" style="margin-top:10px"><b>Right-away notices</b>'
            '<div class="muted"><b>At the next 15-minute run</b> sends one email per person listing everything new for them, skipping '
            'what they already opened. <b>Right away</b> sends it the moment it happens, one email each. The daily list above still '
            'catches anything left waiting.</div>%s</div>') % ''.join(
        '<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">%s%s</div>' % (
            chk('notify_' + k, '%s <span class="muted">(goes to %s)</span>' % (esc(w), esc(who)), k != 'paid'),
            when_select('nw_' + k, when.get(k) or 'run', NOTICE_WHEN)) for k, w, who, _ in NOTICES)


def who_gets_what(st):
    """Every email the tool sends, who gets it and when: for Settings and How it works."""
    when = st.get('notify_when') or {}
    rows = [('The daily list', 'Main routers: what is waiting on them. Committee members: only when something is theirs.', 'Once a day at the set time'),
            ('Secure message notices', 'Whoever a team or private message went to',
             'Right away if they are not in the tool (at most one every 30 minutes per chat)' if st.get('email_secure_when') == 'now'
             else 'About 10 to 25 minutes later, only if still unread and they have not been back')]
    rows += [(w, who[:1].upper() + who[1:], 'Right away' if when.get(k) == 'now' else 'At the next 15-minute run')
             for k, w, who, _ in NOTICES if st.get('notify_' + k, k != 'paid')]
    return ('<table class="cmp"><thead><tr><th>When</th><th>Who gets it</th><th>Sent</th></tr></thead><tbody>%s</tbody></table>') % ''.join(
        '<tr><td data-l="When">%s</td><td data-l="Who">%s</td><td data-l="Sent">%s</td></tr>' % (esc(a), esc(b), esc(c)) for a, b, c in rows)


def page_emails_section(st):
    """Settings, Emails: on/off, when, who gets what, and a preview."""
    state = sched_state()
    on = bool(st.get('email_on'))
    status = ('<span class="pill p-ok">Running</span> ScheduledTasks reached it %s.' % days_ago(st.get('email_last_tick')) if on and state == 'ok' and st.get('email_last_tick')
              else '<span class="pill p-warn">Waiting</span> ScheduledTasks has not reached it yet; give it 15 minutes.' if on and state == 'ok'
              else '<span class="pill p-bad">Out of date</span> Press Save to update its line in ScheduledTasks.' if on and state == 'old'
              else '<span class="pill p-bad">Not in ScheduledTasks</span> Press Save to add it.' if on
              else '<span class="pill p-grey">Off</span>')
    if st.get('mode') != 'live':
        status += ' <b>Testing mode: nothing is sent until you go live.</b> Use the preview below.'
    people = team_people(st)
    opts = ''.join('<option value="%d"%s>%s</option>' % (p, ' selected' if p == int(ME or 0) else '', esc(n)) for p, n in sorted(people.items(), key=lambda x: x[1]))
    chk = lambda k, label, d=True: '<label class="pick"><input type="checkbox" name="%s"%s> %s</label>' % (k, ' checked' if st.get(k, d) else '', label)
    log = (model.TextContent(EMAIL_LOG) or '').splitlines()[-6:]
    return ('<div class="bvc" id="setf" data-sec="emails" data-reload="1"><h2>Emails</h2><div>%s</div>'
            '%s'
            '<label>Send the daily emails at <input type="time" name="email_time" value="%s"></label>'
            '<label>On <select name="email_days"><option value="weekdays"%s>Weekdays</option><option value="all"%s>Every day</option></select></label>'
            '<div><b>Who gets what</b></div>%s%s'
            '<label>Backup routers hear about anything waiting <input type="text" name="email_backup_days" value="%s" size="3"> days or more</label>%s'
            '<label>From name <input type="text" name="email_from_name" value="%s" size="20"></label>'
            '<label>From email <input type="text" name="email_from" value="%s" size="30" placeholder="%s"></label>'
            '<label style="display:block">Tag line at the bottom of every email <span class="muted">(for example, that this mailbox is not '
            'monitored; leave empty for none)</span><textarea name="email_tagline" rows="2" style="width:100%%">%s</textarea></label>'
            '<div class="muted">Emails show request numbers and how long things have waited, never names, needs, amounts or message text.</div>'
            '<div class="bvr" style="justify-content:flex-start"><button class="btn" id="savest">Save</button><span id="savestst"></span></div></div>'
            '<div class="bvc"><h2>Who gets notified, and when</h2>%s</div>'
            '<div class="bvc"><h2>Preview</h2><div class="bvr" style="justify-content:flex-start"><label>What would <select id="pvwho">%s</select> get today?</label>'
            '<button type="button" class="btn ghost" id="pvgo">Show it</button><button type="button" class="btn ghost" id="pvme">Email this preview to me</button>'
            '<span id="pvst" class="muted"></span></div><div id="pvout"></div></div>'
            '%s'
            '<script>(function(){var who=document.getElementById("pvwho"),st=document.getElementById("pvst");'
            'function go(send){st.textContent=send?"Sending...":"Loading..."; bvPost({action:"email_preview",who:who.value,send:send?"1":""},function(r){'
            ' if(!r.success){st.textContent=r.message;return;} st.textContent=r.message||"";'
            ' document.getElementById("pvout").innerHTML=r.subject?"<div class=\\"bvc\\" style=\\"background:#fff\\"><b>"+r.subject_html+"</b><hr>"+r.html+"</div>":"<div class=\\"muted\\">Nothing for them today, so no email.</div>";});}'
            'document.getElementById("pvgo").onclick=function(){go(false);}; document.getElementById("pvme").onclick=function(){go(true);};})();</script>') % (
        status,
        chk('email_on', '<b>Send emails</b> (adds a line to TouchPoint\'s ScheduledTasks; turning it off removes it)', False),
        esc(st.get('email_time') or '07:00'), ' selected' if st.get('email_days', 'weekdays') == 'weekdays' else '',
        ' selected' if st.get('email_days') == 'all' else '',
        chk('email_router', 'Main routers: a daily list of what is waiting on them and anything with no change in a week'),
        chk('email_committee', 'Committee members: a gentle daily note, only when something is theirs'),
        esc(unicode(st.get('email_backup_days') or 3)),
        '<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">%s%s</div>' % (
            chk('email_secure', 'Secure message notices: "You have a secure message from ..." with a link'),
            when_select('email_secure_when', st.get('email_secure_when') or 'later',
                        [('later', 'About 10 to 25 minutes later, if still unread'), ('now', 'Right away, if they are not in the tool')]))
        + notices_block(st, chk),
        esc(st.get('email_from_name') or 'Benevolence'), esc(st.get('email_from') or ''), esc(model.Setting('AdminMail', '') or 'church admin address'),
        esc(st.get('email_tagline') or ''),
        who_gets_what(st), opts,
        ('<div class="bvc"><h2>Recent sends</h2><pre style="white-space:pre-wrap;margin:0">%s</pre></div>' % esc('\n'.join(log))) if log else '')


def act_email_preview(st):
    pid = int(_param('who', '0') or 0)
    people = team_people(st)
    if pid not in people:
        raise Exception('Pick someone on the team.')
    records = all_records()
    by_status = {}
    for s in current_subs(st, records, 300):
        r = (records.get(s.Pid) or {}).get(s.Rid)
        by_status.setdefault((r or {}).get('status') or 'new', []).append((s, r))
    e = daily_for(st, pid, people[pid].split(' ')[0], records, by_status, datetime.date.today().strftime('%Y-%m-%d'))
    if not e:
        return {'success': True, 'subject': '', 'html': ''}
    msg = ''
    if _param('send') == '1':
        send_email(st, int(ME), '[Preview for %s] %s' % (people[pid], e[0]), e[1])
        msg = 'Sent to you.'
    return {'success': True, 'subject': e[0], 'subject_html': esc(e[0]), 'html': with_tagline(st, e[1]), 'message': msg}


# ---------------------------------------------------------------- clearing test data (admins)
# Only what this tool saved is removed (the Benevolence extra value). The form
# submissions, their answers and files are TouchPoint's and are never touched.
def act_set_mode(st):
    want = _param('mode')
    if want == 'live':
        if form_unlocked(st):
            raise Exception("Lock the request form's involvement first (see Who can see the form's answers, above).")
        st['mode'] = 'live'
        st['live_since'] = now_str()[:10]
        if not st.get('go_live'):
            st['go_live'] = now_str()[:10]
    elif want == 'testing':
        if _param('confirm') != 'TESTING':
            raise Exception('Type TESTING to go back.')
        st['mode'] = 'testing'
    else:
        raise Exception('Pick a mode.')
    save_settings(st)
    return {'success': True}


def act_reset_one(st):
    if st.get('mode') == 'live':
        raise Exception('Clearing data is only allowed in Testing mode.')
    s = find_submission(st, _param('rid'))
    if not s:
        raise Exception('Request not found.')
    rec = load_record(s.Pid)
    if rec['requests'].pop(s.Rid, None) is None:
        return {'success': True, 'message': 'Nothing was saved for it.'}
    if rec['requests']:
        save_record(s.Pid, rec)
    else:
        model.DeleteExtraValue(int(s.Pid), EV_FIELD)
    return {'success': True, 'message': 'Started over: it is New again.'}


def act_reset_all(st):
    if st.get('mode') == 'live':
        raise Exception('Clearing data is only allowed in Testing mode.')
    if _param('confirm') != 'RESET':
        raise Exception('Type RESET to clear everything.')
    pids = list(set(r.PeopleId for r in q.QuerySql("""
        SELECT PeopleId FROM PeopleExtra WHERE Field IN (%s)""" % ','.join("'%s'" % f for f in REQ_NAMES + TEAM_NAMES + QUICK_NAMES))))
    for pid in pids:
        for f in REQ_NAMES + TEAM_NAMES + QUICK_NAMES:
            model.DeleteExtraValue(int(pid), f)
    return {'success': True, 'message': 'Cleared what this tool saved for %d people. Every request is back to New.' % len(pids)}


def handle_post(st, role):
    action = _param('action')
    try:
        if action not in ACTION_ACCESS:
            raise Exception('Unknown action')
        if not has_access(role, ACTION_ACCESS[action]):
            raise Exception('You do not have access to do that.')
        if action == 'protect':
            protect_now()
            res = {'success': True}
        elif action == 'apply_update':
            res = apply_update()
        elif action == 'save_settings':
            res = act_save_settings(st)
        elif action == 'assign':
            res = act_assign(st, role)
        elif action == 'assign_note':
            res = act_assign_note(st, role)
        elif action == 'walkin_match':
            res = act_walkin_match(st, role)
        elif action == 'walkin_save':
            res = act_walkin_save(st, role)
        elif action == 'walkin_link':
            res = act_walkin_link(st, role)
        elif action == 'sort_old':
            res = act_sort_old(st, role)
        elif action == 'docs_place':
            res = act_docs_place(st, role)
        elif action == 'form_preview':
            res = act_form_preview(st)
        elif action in ('rdec_save', 'rdec_send'):
            res = act_rdec(st, role, action == 'rdec_send')
        elif action == 'nudge':
            res = act_nudge(st, role)
        elif action == 'take_over':
            res = act_take_over(st, role)
        elif action == 'combine':
            res = act_combine(st, role)
        elif action == 'separate':
            res = act_separate(st, role)
        elif action == 'uncombine':
            res = act_uncombine(st, role)
        elif action in ('handoff_ask', 'handoff_cancel', 'handoff_keep'):
            res = {'handoff_ask': act_handoff_ask, 'handoff_cancel': act_handoff_cancel, 'handoff_keep': act_handoff_keep}[action](st, role)
        elif action in ('team_post', 'team_remove', 'req_msg', 'chat_poll', 'req_thread'):
            res = {'team_post': act_team_post, 'team_remove': act_team_remove, 'req_msg': act_req_msg,
                   'chat_poll': act_chat_poll, 'req_thread': act_req_thread}[action](st, role)
        elif action == 'contact_log':
            res = act_contact_log(st, role)
        elif action in ('text_secure', 'text_docs'):
            res = {'text_secure': act_text_secure, 'text_docs': act_text_docs}[action](st, role)
        elif action in ('share_add', 'share_remove'):
            res = {'share_add': act_share_add, 'share_remove': act_share_remove}[action](st, role)
        elif action in ('pay_confirm', 'pickup_done', 'pickup_move'):
            res = {'pay_confirm': act_pay_confirm, 'pickup_done': act_pickup_done, 'pickup_move': act_pickup_move}[action](st, role)
        elif action == 'call_save':
            res = act_call_save(st, role)
        elif action == 'call_notes':
            res = act_call_notes(st, role)
        elif action == 'email_preview':
            res = act_email_preview(st)
        elif action == 'set_mode':
            res = act_set_mode(st)
        elif action == 'reset_one':
            res = act_reset_one(st)
        elif action == 'reset_all':
            res = act_reset_all(st)
        elif action in ('dec_save', 'dec_send', 'affirm_set', 'affirm_note'):
            res = {'dec_save': act_dec_save, 'dec_send': act_dec_send, 'affirm_set': act_affirm_set,
                   'affirm_note': act_affirm_note}[action](st, role)
        else:
            res = {'success': False, 'message': 'Unknown action'}
    except Exception as e:
        res = {'success': False, 'message': unicode(e)}
    print safe_json(res)


# ---------------------------------------------------------------- help along the way
# Everyone is new to this at first, and people come and go, so every page
# explains itself. The panel opens on a first visit and remembers being closed.
PAGE_HELP = {
    ('home', 'router'): [
        'You give out each new request and confirm what is paid. Everything else is the committee\'s.',
        'When you give one out, you land back here with a line saying who has it, ready for the next one.',
        '<b>&#128227; New announcement</b> at the top posts a note to the whole team. It shows here and above People I\'m helping for two weeks; Remove takes it down.',
        '<b>Waiting on you</b> (the first box) is your whole to-do list: new requests to give out, payouts to confirm, handoff asks, new messages and pickups due. The other boxes count every request by stage.',
        '<b>The stage boxes</b> count every request by where it is. Blue ones are yours to act on. Click a box to see everyone in it; the boxes stay put, so you can click from one to the next. <b>Waiting on you</b> brings you back here.',
        '<b>Waiting on you</b> is new requests to give to a committee member, approved ones to confirm payouts for, and pickups due today or missed.',
        '<b>Pickups</b> lists everything set aside, soonest first, so you can tell the pickup place who is coming.',
        '<b>Who has what</b> shows how many each committee member is holding. Amber means nothing has moved on one of theirs for a week, a good time for a gentle nudge.',
        '<b>With someone else</b> is every open request you are not holding, oldest first.',
        '<b>Someone at the door?</b> Type their name or phone number. It finds everyone who has sent a request and sorts them by where it is: waiting for pickup, open, or asked before. Click a name to open the request, with where it is and everything that happened. Not there? <b>Search everyone in TouchPoint</b> also shows forms started and not sent, family members\' requests, and people with nothing yet.'],
    ('all', 'router'): [
        'Every request since the tool started, newest first. Search by name or phone number, or pick a status.',
        'Under <b>Where it is</b> you see who has it, what stage it is at and how long since anything happened.',
        'Requests from before the tool are under <b>Before this tool</b>. Sort them once so the family history is right.'],
    ('older', 'router'): [
        'These came in before the tool started. Mark each one Helped, Declined or Bring it in, so a family\'s earlier requests show correctly.',
        '<b>Bring it in</b> puts it into the normal process as a new request.'],
    ('mine', 'committee'): [
        'These are the people the router gave you. The blue button on each is the next thing to do.',
        'The usual steps: <b>contact them</b>, <b>decide</b> what we will help with, <b>run it by a teammate</b> if needed, then <b>Send to the router</b>.',
        'Everything saves as you type. You can stop and come back another day; it opens where you left off.',
        '<b>Can\'t take this one?</b> on a card asks the router to give it to someone else, without opening it.',
        '<b>Cases you were brought in on</b> are a teammate\'s cases they asked you to help with or affirm. Read it and add a note; the decision stays theirs.',
        '<b>Documents came in: take a look</b> on a card means they sent what you asked for. <b>Waiting on them for documents</b> means you asked and nothing has come yet.',
        '<b>Your recent cases</b> shows what happened after you sent your decision (approved, ready for pickup, done or declined) for the last four months, so you can answer if they call you.'],
    ('req', 'router'): [
        '<b>Membership and family</b> shows who is in the home and any earlier requests. <b>Compare</b> lines this request up with an earlier one.',
        'For a new request, pick a committee member below. The next in turn is already picked, but you can choose anyone.',
        'Once the committee approves, <b>Confirm payouts</b> appears: tick each line as you write the check (with its number) or set the card aside, and pick what happens to it: <b>Mailed or handed over</b>, or <b>Held for pickup</b>.',
        'Anything held for pickup needs a place, day and time. The request then waits in <b>Waiting for pickup</b> until you press <b>Mark picked up</b>, so nothing sitting in the box is forgotten.',
        'Picked-up help then waits for pickup. When they come, tick each thing they had to show, type who picked it up and press <b>Mark picked up</b>.',
        'Missed the pickup? Give them a call and use <b>Change the day or time</b>.',
        'If a member asks to hand a request off, it shows in Waiting on you: give it to someone else, or keep it with them. <b>Give it to someone else</b> works any time something is stuck. A request with no change for a week (Settings, Decisions) shows in Waiting on you too: <b>Check on it</b> writes to the case worker, <b>Take it over</b> makes you the case worker (they stay on it as a teammate). Depending on Settings, Decisions, a router can also decline a new request, or decide it either way, without giving it out.',
        '<b>Case management</b> shows the case worker, anyone brought in, and under <b>Before</b>, everyone who had it or helped. You can bring someone in or take them off.',
        'When giving one out, <b>Helping on</b> counts the open cases each member was brought in on, beside the ones they own.'],
    ('call', 'committee'): [
        '<b>Call notes</b> are always open. The <b>Guide</b> is folded away: open it to tick topics off as you cover them, or pick <b>Walk through the questions</b> to go one topic at a time. It stays open next time if you leave it open.',
        '<b>Tries to reach them</b>: each time you call, text or email, pick what happened and press <b>Log it</b>. The router sees how many tries you have made.',
        'Answer what you can and skip the rest. The topic buttons turn green as you finish them.',
        '<b>What they wrote on the form</b> is beside the call (below it on a phone). Open just the parts you want to check while you talk.',
        'Everything saves to TouchPoint as you type. If you get signed out, the page keeps what you typed and saves it when you sign back in.',
        'When you hang up, press <b>Go to the decision</b>. <b>&larr; Back to the case</b> stays at the top the whole time.',
        '<b>Team chat</b> (the big button, bottom right) works here too, so you can ask the team something mid-call.'],
    ('req', 'committee'): [
        '<b>Team chat</b> is the big button at the bottom right. It turns orange with a count when someone writes, and opens the whole conversation large; <b>Close</b> puts you back where you were. It updates by itself, so you never need to reload.',
        '<b>Case management</b> at the top folds away who is on the case, a button to open the team chat, and <b>Can\'t take this one?</b>. Shut, it shows one line: who else is on it and how many messages. It opens by itself when there is a new message.',
        '<b>Can\'t take this one?</b> (in Case management, and on each card in People I\'m helping) asks the router to give it to someone else; it stays yours until they do, and your work goes with it.',
        'Two steps, top to bottom: <b>Step 1</b> contact them, <b>Step 2</b> decide.',
        '<b>Documents they sent</b>, just above Step 2, shows everything they sent on the documents form, newest first, in the form\'s own order: each file with <b>Open</b>, and any note they wrote about it. New ones are marked <b>New</b>, and the top of the page says <b>Documents (1 new)</b>.',
        'When documents come in on a request waiting on them, it comes back to you by itself: your list says <b>Documents came in: take a look</b>, and so does your daily email.',
        'Their phone, email and address are private. Please keep them that way.',
        '<b>Your decision:</b> tick each thing we will help with, type the amount and pick how. <b>+ Add other help</b> is for anything they did not ask for.',
        'If something is picked up, agree a place, day and time with them. What they need to bring shows on each line.',
        'Not ready to send? Just leave. Your decision saves as you go and the case stays on your list.',
        '<b>Step 1: Contact them</b> has their phone, email and address, the call, <b>Message</b> and their name (folded; open it to send them a secure message or the documents link by text, email or both), every try to reach them (log a call, text or email right there) and what you have from the call so far. Everything that reaches the person asking is here.',
        'In <b>Case management</b>, <b>Team</b> lists who is on it. <b>+ Bring in a teammate</b>: pick <b>To affirm my decision</b> (picked first) or <b>To help with it</b> (they read it, write in the conversation and add notes, but cannot call, decide or hand it off).',
        'To affirm, leave <b>Share the case</b> ticked and they read it here and leave a note. Untick it if you just talked it over by phone or in person; only their name goes on the case.',
        'The <b>Team chat</b> (in Case management) never reaches the person asking. It stays with this person and shows on their timeline. Once someone else is on the case, pick who a message goes to by name: everyone, only the router, or only your teammates without the router.',
        'Two places to write, never mixed up: <b>Message</b> and their name (Step 1, green) is what the person asking reads. <b>Team chat</b> (Case management) is only the team. <b>Message</b> and their name sends a secure message or the documents link by text, email or both: pick under <b>Send by</b>. Only the ways on their record show. The text or email just says there is a message, with a link; what you wrote stays here.'],
    ('req', 'shared'): [
        'A teammate brought you in on this case. You can read it, see their plan so far, and write in the <b>Conversation</b>. The decision stays theirs.',
        'If they asked you to affirm, write your note in the blue box under <b>Case management</b>, Team; it goes with their decision to the router.',
        '<b>Step off</b> under Case management when you are done.'],
    ('req', 'other'): ['You can read this request, but it is not given to you, so there is nothing for you to do here.'],
    ('team', None): [
        '<b>Everyone</b> is the whole team chat: the routers and the committee.',
        '<b>+ New private chat</b>: pick one or more teammates. Private chats (&#128274;) show only to the people in them.',
        'Chats update by themselves while open. Ctrl+Enter (or Cmd+Enter) sends.',
        'For something about one person, use the <b>Team chat</b> on their request, in Case management, instead, so it stays with them.'],
    ('compare', None): [
        'This request on the left, the earlier one on the right. By default only the answers that changed are shown; <b>Show every answer</b> shows all of them.',
        'How it works has the full guides: an overview, and one each for routers, the committee and admins.'],
    ('reports', None): [
        'Pick a range at the top: this month, last month, this year, the last 12 months, or your own dates.',
        '<b>How many</b>, <b>How much</b>, <b>For what</b>, <b>By month</b> and <b>How fast</b> count requests by the day they came in. <b>How fast</b> is the usual number of days (half are quicker, half slower), the average and the longest.',
        '<b>Asked for</b> adds up the amount they gave for each need. Only when they gave none does it use the total they typed, which is free text and sometimes has typos. It says how many gave an amount at all.',
        '<b>How we helped</b> totals each kind of help, so gift cards can be matched against what was bought.',
        '<b>Payments made</b> lists every payment by the day it was paid. <b>Download as a spreadsheet</b> saves it as a file you can open in Excel or hand to someone who asks.',
        '<b>Print</b> prints the report on its own page. Type in <b>Notes to print with it</b> first to add a note at the top.'],
    ('timeline', None): ['Everything that has happened, newest first: who had it, who was brought in, what they decided, notes, and the conversation.'],
    ('settings', None): [
        '<b>Mode</b>: stay in Testing while the team tries it. Nothing is emailed and test data can be cleared. Press Go live when ready.',
        '<b>Protecting the data</b> happens by itself whenever an admin opens the tool, as long as the access role exists.',
        '<b>Access role</b> (in Forms and people): the TouchPoint role that marks the routers and hides the data. It is Benevolence unless you pick a role your church already uses.',
        '<b>Who can get in</b> spells out the three gates (routers, the committee involvement, admins) with today\'s names, and flags a router missing the access role.',
        '<b>Who can see the form\'s answers</b> checks each form\'s involvement is limited to the access role, that its directory and chat are off, and who has the role. Going live waits until the forms are locked.',
        '<b>Forms and people</b>: the forms, the committee involvement, the access role, and who routes. Routers also need the access role.',
        '<b>Reading your form</b>: words the tool looks for in your form\'s questions. Check the preview after any change to the form.',
        'Settings opens on <b>Getting started</b> until setup is done. The other sections: <b>Setup</b> (mode, forms, people), <b>Reading your form</b>, <b>Documents form</b>, <b>The call</b> (the phone questions), <b>Decisions</b> (kinds of help, pickup, reasons, what routers can decide) and <b>Emails</b> (when they go out, who gets what, and a preview), and <b>Texting them</b>.',
        'Each section saves on its own with its Save button.'],
}

ROLE_STEPS = {
    'router': 'Give each new request to a committee member, keep an eye on where things are, and confirm what is paid.',
    'committee': 'Call the people you are given, decide what we will help with, and send it to the router.',
    'admin': 'Set the tool up in Settings and keep it running.'}


def help_panel(view, role):
    if view == 'req':
        # What this person does on this request, not just their role.
        role = REQ_RELATION[0] or role
    if view == 'call':
        role = 'committee'
    lines = PAGE_HELP.get((view, role)) or PAGE_HELP.get((view, None))
    if not lines:
        return ''
    return ('<div class="help" id="helpbox" data-k="%s_%s" hidden><b>How this page works</b><ul>%s</ul>'
            '<div class="muted"><a href="?view=help">How the whole process works</a></div></div>') % (
        view, role, ''.join('<li>%s</li>' % x for x in lines))


# ---------------------------------------------------------------- reports
# What the team said it wants answered any day: how many, how much, for what,
# how fast, plus how the help was given (gift cards by kind, for
# reconciliation) and every payment the router made, downloadable (finance is
# not part of the tool yet; the router handles every payout). Requests are
# counted by the day they came in; payments by the day they were paid.
REPORT_RANGES = [('month', 'This month'), ('lastmonth', 'Last month'), ('year', 'This year'),
                 ('12m', 'Last 12 months'), ('custom', 'Pick dates')]


def report_range():
    """(from 'YYYY-MM-DD', to 'YYYY-MM-DD' inclusive, preset key)."""
    today = datetime.date.today()
    p = _param('p') if _param('p') in dict(REPORT_RANGES) else 'year'
    if p == 'month':
        a, b = today.replace(day=1), today
    elif p == 'lastmonth':
        b = today.replace(day=1) - datetime.timedelta(days=1)
        a = b.replace(day=1)
    elif p == '12m':
        a, b = today - datetime.timedelta(days=365), today
    elif p == 'custom':
        try:
            a = datetime.datetime.strptime(_param('from'), '%Y-%m-%d').date()
            b = datetime.datetime.strptime(_param('to'), '%Y-%m-%d').date()
        except Exception:
            a, b = today.replace(month=1, day=1), today
        if b < a:
            a, b = b, a
    else:
        a, b = today.replace(month=1, day=1), today
    return a.strftime('%Y-%m-%d'), b.strftime('%Y-%m-%d'), p


def _days(a, b):
    try:
        return (datetime.datetime.strptime(b[:10], '%Y-%m-%d') - datetime.datetime.strptime(a[:10], '%Y-%m-%d')).days
    except Exception:
        return None


def _median(xs):
    xs = sorted(xs)
    if not xs:
        return None
    m = int(len(xs) / 2)
    return xs[m] if len(xs) % 2 else (xs[m - 1] + xs[m]) / 2.0


def membership_dates(pids):
    """{pid: (joined 'YYYY-MM-DD' or '', dropped or '', member now)}."""
    pids = sorted(set(int(p) for p in pids if p))
    out = {}
    for i in range(0, len(pids), 900):
        for x in q.QuerySql("""
                SELECT PeopleId Pid, CONVERT(varchar(10), JoinDate, 120) J, CONVERT(varchar(10), DropDate, 120) D, MemberStatusId Ms
                FROM People WHERE PeopleId IN (%s)""" % ','.join(str(p) for p in pids[i:i + 900])):
            out[int(x.Pid)] = (x.J or '', x.D or '', x.Ms == 10)
    return out


def count_member(d, m, md, day):
    """Was the person a member on the day they asked (joined on or before
    it, not dropped before it)? Or did they join after asking?"""
    if not md:
        return
    joined, dropped, now = md
    if joined and joined <= day and (not dropped or dropped > day) and (now or dropped):
        d['members'] += 1
        m['members'] += 1
    elif joined and joined > day and now:
        d['joined'] += 1
        m['joined'] += 1


def report_data(st, a, b):
    """Everything the report shows, from the form and the records."""
    records = all_records()
    subs = submissions(st, "AND ISNULL(rp.CompletedDate, r.CompletedDate) >= '%s' AND ISNULL(rp.CompletedDate, r.CompletedDate) < DATEADD(day, 1, '%s')" % (
        sql_str(a), sql_str(b)), 5000)
    cache = answers_for(st, [x.Rid for x in subs])
    cut = cutoff(st)
    walks_all = all_walkins() if QUICK_PROTECTED else []
    mem = membership_dates([x.Pid for x in subs] + [e['pid'] for e in walks_all if e.get('pid')])
    d = {'n': len(subs), 'members': 0, 'joined': 0, 'helped': 0, 'declined': 0, 'open': 0, 'unsorted': 0, 'asked': 0.0, 'asked_n': 0,
         'paid': 0.0, 'paid_n': 0, 'approved_unpaid': 0.0, 'people': set(), 'needs': {}, 'months': {}, 'hows': {},
         'fast_give': [], 'fast_decide': [], 'fast_paid': []}
    for x in subs:
        r = (records.get(x.Pid) or {}).get(x.Rid) or (records.get(x.Pid) or {}).get(x.Rid.upper()) or {}
        f = cache.get(x.Rid.upper()) or {'needs': [], 'total': None}
        stt = r.get('status') or ('unsorted' if x.T[:10] < cut else 'new')
        if stt == 'combined':
            d['n'] -= 1          # the same request sent twice counts once
            continue
        d['people'].add(x.Pid)
        m = d['months'].setdefault(x.T[:7], {'n': 0, 'helped': 0, 'declined': 0, 'open': 0, 'unsorted': 0, 'paid': 0.0, 'members': 0, 'joined': 0})
        m['n'] += 1
        count_member(d, m, mem.get(int(x.Pid)), x.T[:10])
        # The amounts listed for each need first: the typed total is free text
        # and has real typos (27477 typed for needs adding up to 2747.75).
        asked = sum(to_number(n.get('amount')) or 0 for n in f['needs']) or to_number(f.get('total'))
        if asked:
            d['asked'] += asked
            d['asked_n'] += 1
        pays = r.get('payouts') or []
        paid = sum(float(p.get('amount') or 0) for p in pays)
        if stt in ('approved', 'pickup', 'done'):
            d['helped'] += 1
            m['helped'] += 1
        elif stt == 'declined':
            d['declined'] += 1
            m['declined'] += 1
        elif stt == 'unsorted':
            d['unsorted'] += 1
            m['unsorted'] += 1
        else:
            d['open'] += 1
            m['open'] += 1
        if stt == 'approved':
            d['approved_unpaid'] += sum(to_number(i.get('amount')) or 0 for i in (r.get('decision') or {}).get('items') or [] if i.get('help'))
        if paid:
            d['paid'] += paid
            d['paid_n'] += 1
            m['paid'] += paid
        paid_by_need = {}
        for p in pays:
            paid_by_need[p.get('need') or 'Other'] = paid_by_need.get(p.get('need') or 'Other', 0) + float(p.get('amount') or 0)
            h = d['hows'].setdefault(p.get('how') or 'Before this tool', {'n': 0, 'total': 0.0})
            h['n'] += 1
            h['total'] += float(p.get('amount') or 0)
        for n in f['needs'] or [{'need': 'Other', 'amount': None}]:
            k = d['needs'].setdefault(n['need'] or 'Other', {'n': 0, 'asked': 0.0, 'helped': 0, 'paid': 0.0})
            k['n'] += 1
            k['asked'] += to_number(n.get('amount')) or 0
            if n['need'] in paid_by_need:
                k['helped'] += 1
                k['paid'] += paid_by_need.pop(n['need'])
        for left, amt in paid_by_need.items():   # help added beyond what they asked for
            k = d['needs'].setdefault(left, {'n': 0, 'asked': 0.0, 'helped': 0, 'paid': 0.0})
            k['helped'] += 1
            k['paid'] += amt
        if not r.get('older'):
            asg = r.get('assigned') or []
            if asg:
                d['fast_give'].append(_days(x.T, asg[0].get('t')))
            if (r.get('decision') or {}).get('sent'):
                d['fast_decide'].append(_days(x.T, r['decision']['sent']))
            if pays and pays[0].get('at'):
                d['fast_paid'].append(_days(x.T, min(p.get('at') or '9' for p in pays)))
    for k in ('fast_give', 'fast_decide', 'fast_paid'):
        d[k] = [v for v in d[k] if v is not None and v >= 0]
    # Walk-ins (asked without the form): counted with the rest; what was
    # given is a payment on the day it was written down.
    walks = [e for e in walks_all if a <= (e.get('t') or '')[:10] <= b]
    d['walkins'] = len(walks)
    wrows = []
    for e in walks:
        d['n'] += 1
        d['people'].add(('w', int(e['pid'])) if e.get('pid') else ('w', e['id']))
        m = d['months'].setdefault(e['t'][:7], {'n': 0, 'helped': 0, 'declined': 0, 'open': 0, 'unsorted': 0, 'paid': 0.0, 'members': 0, 'joined': 0})
        m['n'] += 1
        if e.get('pid'):
            count_member(d, m, mem.get(int(e['pid'])), e['t'][:10])
        o = walkin_outcome(e)[0]
        k = {'paid': 'helped', 'declined': 'declined'}.get(o, 'open')
        d[k] += 1
        m[k] += 1
        asked = sum(to_number(x.get('amount')) or 0 for x in e.get('asks') or [])
        if asked:
            d['asked'] += asked
            d['asked_n'] += 1
        paid = sum(to_number(g.get('amount')) or 0 for g in e.get('gives') or [])
        if paid:
            d['paid'] += paid
            d['paid_n'] += 1
            m['paid'] += paid
        given = {}
        for g in e.get('gives') or []:
            amt = to_number(g.get('amount')) or 0
            given[g.get('need') or 'Other'] = given.get(g.get('need') or 'Other', 0) + amt
            h = d['hows'].setdefault(g.get('how') or 'Other', {'n': 0, 'total': 0.0})
            h['n'] += 1
            h['total'] += amt
            wrows.append({'at': e['t'][:10], 'num': e.get('num') or '', 'pid': None, 'name': walkin_name(e), 'need': g.get('need') or '',
                          'payee': '', 'how': g.get('how') or '', 'ref': g.get('ref') or '', 'amount': amt})
        for x in e.get('asks') or [{'need': 'Other', 'amount': None}]:
            k2 = d['needs'].setdefault(x.get('need') or 'Other', {'n': 0, 'asked': 0.0, 'helped': 0, 'paid': 0.0})
            k2['n'] += 1
            k2['asked'] += to_number(x.get('amount')) or 0
            if x.get('need') in given:
                k2['helped'] += 1
                k2['paid'] += given.pop(x.get('need'))
        for left, amt in given.items():
            k2 = d['needs'].setdefault(left, {'n': 0, 'asked': 0.0, 'helped': 0, 'paid': 0.0})
            k2['helped'] += 1
            k2['paid'] += amt
    # Payments made in the range, whenever the request came in.
    rows = []
    pids = []
    for pid, reqs in records.items():
        for rid, r in reqs.items():
            items = (r.get('decision') or {}).get('items') or []
            for p in r.get('payouts') or []:
                at = (p.get('at') or '')[:10]
                if not at or not (a <= at <= b):
                    continue
                i = p.get('i')
                item = items[i] if isinstance(i, int) and 0 <= i < len(items) else {}
                rows.append({'at': at, 'num': r.get('num') or '', 'pid': pid, 'need': p.get('need') or item.get('need') or '',
                             'payee': item.get('payee') or '', 'how': p.get('how') or '', 'ref': p.get('ref') or '',
                             'amount': float(p.get('amount') or 0)})
                pids.append(pid)
    nm = names_for(pids)
    for x in rows:
        x['name'] = nm.get(int(x['pid']), '')
    rows += wrows
    rows.sort(key=lambda x: (x['at'], unicode(x['num'])))
    d['payments'] = rows
    return d


def page_reports(st, role):
    a, b, p = report_range()
    d = report_data(st, a, b)

    def box(label, value, note=''):
        return ('<div class="rpbox"><div class="muted">%s</div><b>%s</b>%s</div>' % (
            esc(label), value, ('<div class="muted">%s</div>' % note) if note else ''))

    def days_line(label, xs):
        if not xs:
            return '<tr><td>%s</td><td colspan="3" class="muted">None yet in this range</td></tr>' % esc(label)
        avg = float(sum(xs)) / len(xs)
        return '<tr><td>%s</td><td data-l="Usually">%s</td><td data-l="Average">%.1f days</td><td data-l="Longest">%d days</td></tr>' % (
            esc(label), '%g days' % _median(xs), avg, max(xs))
    picks = ''.join('<a class="btn%s" href="?view=reports&amp;p=%s">%s</a>' % ('' if k == p else ' ghost', k, esc(t)) for k, t in REPORT_RANGES if k != 'custom')
    custom = ('<form class="bvr" style="justify-content:flex-start;flex-wrap:wrap" onsubmit="bvGo(\'?view=reports&p=custom&from=\'+this.f.value+\'&to=\'+this.t.value); return false;">'
              '<label>From <input type="date" name="f" value="%s"></label><label>To <input type="date" name="t" value="%s"></label>'
              '<button class="btn%s">Show</button></form>') % (a, b, '' if p == 'custom' else ' ghost')
    title = '%s to %s' % (nice_date(a), nice_date(b))
    avg = (d['paid'] / d['paid_n']) if d['paid_n'] else None
    needs = sorted(d['needs'].items(), key=lambda kv: (-kv[1]['n'], kv[0]))
    months = sorted(d['months'].items())
    hows = sorted(d['hows'].items(), key=lambda kv: -kv[1]['total'])
    pay_total = sum(x['amount'] for x in d['payments'])
    body = (
        '<div id="rpt"><h2>%s</h2>'
        '<h3>How many</h3><div class="rpgrid">%s%s%s%s</div>%s'
        '<h3>How much</h3><div class="rpgrid">%s%s%s%s</div>'
        '<h3>For what</h3><div class="tw"><table><thead><tr><th>Need</th><th>Asked</th><th>Amount asked</th><th>Helped</th><th>Paid</th></tr></thead><tbody>%s</tbody></table></div>'
        '<div class="muted">A request can list several needs, so these add up to more than the number of requests.</div>'
        '<h3>By month</h3><div class="tw"><table><thead><tr><th>Month</th><th>Requests</th><th>Members</th><th>Joined later</th><th>Helped</th><th>Declined</th><th>Open</th>%s<th>Paid</th></tr></thead><tbody>%s</tbody></table></div>'
        '<div class="muted">By the month the request came in. <b>Members</b>: a church member on the day they asked (joined on or before it). '
        '<b>Joined later</b>: became a member after asking.%s</div>'
        '<h3>How fast</h3><div class="tw"><table><thead><tr><th>From the day it came in</th><th>Usually</th><th>Average</th><th>Longest</th></tr></thead><tbody>%s%s%s</tbody></table></div>'
        '<h3>How we helped</h3><div class="tw"><table><thead><tr><th>Given as</th><th>Payments</th><th>Total</th></tr></thead><tbody>%s</tbody></table></div>'
        '<div class="muted">Gift cards and other kinds you set up in Settings, each with its total, to match against what was bought.</div>'
        '</div>') % (
        esc(title),
        box('Requests', '%d' % d['n'], '%d %s%s' % (len(d['people']), 'person' if len(d['people']) == 1 else 'people',
                                                   (', including %d walk-in%s' % (d['walkins'], '' if d['walkins'] == 1 else 's')) if d.get('walkins') else '') +
            ('; %d from members when they asked' % d['members'])),
        box('Helped', '%d' % d['helped']), box('Declined', '%d' % d['declined']), box('Still open', '%d' % d['open']),
        ('<div class="muted">%d more came in before this tool and were never sorted (All requests, Before this tool).</div>' % d['unsorted']) if d['unsorted'] else '',
        box('Asked for', money(d['asked']) or '$0', '%d of %d gave an amount; uses the amount on each need first' % (d['asked_n'], d['n'])),
        box('Paid', money(d['paid']) or '$0', '%d %s helped' % (d['paid_n'], 'request' if d['paid_n'] == 1 else 'requests')),
        box('Average help', money(round(avg, 2)) if avg else 'none yet', 'per request paid'),
        box('Approved, not paid yet', money(d['approved_unpaid']) or '$0'),
        ''.join('<tr><td><b>%s</b></td><td data-l="Asked">%d</td><td data-l="Amount asked">%s</td><td data-l="Helped">%d</td><td data-l="Paid">%s</td></tr>' % (
            esc(k), v['n'], money(v['asked']) or '', v['helped'], money(v['paid']) or '') for k, v in needs) or '<tr><td colspan="5" class="muted">No requests in this range.</td></tr>',
        '<th>Not sorted</th>' if d['unsorted'] else '',
        ''.join('<tr><td><b>%s</b></td><td data-l="Requests">%d</td><td data-l="Members">%d</td><td data-l="Joined later">%d</td><td data-l="Helped">%d</td>'
                '<td data-l="Declined">%d</td><td data-l="Open">%d</td>%s<td data-l="Paid">%s</td></tr>' % (
            datetime.datetime.strptime(k + '-01', '%Y-%m-%d').strftime('%B %Y'), v['n'], v.get('members', 0), v.get('joined', 0), v['helped'], v['declined'], v['open'],
            ('<td data-l="Not sorted">%d</td>' % v['unsorted']) if d['unsorted'] else '', money(v['paid']) or '')
            for k, v in months) or '<tr><td colspan="9" class="muted">No requests in this range.</td></tr>',
        ' <b>Not sorted</b>: came in before this tool and never sorted.' if d['unsorted'] else '',
        days_line('Given to a committee member', d['fast_give']), days_line('Decision sent', d['fast_decide']), days_line('Paid', d['fast_paid']),
        ''.join('<tr><td><b>%s</b></td><td data-l="Payments">%d</td><td data-l="Total">%s</td></tr>' % (esc(k), v['n'], money(v['total'])) for k, v in hows)
        or '<tr><td colspan="3" class="muted">No payments for these requests yet.</td></tr>')
    pays = ''.join('<tr><td>%s</td><td data-l="#">%s</td><td data-l="For">%s</td><td data-l="Need">%s</td><td data-l="Paid to">%s</td>'
                   '<td data-l="How">%s</td><td data-l="Check #">%s</td><td data-l="Amount">%s</td></tr>' % (
                       nice_date(x['at']), esc(unicode(x['num'])), esc(x['name']), esc(x['need']), esc(x['payee']), esc(x['how']), esc(x['ref']), money(x['amount']))
                   for x in d['payments'])
    csv = [['Date paid', 'Request', 'For', 'Need', 'Paid to', 'How', 'Check number', 'Amount']] + [
        [x['at'], unicode(x['num']), x['name'], x['need'], x['payee'], x['how'], x['ref'], '%.2f' % x['amount']] for x in d['payments']]
    payments = ('<div class="bvc" id="rppay"><div class="bvr"><h2>Payments made %s</h2><button type="button" class="btn ghost" id="rpcsv">Download as a spreadsheet</button></div>'
                '<div class="tw"><table><thead><tr><th>Date</th><th>#</th><th>For</th><th>Need</th><th>Paid to</th><th>How</th><th>Check #</th><th>Amount</th></tr></thead>'
                '<tbody>%s</tbody></table></div><div class="bvr"><span class="muted">%d %s, by the day they were paid, whenever the request came in.</span><b>Total %s</b></div></div>') % (
        esc(title), pays or '<tr><td colspan="8" class="muted">No payments in this range.</td></tr>',
        len(d['payments']), 'payment' if len(d['payments']) == 1 else 'payments', money(pay_total) or '$0')
    return ('<div class="bvr"><h1>Reports</h1><button type="button" class="btn ghost" id="rpprint">Print</button></div>'
            '<div class="bvr" style="justify-content:flex-start;flex-wrap:wrap">%s</div>%s'
            '<details class="bvc flat" id="rpnotes"><summary><b>Notes to print with it</b> <span class="muted" style="font-weight:normal">for a meeting or a report to leadership</span></summary>%s</details>'
            '<div class="bvc">%s</div>%s'
            '<script>(function(){var CSV=%s, NAME=%s;'
            'document.getElementById("rpcsv").onclick=function(){var t=CSV.map(function(r){return r.map(function(v){v=String(v==null?"":v);'
            ' return /[",\\n]/.test(v)?\'"\'+v.replace(/"/g,\'""\')+\'"\':v;}).join(",");}).join("\\r\\n");'
            ' var a=document.createElement("a"); a.href=URL.createObjectURL(new Blob(["\\ufeff"+t],{type:"text/csv"})); a.download=NAME; document.body.appendChild(a); a.click(); a.remove();};'
            'document.getElementById("rpprint").onclick=function(){var css="*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}'
            'body{margin:0;padding:24px;font-family:Verdana,sans-serif;color:#1C2433;font-size:13px}h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:14px 0 6px}h3{font-size:15px;margin:14px 0 6px}'
            'table{width:100%%;border-collapse:collapse;margin:4px 0}th,td{border-bottom:1px solid #ddd;padding:4px 6px;text-align:left}th{background:#F2F0EB}'
            '.rpgrid{display:flex;gap:10px;flex-wrap:wrap}.rpbox{border:1px solid #ccc;border-radius:8px;padding:8px 12px;min-width:120px}.rpbox b{font-size:18px;display:block}.muted{color:#5A6478;font-size:12px}'
            '.notes{border:1px solid #ccc;border-radius:8px;padding:8px 12px;margin:8px 0}.pb{page-break-before:always}";'
            ' var n=document.querySelector("#rpnotes .rt"), notes=n&&n.textContent.trim()?\'<div class="notes"><b>Notes</b>\'+n.innerHTML+"</div>":"";'
            ' var w=window.open("","_blank"); if(!w){alert("Please allow pop-ups to print."); return;}'
            ' w.document.write("<!DOCTYPE html><html><head><title>Benevolence report</title><style>"+css+"</style></head><body><h1>Benevolence report</h1>"+notes'
            '  +document.getElementById("rpt").innerHTML+\'<div class="pb"></div>\'+document.getElementById("rppay").innerHTML.replace(/<button[^>]*>.*?<\\/button>/,"")+"</body></html>");'
            ' w.document.close(); w.focus(); setTimeout(function(){w.print();},300);};})();</script>') % (
        picks, custom, rich_box('', '', 80), body, payments, safe_json(csv), safe_json('benevolence-payments-%s-to-%s.csv' % (a, b)))



GUIDE_TABS = [('overview', 'Overview'), ('router', 'For routers'), ('committee', 'For the committee'), ('admin', 'For admins')]


def _sec(title, *parts):
    return '<div class="bvc"><h2>%s</h2>%s</div>' % (title, ''.join(parts))


def _p(t):
    return '<p>%s</p>' % t


def _ul(items):
    return '<ul>%s</ul>' % ''.join('<li>%s</li>' % x for x in items)


def _ol(items):
    return '<ol class="steps">%s</ol>' % ''.join('<li>%s</li>' % x for x in items)


def guide_facts(st):
    """The settings the guides quote, in plain words."""
    kinds = ', '.join(esc(k['name']) for k in st.get('help_kinds') or []) or 'the kinds of help set in Settings'
    rd = st.get('router_decide', 'decline')
    router_can = {'none': 'always give it to a committee member (routers do not decide here)',
                  'decline': 'give it to a committee member, or decline it right away',
                  'both': 'give it to a committee member, or decide it right away, either way'}.get(rd)
    if st.get('affirm_on') and st.get('affirm_required'):
        aff = 'Every decision is run by a second committee member before it goes to the router.'
    elif st.get('affirm_on') and float(st.get('affirm_over') or 0):
        aff = 'Help over %s is run by a second committee member before it goes to the router.' % money(float(st['affirm_over']))
    elif st.get('affirm_on'):
        aff = 'A committee member can run a decision by a teammate when they want a second opinion.'
    else:
        aff = 'Decisions go straight to the router (affirmation by a teammate is turned off).'
    return {'kinds': kinds, 'router_can': router_can, 'aff': aff, 'stuck': stuck_days(st),
            'pickup_req': bool(st.get('req_pickup')), 'earlier': st.get('earlier_on', True),
            'live': bool(st.get('email_on')) and st.get('mode') == 'live'}


def guide_overview(st):
    g = guide_facts(st)
    return ''.join([
        _sec('What this is',
             _p('Benevolence handles a request for help from the moment someone fills out your request form to the day the check is mailed or the help is picked up. '
                'It runs inside TouchPoint, on your own form, your own committee involvement and your own people. There is nothing else to sign up for.'),
             _p('It keeps everyone working from the same page: who has the request, what was learned on the call, what was decided and why, and what was paid. '
                'And it keeps what people share private.')),
        _sec('A request, start to finish',
             _ol([
                 '<b>Someone asks for help.</b> They fill out your request form online, or someone fills it out with them at the door. '
                 'It shows up for the router as New. If they send the form twice, the router sees one card and decides whether it is the same request or a new one.',
                 '<b>The router looks first.</b> The request page shows the family, membership, every earlier request from anyone in the home%s, '
                 'and every answer they gave. The router can %s.' % (
                     ' and what was decided before and why' if g['earlier'] else '', g['router_can']),
                 '<b>A committee member reaches out.</b> They call with your questions in front of them, log each try, and can send a secure message '
                 'or a link to the documents form by text or email. Documents they send land on their request by themselves.',
                 '<b>They decide.</b> Which needs we help with, how much and how (%s), or decline with a reason and somewhere to send them. They write the Why. %s' % (
                     g['kinds'], g['aff']),
                 '<b>The router confirms the payouts.</b> Each line is checked off as the check is written or the card is set aside, mailed or held for pickup%s.' % (
                     ' at the place, day and time the committee member set' if g['pickup_req'] else ''),
                 '<b>Picked up.</b> Whoever hands it out sees what the person needs to bring, notes who picked it up, and marks it done.',
                 '<b>It stays on the record.</b> The timeline keeps every step, note and message, so the next time this family asks, the team knows the story.'])),
        _sec('Who does what',
             _ul([
                 '<b>The person asking</b> fills out the form, may be called or messaged, and sends documents if asked. They never see the team\'s notes.',
                 '<b>Routers</b> see every request. They give each one out in turn (or decline it, if your settings allow), keep things moving, '
                 'confirm what is paid, and handle pickups.',
                 '<b>Committee members</b> see only the requests given to them, or ones a teammate brings them in on. They call, decide and write the Why.',
                 '<b>Admins</b> set it up and keep it running in Settings. Being an admin does not open anyone\'s request; only routers and the committee member helping them see it.'])),
        _sec('Private by design',
             _ul([
                 'Requests are hidden behind one TouchPoint role that only routers hold. Everyone else, including other staff, sees nothing.',
                 'The request form\'s answers are locked to that role too, so they cannot be read from the involvement.',
                 'Emails never hold a name, a need, an amount or a message: just a request number and a link. Secure messages to the person asking are read only after they sign in.',
                 'The team chat on a request never reaches the person asking. Answers marked private in Reading your form show only to routers.'])),
        _sec('Keeping it moving',
             _ul([
                 '<b>Waiting on you</b> is the router\'s whole to-do list: new requests, handoff asks, new messages, payouts to confirm, pickups due, and anything with no change for %d days.' % g['stuck'],
                 '<b>People I\'m helping</b> is each committee member\'s list, with the next step on every card.',
                 '<b>Emails</b>: a daily list of what is waiting on each person, notices when something is given to them or comes back, and secure message notices. '
                 'Admins choose which go out and whether right away or at the next 15-minute run.',
                 '<b>Live chat</b> on each request and for the whole team, which shows who is there right now.',
                 'A committee member who cannot take one asks the router to hand it off. A router can check on a stalled one, take it over, or give it to someone else.'])),
        _sec('Reports',
             _p('How many asked, helped and declined; what was asked for and what was given, by need and by kind of help; how fast requests move; '
                'and every payment in a date range, ready to print or download for your finance team.')),
        _sec('What it does not do',
             _ul(['It does not write checks or move money. It records what your church paid.',
                  'It does not text anyone unless an admin picks a TouchPoint texting group.',
                  'It does not keep request details anywhere TouchPoint shows to the public.']))])


def guide_router(st):
    g = guide_facts(st)
    return ''.join([
        _sec('Your part',
             _p('You see every request. You give each new one to a committee member, keep an eye on where things are, and confirm what is paid. '
                'The decision itself belongs to the committee member%s.' % ('' if st.get('router_decide', 'decline') == 'none' else ', unless you decide it yourself right away'))),
        _sec('Waiting on you',
             _ul(['Your whole to-do list, newest first: new requests, handoff asks, new messages, payouts to confirm, pickups due today or missed, and anything stalled.',
                  'The stage boxes at the top count every request by where it is. Click one to list it; <b>Waiting on you</b> brings you back.',
                  '<b>Someone at the door?</b> Type a name or phone number. It finds everyone who has sent a request, sorted by where it is, or was. Click a name to open it with its whole history.',
                  '<b>&#128227; New announcement</b> posts a note to the whole team for two weeks.',
                  '<b>Documents that match no request</b> lists documents sent by someone with no request before them. Attach each to the right request, or set it aside.'])),
        _sec('A new request',
             _ol(['Read the form and <b>Membership and family</b>: who is in the home and every earlier request from anyone in it. <b>Compare</b> lines it up with an earlier one.%s' % (
                      ' <b>Earlier help to this family</b> shows what was decided before, the Why and the call notes.' if g['earlier'] else ''),
                  {'none': 'Pick a committee member.', 'decline': 'In <b>Your decision</b>, pick <b>Assign</b> or <b>Decline</b>.',
                   'both': 'In <b>Your decision</b>, pick <b>Decline</b>, <b>Assign</b> or <b>Help</b>.'}[st.get('router_decide', 'decline')] +
                  ' <b>Assign</b> gives it to a committee member: the next in turn is already picked; <b>Helping on</b> counts cases they were brought in on. Add a note if you like, and <b>Give it to them</b>.',
                  {'none': 'Your church has routers give every request out; you do not decide them here.',
                   'decline': '<b>Decline</b>: pick a reason, somewhere to send them, and write the Why. It closes as declined.',
                   'both': '<b>Help</b> or <b>Decline</b> decides it yourself, right away: tick what we will give and how, or pick a reason, and write the Why.'}[st.get('router_decide', 'decline')]])),
        _sec('Someone without a form (walk-ins)',
             _p('People often ask in person or on the phone without filling out the form, and many are not in TouchPoint. '
                '<b>+ Someone without a form</b> (on Waiting on you, or the <b>Walk-ins</b> tab) writes them down:'),
             _ol(['Type who they are: name, phone, email, address. As you type, it looks in TouchPoint and in earlier walk-ins. '
                  'If they are in TouchPoint, pick <b>This is them</b> to link it.',
                  'What they asked for, picked from your form\'s needs, with an amount.',
                  'What we decided: <b>Helped</b> (what we gave, how, and the check or card number), <b>Declined</b> (why, and where we sent them), '
                  'or <b>Not decided yet</b>. Write the Why.']),
             _ul(['Not in TouchPoint? It is kept as a walk-in with what you typed. Nothing is added to TouchPoint.',
                  'When a form comes in later from someone who matches a walk-in (same phone, email, or first and last name), the request shows '
                  '<b>They may have asked before without the form</b>. <b>Same person: link it</b>, or <b>Not them</b>.',
                  'Linked walk-ins show with the person\'s other requests, on the family timeline and in <b>Someone at the door?</b>. '
                  'Ones not linked show there too, by name or phone.',
                  'Reports count walk-ins with the rest, and what was given is in the payments.',
                  'Only routers see walk-ins.'])),
        _sec('They sent the form twice',
             _p('When the same person sends the form again within %d days and both are still open, you see it on Waiting on you, even when both have already been given out. '
                'On the request, every answer from both forms is side by side with the changes highlighted. Then pick:' % REPEAT_DAYS),
             _ul(['<b>Same request: combine them.</b> Keep one (the newer is picked). The other closes as Combined: not declined, not counted twice, and its documents move over. '
                  'If the other is already with a committee member, it joins their case and they are told. If both were given out, pick whose case to keep; '
                  'the other member is told they can stop working theirs.',
                  'Until you decide, the committee member on either one sees a note on their case that the form came in twice and a router is deciding.',
                  '<b>Separate requests: keep both.</b> For a different need. The question goes away.',
                  'Changed your mind? <b>Undo</b> on the combined one makes it its own request again.'])),
        _sec('Keeping it moving',
             _ul(['A request with no change for %d days shows on Waiting on you. On it, <b>Keeping it moving</b> has:' % g['stuck'],
                  '<b>Check on it</b>: writes your note in the team chat and lets the case worker know.',
                  '<b>Take it over</b>: you become the case worker and get the call and the decision. They stay on it as a teammate, so nothing they learned is lost.',
                  '<b>Give it to someone else</b>: moves it, with everything done so far.',
                  'If a member asks to hand one off, it shows on Waiting on you: give it to someone else, or keep it with them.'])),
        _sec('Confirming payouts and pickups',
             _ul(['Once approved, <b>Confirm payouts</b> appears. Tick each line as you write the check (with its number) or set the card aside, and pick <b>Mailed or handed over</b> or <b>Held for pickup</b>.',
                  'Held for pickup needs a place, day and time. It waits in <b>Waiting for pickup</b> until it is picked up.',
                  'When they come, tick what they had to show, type who picked it up and press <b>Mark picked up</b>. Missed it? Call them and <b>Change the day or time</b>.'])),
        _sec('All requests and Reports',
             _ul(['<b>All requests</b> is every request ever sent, newest first. Search by name or phone, or pick a status, including <b>Before this tool</b> for older ones to sort once.',
                  '<b>Reports</b> covers any date range: how many, how much, for what, by month, how fast, and every payment, ready to print or download.']))])


def guide_committee(st):
    g = guide_facts(st)
    return ''.join([
        _sec('Your part',
             _p('The router gives you a request. You reach the person, learn what is going on, decide what we will help with and write why. '
                'You only see the requests given to you, and ones a teammate brings you in on.')),
        _sec('People I\'m helping',
             _ul(['Your list. The blue button on each card is the next thing to do.',
                  '<b>Can\'t take this one?</b> asks the router to give it to someone else. It stays yours until they do, and your work goes with it.',
                  '<b>Documents came in: take a look</b> means they sent what you asked for.',
                  '<b>Cases you were brought in on</b> are a teammate\'s cases. Read it and add a note; the decision stays theirs.',
                  '<b>Your recent cases</b> shows how your last four months turned out, so you can answer if they call.',
                  '<b>They sent the form again</b> at the top of a case means the same person sent another form within %d days. A router decides whether it is the same request; '
                  'if it is, the two become one case and you are told which one to work.' % REPEAT_DAYS])),
        _sec('Step 1: Contact them',
             _ul(['Their phone, email and address are at the top. Please keep them private.',
                  '<b>Start the call</b> opens the call page: your notes on one side, what they wrote on the form on the other, and the call questions folded away. '
                  'Everything saves as you type.',
                  'Log every try to reach them, called, texted or emailed, and what happened.',
                  '<b>Message</b> sends them a secure message or the link to the documents form by text, email or both. They sign in to read it.',
                  'Documents they send show under <b>Documents they sent</b>, newest first, and your list tells you when new ones arrive.'])),
        _sec('Step 2: Your decision',
             _ul(['Tick each thing we will help with, type the amount and pick how (%s). <b>+ Add other help</b> covers anything they did not ask for.' % g['kinds'],
                  ('Pickup place, day and time are required when anything is picked up.' if g['pickup_req']
                   else 'If something is picked up, agree a place, day and time with them, or leave it for the router.'),
                  'Or <b>Decline</b>: pick a reason, somewhere to send them, and write the Why.',
                  g['aff'],
                  'The checklist under the decision shows what is still needed. When it is all ticked, <b>Send to the router</b>.',
                  'Not ready? Just leave. It saves as you go.%s' % (
                      ' <b>Earlier help to this family</b>, near the top, shows what was decided before and why.' if g['earlier'] else '')])),
        _sec('Working with the team',
             _ul(['<b>Team chat</b> (the big button at the bottom right) is only the team; the person asking never sees it. It updates by itself and shows who is there right now.',
                  'Once a teammate is on the case, pick who a message goes to: just the people on the case, the case and the router, or only the router.',
                  '<b>+ Bring in a teammate</b> to affirm your decision or to help. They can read and write in the chat, but the call and the decision stay yours.',
                  '<b>Messages</b> at the top is the whole team chat and private chats with teammates.']))])


def guide_admin(st):
    return ''.join([
        _sec('Your part',
             _p('You set the tool up and keep it running. Settings opens on <b>Getting started</b>, a checklist that walks you through setup in order and '
                'shows what is done. Being an admin does not open anyone\'s request.')),
        _sec('Getting started, in order',
             _ol(['<b>Make the access role</b> and protect the data. Only routers get it.',
                  '<b>Pick your request form and lock it</b>, so its answers are hidden behind the access role.',
                  '<b>Check how the tool reads your form</b> (Reading your form): which questions are needs, amounts, people in the home and private answers, with a live preview '
                  'that follows your form\'s skip logic.',
                  '<b>Your documents form</b> (optional), read the same way.',
                  '<b>Pick the committee</b> involvement and <b>choose your routers</b>.',
                  '<b>Write the call questions</b> committee members ask on the phone.',
                  '<b>Set the kinds of help</b> and how decisions work.',
                  '<b>Turn on the emails</b>, and texting if you want it.',
                  '<b>Try it, then go live.</b> In Testing nothing is emailed and test data can be cleared.'])),
        _sec('What each part of Settings changes',
             _ul(['<b>Setup</b>: Testing or Live, the forms, the committee involvement, the access role, and who routes.',
                  '<b>Reading your form</b> and <b>Documents form</b>: how the tool understands each question. Check the preview after any change to the form.',
                  '<b>The call</b>: the questions on the call page.',
                  '<b>Decisions</b>: kinds of help and what each needs at pickup, pickup places, decline reasons and referrals, what must be done before a decision is sent '
                  '(the Why, talking with them first, pickup details), teammate affirmation, what a router can decide on a new request, '
                  'how many days before a request counts as stalled, and whether earlier help is shown.',
                  '<b>Emails</b>: the daily list (time and days), secure message notices, and each right-away notice with who gets it and whether it goes right away '
                  'or at the next 15-minute run. A preview shows exactly what goes out.',
                  '<b>Texting them</b>: the TouchPoint texting group used to text the person asking. Off means no texts.'])),
        _sec('Good to know',
             _ul(['Emails and notices are sent by a TouchPoint scheduled task that runs about every 15 minutes, and only when the mode is Live and emails are on.',
                  'Each Settings section saves on its own with its Save button.',
                  'Settings are stored where TouchPoint can show them to others, so never put names, phone numbers or other private details in them. '
                  'Request data is kept in protected records instead.']))])


def page_help(st, role):
    """How it works: an overview, then a guide for each role."""
    default = {'router': 'router', 'committee': 'committee'}.get(role, 'overview')
    # Each person sees the guides for their part: the committee only theirs
    # (no overview: it covers routing, payouts and setup they do not need);
    # routers theirs and the committee's (they hand cases over); admins all.
    mine = {'committee': ('committee',), 'router': ('overview', 'router', 'committee')}.get(role, ('overview', 'router', 'committee'))
    if IS_ADMIN and not VIEW_AS:
        mine += ('admin',)
    tabs = [(k, l) for k, l in GUIDE_TABS if k in mine]
    tab = _param('tab') if _param('tab') in mine else default
    you = ROLE_STEPS.get(role, '')
    if role == 'router' and IS_MEMBER:
        you += ' You are on the committee too, so requests given to you show under People I\'m helping.'
    body = {'overview': guide_overview, 'router': guide_router, 'committee': guide_committee, 'admin': guide_admin}[tab](st)
    if tab == 'overview':
        body += ('<div class="bvc"><h2>Who gets emailed, and when</h2><div class="muted">Emails never hold the name of the person asking, '
                 'their needs, amounts or what anyone wrote: just the request number and a link. Admins choose which of these go out%s.</div>%s</div>') % (
            '' if st.get('email_on') and st.get('mode') == 'live' else ' (emails are not on yet)', who_gets_what(st))
    return ('<h1>How it works</h1><div class="bvc bvblue"><b>Your part</b><div>%s</div></div>'
            '%s%s'
            '<div class="bvc"><b>Good to know</b><ul>'
            '<li>Everything saves as you type. If you are signed out, the page keeps what you typed and saves it once you sign in again.</li>'
            '<li>The round <b>?</b> button at the top of each page explains that page.</li>'
            '<li>Text size buttons at the top make everything bigger or smaller, and it remembers your choice.</li></ul></div>') % (
        esc(you), ('<div class="subnav">%s</div>' % ''.join('<a href="?view=help&amp;tab=%s" class="%s">%s</a>' % (k, 'on' if k == tab else '', esc(l)) for k, l in tabs))
        if len(tabs) > 1 else '', body)


REQ_RELATION = ['']     # set by page_request: router, committee (it is theirs), shared or other

# Admins: add &debug=1 to see how wide the page is on this device (for
# layout problems a desktop browser does not show, like an iPhone gap).
WIDTH_DEBUG_JS = """<script>window.addEventListener('load',function(){setTimeout(function(){
var W=window.innerWidth,o=['screen '+W+' wide, page '+document.documentElement.scrollWidth+' wide, zoom '+(window.visualViewport?Math.round(visualViewport.scale*100)/100:'?')];
var e=document.getElementById('bv');while(e&&e!==document.documentElement){var r=e.getBoundingClientRect(),c=getComputedStyle(e);
o.push((e.id?'#'+e.id:e.tagName.toLowerCase()+'.'+String(e.className).split(' ')[0])+' '+Math.round(r.left)+' to '+Math.round(r.right)+' pad '+c.paddingLeft+'/'+c.paddingRight+' mar '+c.marginLeft+'/'+c.marginRight);e=e.parentElement;}
var wide=[];document.querySelectorAll('body *').forEach(function(x){var r=x.getBoundingClientRect();if(r.width&&r.right>W+1&&!x.closest('nav')){var p=x.parentElement;if(p&&p.getBoundingClientRect().right<=W+1)wide.push(x.tagName.toLowerCase()+'.'+String(x.className).split(' ')[0]+' to '+Math.round(r.right));}});
o.push('wider than the screen: '+(wide.slice(0,8).join(', ')||'nothing'));
var d=document.createElement('pre');d.style.cssText='position:fixed;left:4px;right:4px;bottom:4px;z-index:99999;background:#1C2433;color:#fff;font:12px/1.4 monospace;padding:8px;white-space:pre-wrap;border-radius:8px;max-height:45vh;overflow:auto';
d.textContent=o.join('\\n');document.body.appendChild(d);},800);});</script>"""

HELP_BOX = ['']     # this page's help, shown under the ? button

HELP_JS = """<script>(function(){var b=document.getElementById('helpq'),d=document.getElementById('helpbox'); if(!b||!d) return;
var k='bv.help.'+d.getAttribute('data-k');
function set(open){d.hidden=!open; b.className='qbtn'+(open?' on':''); b.setAttribute('aria-expanded',open?'true':'false');}
try{set(localStorage.getItem(k)==='open');}catch(e){set(false);}
b.onclick=function(){var open=d.hidden; set(open); try{localStorage.setItem(k,open?'open':'closed');}catch(e){}};})();</script>"""


# ---------------------------------------------------------------- entry
REDIRECT_JS = """<script>
if (location.pathname.indexOf('/PyScript/') > -1)
  location.replace(location.pathname.replace('/PyScript/', '/PyScriptForm/') + location.search + location.hash);
</script>"""

SETTINGS = load_settings()
# An admin opening the tool protects the data automatically: it registers our
# extra values in StandardExtraValues2 as hidden to the access role (a
# church-wide setting only admins can change). It needs the role to exist.
AUTO_PROTECT_ERROR = ''
if IS_ADMIN:
    _ok, _missing, _role_ok = protection_state()
    if not _ok and _role_ok:
        try:
            protect_now()
        except Exception as e:
            AUTO_PROTECT_ERROR = unicode(e)
PROTECTED = protection_state(REQ_NAMES)[0]
TEAM_PROTECTED = protection_state(TEAM_NAMES)[0]
QUICK_PROTECTED = protection_state(QUICK_NAMES)[0]
ROLE = who_am_i(SETTINGS)
# On the committee too (a router can be both).
IS_MEMBER = bool(ME) and any(m['pid'] == int(ME) for m in committee_members(SETTINGS))
# Committee view: a router who is also on the committee can see the tool as a
# committee member does (their own assignments, no router pages). The choice
# is kept in their own team value so it lasts from page to page.
REAL_ROLE = ROLE
VIEW_AS = False
if ROLE == 'router' and IS_MEMBER and TEAM_PROTECTED and my_team().get('view_as') == 'committee':
    ROLE, VIEW_AS = 'committee', True


def act_view_as():
    if REAL_ROLE != 'router':
        raise Exception('Only routers switch views.')
    want = _param('as')
    if want == 'committee' and not IS_MEMBER:
        raise Exception('Add yourself to the committee involvement first.')
    d = my_team()
    d['view_as'] = 'committee' if want == 'committee' else ''
    save_my_team(d)
    return {'success': True}

if _param('bvtick') == 'tick' and (not ME or IS_ADMIN):
    # From ScheduledTasks, which runs in TouchPoint's batch process with no one
    # signed in. A signed-in non-admin adding ?bvtick=tick (and a made-up
    # bvslot time) just gets the normal page, so no one can trigger the emails
    # early. Never raise (TouchPoint emails admins about every
    # failed run, every 15 minutes); log instead.
    try:
        email_tick(SETTINGS)
    except Exception as e:
        email_log('scheduled run failed: %s' % unicode(e)[:200])
elif model.HttpMethod == 'post' and _param('action') == 'view_as':
    try:
        print safe_json(act_view_as())
    except Exception as e:
        print safe_json({'success': False, 'message': unicode(e)})
elif model.HttpMethod == 'post':
    handle_post(SETTINGS, ROLE)
else:
    if _param('view') == 'letter':
        # The person who asked, from the link in a text: only their own messages.
        model.Form = REDIRECT_JS + page_letter(SETTINGS)
    elif ROLE == 'none':
        model.Form = REDIRECT_JS + page_denied()
    else:
        view = _param('view') or {'router': 'home', 'committee': 'mine'}.get(ROLE, 'settings')
        pages = {'home': page_home, 'all': page_all, 'req': page_request, 'timeline': page_timeline,
                 'find': page_find, 'settings': page_settings, 'older': page_older, 'compare': page_compare,
                 'mine': page_mine, 'help': page_help, 'call': page_call, 'team': page_team, 'reports': page_reports,
                 'walkins': page_walkins, 'walkin': page_walkin}
        fn = pages.get(view)
        try:
            if view not in PAGE_ACCESS or not has_access(ROLE, PAGE_ACCESS[view]):
                body = '<div class="bvc"><b>You don\'t have access to this page.</b></div>'
            else:
                body = fn(SETTINGS, ROLE) if fn else '<div class="bvc">Coming next: the committee screens.</div>'
        except Exception as e:
            body = '<div class="bvc bvbad"><b>Something went wrong.</b><span>%s</span></div>' % esc(repr(e))
        if view in PAGE_ACCESS and has_access(ROLE, PAGE_ACCESS[view]):
            HELP_BOX[0] = help_panel(view, ROLE)
        model.Form = REDIRECT_JS + shell(ROLE, view, body) + HELP_JS + (WIDTH_DEBUG_JS if IS_ADMIN and _param('debug') == '1' else '')
